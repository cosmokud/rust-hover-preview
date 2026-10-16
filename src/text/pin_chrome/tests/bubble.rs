use super::*;

/// The side every bubble here is painted at: the size `PIN_BUBBLE_PIXELS` is drawn at on a
/// display at 100%, and a size at which a face, a hairline of ring and a glyph are all more than
/// one pixel across — a smaller bubble would be testing rounding rather than drawing.
const SIDE: u32 = 64;

/// A palette with three colours nothing can confuse for one another, so that a pixel read back is
/// unambiguously the ring, the face or the glyph.
fn palette() -> ChromePalette {
    ChromePalette {
        background: [24, 26, 32],
        foreground: [240, 240, 240],
        accent: [97, 175, 239],
        dark: true,
    }
}

fn paint(art: BubbleArt<'_>, mark: BubbleMark) -> Vec<u8> {
    let mut buffer = vec![0u8; (SIDE * SIDE * 4) as usize];
    paint_bubble(&mut buffer, SIDE, SIDE, &palette(), art, mark);
    buffer
}

/// One pixel of a painted bubble, in the premultiplied BGRA the buffer is in.
fn pixel(buffer: &[u8], x: u32, y: u32) -> [u8; 4] {
    let index = (y as usize * SIDE as usize + x as usize) * 4;
    [
        buffer[index],
        buffer[index + 1],
        buffer[index + 2],
        buffer[index + 3],
    ]
}

/// How far a pixel's own middle is from the bubble's, which is the only question the circle is
/// drawn by (see `paint_bubble`).
fn distance(x: u32, y: u32) -> f32 {
    let centre = SIDE as f32 / 2.0;
    let dx = x as f32 + 0.5 - centre;
    let dy = y as f32 + 0.5 - centre;
    (dx * dx + dy * dy).sqrt()
}

/// A picture of one colour, `side × side` BGRA, as every thumbnail a bubble is given is.
fn solid(bgra: [u8; 4]) -> Vec<u8> {
    let mut picture = Vec::with_capacity((SIDE * SIDE * 4) as usize);
    for _ in 0..SIDE * SIDE {
        picture.extend_from_slice(&bgra);
    }
    picture
}

/// The colour of a pixel a mark is drawn in, off the palette: a pixel written at full coverage is
/// exactly the colour it was given, in the buffer's own order (see `paint_bubble`).
fn glyph() -> [u8; 4] {
    let colour = palette().foreground;
    [colour[2], colour[1], colour[0], 255]
}

fn face() -> [u8; 4] {
    let colour = palette().background;
    [colour[2], colour[1], colour[0], 255]
}

/// A painted bubble over a colour, as the desktop would show it: premultiplied over an opaque
/// backdrop, which is the only way the antialiasing of an edge is a thing that can be looked at.
fn over(buffer: &[u8], width: u32, height: u32, backdrop: [u8; 3]) -> Vec<u8> {
    let mut out = vec![255u8; (width * height * 4) as usize];
    for (pixel, source) in out.as_chunks_mut::<4>().0.iter_mut().zip(buffer.chunks(4)) {
        let alpha = source[3] as u32;
        for channel in 0..3 {
            pixel[channel] =
                (source[channel] as u32 + backdrop[channel] as u32 * (255 - alpha) / 255) as u8;
        }
    }
    out
}

/// The bubbles a collapsed pin can become, side by side at both scales, over a desktop's own
/// grey: the design under review.
///
/// What a picture of them shows and a test on their pixels cannot is whether a circle reads as a
/// circle: the ring, the face and the picture inside meet along three curves, and any of the three
/// drawn a whole pixel at a time can pass a per-pixel test and still be a staircase to look at.
/// The faces, left to right: a page's play, a paused sound, a page's own mark, a picture whose
/// frame is not in hand, a playing sound's meter, and a frame with the pause mark over it.
#[test]
fn draws_the_bubbles_faces() {
    let dir = scratch("bubble-faces");
    let backdrop = [128u8, 128, 128];

    for (dpi, tag) in [(96u32, "96dpi"), (192u32, "192dpi")] {
        // The size the window paints the circle at: `PIN_BUBBLE_PIXELS` on a display at this
        // scale (`logical_px`).
        let side = (44.0 * dpi as f32 / 96.0).round() as u32;
        // A picture with something in it to lose: a shade per column, which is a picture shown
        // smaller than it was drawn if anything is (see `picture_pixel`).
        let picture = (0..side * side)
            .flat_map(|index| {
                let shade = ((index % side) * 255 / side) as u8;
                [shade, shade.wrapping_sub(128), 128, 255]
            })
            .collect::<Vec<u8>>();

        let faces = [
            (BubbleArt::None, BubbleMark::Play),
            (BubbleArt::None, BubbleMark::Pause),
            (BubbleArt::None, BubbleMark::Page),
            (BubbleArt::None, BubbleMark::Picture),
            (
                BubbleArt::Audio {
                    level: 0.7,
                    phase: 0.0,
                },
                BubbleMark::Play,
            ),
            (
                BubbleArt::Picture {
                    pixels: &picture,
                    width: side,
                    height: side,
                },
                BubbleMark::Pause,
            ),
        ];

        let width = side * faces.len() as u32;
        let mut strip = image_buffer(width, side);
        for (cell, (art, mark)) in faces.into_iter().enumerate() {
            let mut buffer = image_buffer(side, side);
            paint_bubble(&mut buffer, side, side, &palette(), art, mark);
            let painted = over(&buffer, side, side, backdrop);
            blit_cell(&mut strip, width, side, &painted, side, cell as u32 * side);
        }

        let path = dir.join(format!("bubbles-{tag}.png"));
        write_png(path.clone(), &strip, width, side);
        println!("{}", path.display());

        // The whole strip again at four times the size, which is what a review of an edge is:
        // a circle forty-four pixels across has no edge to look at on its own.
        if dpi == 96 {
            let (large, wide, tall) = zoom(&strip, width, side, 4);
            let path = dir.join(format!("bubbles-{tag}-x4.png"));
            write_png(path.clone(), &large, wide, tall);
            println!("{}", path.display());
        }
    }
}

#[test]
fn the_face_and_the_ring_meet_at_a_ramp_rather_than_a_step() {
    // The ring is a hairline of the accent around a face of the theme's own colour, and the two
    // meet along a circle: a pixel on that circle holds some of each, and a painter that gives it
    // one or the other leaves a staircase where the two colours change (see `paint_bubble`).
    let painted = paint(BubbleArt::None, BubbleMark::Play);
    let face = face();
    let ring = [
        palette().accent[2],
        palette().accent[1],
        palette().accent[0],
        255,
    ];

    // Outwards from the middle of the bubble on its right, which is the bare face: the play mark
    // is a glyph across the middle of the face and does not reach the ring on this row.
    let row = SIDE / 2;
    let colours = (SIDE / 2..SIDE)
        .map(|x| pixel(&painted, x, row))
        .collect::<Vec<_>>();
    let last_face = colours
        .iter()
        .rposition(|&colour| colour == face)
        .expect("the face is drawn");
    let first_ring = colours
        .iter()
        .position(|&colour| colour == ring)
        .expect("the ring is drawn");

    assert!(
        first_ring > last_face + 1,
        "the face runs to {last_face} and the ring starts at {first_ring} with nothing between \
         them: a pixel on the edge of either is both of them in the shares the circle leaves"
    );
}

#[test]
fn the_pictures_round_edge_is_a_ramp_rather_than_a_cut() {
    // The picture a bubble carries is cut to the circle the ring leaves, which is a curve like
    // any other in here: a pixel on it holds some of the picture and some of the ring, and a
    // painter that gives it the one or the other leaves an edge a hand can see the steps in
    // (see `paint_bubble`).
    let upper = [10, 200, 30, 255];
    let lower = [242, 40, 100, 255];
    let mut halves = Vec::with_capacity((SIDE * SIDE * 4) as usize);
    for row in 0..SIDE {
        let band = if row < SIDE / 2 { upper } else { lower };
        for _ in 0..SIDE {
            halves.extend_from_slice(&band);
        }
    }

    let mean = average_color(&halves).expect("a picture of two colours has a mean");
    let ring = [mean[2], mean[1], mean[0], 255];
    let picture = upper;
    let painted = paint(
        BubbleArt::Picture {
            pixels: &halves,
            width: SIDE,
            height: SIDE,
        },
        BubbleMark::Page,
    );

    // A row through the upper half of the picture and out to its edge, where the two colours a
    // pixel can be are the picture's own and the ring's mean — and nothing else.
    let row = SIDE / 4;
    let colours = (SIDE / 2..SIDE)
        .map(|x| pixel(&painted, x, row))
        .collect::<Vec<_>>();
    let last_picture = colours
        .iter()
        .rposition(|&colour| colour == picture)
        .expect("the picture is drawn");
    let first_ring = colours
        .iter()
        .position(|&colour| colour == ring)
        .expect("the ring is drawn");

    assert!(
        first_ring > last_picture + 1,
        "the picture runs to {last_picture} and the ring starts at {first_ring} with nothing \
         between them: the circle the picture is cut to passes between the two"
    );
}

#[test]
fn a_picture_finer_than_the_bubble_is_averaged_rather_than_sampled() {
    // A picture cut to a square is wider than the round face it is put in, so nearly every pixel
    // of the face asks the picture for a point that falls between four of its pixels. What the
    // point falls among is what the pixel is (see `paint_bubble`); the one nearest it is what
    // leaves a picture of a pixel-wide pattern looking like static rather than like a picture.
    let mut checkerboard = Vec::with_capacity((SIDE * SIDE * 4) as usize);
    let ink = [[255u8, 255, 255, 255], [0, 0, 0, 255]];
    for y in 0..SIDE {
        for x in 0..SIDE {
            checkerboard.extend_from_slice(&ink[((x + y) % 2) as usize]);
        }
    }

    let painted = paint(
        BubbleArt::Picture {
            pixels: &checkerboard,
            width: SIDE,
            height: SIDE,
        },
        BubbleMark::Page,
    );

    // The face, and not the ring or the edge of the circle: every pixel of it is a pixel of the
    // picture, so every one of them is a black, a white, or the grey a pixel between them is.
    let mut between = 0;
    for y in 0..SIDE {
        for x in 0..SIDE {
            if distance(x, y) > SIDE as f32 / 2.0 - 6.0 {
                continue;
            }
            if !matches!(pixel(&painted, x, y), [255, 255, 255, 255] | [0, 0, 0, 255]) {
                between += 1;
            }
        }
    }

    assert!(
        between > 100,
        "{between} pixels of the face are neither of the two colours the picture is made of: a \
         picture shown smaller than it was drawn is the average of what it covers"
    );
}

#[test]
fn the_marks_slanted_edge_is_a_share_of_the_ink_rather_than_all_of_it() {
    // The play triangle is the one mark in a bubble whose edge runs across the pixel grid: the
    // pixels it crosses are some of the face and some of the ink, and a painter that gives each of
    // them the one or the other draws a staircase (see `paint_mark`). The mark is one colour laid
    // down at one opacity, so the share of it a pixel carries is read back off the pixel.
    let painted = paint(BubbleArt::None, BubbleMark::Play);
    let face = face();
    // The ink at full coverage, read out of the triangle where nothing else can be: a column of
    // the mark off its own middle row, which every pixel of its edge is a share of rather than
    // the whole of.
    let solid = pixel(&painted, SIDE / 2, SIDE / 2 - 3);
    assert_ne!(solid, face, "the ink of the mark is drawn");

    let share =
        |pixel: [u8; 4]| (pixel[0] as f32 - face[0] as f32) / (solid[0] as f32 - face[0] as f32);

    let mut partial = 0;
    let mut rows = Vec::new();
    for y in SIDE / 2 - 8..=SIDE / 2 + 8 {
        for x in SIDE / 2 - 12..=SIDE / 2 + 12 {
            let pixel = pixel(&painted, x, y);
            if pixel == face || pixel == solid {
                continue;
            }

            let carried = share(pixel);
            if (0.02..0.98).contains(&carried) {
                partial += 1;
                if !rows.contains(&y) {
                    rows.push(y);
                }
            }
        }
    }

    assert!(
        partial > 20,
        "only {partial} pixels of the triangle carry a part of the ink rather than all of it: a \
         pixel the edge crosses holds the area of the triangle that is in it"
    );
    assert!(
        rows.len() > 8,
        "the pixels carrying part of the ink are on {} rows out of the {} the triangle is drawn \
         across, which is an edge rounded off rather than drawn",
        rows.len(),
        SIDE / 2 + 8 - (SIDE / 2 - 8) + 1
    );
}

#[test]
fn a_picture_gives_the_ring_its_own_average_colour() {
    // The ring of a bubble over a picture is the picture's own mean rather than the theme's
    // accent, which is the theme coming to the edge of somebody else's file (see `paint_bubble`).
    let colour = [10, 200, 30, 255];
    let picture = solid(colour);
    let painted = paint(
        BubbleArt::Picture {
            pixels: &picture,
            width: SIDE,
            height: SIDE,
        },
        BubbleMark::Picture,
    );

    // The mean of a picture of one colour is that colour, in a palette's own order.
    assert_eq!(
        average_color(&picture),
        Some([colour[2], colour[1], colour[0]])
    );

    // A few columns in from the right edge is the hairline and not the face, and a pixel there is
    // at full coverage — the one place in a painted bubble where a colour is the colour it was
    // given, untouched by the antialiasing of an edge. The middle is the picture itself.
    assert_eq!(pixel(&painted, SIDE - 3, SIDE / 2), colour);
    assert_ne!(pixel(&painted, SIDE - 3, SIDE / 2), face());
    assert_eq!(pixel(&painted, SIDE / 2, SIDE / 2), colour);

    // And the two halves of a picture of two colours, so that the ring is provably the mean of the
    // whole picture rather than the colour of the pixel it happens to be drawn beside: the ring
    // pixel read below sits in the bottom half, and what is there is neither half's colour.
    let mut halves = Vec::with_capacity((SIDE * SIDE * 4) as usize);
    for row in 0..SIDE {
        let band = if row < SIDE / 2 {
            colour
        } else {
            [242, 40, 100, 255]
        };
        for _ in 0..SIDE {
            halves.extend_from_slice(&band);
        }
    }

    let mean = average_color(&halves).expect("a picture of two colours has a mean");
    let painted = paint(
        BubbleArt::Picture {
            pixels: &halves,
            width: SIDE,
            height: SIDE,
        },
        BubbleMark::Picture,
    );

    assert_eq!(
        pixel(&painted, SIDE - 3, SIDE / 2),
        [mean[2], mean[1], mean[0], 255]
    );
    assert_ne!(mean, [colour[2], colour[1], colour[0]]);
    assert_ne!(mean, [100, 40, 242]);
}

#[test]
fn the_average_of_a_picture_is_the_average_of_what_is_drawn_in_it() {
    // A frame with a hole in it. A transparent pixel is not a pixel of the picture, so it is not an
    // input to its mean: counted as black, it would drag the ring of a cut-out picture towards the
    // dark — a colour the picture is not.
    let opaque = [40, 80, 120, 255];
    let mut picture = solid(opaque);
    picture[..4].copy_from_slice(&[255, 255, 255, 0]);

    assert_eq!(average_color(&picture), Some([120, 80, 40]));

    // And a picture that is transparent everywhere has no colour to give at all, which the caller
    // reads as *take the theme's accent* rather than as black (see `paint_bubble`).
    let clear = vec![0u8; 4 * 4];
    assert_eq!(average_color(&clear), None);
    assert_eq!(average_color(&[]), None);
}

#[test]
fn the_bars_of_a_sound_are_drawn_inside_the_circle() {
    // The one thing a bubble must never do is write outside itself: what is behind it is the
    // desktop, so a bar reaching past the circle is a bar painted over whatever the user has there.
    // The two paints differ in the level alone, so every pixel that differs between them is a pixel
    // of the bars — the face, the ring and the glyph are the same in both.
    let quiet = paint(
        BubbleArt::Audio {
            level: 0.0,
            phase: 1.1,
        },
        BubbleMark::Pause,
    );
    let loud = paint(
        BubbleArt::Audio {
            level: 1.0,
            phase: 1.1,
        },
        BubbleMark::Pause,
    );

    // The circle the bubble is: a disc of the window's own width, less the pixel its edge is
    // antialiased into (see `paint_bubble`), which is the bound every pixel written has to be
    // within (see `draw_histogram`).
    let circle = SIDE as f32 / 2.0 - 1.0;
    let mut drawn = 0;

    for y in 0..SIDE {
        for x in 0..SIDE {
            if pixel(&loud, x, y) == pixel(&quiet, x, y) {
                continue;
            }

            drawn += 1;
            assert!(
                distance(x, y) <= circle,
                "the bars reached {x},{y}, which is outside the circle"
            );
        }
    }

    // A test that asserts nothing was drawn draws nothing: the bars are in the face and not only
    // absent, which is what the assertion above would also be answered by.
    assert!(drawn > 0, "the bars were not drawn at all");
}

#[test]
fn a_playing_file_and_a_paused_one_are_two_different_marks() {
    // The middle row of the two marks: a triangle is solid across its own middle, and two bars
    // leave a gap there — which is the whole of what tells playing from paused at a glance.
    let playing = paint(BubbleArt::None, BubbleMark::Play);
    let paused = paint(BubbleArt::None, BubbleMark::Pause);

    let row = |buffer: &Vec<u8>| {
        (0..SIDE)
            .map(|x| pixel(buffer, x, SIDE / 2))
            .collect::<Vec<_>>()
    };

    assert_ne!(row(&playing), row(&paused));

    let middle = (SIDE / 2) as usize;
    // The middle of the triangle is covered by the mark, so it is no longer the face under it —
    // the mark is drawn translucent, so it is not the glyph's own colour either (see `paint_mark`).
    // The middle of the pause mark is the gap between the two bars, which leaves the face showing.
    assert_ne!(row(&playing)[middle], face());
    assert_ne!(row(&playing)[middle], glyph());
    assert_eq!(row(&paused)[middle], face());
}

#[test]
fn a_mark_goes_over_a_picture_only_where_the_picture_can_be_playing() {
    // What a bubble carries over its own picture is the play/pause of the file on screen, and
    // nothing else: a page's mark and a picture's mark are the marks of a bubble with *no* art, and
    // drawing either over a photograph would be a glyph standing in for a picture that is there.
    let colour = [200, 40, 60, 255];
    let picture = solid(colour);
    let painted = |mark: BubbleMark| {
        paint(
            BubbleArt::Picture {
                pixels: &picture,
                width: SIDE,
                height: SIDE,
            },
            mark,
        )
    };

    let middle = (SIDE / 2, SIDE / 2);
    for mark in [BubbleMark::Page, BubbleMark::Picture] {
        // The picture itself, untouched: no glyph and no scrim behind one.
        assert_eq!(pixel(&painted(mark), middle.0, middle.1), colour);
    }

    for mark in [BubbleMark::Play, BubbleMark::Pause] {
        assert_ne!(pixel(&painted(mark), middle.0, middle.1), colour);
    }
}

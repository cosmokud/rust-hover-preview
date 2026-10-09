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
    assert_eq!(row(&playing)[middle], glyph());
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

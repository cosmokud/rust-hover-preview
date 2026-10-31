use super::*;

/// The band a card is painted over in these tests: a picture's own box, big enough that a panel
/// and the two lines in it are drawn rather than rounded away.
const BAND: (u32, u32) = (800, 450);

/// A palette with three colours nothing can confuse for one another, so that a pixel read back
/// is unambiguously the scrim, the panel or a letter.
fn palette() -> ChromePalette {
    ChromePalette {
        background: [24, 26, 32],
        foreground: [240, 240, 240],
        accent: [97, 175, 239],
        dark: true,
    }
}

/// A band painted with one frame of the card at a coverage of `coverage` and a display at `dpi`.
fn painted(coverage: f32, dpi: u32) -> DibSurface {
    let surface = DibSurface::create(BAND.0, BAND.1).expect("a surface of the band's own size");
    let blanked = unsafe {
        std::slice::from_raw_parts_mut(surface.bits(), BAND.0 as usize * BAND.1 as usize * 4)
    };
    blanked.fill(0);
    paint_corruption_card(&surface, &palette(), coverage, dpi);
    surface
}

/// One pixel of a painted band, in the premultiplied BGRA the outside of it is in.
fn pixel(surface: &DibSurface, x: u32, y: u32) -> [u8; 4] {
    let index = (y as usize * surface.width as usize + x as usize) * 4;
    let pixels = unsafe {
        std::slice::from_raw_parts(surface.bits(), BAND.0 as usize * BAND.1 as usize * 4)
    };
    [
        pixels[index],
        pixels[index + 1],
        pixels[index + 2],
        pixels[index + 3],
    ]
}

/// How many rows of the band's middle column are opaque, which is how tall the panel is.
fn panel_rows(surface: &DibSurface) -> u32 {
    (0..BAND.1)
        .filter(|y| pixel(surface, BAND.0 / 2, *y)[3] == 255)
        .count() as u32
}

#[test]
fn nothing_is_drawn_while_the_card_is_not_on_screen() {
    // The first frame of the run, and the last: both are a coverage of nothing, and neither is a
    // card the strip is to carry.
    for coverage in [0.0, 0.0] {
        let surface = painted(coverage, 96);
        for y in (0..BAND.1).step_by(17) {
            for x in (0..BAND.0).step_by(17) {
                assert_eq!(
                    pixel(&surface, x, y),
                    [0, 0, 0, 0],
                    "a card at no coverage leaves the band as it found it"
                );
            }
        }
    }
}

#[test]
fn the_picture_is_dimmed_and_the_card_stands_in_the_middle_of_it() {
    let surface = painted(1.0, 96);

    // The band's own corner is the scrim and nothing else: black at its own coverage, which is
    // what dims the film under it rather than hiding it.
    let corner = pixel(&surface, 2, 2);
    assert_eq!(corner[3], 97, "the scrim is 38% of the way to black");
    assert_eq!((corner[0], corner[1], corner[2]), (0, 0, 0));

    // The middle of the band is the panel: opaque, and in the theme's own page colour rather than
    // the film's.
    let middle = pixel(&surface, BAND.0 / 2, BAND.1 / 2);
    assert_eq!(middle[3], 255, "the panel is opaque over the dimmed band");

    // And the panel is a card rather than a band of its own: it stops well short of the edges.
    assert_eq!(pixel(&surface, 0, BAND.1 / 2)[3], 97);
    assert_eq!(pixel(&surface, BAND.0 - 1, BAND.1 / 2)[3], 97);
}

#[test]
fn the_card_is_drawn_larger_on_a_display_at_a_higher_dpi() {
    // Every size in the card is a logical pixel scaled by the display, so the same band on a
    // display at 200% draws a taller panel — the one thing that would be wrong if a literal had
    // been written where `text_paint::scaled` belongs.
    let at_96 = panel_rows(&painted(1.0, 96));
    let at_192 = panel_rows(&painted(1.0, 192));

    assert!(at_96 > 0, "a card is drawn at 100%");
    assert!(
        at_192 > at_96,
        "a card on a display at 200% is larger than one at 100%: {at_192} against {at_96}"
    );
}

#[test]
fn the_card_fades_rather_than_blinking_out() {
    // Halfway up the fade the panel is half as opaque as it is whole, which is the one number the
    // three beats of the run are drawn from.
    let half = pixel(&painted(0.5, 96), BAND.0 / 2, BAND.1 / 2);
    let whole = pixel(&painted(1.0, 96), BAND.0 / 2, BAND.1 / 2);

    assert!(
        (126..=129).contains(&half[3]),
        "halfway through the fade the panel is half opaque: {}",
        half[3]
    );
    assert_eq!(whole[3], 255);
}

//! The card a page's find bar floats in: where its two rows sit, what a point on it means, and what it
//! draws (see `find_bar`).

use super::*;

fn palette() -> ChromePalette {
    ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    }
}

fn card_state(query: &str, current: Option<usize>, total: usize) -> FindBarPaint {
    FindBarPaint {
        query: query.to_string(),
        hovered: None,
        pressed: None,
        current,
        total,
        case_sensitive: false,
        whole_word: false,
    }
}

/// The card is drawn against the page's top-right corner, under the caption, and its two rows are one
/// under the other inside it: the row a query is typed into, and the row of switches and commands
/// below that (see `find_bar_layout`).
#[test]
fn the_card_sits_against_the_pages_top_right_corner_under_its_caption() {
    let card = find_bar_layout(400, 300, 30, 96);

    assert_eq!(
        card.panel.top,
        30 + 10,
        "the card keeps its margin under the caption"
    );
    assert_eq!(
        card.panel.right,
        400 - 10,
        "and the same margin from the page's right edge"
    );
    assert!(
        card.panel.left > 0,
        "the card is inside the page's left edge too"
    );
    assert!(card.panel.bottom < 30 + 300, "and above its bottom edge");

    assert!(
        card.query.right <= card.count.left,
        "the count keeps its own room at the right end of the query's row"
    );
    assert!(
        card.case.top >= card.query.bottom,
        "the switches are on the row under the query"
    );
    for rect in [
        card.case,
        card.whole_word,
        card.previous,
        card.next,
        card.close,
    ] {
        assert!(
            rect.top >= card.panel.top && rect.bottom <= card.panel.bottom,
            "every control is inside the card it is drawn on"
        );
    }
}

/// A pad, a row, a glyph and the shadow are all the display's: the card at 200% is the card, twice the
/// size, rather than a small card in a large corner (see `find_bar_layout`).
#[test]
fn every_metric_of_the_card_scales_with_the_display() {
    let at_96 = find_bar_layout(800, 600, 0, 96);
    let at_192 = find_bar_layout(800, 600, 0, 192);

    let width = |card: FindBar| card.panel.right - card.panel.left;
    let height = |card: FindBar| card.panel.bottom - card.panel.top;
    assert_eq!(
        width(at_192),
        width(at_96) * 2,
        "the card's own width doubles"
    );
    assert_eq!(height(at_192), height(at_96) * 2, "and so does its height");

    let side = |rect: RECT| rect.right - rect.left;
    assert_eq!(
        side(at_192.close),
        side(at_96.close) * 2,
        "and the boxes a hand answers in"
    );
}

/// The five controls answer at their own boxes and nowhere else, and the card's own face — the query's
/// row and the gaps between its buttons — is the card's without being a control's: a press there is not
/// a press on the page under it, and there is no caret in the query for it to place (see
/// `FindBar::part_at`).
#[test]
fn every_control_answers_at_its_own_box_and_the_card_holds_the_rest() {
    let card = find_bar_layout(400, 300, 0, 96);

    for (rect, part) in [
        (card.case, FindBarPart::Case),
        (card.whole_word, FindBarPart::WholeWord),
        (card.previous, FindBarPart::Previous),
        (card.next, FindBarPart::Next),
        (card.close, FindBarPart::Close),
    ] {
        let (x, y) = ((rect.left + rect.right) / 2, (rect.top + rect.bottom) / 2);
        assert_eq!(
            card.part_at(x, y),
            Some(part),
            "{part:?} is answered at its middle"
        );
    }

    let (x, y) = (card.query.left + 1, card.query.top + 1);
    assert_eq!(card.part_at(x, y), None, "the query's row is no control");
    assert!(
        card.holds(x, y),
        "but it is the card's, and the page under it is not pressed"
    );

    assert!(
        !card.holds(card.panel.left - 1, card.panel.top + 1),
        "a point beside the card is the page's"
    );
    assert_eq!(card.part_at(card.panel.left - 1, card.panel.top + 1), None);
}

/// Painting the card puts a panel over the page at its corner, and what it says is what a second paint
/// with another query and another switch draw differently: the count and a pressed switch are the two
/// things on it that change without the card moving (see `paint_find_bar`).
#[test]
fn the_card_is_painted_over_the_page_and_its_pressed_switches_and_count_change_it() {
    let (page_width, page_height) = (400, 300);
    let card = find_bar_layout(page_width, page_height, 0, 96);
    let surface = DibSurface::create(
        (card.panel.right - card.panel.left) as u32,
        (card.panel.bottom - card.panel.top) as u32,
    )
    .expect("a surface");

    let paint = |state: &FindBarPaint| {
        let mut buffer = image_buffer(page_width as u32, page_height as u32);
        paint_find_bar(
            &mut buffer,
            page_width,
            &palette(),
            &card,
            state,
            &surface,
            96,
        );
        buffer
    };
    let pixel = |buffer: &[u8], x: i32, y: i32| {
        let at = (y as usize * page_width as usize + x as usize) * 4;
        [buffer[at], buffer[at + 1], buffer[at + 2], buffer[at + 3]]
    };

    let first = paint(&card_state("needle", Some(0), 2));
    assert_ne!(
        pixel(&first, card.panel.left + 2, card.panel.top + 2)[3],
        0,
        "the card's own corner is painted, and it is not a hole onto the page"
    );
    assert_eq!(
        pixel(&first, 0, page_height - 1)[3],
        0,
        "the page away from the card is left as it was"
    );

    let other = paint(&card_state("another", Some(1), 7));
    assert_ne!(
        first, other,
        "another query and another count are another card drawn"
    );

    let held = paint(&FindBarPaint {
        case_sensitive: true,
        whole_word: true,
        ..card_state("needle", Some(0), 2)
    });
    assert_ne!(
        first, held,
        "and a switch that is on is drawn as pressed rather than as off"
    );

    // The card as it is drawn, so a render can be looked at rather than only asserted (see
    // `scratch`).
    write_png(
        scratch("find-bar").join("the-card.png"),
        &first,
        page_width as u32,
        page_height as u32,
    );
}

/// The card painted over an empty page, at the scale a display is drawn at: what a test asks its
/// pixels of, and the card it stood on (see `paint_find_bar`).
fn painted_card(state: &FindBarPaint, dpi: u32) -> (Vec<u8>, FindBar) {
    let (page_width, page_height) = (400, 300);
    let card = find_bar_layout(page_width, page_height, 0, dpi);
    let surface = DibSurface::create(
        (card.panel.right - card.panel.left) as u32,
        (card.panel.bottom - card.panel.top) as u32,
    )
    .expect("a surface");

    let mut buffer = image_buffer(page_width as u32, page_height as u32);
    paint_find_bar(
        &mut buffer,
        page_width,
        &palette(),
        &card,
        state,
        &surface,
        dpi,
    );
    (buffer, card)
}

/// One pixel of a painted page, in the four channels the surface is written in — premultiplied, which
/// is why a colour drawn at full coverage is the colour it was asked for in the two channels swapped.
fn pixel_at(buffer: &[u8], width: i32, x: i32, y: i32) -> [u8; 4] {
    let at = (y as usize * width as usize + x as usize) * 4;
    [buffer[at], buffer[at + 1], buffer[at + 2], buffer[at + 3]]
}

/// A step mark is a chevron and not a narrow `v`: wider than it is tall, its two legs closing on one
/// apex, and that apex on a pixel centre. The mark this card was drawn with before was ten pixels
/// square — as tall as it was wide, which is what made it read as a spike rather than as a chevron —
/// and it fails the first of these (see `step_mark_legs` and `step_mark_shape`).
#[test]
fn the_step_mark_is_wider_than_it_is_tall_and_closes_on_an_apex_on_a_pixel() {
    let (width, height) = step_mark_shape(1.0);
    assert!(
        width > height,
        "the mark is {width} across and {height} tall: a chevron is wider than it is tall"
    );

    let legs = step_mark_legs(100, 50, width, height, false);
    let (left_end, left_apex) = legs[0];
    let (right_end, right_apex) = legs[1];
    assert_eq!(
        left_apex, right_apex,
        "the two legs end at the one apex the mark points with"
    );

    // The apex is on a pixel centre, and both free ends are a whole number of pixels from it: a mark
    // whose apex sat between two pixels would draw its join as two half-lit ones rather than as one.
    assert_eq!(
        (left_apex.0.fract(), left_apex.1.fract()),
        (0.5, 0.5),
        "the apex is on a pixel centre"
    );
    assert!(
        left_apex.1 > left_end.1 && left_apex.1 > right_end.1,
        "a mark pointing down has its apex under its two free ends"
    );
    assert_eq!(
        left_apex.0 - left_end.0,
        right_end.0 - left_apex.0,
        "and the two legs reach the same distance either side of it"
    );
    assert_eq!(
        left_end.1, right_end.1,
        "the free ends are on one row, so the mark's box is a box"
    );

    let up = step_mark_legs(100, 50, width, height, true);
    assert!(
        up[0].1 .1 < up[0].0 .1,
        "a mark pointing up has its apex over its free ends"
    );

    // The mark is the display's like everything else on this card, and stays a chevron at a scale.
    let (wide, tall) = step_mark_shape(1.5);
    assert!(
        wide > width && tall > height,
        "the mark grows with the display"
    );
    assert!(wide > tall, "and is still a chevron at 150%");
}

/// The caret is placed where the query was really drawn. A query wider than its field is drawn from
/// the right, so a caret counted from the field's left edge is nowhere near the text — the fault this
/// asks after — and a caret placed after the text's own end stands at the field's right instead (see
/// `paint_find_bar`).
#[test]
fn the_caret_stands_after_the_last_character_of_a_query_too_long_for_its_field() {
    let query = "the needle that is far too long for a field of this card to hold in its row";
    let state = card_state(query, Some(0), 1);

    // The text, measured the way the paint measures it, is wider than the room the field has.
    let probe = find_bar_layout(400, 300, 0, 96);
    let surface = DibSurface::create(
        (probe.panel.right - probe.panel.left) as u32,
        (probe.panel.bottom - probe.panel.top) as u32,
    )
    .expect("a surface");
    let measured = measure_text(&surface, &caption_style(palette().foreground), query, 1.0);
    assert!(
        measured > probe.query.right - probe.query.left,
        "the query is {measured} wide and its field is {} across",
        probe.query.right - probe.query.left
    );

    let (buffer, card) = painted_card(&state, 96);
    let accent = palette().accent;
    let lit: Vec<(i32, i32)> = (card.query.left..card.query.right + 1)
        .flat_map(|x| (card.query.top..card.query.bottom).map(move |y| (x, y)))
        .filter(|(x, y)| pixel_at(&buffer, 400, *x, *y) == [accent[2], accent[1], accent[0], 255])
        .collect();

    let mut columns: Vec<i32> = lit.iter().map(|(x, _)| *x).collect();
    columns.dedup();
    assert!(
        columns.contains(&card.query.right),
        "the caret is at the end of the text, at the field's right edge"
    );
    assert!(
        !columns.contains(&card.query.left),
        "and not at the field's left edge, where the text does not begin"
    );
    assert_eq!(columns.len(), 1, "the caret is one column and not a band");

    // Its rows are the text's own line and not the row's: the run's height, inset a pixel at each end.
    let rows: Vec<i32> = lit.iter().map(|(_, y)| *y).collect();
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        1.0,
    );
    let line_top = card.query.top + (card.query.bottom - card.query.top - cell) / 2;
    assert_eq!(
        (rows.iter().min().copied(), rows.iter().max().copied()),
        (Some(line_top + 1), Some(line_top + cell - 2)),
        "the caret spans the text's line box, less a pixel at either end"
    );
}

/// The control under the pointer is washed under its own mark, a press is washed further with the
/// mark in the accent, and a switch that is on keeps its own wash under either of them: the two states
/// a caption's button is drawn in, on this card's five controls (see `paint_find_bar`).
#[test]
fn the_control_under_the_pointer_is_washed_and_a_press_is_washed_further() {
    let plain = painted_card(&card_state("needle", Some(0), 2), 96);
    let hovered = painted_card(
        &FindBarPaint {
            hovered: Some(FindBarPart::Previous),
            ..card_state("needle", Some(0), 2)
        },
        96,
    );
    let pressed = painted_card(
        &FindBarPaint {
            hovered: Some(FindBarPart::Previous),
            pressed: Some(FindBarPart::Previous),
            ..card_state("needle", Some(0), 2)
        },
        96,
    );

    assert_ne!(plain.0, hovered.0, "a pointer on a control draws it lit");
    assert_ne!(
        hovered.0, pressed.0,
        "and a press held on it draws it pressed rather than merely lit"
    );

    // Only the box the pointer is on is washed: the button beside it is the card's own face still.
    let beside = (
        (plain.1.next.left + plain.1.next.right) / 2,
        (plain.1.next.top + plain.1.next.bottom) / 2,
    );
    assert_eq!(
        pixel_at(&plain.0, 400, beside.0, beside.1),
        pixel_at(&hovered.0, 400, beside.0, beside.1),
        "the control the pointer is not on is drawn exactly as it was"
    );

    // A switch that is on is washed with the accent, and stays so when the pointer is also on it: the
    // pointer's wash goes over the switch's rather than in place of it.
    let on = painted_card(
        &FindBarPaint {
            case_sensitive: true,
            ..card_state("needle", Some(0), 2)
        },
        96,
    );
    let on_and_hovered = painted_card(
        &FindBarPaint {
            case_sensitive: true,
            hovered: Some(FindBarPart::Case),
            ..card_state("needle", Some(0), 2)
        },
        96,
    );
    let corner = (plain.1.case.left + 2, plain.1.case.top + 2);
    let lit_pixel = pixel_at(&on.0, 400, corner.0, corner.1);
    let hovered_pixel = pixel_at(&on_and_hovered.0, 400, corner.0, corner.1);
    assert!(
        lit_pixel[0] > lit_pixel[2],
        "a switch that is on is washed with the accent's own blue"
    );
    assert_ne!(
        lit_pixel, hovered_pixel,
        "and a pointer on it washes it a step further still"
    );
    assert!(
        hovered_pixel[0] > hovered_pixel[2],
        "while the accent's wash reads through the pointer's"
    );
}

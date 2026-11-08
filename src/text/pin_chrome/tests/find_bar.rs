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
        // A query that has just been typed: the caret after its last character and nothing selected,
        // which is the state the caret's own placement is asked about.
        caret: query.len(),
        anchor: query.len(),
        caret_lit: true,
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
    let card = find_bar_layout(400, 300, 30, 96, "1/3", false);

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
        card.query.right <= card.case.left,
        "the query is what is left of the row the switches are on"
    );
    for rect in [card.case, card.whole_word, card.count, card.previous] {
        assert!(
            card.query.right <= rect.left,
            "and every control of the row sits to the right of it"
        );
    }
    for rect in [
        card.case,
        card.whole_word,
        card.count,
        card.previous,
        card.next,
        card.close,
    ] {
        assert!(
            rect.top >= card.panel.top
                && rect.bottom <= card.panel.bottom
                && rect.top <= card.query.bottom
                && rect.bottom >= card.query.top,
            "every one of them is on the card's one row"
        );
    }
}

/// A pad, a row, a glyph and the shadow are all the display's: the card at 200% is the card, twice the
/// size, rather than a small card in a large corner (see `find_bar_layout`).
#[test]
fn every_metric_of_the_card_scales_with_the_display() {
    let at_96 = find_bar_layout(800, 600, 0, 96, "1/3", false);
    let at_192 = find_bar_layout(800, 600, 0, 192, "1/3", false);

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
    let card = find_bar_layout(400, 300, 0, 96, "1/3", false);

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
    let card = find_bar_layout(page_width, page_height, 0, 96, "1/3", false);
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
    painted_card_on((400, 300), state, dpi)
}

/// The same card painted on a page of a size of the test's own: a row wide enough to hold the query
/// and every control needs a page that is, and the narrow page is the case a test paints on purpose
/// (see `find_bar_layout`).
fn painted_card_on(
    (page_width, page_height): (i32, i32),
    state: &FindBarPaint,
    dpi: u32,
) -> (Vec<u8>, FindBar) {
    let card = find_bar_layout(page_width, page_height, 0, dpi, "1/3", false);
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

/// The caret is placed where the query was really drawn. A query wider than its field is scrolled by
/// the caret, so a caret counted from the field's left edge is nowhere near the text — the fault this
/// asks after — and a caret placed after the text's own end stands a margin inside the field's right
/// edge rather than on it (see `query_run` and `paint_find_bar`).
#[test]
fn the_caret_stands_after_the_last_character_of_a_query_too_long_for_its_field() {
    let query = "the needle that is far too long for a field of this card to hold in its row";
    let state = card_state(query, Some(0), 1);

    // The text, measured the way the paint measures it, is wider than the room the field has.
    let probe = find_bar_layout(400, 300, 0, 96, "1/3", false);
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
        columns.contains(&(card.query.right - text_paint::scaled(FIND_BAR_CARET_MARGIN, 1.0))),
        "the caret is at the end of the text, a margin inside the field's right edge"
    );
    assert!(
        !columns.contains(&card.query.right),
        "and never on the field's own last column, which is the count's side of it"
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

/// A press in the query names a place, and the place is counted from the column the text is really
/// drawn at: a short query begins at the field's left edge, and one too wide for its row begins to
/// the left of it, so a press at the field's own right edge is the end of the query in both cases
/// and a press at its left edge is not the beginning of a query that has been cut (`query_origin`,
/// `query_index_at`).
#[test]
fn a_press_names_a_place_in_the_query_from_the_column_it_is_drawn_at() {
    let card = find_bar_layout(400, 300, 0, 96, "1/3", false);

    // The boundary nearest the point: a point on a character's left half is before it and one on its
    // right half after it, which is the same rule for a real run and for these columns.
    let columns = [0, 10, 20, 30];
    assert_eq!(index_at_column(&columns, 100, 99), 0, "before the run");
    assert_eq!(
        index_at_column(&columns, 100, 104),
        0,
        "left half of the first"
    );
    assert_eq!(index_at_column(&columns, 100, 106), 1, "right half of it");
    assert_eq!(index_at_column(&columns, 100, 124), 2);
    assert_eq!(index_at_column(&columns, 100, 130), 3, "the run's end");
    assert_eq!(index_at_column(&columns, 100, 400), 3, "and past it");

    // A short query starts at the field's left edge, so the left edge is its beginning and the right
    // edge is past its end.
    let short = "needle";
    assert_eq!(
        query_index_at(&card, short, short.len(), card.query.left, 96),
        0
    );
    assert_eq!(
        query_index_at(&card, short, short.len(), card.query.right - 1, 96),
        short.len(),
        "a point past the end of the run is the caret after its last character"
    );

    // The same query with a character outside the basic plane in it: the place is a byte index, so
    // the caret after `😀` is four bytes into the query rather than one.
    assert_eq!(
        query_index_at(&card, "a😀b", "a😀b".len(), card.query.right - 1, 96),
        "a😀b".len()
    );

    // Too wide for the field: the run is drawn from the right, so the field's own left edge is not
    // where the text begins, and a press there is not the caret at the start of the query.
    let long = "the needle that is far too long for a field of this card to hold in its row";
    let surface = DibSurface::create(
        (card.panel.right - card.panel.left) as u32,
        (card.panel.bottom - card.panel.top) as u32,
    )
    .expect("a surface");
    let measured = measure_text(&surface, &caption_style(palette().foreground), long, 1.0);
    assert!(measured > card.query.right - card.query.left);
    assert!(
        query_origin(&card, measured) < card.query.left,
        "the run begins to the left of the field it is cut by"
    );
    let at_left = query_index_at(&card, long, long.len(), card.query.left, 96);
    assert!(at_left > 0, "the field's left edge is inside the run");
    assert_eq!(
        query_index_at(&card, long, long.len(), card.query.right, 96),
        long.len(),
        "and its right edge is the end of the query"
    );
}

/// The field's two additions to the card are drawn as the page draws them: a selection is washed
/// between its two ends, and a caret in the off half of its blink draws no column at all — so a
/// repaint that changed only the blink is a card a pixel test can tell apart (see `paint_find_bar`).
#[test]
fn a_selection_is_washed_in_the_field_and_the_caret_obeys_its_blink() {
    let plain = painted_card(&card_state("needle", Some(0), 2), 96);

    let selected = painted_card(
        &FindBarPaint {
            anchor: 0,
            caret: 3,
            ..card_state("needle", Some(0), 2)
        },
        96,
    );
    assert_ne!(
        plain.0, selected.0,
        "a selection is drawn rather than left as the card's own face"
    );

    // The wash is inside the field's own row, and the card beside the field is left exactly as it
    // was: a selection drawn past the field would be a wash over the count.
    let at = (
        plain.1.query.left + 1,
        (plain.1.query.top + plain.1.query.bottom) / 2,
    );
    assert_ne!(
        pixel_at(&plain.0, 400, at.0, at.1),
        pixel_at(&selected.0, 400, at.0, at.1),
        "the selected columns are washed"
    );
    let beside = (plain.1.panel.left + 1, at.1);
    assert_eq!(
        pixel_at(&plain.0, 400, beside.0, beside.1),
        pixel_at(&selected.0, 400, beside.0, beside.1),
        "and nothing outside the field is"
    );

    // The caret in the other half of its blink: the accent column is gone and everything else is
    // the card as it was.
    let unlit = painted_card(
        &FindBarPaint {
            caret_lit: false,
            ..card_state("needle", Some(0), 2)
        },
        96,
    );
    let accent = palette().accent;
    let lit_pixels = |buffer: &[u8]| {
        let card = painted_card(&card_state("needle", Some(0), 2), 96).1;
        (card.query.left..card.query.right)
            .flat_map(|x| (card.query.top..card.query.bottom).map(move |y| (x, y)))
            .filter(|(x, y)| {
                pixel_at(buffer, 400, *x, *y) == [accent[2], accent[1], accent[0], 255]
            })
            .count()
    };
    assert!(
        lit_pixels(&plain.0) > 0,
        "the lit caret is an accent column"
    );
    assert_eq!(
        lit_pixels(&unlit.0),
        0,
        "and the unlit half of the blink draws no caret at all"
    );
}

/// The pixels of the query's own row that are outside the field, as `(x, y)`: what a run drawn wider
/// than its field — or a wash or a caret welded past its edge — would ink over the card's padding
/// (see `paint_find_bar`).
fn strays_on_the_query_row(
    buffer: &[u8],
    plain: &[u8],
    card: FindBar,
    width: i32,
) -> Vec<(i32, i32)> {
    (card.query.top..card.query.bottom)
        .flat_map(|y| {
            (card.panel.left..card.query.left)
                .chain(card.query.right..card.panel.right)
                .map(move |x| (x, y))
        })
        .filter(|(x, y)| pixel_at(buffer, width, *x, *y) != pixel_at(plain, width, *x, *y))
        .collect()
}

/// A query too wide for its field scrolls under its caret, and the field is what cuts it: at `Home`
/// the run's beginning is what the reader sees, a margin inside the field with the caret before its
/// first character; at `End` the run's end, a margin inside the other edge; and a caret whose own
/// column the field already shows leaves the run exactly where it rests. In none of the three does a
/// glyph reach the card's padding beside the field (`query_run`, `paint_find_bar`).
#[test]
fn a_long_query_scrolls_under_its_caret_and_is_cut_at_the_field() {
    const LONG: &str =
        "WWW the same word runs on far past the width of this card's own field and keeps going";
    // The card with no query at all, and the same `1/3` count: every pixel a long query changes
    // outside the field is a pixel that differs from this card, and the count is not one of them.
    let (plain, card) = painted_card(&card_state("", Some(0), 3), 96);
    let accent = palette().accent;
    let accent_column = [accent[2], accent[1], accent[0], 255];
    let margin = text_paint::scaled(FIND_BAR_CARET_MARGIN, 1.0);

    let inked = |buffer: &[u8], x: i32| {
        (card.query.top..card.query.bottom)
            .any(|y| pixel_at(buffer, 400, x, y) != pixel_at(&plain, 400, x, y))
    };
    let caret_columns = |buffer: &[u8]| -> Vec<i32> {
        (card.query.left..card.query.right)
            .filter(|x| {
                (card.query.top..card.query.bottom)
                    .any(|y| pixel_at(buffer, 400, *x, y) == accent_column)
            })
            .collect()
    };

    // `Home` on a query typed to its end: the caret's own character is off the field where the run
    // rests, so the run scrolls to show the beginning of the query.
    let home = painted_card(
        &FindBarPaint {
            caret: 0,
            anchor: 0,
            ..card_state(LONG, Some(0), 3)
        },
        96,
    );
    assert_eq!(
        strays_on_the_query_row(&home.0, &plain, card, 400),
        Vec::new(),
        "the beginning the field has no room for is cut at its edge, not drawn over the padding"
    );
    assert_eq!(
        caret_columns(&home.0),
        vec![card.query.left + margin],
        "the caret at Home stands a margin inside the field"
    );
    assert!(
        (card.query.left + margin + 1..card.query.left + margin + 12).any(|x| inked(&home.0, x)),
        "with the query's first character drawn beside it"
    );

    // `End`, where a reader typing has the caret: the run's own end stays on the field, a margin
    // inside its right edge, and it is the run's beginning that is cut.
    let end = painted_card(&card_state(LONG, Some(0), 3), 96);
    assert_eq!(
        strays_on_the_query_row(&end.0, &plain, card, 400),
        Vec::new(),
        "a run that overflows the field is cut at its right edge too"
    );
    assert_eq!(
        caret_columns(&end.0),
        vec![card.query.right - margin],
        "the caret at End keeps a margin from the field's right edge"
    );
    assert!(
        (card.query.left..card.query.left + 8).any(|x| inked(&end.0, x)),
        "and the run is drawn up to the field's left edge"
    );

    // A caret whose column the field already shows leaves the run where it rests: the text under the
    // reader does not shuffle for an arrow key they can see the caret move on.
    let surface = DibSurface::create(
        (card.panel.right - card.panel.left) as u32,
        (card.panel.bottom - card.panel.top) as u32,
    )
    .expect("a surface");
    let columns = measure_run(&surface, &caption_style(palette().foreground), LONG, 1.0);
    let measured = columns.last().copied().unwrap_or(0);
    assert!(
        measured > card.query.right - card.query.left,
        "the query is wider than its field"
    );
    let rest = query_origin(&card, measured);
    let in_view: Vec<usize> = (0..=LONG.chars().count())
        .filter(|index| {
            let at = rest + columns[*index];
            at > card.query.left + 3 * margin && at < card.query.right - 3 * margin
        })
        .collect();
    assert!(
        in_view.len() > 2,
        "the field shows a stretch of the run at rest"
    );
    for index in [in_view[0], in_view[in_view.len() - 1]] {
        assert_eq!(
            query_run(&card, columns.clone(), index, 96).origin,
            rest,
            "a caret already in view does not move the run"
        );
    }

    let painted = |index: usize| {
        let at = LONG
            .char_indices()
            .nth(index)
            .map_or(LONG.len(), |(at, _)| at);
        painted_card(
            &FindBarPaint {
                caret: at,
                anchor: at,
                ..card_state(LONG, Some(0), 3)
            },
            96,
        )
    };
    let column = |buffer: &[u8]| -> Vec<[u8; 4]> {
        (card.query.top..card.query.bottom)
            .map(|y| pixel_at(buffer, 400, card.query.left, y))
            .collect()
    };
    let first_caret = painted(in_view[0]);
    let last_caret = painted(in_view[in_view.len() - 1]);
    assert_eq!(
        column(&first_caret.0),
        column(&last_caret.0),
        "the run is drawn from the column it rested at, whichever of its characters the caret is on"
    );
    for painted in [&first_caret, &last_caret] {
        assert_eq!(
            strays_on_the_query_row(&painted.0, &plain, card, 400),
            Vec::new(),
            "and is still cut at the field"
        );
    }

    // The margin is the display's like every other metric of the card, and the cut is the field's at
    // any scale (see `find_bar_layout`).
    // The card at 200% needs a page twice as wide to hold one row, so this leg is painted on one: a
    // 400 px page at 200% is a card 200 logical pixels across, which is a field of nothing.
    let page = (800, 600);
    let (plain_200, card_200) = painted_card_on(page, &card_state("", Some(0), 3), 192);
    let home_200 = painted_card_on(
        page,
        &FindBarPaint {
            caret: 0,
            anchor: 0,
            ..card_state(LONG, Some(0), 3)
        },
        192,
    );
    let caret_200: Vec<i32> = (card_200.query.left..card_200.query.right)
        .filter(|x| {
            (card_200.query.top..card_200.query.bottom)
                .any(|y| pixel_at(&home_200.0, page.0, *x, y) == accent_column)
        })
        .collect();
    assert_eq!(
        caret_200.first().copied(),
        Some(card_200.query.left + text_paint::scaled(FIND_BAR_CARET_MARGIN, 2.0)),
        "the margin is the display's: twice as wide at 200%"
    );
    assert_eq!(
        strays_on_the_query_row(&home_200.0, &plain_200, card_200, page.0),
        Vec::new(),
        "and the field still cuts the run at 200%"
    );
}

/// A press on a scrolled run is answered from the column the caret left the run at: the first
/// character the field shows is the one a press on it names — not the query's first, which the field
/// has cut — and the field's own last column names the character drawn under it (`query_run`,
/// `query_index_at`).
#[test]
fn a_press_on_a_scrolled_run_names_the_character_under_the_pointer() {
    let card = find_bar_layout(400, 300, 0, 96, "1/3", false);
    let long =
        "When the same word runs on far past the width of this card's own field it keeps going";
    let surface = DibSurface::create(
        (card.panel.right - card.panel.left) as u32,
        (card.panel.bottom - card.panel.top) as u32,
    )
    .expect("a surface");
    let columns = measure_run(&surface, &caption_style(palette().foreground), long, 1.0);
    let byte_at = |index: usize| {
        long.char_indices()
            .nth(index)
            .map_or(long.len(), |(at, _)| at)
    };

    // The caret at the end of the run, where a reader typing a long query has it, so the run is
    // scrolled to its end and its beginning is what the field cut away.
    let typed = query_run(&card, columns.clone(), long.chars().count(), 96);
    assert!(
        typed.origin < card.query.left,
        "the run begins off the field"
    );
    let first = (1..columns.len())
        .find(|index| typed.origin + columns[*index] > card.query.left)
        .expect("a character the field shows");
    assert!(
        first > 0,
        "the field's first character is not the query's first"
    );
    assert_eq!(
        query_index_at(&card, long, long.len(), typed.origin + columns[first], 96),
        byte_at(first),
        "a press on the first character the field shows names that character"
    );

    // The caret at the run's beginning: the run is scrolled to show its start, so the field's own last
    // column still has a character under it.
    let home = query_run(&card, columns.clone(), 0, 96);
    assert!(
        home.origin > card.query.left,
        "the run's beginning is inside the field at Home"
    );
    let at = card.query.right - 1 - home.origin;
    let under = (1..columns.len())
        .find(|index| columns[index - 1] <= at && at < columns[*index])
        .expect("a character under the field's last column");
    let named = query_index_at(&card, long, 0, card.query.right - 1, 96);
    assert!(
        named == byte_at(under - 1) || named == byte_at(under),
        "a press on the field's last column names the character drawn under it"
    );
    assert_ne!(named, long.len(), "which is not the end of the query");
}

/// `Ctrl`+`A` on a query too wide for its field washes the part of the run the field holds and nothing
/// beyond it: a wash as wide as the run would run over the card's padding and, sealed as wide as the
/// run, carry a band of another colour across the panel (see `paint_find_bar`).
#[test]
fn select_all_on_a_long_query_washes_only_the_part_of_the_run_the_field_holds() {
    const LONG: &str =
        "WWW the same word runs on far past the width of this card's own field and keeps going";
    let (plain, card) = painted_card(&card_state("", Some(0), 3), 96);
    // The same query with nothing selected, and the same caret at its end: the wash is then the only
    // difference between the two cards.
    let unselected = painted_card(&card_state(LONG, Some(0), 3), 96);
    let all = painted_card(
        &FindBarPaint {
            anchor: 0,
            ..card_state(LONG, Some(0), 3)
        },
        96,
    );

    let washed = |x: i32| {
        (card.query.top..card.query.bottom)
            .any(|y| pixel_at(&all.0, 400, x, y) != pixel_at(&unselected.0, 400, x, y))
    };
    assert!(
        washed(card.query.left + 1),
        "the wash reaches the field's left edge, where the run is cut"
    );
    assert!(
        washed(card.query.right - 3 * text_paint::scaled(FIND_BAR_CARET_MARGIN, 1.0)),
        "and the columns up to the run's own end"
    );
    assert_eq!(
        strays_on_the_query_row(&all.0, &plain, card, 400),
        Vec::new(),
        "and no column of it is the card's padding"
    );
}

/// A count wider than the room a small one is given is drawn whole: the count's own column is laid
/// out from the count's text, and the field beside it is what gives up the room — a count cut at its
/// last digits is a count that is wrong, and `1210/19000` is the count a page of nineteen thousand
/// places really has (see `find_bar_layout` and `count_text`).
#[test]
fn a_count_too_wide_for_its_room_takes_the_room_rather_than_being_cut() {
    let (page_width, page_height) = (400, 300);
    let small = find_bar_layout(page_width, page_height, 0, 96, "1/3", false);
    let long = find_bar_layout(page_width, page_height, 0, 96, "1210/19000", false);

    let surface = DibSurface::create(1, 1).expect("a surface");
    let style = caption_style([30, 30, 30]);
    let needed = measure_text(&surface, &style, "1210/19000", 1.0);
    assert!(
        long.count.right - long.count.left >= needed,
        "the count's column is as wide as the count it holds"
    );
    assert!(
        long.query.left < long.query.right,
        "and the field beside it is still a field"
    );
    assert_eq!(
        small.count.right - small.count.left,
        text_paint::scaled(FIND_BAR_COUNT_WIDTH, 1.0),
        "a count that fits its room moves nothing"
    );

    // The card a page of nineteen thousand places draws: the count is whole on it, rather than the
    // first digits of it in a column laid out for a count of three.
    let surface = DibSurface::create(
        (long.panel.right - long.panel.left) as u32,
        (long.panel.bottom - long.panel.top) as u32,
    )
    .expect("a surface");
    let mut buffer = image_buffer(page_width as u32, page_height as u32);
    paint_find_bar(
        &mut buffer,
        page_width,
        &palette(),
        &long,
        &card_state("needle", Some(1209), 19000),
        &surface,
        96,
    );
    let drawn = (long.count.left..long.count.right)
        .filter(|x| {
            (long.count.top..long.count.bottom)
                .any(|y| pixel_at(&buffer, page_width, *x, y)[3] != 0)
        })
        .count();
    assert!(
        drawn as i32 >= needed,
        "the whole count is on the card, not the first digits of it"
    );
}

/// A switch says what it is while a pointer rests on it: a name hanging below the card, centred under
/// the switch it belongs to, drawn only where the surface the bar is written on keeps the row for it
/// (see `switch_tooltip` and `find_bar_surface_size`).
#[test]
fn a_switch_names_itself_while_the_pointer_rests_on_it() {
    let (page_width, page_height) = (400, 300);
    let card = find_bar_layout(page_width, page_height, 0, 96, "1/3", false);
    let (surface_width, surface_height) = find_bar_surface_size(&card, 96);
    let surface = DibSurface::create(surface_width, surface_height).expect("a surface");

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
    let plain = paint(&card_state("needle", Some(0), 2));
    let named = paint(&FindBarPaint {
        hovered: Some(FindBarPart::Case),
        ..card_state("needle", Some(0), 2)
    });
    let other = paint(&FindBarPaint {
        hovered: Some(FindBarPart::WholeWord),
        ..card_state("needle", Some(0), 2)
    });

    // Below the card, where the page is: nothing while no pointer rests on a switch, and a panel under
    // the switch the pointer is on once one does.
    let under_the_card = card.panel.bottom + 8;
    let named_below = |buffer: &[u8]| {
        (0..page_width).find(|x| pixel_at(buffer, page_width, *x, under_the_card)[3] != 0)
    };
    assert_eq!(
        named_below(&plain),
        None,
        "a card no pointer rests on is a card and nothing below it"
    );
    let case_left = named_below(&named).expect("a switch under the pointer says its name");
    let word_left = named_below(&other).expect("and so does the other one");
    assert!(
        case_left < word_left,
        "each name hangs under its own switch"
    );
    let styled = caption_style(palette().foreground);
    let name_width = measure_text(&surface, &styled, "Match Case", 1.0);
    let panel_width = (0..page_width)
        .filter(|x| pixel_at(&named, page_width, *x, under_the_card)[3] != 0)
        .count();
    assert!(
        panel_width as i32 > name_width,
        "and the panel under it is the name and the room it keeps about it"
    );

    // The card's own controls are not touched by it: a name is a card below the card, not a wash over
    // the row it hangs from.
    assert_ne!(
        pixel_at(&named, page_width, card.case.left + 1, card.case.top + 2),
        pixel_at(&plain, page_width, card.case.left + 1, card.case.top + 2),
        "the switch under the pointer is still drawn as the one it is"
    );
}

/// A page longer than its frame draws its scrollbar in the page's own right margin, and the card keeps
/// clear of it: on such a page the card's margin is measured from the scrollbar's left edge rather than
/// from the frame's, with a gap of its own beyond that so the two do not read as one column — the same
/// rule at every scale and every width a pin is resized to (see `find_bar_layout` and
/// `scrollbar_room`).
#[test]
fn the_card_keeps_clear_of_the_pages_scrollbar_at_every_scale_and_width() {
    for dpi in [96, 120, 144, 192] {
        let scale = dpi as f32 / 96.0;
        let scrollbar_left = |width: i32| {
            width
                - text_paint::scaled(text_paint::PADDING_PIXELS as i32, scale)
                - text_paint::scaled(text_paint::SCROLLBAR_WIDTH_PIXELS, scale).max(1)
        };

        for width in [420, 640, 1200, 2400] {
            let scrolling = find_bar_layout(width, 600, 0, dpi, "1/3", true);
            let plain = find_bar_layout(width, 600, 0, dpi, "1/3", false);

            assert!(
                scrolling.panel.right <= scrollbar_left(width),
                "{dpi} dpi in {width} px: the card ends left of the scrollbar's column"
            );
            assert_eq!(
                scrollbar_left(width) - scrolling.panel.right,
                text_paint::scaled(FIND_BAR_MARGIN as i32, scale)
                    + text_paint::scaled(FIND_BAR_SCROLLBAR_GAP as i32, scale),
                "{dpi} dpi in {width} px: and stands off it by its own margin and the gap"
            );
            assert!(
                scrolling.panel.right < plain.panel.right,
                "{dpi} dpi in {width} px: which moves it left of where a page that does not scroll puts it"
            );
            assert_eq!(
                (scrolling.panel.top, scrolling.query.bottom),
                (plain.panel.top, plain.query.bottom),
                "{dpi} dpi in {width} px: the card moves along the row, not down it"
            );
        }
    }
}

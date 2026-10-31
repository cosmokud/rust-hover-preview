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

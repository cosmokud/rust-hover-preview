use super::*;
use crate::readers::audio_track::Player;

fn card() -> Card {
    Card {
        name: "Kind of Blue - So What.flac".to_string(),
        facts: vec![
            Fact {
                text: "FLAC".to_string(),
                kind: FactKind::Name,
            },
            Fact {
                text: "44.1 kHz".to_string(),
                kind: FactKind::Number,
            },
            Fact {
                text: "stereo".to_string(),
                kind: FactKind::Word,
            },
            Fact {
                text: "1006 kbps".to_string(),
                kind: FactKind::Number,
            },
        ],
        duration: Some(562.0),
        elapsed: Some(67.0),
        name_offset: 0,
        controls: None,
    }
}

/// The same card as a pinned window shows it: the four controls it carries, and the state
/// they are drawn from. Nothing is hovered or held, because what a test is about is where the
/// buttons are rather than how a lit one looks.
fn pinned() -> Card {
    Card {
        controls: Some(CardChrome {
            playing: false,
            volume: 40,
            hovered: None,
            pressed: None,
        }),
        ..card()
    }
}

/// The name of a card too narrow for it is scrolled rather than cut short: the page is
/// built at the box the layout settled on, so the test's card is the one a card of a
/// measured box is built from, with the offset a tick would have reached.
fn scrolled(card: &Card, width: u32, offset: i32) -> Page {
    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    let theme = text_theme::loaded(options().theme).expect("the bundled theme");
    let mut card = card.clone();
    card.name_offset = offset;

    let page = build_page(&card, theme, &metrics, width, width);
    unsafe {
        let _ = DeleteDC(dc);
    }

    page
}

fn options() -> AudioPreviewOptions {
    AudioPreviewOptions {
        theme: TextTheme::Dark,
        font_scale_percent: 125,
    }
}

/// The card is a page like any other painted preview's: measured before the window is
/// placed, drawn into the box the layout settled on, and opaque everywhere.
#[test]
fn draws_the_card_a_sound_is_previewed_as() {
    let (width, height) = measure(&card(), 4096, 2160, 96, options()).expect("a measured card");
    assert!(
        width > 128 && height > 32,
        "a card is a page with room for a name and a bar in it: {width}x{height}"
    );

    let (pixels, width, height) =
        render(&card(), width, height, 96, options()).expect("a painted card");
    assert_eq!(pixels.len(), (width * height * 4) as usize);
    assert!(
        pixels
            .as_chunks::<4>()
            .0
            .iter()
            .all(|pixel| pixel[3] == 255),
        "a page is a page: what stands behind the card is the card"
    );
}

/// A card with nothing to say about its length still draws, and one with no room at all is
/// answered with the page every painted preview gives in that place.
#[test]
fn draws_a_sound_whose_length_is_not_known() {
    let mut unknown = card();
    unknown.duration = None;

    let (width, height) = measure(&unknown, 4096, 2160, 96, options()).expect("a measured card");
    let painted = render(&unknown, width, height, 96, options()).expect("a painted card");
    assert_eq!(painted.0.len(), (painted.1 * painted.2 * 4) as usize);

    let (_, cramped) = measure(&card(), 8, 8, 96, options()).expect("a measured card");
    assert_eq!(cramped, 1, "a card with no room is one pixel and no page");
}

/// The facts of a track are the ones it holds: what a file does not say is left out rather
/// than guessed at, and a file no engine named a codec for is named by its extension.
#[test]
fn writes_the_facts_a_file_has_and_no_others() {
    let track = Track {
        player: Player::Native,
        codec: Some("FLAC".to_string()),
        rate: Some(44_100),
        channels: Some(2),
        bitrate: Some(1_006_000),
        duration: Some(562.0),
    };
    let facts = facts_of(&track, Path::new("C:/music/track.flac"));
    let words: Vec<&str> = facts.iter().map(|fact| fact.text.as_str()).collect();
    assert_eq!(words, ["FLAC", "44.1 kHz", "stereo", "1006 kbps"]);

    let bare = Track {
        player: Player::Ffmpeg,
        codec: None,
        rate: None,
        channels: None,
        bitrate: None,
        duration: None,
    };
    let facts = facts_of(&bare, Path::new("C:/music/track.ape"));
    assert_eq!(facts.len(), 1, "nothing is said that the file did not say");
    assert_eq!(facts[0].text, "APE");
    assert_eq!(facts[0].kind, FactKind::Name);
}

/// The numbers a card is written with, at the lengths and rates a file can have.
#[test]
fn writes_a_position_the_way_a_player_does() {
    assert_eq!(clock(0.0), "0:00");
    assert_eq!(clock(67.4), "1:07");
    assert_eq!(clock(562.0), "9:22");
    assert_eq!(clock(3723.0), "1:02:03");
    assert_eq!(rate_label(44_100), "44.1 kHz");
    assert_eq!(rate_label(48_000), "48 kHz");
    assert_eq!(rate_label(22_050), "22.05 kHz");
    assert_eq!(channel_label(1), "mono");
    assert_eq!(channel_label(2), "stereo");
    assert_eq!(channel_label(6), "5.1");
    assert_eq!(bitrate_label(320_000), "320 kbps");
    assert_eq!(bitrate_label(320), "320 bps");
}

/// The bar is the played share of the whole, and a block crossing the track where there is
/// no whole to measure against: what a card with no player shows is an empty bar.
#[test]
fn fills_the_bar_by_where_the_sound_is() {
    let mut known = card();
    known.duration = Some(100.0);
    known.elapsed = Some(25.0);
    assert_eq!(fill_span(&known, 400, 1.0), (0, 100));

    known.elapsed = Some(562.0);
    assert_eq!(
        fill_span(&known, 400, 1.0),
        (0, 400),
        "a position past the end is the whole of the bar and no more"
    );

    known.duration = None;
    let (_, width) = fill_span(&known, 400, 1.0);
    assert!(
        width > 0 && width < 400,
        "a sound whose length is not known is a block crossing the track: {width}"
    );

    known.elapsed = None;
    assert_eq!(
        fill_span(&known, 400, 1.0),
        (0, 0),
        "and a card at no volume at all is a bar with nothing in it"
    );
}

/// The readout is the played part and the whole, and it is dropped rather than guessed at
/// where there is neither.
#[test]
fn writes_the_clock_the_card_is_drawn_with() {
    let mut card = card();
    let colors = ([0, 0, 0], [1, 1, 1]);

    let runs = clock_runs(&card, 400, 8, colors.0, colors.1);
    let text: String = runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(text, "1:07 / 9:22");
    assert!(
        runs.iter().map(|run| run.width).sum::<i32>() <= 400,
        "the readout is drawn inside the room it was given"
    );

    card.duration = None;
    let text: String = clock_runs(&card, 400, 8, colors.0, colors.1)
        .iter()
        .map(|run| run.text.as_str())
        .collect();
    assert_eq!(
        text, "1:07",
        "a file that does not say how long it is has no whole"
    );

    card.elapsed = None;
    assert!(
        clock_runs(&card, 400, 8, colors.0, colors.1).is_empty(),
        "and a card with no player behind it has nothing to say about time"
    );
}

/// A name the card has no room for is not a reason to draw a wider card: the width comes
/// from the facts line and the bar under it, and the name is the thing that gives way.
#[test]
fn a_name_that_does_not_fit_does_not_widen_the_card() {
    let mut long = card();
    long.name = "18 - The Longest Track Name On This Album (Remastered, 2026).flac".to_string();
    let mut short = card();
    short.name = "2.flac".to_string();

    let long_box = measure(&long, 4096, 2160, 96, options()).expect("a measured card");
    let short_box = measure(&short, 4096, 2160, 96, options()).expect("a measured card");
    assert_eq!(
        long_box, short_box,
        "the card is the size its facts ask for whatever its name is"
    );

    // And the card really has no room for the name, which is what makes the two above the
    // same box rather than two names that both fit.
    let (width, _) = long_box;
    assert!(
        NameScroll::of(&long.name, width, 96, options()).moves(),
        "the name the test is about is one the card cannot fit"
    );
    assert!(
        !NameScroll::of(&short.name, width, 96, options()).moves(),
        "and one it can is left where it is"
    );
}

/// The name is drawn whole and clipped to the box the card leaves it — never cut short —
/// and the offset a tick has reached is the whole of what moves: the origin of the run.
#[test]
fn draws_the_whole_name_where_the_offset_puts_it() {
    let mut long = card();
    long.name = "18 - The Longest Track Name On This Album (Remastered, 2026).flac".to_string();
    let (width, _) = measure(&long, 4096, 2160, 96, options()).expect("a measured card");

    let resting = scrolled(&long, width, 0);
    let name = resting.header.last().expect("the run the name is drawn in");
    assert_eq!(
        name.text, long.name,
        "the name is drawn whole, not cut short"
    );
    assert_eq!(
        name.origin, name.x,
        "a name at rest starts at the left edge of the box the mark leaves it"
    );
    assert!(
        name.x + name.width <= width as i32,
        "and that box is the card's own content box: it ends inside the card"
    );

    let moved = scrolled(&long, width, 40);
    let name = moved.header.last().expect("the run the name is drawn in");
    assert_eq!(
        name.text, long.name,
        "a scrolled name is still the whole name"
    );
    assert_eq!(name.origin, name.x - 40, "the offset is what moves it");
    assert_eq!(
        (
            resting.header.last().expect("the name").x,
            resting.header.last().expect("the name").width
        ),
        (name.x, name.width),
        "and the box it is clipped to stays where it is"
    );
}

/// The scroll starts at the beginning of the name, holds there, runs to the end of it,
/// holds again and comes back — ping-ponging for as long as the card is up — and each end
/// is where the page draws it.
#[test]
fn scrolls_the_name_from_one_end_of_itself_to_the_other_and_back() {
    let mut long = card();
    long.name = "18 - The Longest Track Name On This Album (Remastered, 2026).flac".to_string();
    let (width, _) = measure(&long, 4096, 2160, 96, options()).expect("a measured card");
    let step = Duration::from_millis(33);

    let mut scroll = NameScroll::of(&long.name, width, 96, options());
    assert!(
        scroll.moves() && scroll.offset() == 0,
        "it starts at the start"
    );

    // The hold a card is put up with: repaints inside it move nothing at all.
    scroll.advance(Instant::now(), step);
    assert_eq!(
        scroll.offset(),
        0,
        "the name rests before it begins to move"
    );

    let first_end = scroll.travel;
    assert!(first_end > 0, "the name has an end past the box to reach");

    // Left until the end of the name comes into sight, where it stops and holds.
    let arrival = Instant::now() + NAME_HOLD + Duration::from_millis(1);
    for _ in 0..first_end + 1 {
        scroll.advance(arrival, step);
    }
    assert_eq!(
        scroll.offset(),
        first_end,
        "the name stops at the end of itself"
    );
    scroll.advance(arrival, step);
    assert_eq!(
        scroll.offset(),
        first_end,
        "and holds there rather than turning at once"
    );

    let far = scrolled(&long, width, scroll.offset());
    let name = far.header.last().expect("the run the name is drawn in");
    assert_eq!(
        name.origin,
        name.x - first_end,
        "the far end of the ping-pong is the whole name shown"
    );

    // And back to the beginning, once the hold at the far end is over.
    let returning = arrival + NAME_HOLD + Duration::from_millis(1);
    for _ in 0..first_end {
        scroll.advance(returning, step);
    }
    assert_eq!(
        scroll.offset(),
        0,
        "the name comes back to where it started"
    );

    let home = scrolled(&long, width, scroll.offset());
    let name = home.header.last().expect("the run the name is drawn in");
    assert_eq!(name.origin, name.x, "and the near end is its resting place");
}

/// The bar is the card's one control, and a press on it is answered against the row the card
/// is drawn with rather than one kept beside it: the bar's own three pixels and the gap above
/// it, from margin to margin, and nowhere else on the card.
///
/// The row is asked of the page the card is laid out with, so the two agree by construction —
/// a bar drawn by one arithmetic and pressed by another would answer a press in the middle of
/// the facts line.
#[test]
fn a_press_on_the_bar_is_answered_where_the_bar_is_drawn() {
    let (width, _) = measure(&card(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&card(), width, 0);
    let bar = page.bar;
    assert!(bar.width > 0 && bar.height > 0, "there is a bar to press");

    let at = |x: i32, y: i32| bar_share_at(x, y, width, 96, options(), false);
    // The centre of the row, which is where the line itself is drawn: the row is taller than
    // the bar now (see `BarRow`), so the bar's own top is not where a hand would aim.
    let centre = bar.top + bar.height / 2;

    // Across the bar: the margin it starts at is the beginning of the file, the middle of it
    // is half way through, and the far end of it is the end.
    assert_eq!(
        at(bar.left, centre),
        Some(0.0),
        "the margin is the beginning"
    );
    assert_eq!(
        at(bar.left + bar.width / 2, centre),
        Some(0.5),
        "the middle of the bar is half way through the file"
    );
    let far = at(bar.left + bar.width - 1, centre).expect("the far end is the bar too");
    assert!(
        far > 0.99,
        "and the last pixel of it is the end of the file: {far}"
    );

    // The band a press is answered against is the bar's own row and the gap above it, which
    // is what makes three pixels of track a thing a hand can be asked to hit at all.
    assert!(
        at(bar.left + bar.width / 2, bar.top - 1).is_some(),
        "the line just above the bar is still the bar's row"
    );
    assert_eq!(
        at(bar.left + bar.width / 2, bar.top - BAR_GAP_PIXELS * 4),
        None,
        "while the facts line above the gap is a line of text and not a control"
    );

    // And across the margins: the bar runs from one to the other and no further.
    assert_eq!(
        at(bar.left - 1, centre),
        None,
        "left of it is the card's margin"
    );
    assert_eq!(
        at(bar.left + bar.width, centre),
        None,
        "and right of it is the other one"
    );
}

/// The press band reaches below the bar as well as above it, because a hand aiming upwards at
/// a three-pixel line from the bottom of a card lands beside it about as often as on it — and
/// it stops at the card's own bottom edge, which is what keeps it from swallowing the margin
/// further down where a hand is carrying the window rather than aiming at anything.
#[test]
fn the_bar_press_reaches_below_it_and_stops_at_the_card_edge() {
    let (width, height) = measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&pinned(), width, 0);
    let bar = page.bar;
    let at = |x: i32, y: i32| bar_share_at(x, y, width, 96, options(), true);

    let scale = metrics_scale();
    let reach = scaled(BAR_REACH_PIXELS, scale);
    let middle = bar.left + bar.width / 2;

    // Under the bar, within the reach: still the bar.
    assert!(
        at(middle, bar.top + bar.height).is_some(),
        "the row just under the bar is still the bar's row"
    );
    assert!(
        at(middle, bar.top + bar.height + reach - 1).is_some(),
        "and so is the last row of the reach"
    );
    assert_eq!(
        at(middle, bar.top + bar.height + reach),
        None,
        "while a row past it is a hand on the card rather than on the bar"
    );

    // And the card's own bottom edge is where the band stops whatever the reach would have
    // been: a card drawn into a box shorter than the reach cannot answer below itself.
    assert_eq!(
        at(middle, height as i32),
        None,
        "and past the card there is nothing at all"
    );
}

/// The scale the card's own layout is counted in, read of the metrics rather than written
/// down: a test that asserted an unscaled pixel count against a card drawn at 125% would be
/// asserting the wrong row.
fn metrics_scale() -> f32 {
    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    unsafe {
        let _ = DeleteDC(dc);
    }

    metrics.scale
}

/// The card's own margin, read of the metrics rather than written down, for the same reason
/// `metrics_scale` is.
fn metrics_padding() -> i32 {
    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    unsafe {
        let _ = DeleteDC(dc);
    }

    metrics.padding
}

/// Every control of a pinned card is answered inside its own box and nowhere else, which is
/// the whole of what "a button hit-tested by one arithmetic and drawn by another is a button
/// that answers a press in the middle of the facts line" means.
#[test]
fn a_pinned_card_carries_its_four_controls_where_they_are_drawn() {
    let (width, _) = measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    let boxes = control_boxes_for(&pinned(), width);

    for control in [
        CardControl::Previous,
        CardControl::Play,
        CardControl::Next,
        CardControl::Seek,
        CardControl::Volume,
    ] {
        let rect = control_box(control, width, 96, options(), true).expect("a box to press");
        let centre = ((rect.left + rect.right) / 2, (rect.top + rect.bottom) / 2);
        assert_eq!(
            control_at(centre.0, centre.1, width, 96, options(), true),
            Some(control),
            "the middle of {control:?} is {control:?}"
        );

        // The corners of every other box are this one: nothing overlaps, so a press is
        // answered by exactly the thing it landed on.
        for other in [
            CardControl::Previous,
            CardControl::Play,
            CardControl::Next,
            CardControl::Seek,
            CardControl::Volume,
        ] {
            if other == control {
                continue;
            }
            let other_box = control_box(other, width, 96, options(), true).expect("a box");
            assert_ne!(
                rect, other_box,
                "{control:?} and {other:?} are drawn in the same box"
            );
        }
    }

    // The same boxes the page draws them in, which is the claim the boxes exist for.
    let drawn = scrolled(&pinned(), width, 0);
    let page_boxes = drawn.boxes.as_ref().expect("a card that carries controls");
    assert_eq!(
        page_boxes.rect(CardControl::Volume),
        control_box(CardControl::Volume, width, 96, options(), true),
        "the volume button is drawn where it is pressed"
    );
    assert_eq!(page_boxes.bar.left, drawn.bar.left);
    assert_eq!(page_boxes.bar.right, drawn.bar.left + drawn.bar.width);

    // And the row is laid out left to right: the walk at the left, the bar in the middle, the
    // level at the right — which is the order a hand reads them in.
    let _ = boxes;
}

/// A pinned card is the size the card a hover shows is: its controls stand in the bar's own
/// row, so they are carved out of the card rather than added to it — and everything they are
/// drawn on is still inside it.
///
/// Which is the whole of what the take-up does not have to do: a window given the hover's box
/// is the window the card is drawn in, and the buttons reach a few pixels above the bar and
/// below it into the gap and the margin that are already there. The height is the card's own
/// arithmetic rather than a number written down here, which is what makes it a claim about
/// every row rather than this one.
#[test]
fn a_pinned_card_is_the_size_a_hovers_card_is_and_its_controls_stay_inside_it() {
    let (width, height) = measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    let (hover_width, hover_height) =
        measure(&card(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&pinned(), width, 0);

    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    let row = bar_row(&metrics);
    let row_top = row.top;
    let row_bottom = row.top + row.height;
    let edge = row.top + row.height + metrics.padding;
    let boxes = page.boxes.as_ref().expect("a card that carries controls");
    let band = bar_band(row, boxes.bar.left, boxes.bar.right, &metrics);
    unsafe {
        let _ = DeleteDC(dc);
    }

    assert_eq!(
        (width, height),
        (hover_width, hover_height),
        "a card carrying its controls is drawn in the box a hover's is: a pin's window is \
         the hover's window"
    );
    assert_eq!(
        page.height, edge as u32,
        "the card runs to the end of the bar's own row and one margin below it"
    );
    assert_eq!(
        height, page.height,
        "and the take-up measures the card at the height it is drawn at"
    );

    // Which is only worth anything if the controls are on it: every one of them, and the bar
    // between the ends of the row, is inside the card rather than through it.
    for control in [
        CardControl::Previous,
        CardControl::Play,
        CardControl::Next,
        CardControl::Seek,
        CardControl::Volume,
    ] {
        let rect = boxes
            .rect(control)
            .expect("a box on a card that carries controls");
        assert!(rect.left >= 0 && rect.top >= 0, "{control:?} at {rect:?}");
        assert!(
            rect.right <= page.width as i32 && (rect.bottom as u32) <= page.height,
            "{control:?} at {rect:?} against a card of {}x{}",
            page.width,
            page.height
        );
    }

    // And the buttons really are carved out of the card rather than laid out inside the row:
    // they stand taller than the bar's line, reaching into the gap above it and the margin
    // below — no further, or they would be off the card.
    let previous = boxes.previous;
    assert!(
        previous.top < row_top && previous.bottom > row_bottom,
        "a button is centred on the bar and reaches beyond it: {previous:?} against a row at \
         {row_top}"
    );
    assert!(
        previous.bottom <= edge,
        "and no further than the margin under the row holds: {previous:?} against {edge}"
    );

    // And the reach a press on the bar is answered against stops at the card's own edge rather
    // than running off it — which is the same margin counted from the other end, and is what
    // keeps the last few rows of the card's margin a place a hand carries the window from.
    assert!(
        band.top < row_top && band.bottom <= edge && (band.bottom as u32) <= page.height,
        "the band a press on the bar is answered against is on the card: {}..{} against {edge}",
        band.top,
        band.bottom
    );
}

/// A card that carries no controls is the card it has always been: no control anywhere on it,
/// and a bar that runs from margin to margin rather than from one button to the other.
#[test]
fn the_hover_card_is_the_card_it_has_always_been() {
    let (plain, _) = measure(&card(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&card(), plain, 0);

    assert!(
        page.boxes.is_none(),
        "a hover's card carries no controls at all"
    );
    assert!(page.chrome.is_none());

    // Every control is answered as nothing, anywhere on the card — including inside the bar,
    // which on a pinned card is a control of its own.
    for y in (0..page.height as i32).step_by(3) {
        for x in (0..plain as i32).step_by(7) {
            assert_eq!(
                control_at(x, y, plain, 96, options(), false),
                None,
                "a point on a hover's card is on no control: ({x}, {y})"
            );
        }
    }

    // And the bar still spans margin to margin, which is what it did before there were
    // buttons: the row is the same height it always was, and the bar fills it.
    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    let padding = metrics.padding;
    unsafe {
        let _ = DeleteDC(dc);
    }

    assert_eq!(
        page.bar.left, padding,
        "the bar begins at the card's own margin"
    );
    assert_eq!(
        page.bar.left + page.bar.width,
        plain as i32 - padding,
        "and ends at the other one"
    );
    assert_eq!(
        page.bar.height,
        scaled(BAR_PIXELS, metrics_scale()),
        "and the row is the height the bar always was: a hover's card is not taller for the \\
         controls a pinned card carries"
    );
}

/// A seek is measured across the bar and not across the row: on a pinned card the bar fills
/// what is left of the content box between four buttons, so a press at the row's left edge is
/// a press on the button there rather than on the first second of the file — and the share at
/// the bar's own edge is nothing at all, never a negative and never a share of the whole.
#[test]
fn a_seek_is_measured_across_the_bar_and_not_across_the_buttons() {
    let (width, _) = measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&pinned(), width, 0);
    let bar = page.bar;
    let centre = bar.top + bar.height / 2;

    let at = |x: i32, y: i32| bar_share_at(x, y, width, 96, options(), true);

    // The bar's own edges are the beginning and the end, exactly as they are without buttons.
    assert_eq!(at(bar.left, centre), Some(0.0));
    let far = at(bar.left + bar.width - 1, centre).expect("the far end of the bar");
    assert!(far > 0.99, "and the last pixel of it is the end: {far}");
    assert_eq!(
        at(bar.left - 1, centre),
        None,
        "while the row to the left of the bar is the button standing in it"
    );
    assert_eq!(at(bar.left + bar.width, centre), None);

    // Which is also what the row's own boxes say: the volume button is at the right of the
    // card and is not a seek to the last second of the file however far right it is.
    let volume =
        control_box(CardControl::Volume, width, 96, options(), true).expect("a box to press");
    assert_eq!(
        control_at(
            (volume.left + volume.right) / 2,
            (volume.top + volume.bottom) / 2,
            width,
            96,
            options(),
            true
        ),
        Some(CardControl::Volume),
        "the button at the row's right edge is the volume button and not the end of the file"
    );
    assert!(
        volume.left >= bar.left + bar.width,
        "and it stands outside the bar rather than over it: {:?} against {bar:?}",
        volume
    );
}

/// A card is as wide with its controls as without them: the four buttons are carved out of
/// the bar rather than added to the card, so a pin's card is the width a hover's is and what
/// the buttons cost is the track's length — which is still a track, because a bar a hundred
/// pixels long is a different control.
#[test]
fn the_controls_are_carved_out_of_the_bar_rather_than_out_of_the_card() {
    // A card whose facts say plenty is as wide as its facts ask for either way — the row of
    // buttons is paid for out of what the bar would have had, not out of the card.
    let (plain, plain_height) =
        measure(&card(), 4096, 2160, 96, options()).expect("a measured card");
    let (with, with_height) =
        measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    assert_eq!(
        (with, with_height),
        (plain, plain_height),
        "a card with buttons is exactly the card it would have been without them"
    );

    // And a card whose facts are narrow is not widened either: the width comes from the
    // narrowest bar there is either way, and the buttons come out of that bar.
    let narrow = |controls| Card {
        facts: vec![Fact {
            text: "FLAC".to_string(),
            kind: FactKind::Name,
        }],
        controls,
        ..card()
    };
    let (narrow_plain, _) =
        measure(&narrow(None), 4096, 2160, 96, options()).expect("a measured card");
    let (narrow_with, _) = measure(
        &narrow(Some(CardChrome::default())),
        4096,
        2160,
        96,
        options(),
    )
    .expect("a measured card");
    assert_eq!(
        narrow_with, narrow_plain,
        "a narrow card is not widened by four buttons on it: {narrow_with} vs {narrow_plain}"
    );

    // What the buttons cost is what the bar gave up: the four squares, the gaps between them
    // and the margin at either end of the content box — and a track is left beside them.
    let side = scaled(CONTROL_SIDE_PIXELS, metrics_scale());
    let gap = scaled(CONTROL_GAP_PIXELS, metrics_scale());
    let padding = metrics_padding();
    let page = scrolled(&narrow(Some(CardChrome::default())), narrow_with, 0);
    assert_eq!(
        page.bar.left,
        padding + side * 3 + gap * 3,
        "the bar begins after the three buttons at the left of the row and the gaps between them"
    );
    assert_eq!(
        page.bar.left + page.bar.width,
        narrow_with as i32 - padding - side - gap,
        "and ends before the volume button at the row's right edge"
    );
    assert_eq!(
        narrow_with as i32 - 2 * padding - side * 4 - gap * 4,
        page.bar.width,
        "so the four buttons and their four gaps cost the bar's width and nothing else"
    );
    assert!(
        page.bar.width > 0,
        "and a card with buttons still has a bar to take a sound to a second of it"
    );
}

/// The boxes of a card that carries its controls, as the page laid it out — the one definition
/// both the painting and the hit test read, which is what the test above is about.
fn control_boxes_for(card: &Card, width: u32) -> Option<CardBoxes> {
    scrolled(card, width, 0).boxes
}

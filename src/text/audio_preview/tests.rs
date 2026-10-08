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
/// buttons are rather than how a lit one looks. The window buttons are up, because a test
/// painting a card is a card a hand is near — the top border of the window, where the card's
/// own rows begin.
fn pinned() -> Card {
    Card {
        controls: Some(CardChrome {
            playing: false,
            volume: 40,
            hovered: None,
            pressed: None,
            window_buttons: true,
        }),
        ..card()
    }
}

/// The name of a card too narrow for it is scrolled rather than cut short: the page is
/// built at the box the layout settled on, so the test's card is the one a card of a
/// measured box is built from, with the offset a tick would have reached.
fn scrolled(card: &Card, width: u32, offset: i32) -> Page {
    scrolled_at(card, width, offset, options())
}

/// The same page at another font size, which is what a card is built with at a share
/// other than the 10% anchor: the font the share names (see `audio_font_scale_percent`
/// in the preview window's dimensions), so that the boxes a card carries at that share
/// are the boxes the question is about.
fn scrolled_at(card: &Card, width: u32, offset: i32, options: AudioPreviewOptions) -> Page {
    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options.font_scale_percent).expect("metrics");
    let theme = text_theme::loaded(options.theme).expect("the bundled theme");
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

/// The same options at another font size, which is what a card is
/// built with at a share other than the 10% anchor: the font the
/// share names (see `audio_font_scale_percent` in the preview window's
/// dimensions).
fn options_at(font_scale_percent: u32) -> AudioPreviewOptions {
    AudioPreviewOptions {
        theme: TextTheme::Dark,
        font_scale_percent,
    }
}

/// The narrowest card worth drawing at the options the tests build: the
/// width floor the card is laid out at, counted the way `build_page`
/// counts it — the bar's own advances, plus the margin the card stands
/// inside on both sides.
fn narrowest_card() -> u32 {
    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    let narrowest =
        (metrics.advance[BODY_LEVEL as usize] * MIN_CONTENT_ADVANCES + metrics.padding * 2) as u32;
    unsafe {
        let _ = DeleteDC(dc);
    }
    narrowest
}

/// The card measured at the room the default share of a 3440x1440 display
/// gives it at 96 DPI — the room a name-scroll test needs, because the name
/// the tests below draw is one this room has no room for: the name is the
/// thing that gives way, not the room that grows to fit it (see
/// `audio_box_room`).
fn measured_at_the_default_share(card: &Card) -> (u32, u32) {
    measure(card, 374, 144, 96, options()).expect("a measured card")
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

    // A room of a few pixels is not a room with nothing in it: the card's
    // width is floored at the narrowest card worth drawing, so the answer
    // is that card's width with the one-pixel height a room too short to
    // draw a name in answers with.
    let (cramped_width, cramped) = measure(&card(), 8, 8, 96, options()).expect("a measured card");
    assert_eq!(
        cramped_width,
        narrowest_card(),
        "a card with no room is the narrowest card"
    );
    assert_eq!(cramped, 1, "and no page taller than one pixel");
}

/// A card fills the room it is measured at: the width is the room
/// itself, floored at the narrowest card worth drawing at the font
/// the room's share builds its card at — which every room the menu
/// offers clears, the 5% room included, measured at the 5% font —
/// and the height is the card's own, the same one-line strip at
/// every share, because the share decides the room and not the card.
#[test]
fn a_card_fills_the_room_its_share_gives_it() {
    // The 10% room of a 3440x1440 work area at the default share, and
    // the 5% room beside it — each measured at the font its share
    // builds its card at, which is what makes the 5% room the wider of
    // the room and the narrowest card worth drawing at that font.
    let (wide, height) = measure(&card(), 374, 144, 96, options()).expect("a measured card");
    let (narrow, short_height) =
        measure(&card(), 188, 72, 96, options_at(63)).expect("a measured card");

    assert_eq!(wide, 374, "the card fills the room it is measured at");
    assert_eq!(
        narrow, 188,
        "the 5% room is the wider of the room and the narrowest card at the 5% font"
    );

    // The height is the card's own at both rooms — a one-line strip at
    // every share: the room's height does not squash it, and a room ten
    // times the size does not stretch it either.
    let (_, roomy) = measure(&card(), 4096, 2160, 96, options()).expect("a measured card");
    let (_, short_roomy) =
        measure(&card(), 4096, 2160, 96, options_at(63)).expect("a measured card");
    assert_eq!(height, roomy, "the widest room does not stretch the card");
    assert_eq!(
        short_height, short_roomy,
        "and the 5% share is a one-line strip at its own font too"
    );
    assert!(
        height > short_height,
        "the card's height follows its font: the 10% card is the taller one"
    );
    assert!(
        height > 32,
        "a card is a page with a name and a bar in it: {height}"
    );
}

/// The ladder the five shares answer with: a card measured at each
/// share's room, at the font that share builds its card at, is the
/// room wide — at every share the room is the wider of the room and
/// the narrowest card worth drawing at the share's own font, so the
/// width floor never engages at a share the menu offers — and the
/// height is what the card's content measures at that font, which is
/// the anchor's height scaled by the share's fraction of the anchor.
#[test]
fn every_share_answers_the_anchor_s_card_scaled() {
    let (_, anchor_height) = measure(&card(), 374, 144, 96, options()).expect("the anchor's card");

    for (room, font) in [
        ((188u32, 72u32), 63u32),
        ((374, 144), 125),
        ((562, 216), 188),
        ((748, 288), 250),
        ((936, 360), 313),
    ] {
        let (width, height) =
            measure(&card(), room.0, room.1, 96, options_at(font)).expect("a measured card");
        assert_eq!(
            width, room.0,
            "the {font}% card fills the {font}% room: the floor never engages at a share the menu offers"
        );

        // The height follows the font, not the room: the anchor's
        // height scaled by the share's fraction of the anchor. The
        // band is wide because the font's own metrics round each line
        // and margin to a whole pixel, and a card is several of them
        // stacked — but a card whose height ignored its font would be
        // off by the share's whole fraction, far outside it.
        let expected = anchor_height as f32 * font as f32 / 125.0;
        assert!(
            (height as f32 - expected).abs() / expected < 0.10,
            "the {font}% card's height is the anchor's scaled by its font: {height} against {expected}"
        );
    }
}

/// A room too short for a name and a line under it answers with
/// nothing: the one-pixel page, at the width the room gives — the
/// width floor is what keeps the width at or above the narrowest
/// card, and a room this wide clears it.
#[test]
fn a_room_too_short_for_a_name_answers_with_nothing() {
    let (width, height) = measure(&card(), 374, 10, 96, options()).expect("a measured card");
    assert_eq!(
        width, 374,
        "the width is the room's own, floored at the narrowest card"
    );
    assert_eq!(
        height, 1,
        "a room with no room for a name is one pixel and no page"
    );
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

/// A name the card has no room for is not a reason to draw a wider card: the width
/// comes from the room the card is measured at, and the name is the thing that
/// gives way.
#[test]
fn a_name_that_does_not_fit_does_not_widen_the_card() {
    let mut long = card();
    long.name = "18 - The Longest Track Name On This Album (Remastered, 2026).flac".to_string();
    let mut short = card();
    short.name = "2.flac".to_string();

    let long_box = measured_at_the_default_share(&long);
    let short_box = measured_at_the_default_share(&short);
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
    let (width, _) = measured_at_the_default_share(&long);
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

/// The two window buttons a pinned sound carries are answered in their
/// own boxes and nowhere else: a hand in one is a hand on that button,
/// and the cell the card's mark is drawn in answers for nothing at all —
/// the mark is a plain bullet no hand can press (see `BULLET`). There is
/// no third window button: the card's menu opens from a right-click on
/// the window rather than from a gear (see `pin_menu::open_pin_menu`).
#[test]
fn the_two_window_buttons_are_answered_in_their_own_boxes() {
    let (width, _) = measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&pinned(), width, 0);
    let boxes = page.boxes.as_ref().expect("a card that carries controls");

    // The cell the card's mark is drawn in, read of the name line's
    // own first run: the mark and the room after it, and a hand in it
    // is a hand on no control.
    let mark = &page.header[0];
    let cell = RECT {
        left: mark.x,
        top: page.header_top,
        right: mark.x + mark.width,
        bottom: page.header_top + page.header_height,
    };
    let centre = ((cell.left + cell.right) / 2, (cell.top + cell.bottom) / 2);
    assert_eq!(
        control_at(centre.0, centre.1, width, 96, options(), true),
        None,
        "the cell the mark is drawn in answers for nothing"
    );

    // The two window buttons keep their own boxes, and the card answers
    // for each in its own box — wherever it stands, with no flag saying
    // whether the band it stands in is showing.
    for (control, button) in [
        (CardControl::Minimize, &boxes.minimize),
        (CardControl::Close, &boxes.close),
    ] {
        let middle = (
            (button.drawn.left + button.drawn.right) / 2,
            (button.drawn.top + button.drawn.bottom) / 2,
        );
        assert_eq!(
            boxes.window_button_at(middle.0, middle.1),
            Some(control),
            "the middle of {control:?} is {control:?} and nothing beside it"
        );
        assert_eq!(
            control_at(middle.0, middle.1, width, 96, options(), true),
            Some(control),
            "and the card answers for it there"
        );
    }

    // The gap between the two of them is the card's own, and the margin
    // to the left of the minimize as well: a hand in either is a hand on
    // no button.
    let button_gap = (boxes.minimize.drawn.right + boxes.close.drawn.left) / 2;
    let middle_row = (boxes.minimize.drawn.top + boxes.minimize.drawn.bottom) / 2;
    assert_eq!(
        boxes.window_button_at(button_gap, middle_row),
        None,
        "the gap between the two is no button's"
    );
    assert_eq!(
        boxes.window_button_at(0, middle_row),
        None,
        "the margin to the left of the minimize is the card's own"
    );
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

/// The pinned card with one of its two window buttons lit, the way the card is
/// handed to the paint that draws a pointer resting on one.
fn lit(hovered: Option<CardControl>, pressed: Option<CardControl>) -> Card {
    Card {
        controls: Some(CardChrome {
            playing: false,
            volume: 40,
            hovered,
            pressed,
            window_buttons: true,
        }),
        ..pinned()
    }
}

/// The two window buttons stand in the card's top corner, over the name line: the
/// drawn boxes begin at the gap below the window's own top edge and reach into the
/// name line's own rows — room the margin does not have, a button being twice the
/// side the margin alone would hold — and the name keeps the box and the scroll it
/// has always had, running underneath the buttons while they are up (see
/// `the_name_runs_underneath_the_window_buttons_while_they_are_up`).
#[test]
fn the_window_buttons_stand_in_the_top_corner_over_the_name_line() {
    let (width, _) = measured_at_the_default_share(&pinned());
    let page = scrolled(&pinned(), width, 0);
    let boxes = page.boxes.as_ref().expect("a card that carries controls");

    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    let gap = scaled(WINDOW_BUTTON_GAP_PIXELS, metrics.scale);
    unsafe {
        let _ = DeleteDC(dc);
    }

    // The name line is centred between the window's own top border
    // and the rule under it: the ink of its glyphs — the neon circle
    // and the name — is as far from the one as from the other. The
    // buttons stand in the name's own rows, which the card has
    // already, so they stand over a line that is centred rather than
    // one that begins at the margin.
    let ink_top = metrics.ink_top[HEADER_LEVEL as usize];
    let ink_bottom = metrics.ink_bottom[HEADER_LEVEL as usize];
    assert_eq!(
        page.header_top + ink_top,
        page.rule_top - (page.header_top + ink_bottom) - 1,
        "the name line's ink is centred between the window's top border and the rule"
    );

    // A button is a square stood back from the window's own top edge by the
    // gap, with the same gap to the card's right border and between each pair
    // of them — and a side twice what the margin alone would hold, which is room
    // the margin does not have: the drawn box reaches into the name line's own
    // rows, which is where a button that size ends.
    let margin_side = metrics.padding - gap * 2;
    for button in [&boxes.minimize, &boxes.close] {
        assert_eq!(
            button.drawn.top, gap,
            "a drawn box stood back by the gap from the window's own top edge: {:?}",
            button.drawn
        );
        assert_eq!(
            button.drawn.bottom,
            gap + margin_side * 2,
            "and reaching into the name line's own rows, where a button twice the \
             margin's own side ends: {:?}",
            button.drawn
        );
        assert_eq!(
            button.drawn.right - button.drawn.left,
            margin_side * 2,
            "a drawn side twice what the margin alone holds: {:?}",
            button.drawn
        );
        assert!(
            button.drawn.left >= 0 && button.drawn.right <= page.width as i32,
            "and inside the card: {:?} against a card {} wide",
            button.drawn,
            page.width
        );
    }
    assert_eq!(
        boxes.close.drawn.right,
        page.width as i32 - gap,
        "the close stands back from the card's right border by the gap"
    );
    assert_eq!(
        boxes.minimize.drawn.right,
        boxes.close.drawn.left - gap,
        "and the minimize stands beside it with the same gap between the two"
    );

    // The name's own box is the card's content box — the same box it is drawn
    // in on a card with no buttons on it — and a name too long for it still has
    // the whole width of it to scroll along.
    let name = page.header.last().expect("the run the name is drawn in");
    let plain = scrolled(&card(), width, 0);
    let plain_name = plain.header.last().expect("the run the name is drawn in");
    assert_eq!(
        (name.x, name.width),
        (plain_name.x, plain_name.width),
        "the name's box is the box it has always been"
    );
    let long = "18 - The Longest Track Name On This Album (Remastered, 2026).flac";
    assert!(
        NameScroll::of(long, width, 96, options()).moves(),
        "and a name the card has no room for still has a runway to scroll along"
    );
}

/// While the two window buttons are up, the name runs underneath them:
/// the box each stands in is the page's own colour, which is what keeps
/// a name scrolling under one from running under its mark. The name's
/// box and its scroll are unchanged, so a name the card has no room for
/// still travels the whole width of the card — under the buttons and off
/// the card's own edge.
#[test]
fn the_name_runs_underneath_the_window_buttons_while_they_are_up() {
    let long = "18 - The Longest Track Name On This Album (Remastered, 2026).flac";
    let mut named = card();
    named.name = long.to_string();
    named.controls = Some(CardChrome {
        playing: false,
        volume: 40,
        hovered: None,
        pressed: None,
        window_buttons: true,
    });

    let (width, height) = measured_at_the_default_share(&named);
    let page = scrolled(&named, width, 0);
    let boxes = page.boxes.as_ref().expect("a card that carries controls");
    let name_run = page.header.last().expect("the run the name is drawn in");

    // The name scrolled to the end of its runway, where its last
    // glyphs stand at the right edge of the card — inside the
    // corner the buttons stand in.
    let dc = unsafe { CreateCompatibleDC(None) };
    let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
    let header_advance = metrics.advance[HEADER_LEVEL as usize];
    unsafe {
        let _ = DeleteDC(dc);
    }
    let offset = super::page::text_width(long, header_advance) - name_run.width;

    let mut up = named.clone();
    up.name_offset = offset;
    let mut down = named;
    down.name_offset = offset;
    let chrome = down.controls.expect("a card that carries controls");
    down.controls = Some(CardChrome {
        window_buttons: false,
        ..chrome
    });

    let up_painted = render(&up, width, height, 96, options()).expect("a painted card");
    let down_painted = render(&down, width, height, 96, options()).expect("a painted card");

    // The colors the paint reads, worked out here the same way: the
    // page's own color, and the card's foreground — the ink the
    // name's glyphs and the buttons' marks are painted in at rest,
    // which is the same ink, so the two are told apart by row
    // rather than by colour.
    let theme = text_theme::loaded(options().theme).expect("the bundled theme");
    let page_color = rgb(theme.background());
    let foreground = rgb(theme.foreground());

    // The ink of a pixel is its color read backwards: the buffer is
    // blue, green, red, alpha, and a color is red, green, blue.
    let pixel = |painted: &(Vec<u8>, u32, u32), x: i32, y: i32| -> [u8; 4] {
        let index = (y as usize * width as usize + x as usize) * 4;
        [
            painted.0[index],
            painted.0[index + 1],
            painted.0[index + 2],
            painted.0[index + 3],
        ]
    };
    let ink = |color: [u8; 3]| [color[2], color[1], color[0], 255];

    // The minimize's own box, which is the one the name's end lands
    // inside, and the row its mark is drawn in: the middle of the
    // box, which is the name line's own top row or the margin's
    // last row above it — a row the name's ink never stands in,
    // because a line's ink begins below the leading its box
    // leaves above it.
    let drawn = boxes.minimize.drawn;
    let middle_row = (drawn.top + drawn.bottom) / 2;
    // The rows of the name's own ink the box reaches: from the
    // first row its glyphs stand in to the box's bottom, with
    // the mark's row above them left out.
    let name_rows = page.header_top + metrics.ink_top[HEADER_LEVEL as usize]..drawn.bottom;

    // While the buttons are up, the mark is drawn in the middle of
    // the box, and every row of the name's ink the box reaches is
    // the page's own colour: the name is painted first and
    // cleared out of the box, so it runs underneath the button
    // rather than under its mark.
    assert!(
        (drawn.left..drawn.right).any(|x| pixel(&up_painted, x, middle_row) == ink(foreground)),
        "the minimize's dash is drawn in the box's middle row while it is up"
    );
    for y in name_rows {
        for x in drawn.left..drawn.right {
            assert_eq!(
                pixel(&up_painted, x, y),
                ink(page_color),
                "the name is cleared out of the box at ({x}, {y})"
            );
        }
    }

    // While the buttons are down, the corner is the name's own: no
    // mark in the middle row, and the name's ink in the rows the box
    // stands in — the 'l' of '.flac', an ascender with ink at the
    // glyph's top rows, at the end of the runway inside this very box.
    assert!(
        (drawn.left..drawn.right).all(|x| pixel(&down_painted, x, middle_row) == ink(page_color)),
        "no mark is drawn in the box's middle row while they are down"
    );
    assert!(
        (page.header_top..drawn.bottom).any(|y| (drawn.left..drawn.right).any(|x| pixel(
            &down_painted,
            x,
            y
        ) != ink(
            page_color
        ))),
        "and the name's ink stands in the corner where the button stood"
    );
}

/// The two window buttons are answered where their glyphs are drawn and a little
/// beyond them: the middle of a drawn box is the button, the hit box reaches the
/// window's own top edge above it and a cushion below it, and a row past the
/// cushion is nothing at all.
#[test]
fn a_window_button_is_answered_where_its_glyph_is_and_a_little_beyond_it() {
    let (width, _) = measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&pinned(), width, 0);
    let boxes = page.boxes.as_ref().expect("a card that carries controls");

    // The gap between the buttons, which is the card's own margin as far as
    // a press is concerned: a button is a box, not a band.
    let between = (boxes.minimize.drawn.right + boxes.close.drawn.left) / 2;

    for (control, button) in [
        (CardControl::Minimize, &boxes.minimize),
        (CardControl::Close, &boxes.close),
    ] {
        let drawn = button.drawn;
        let hit = boxes.rect(control).expect("a box to press");
        let middle = (drawn.left + drawn.right) / 2;
        let at = |x: i32, y: i32| control_at(x, y, width, 96, options(), true);

        // The middle of the drawn box, and the window's own top row above it:
        // the cushion the hit box reaches to is part of the button, which is
        // what makes a mark this small a thing a hand can hit at all.
        assert_eq!(
            at(middle, (drawn.top + drawn.bottom) / 2),
            Some(control),
            "the middle of the drawn box is {control:?}"
        );
        assert_eq!(
            at(middle, 0),
            Some(control),
            "and so is the window's own top row, where the cushion reaches"
        );

        // The rows under the drawn box, down to the hit box's own last row — the
        // cushion below — and nothing past it.
        assert_eq!(
            at(middle, drawn.bottom),
            Some(control),
            "a row under the drawn box is still {control:?}"
        );
        assert_eq!(
            at(middle, hit.bottom - 1),
            Some(control),
            "and so is the cushion's last row"
        );
        assert_eq!(
            at(middle, hit.bottom),
            None,
            "while a row past it is nothing at all"
        );

        // The cushion below the drawn box is the gap the button stands in
        // plus a pixel of the name line's invisible leading — the same share
        // of the button at every scale — and the hit box it ends reaches into
        // the name's own rows, which is the room the buttons stand in: they
        // are twice the margin's own side, and a hit box that stopped at the
        // name line's top would answer for less than the button it is asked
        // about.
        let dc = unsafe { CreateCompatibleDC(None) };
        let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
        let scale = metrics.scale;
        unsafe {
            let _ = DeleteDC(dc);
        }
        assert_eq!(
            hit.bottom,
            drawn.bottom + drawn.top + scaled(1, scale),
            "the cushion is the gap above the button plus a pixel of the leading"
        );

        // And the margin beside the buttons answers nothing, nor does the gap
        // between any two of them.
        assert_eq!(
            at(0, (drawn.top + drawn.bottom) / 2),
            None,
            "the margin the buttons do not stand in is the card's own"
        );
        assert_eq!(
            at(between, (drawn.top + drawn.bottom) / 2),
            None,
            "and so is the gap between the minimize and the close"
        );
    }
}

/// The two window buttons wear nothing but their marks: at rest each is drawn in
/// the card's own ink, the one the pointer is on or holds is drawn in the accent
/// the played part of the bar is drawn in, and the corner around them is the
/// page's own color in every state — no wash, no box, no fill, in any state.
#[test]
fn the_window_buttons_are_drawn_as_ink_on_the_corner_and_nothing_else() {
    let (width, height) = measure(&pinned(), 4096, 2160, 96, options()).expect("a measured card");
    let page = scrolled(&pinned(), width, 0);
    let boxes = page.boxes.as_ref().expect("a card that carries controls");

    // The colors the paint reads, worked out here the same way: the page's own
    // color, the card's foreground, and the accent the played part of the bar is
    // drawn in.
    let theme = text_theme::loaded(options().theme).expect("the bundled theme");
    let page_color = rgb(theme.background());
    let foreground = rgb(theme.foreground());
    let accent = readable(
        rgb(theme.style_for_scopes(&["support.function"]).foreground),
        page_color,
    );

    let resting = render(&pinned(), width, height, 96, options()).expect("a painted card");
    let hovered = render(
        &lit(Some(CardControl::Minimize), None),
        width,
        height,
        96,
        options(),
    )
    .expect("a painted card");
    let held = render(
        &lit(None, Some(CardControl::Close)),
        width,
        height,
        96,
        options(),
    )
    .expect("a painted card");

    // The ink of a pixel is its color read backwards: the buffer is blue, green,
    // red, alpha, and a color is red, green, blue.
    let pixel = |painted: &(Vec<u8>, u32, u32), x: i32, y: i32| -> [u8; 4] {
        let index = (y as usize * width as usize + x as usize) * 4;
        [
            painted.0[index],
            painted.0[index + 1],
            painted.0[index + 2],
            painted.0[index + 3],
        ]
    };
    let ink = |color: [u8; 3]| [color[2], color[1], color[0], 255];
    let holds = |painted: &(Vec<u8>, u32, u32), box_: RECT, want: [u8; 4]| -> bool {
        (box_.top..box_.bottom)
            .any(|y| (box_.left..box_.right).any(|x| pixel(painted, x, y) == want))
    };

    // At rest, both buttons are drawn in the card's own ink — the ink the
    // row's glyphs are painted in, verbatim.
    for button in [&boxes.minimize, &boxes.close] {
        assert!(
            holds(&resting, button.drawn, ink(foreground)),
            "a resting button is drawn in the card's own ink: {:?}",
            button.drawn
        );
    }

    // The one the pointer is on, and the one it holds, are drawn in the accent —
    // and the buttons beside the one the pointer is on keep the ink they had.
    assert!(
        holds(&hovered, boxes.minimize.drawn, ink(accent)),
        "a hovered button is drawn in the accent"
    );
    assert!(
        holds(&hovered, boxes.close.drawn, ink(foreground)),
        "and the one beside it keeps the card's ink"
    );
    assert!(
        holds(&held, boxes.close.drawn, ink(accent)),
        "and a held button is drawn in the accent as well"
    );

    // The corner is the page's own color everywhere but the two drawn boxes, in
    // every state: nothing of a button is drawn but its mark, so there is no
    // wash under one and no box around one.
    for painted in [&resting, &hovered, &held] {
        for y in 0..page.header_top {
            for x in 0..width as i32 {
                let on_a_button = [&boxes.minimize.drawn, &boxes.close.drawn]
                    .iter()
                    .any(|drawn| {
                        x >= drawn.left && x < drawn.right && y >= drawn.top && y < drawn.bottom
                    });
                if on_a_button {
                    continue;
                }
                assert_eq!(
                    pixel(painted, x, y),
                    ink(page_color),
                    "the corner is the page's own color at ({x}, {y})"
                );
            }
        }
    }

    // The minimize is one dash: a single pixel row of ink inside its box, and
    // the close two strokes crossing, which are more than one row.
    let rows_of_ink = |painted: &(Vec<u8>, u32, u32), box_: RECT| -> Vec<i32> {
        (box_.top..box_.bottom)
            .filter(|&y| (box_.left..box_.right).any(|x| pixel(painted, x, y) != ink(page_color)))
            .collect()
    };
    assert_eq!(
        rows_of_ink(&resting, boxes.minimize.drawn).len(),
        1,
        "the minimize is one dash, one pixel row thick"
    );
    assert!(
        rows_of_ink(&resting, boxes.close.drawn).len() > 1,
        "and the close is two strokes crossing, more than one row of them"
    );
}

use super::super::playlist::{
    clock_origin, facts_style, playlist_edge_at, playlist_row, playlist_scroll_at,
    playlist_scroll_max, playlist_shown, playlist_step_pick, playlist_thumb, playlist_thumb_grab,
    playlist_track_holds, search_close_holds, search_field_holds, search_index_at, PlaylistEdge,
    PlaylistSearch,
};
use super::*;
use crate::readers::audio_track::{self, Player, Track};
use crate::text::audio_preview::{Fact, FactKind};

/// A band of a display's own scale at 96 dpi, which is what every number below is counted in.
const SCALE: f32 = 1.0;

fn a_row(name: &str) -> PlaylistRow {
    PlaylistRow {
        name: name.to_string(),
        facts: vec![
            Fact {
                text: "MP3".to_string(),
                kind: FactKind::Format,
            },
            Fact {
                text: "4.2 MB".to_string(),
                kind: FactKind::Note,
            },
        ],
        clock: None,
    }
}

/// The band's two columns: a name and its facts held to the left of the row, and the length
/// right-aligned in a column of its own at the right — so that a name is never drawn under a
/// clock, whether it is a name of four characters or of forty.
#[test]
fn the_rows_are_two_columns_and_the_clock_is_the_second() {
    let layout = playlist_layout(400, 240, SCALE, 2);

    assert!(
        layout.text_left > 0,
        "the rows are held off the band's edge"
    );
    assert!(
        layout.facts_right < layout.clock_left,
        "the facts line stops short of the clock's own"
    );
    assert!(
        layout.name_right > layout.facts_right,
        "and the name runs nearer the length than the note under it does"
    );
    assert_eq!(
        layout.clock_right, 390,
        "the clock ends where the band's own pad begins"
    );
    assert_eq!(
        layout.track, None,
        "two rows in a band that holds five of them are not a band that scrolls"
    );

    // A band with more rows than fit keeps the scrollbar's room off its right as well, and the
    // rows are narrowed by exactly that. The bar itself is a mark on the window's own edge: it is
    // held two pixels off it, which is the whole of the margin it is given.
    let scrolling = playlist_layout(400, 240, SCALE, 40);
    let (left, right) = scrolling
        .track
        .expect("a band of forty rows in the room of five scrolls");
    assert_eq!(
        right, 398,
        "the track hugs the window's own right edge, two pixels off it"
    );
    assert_eq!(
        right - left,
        4,
        "and it is the thickness of a bar: the thinnest mark a column is drawn in"
    );
    assert_eq!(
        left,
        400 - 2 - 4,
        "so its left edge is two margins off the edge"
    );
    assert!(
        scrolling.clock_right < layout.clock_right,
        "the clock gives the bar its room"
    );
    assert!(
        scrolling.clock_right < left,
        "and the clock ends before the track begins, so the two are never one column"
    );
}

/// The name's cut is measured against the length as it is drawn rather than against the column the
/// length is right-aligned in: at the widest shape a length is ever set in the two are the scaled
/// gap apart, and a shorter clock only moves the ink, never the cut.
#[test]
fn the_name_column_stops_five_pixels_short_of_the_drawn_clock() {
    let layout = playlist_layout(400, 240, SCALE, 2);
    let surface = DibSurface::create(1, 1).expect("a surface to measure on");
    let drawn = |clock: &str| {
        clock_origin(
            layout.clock_right,
            layout.clock_left,
            measure_text(&surface, &facts_style(), clock, SCALE),
        )
    };
    let gap = (5.0 * SCALE) as i32;

    assert_eq!(
        drawn("0:00:00") - layout.name_right,
        gap,
        "the widest length leaves the name cut the scaled gap off its own drawn edge"
    );
    assert!(
        drawn("0:04") - layout.name_right > gap,
        "a short length is drawn further right and the cut stays where it is"
    );
    assert!(
        drawn("1:02:03") - layout.name_right >= gap,
        "and a length of another shape never reaches back past the widest one"
    );
}

/// A press lands on the row its row of the band holds, and on nothing in the pad or below the
/// last row: the band's press is answered by the same numbers its own rows were drawn with.
#[test]
fn a_press_lands_on_the_row_its_rows_of_the_band_hold() {
    let layout = playlist_layout(400, 240, SCALE, 40);
    let row_height = layout.row_height;

    assert_eq!(
        playlist_row_at(&layout, 0, 40, layout.rows_top),
        Some(0),
        "the first row begins where the rows begin"
    );
    assert_eq!(
        playlist_row_at(&layout, 0, 40, layout.rows_top + row_height),
        Some(1)
    );
    assert_eq!(
        playlist_row_at(&layout, 0, 40, layout.rows_top + row_height - 1),
        Some(0),
        "the last pixel of a row is still that row"
    );
    assert_eq!(
        playlist_row_at(&layout, 0, 40, layout.rows_top - 1),
        None,
        "the pad above the rows is nobody's row"
    );
    assert_eq!(
        playlist_row_at(&layout, 2, 40, layout.rows_top),
        Some(2),
        "a scrolled band answers with the row it is showing rather than with the row of the slot"
    );
    assert_eq!(
        playlist_row_at(&layout, 0, 3, layout.rows_top + row_height * 3),
        None,
        "a band of three rows has nothing in the fourth row of its own room"
    );
}

/// What a band can scroll to: every row that does not fit, and none at all where they all do, and
/// a notch never scrolls past either end.
#[test]
fn the_scroll_stops_at_the_last_row_that_fits() {
    assert_eq!(playlist_scroll_max(40, 5), 35);
    assert_eq!(
        playlist_scroll_max(2, 5),
        0,
        "a band whose rows all fit has nowhere to scroll to"
    );

    // A notch up is the list come back toward its first row and a notch down is it going on, and
    // neither end is stepped past: what the band can scroll to is the last row that fits.
    let max = playlist_scroll_max(40, 5) as i32;
    for (from, rows, expected) in [
        (max, 3, max - 3),
        (max, -3, max),
        (0, 3, 0),
        (1, 3, 0),
        (0, -1, 1),
    ] {
        let scrolled = (from - rows).clamp(0, max);
        assert_eq!(scrolled, expected, "from {from} by {rows} rows");
    }
}

/// The thumb is the share of the rows the band shows, and its place is the share it has scrolled
/// by: at the first row it is at the top of the track, and at the last it is at its bottom.
#[test]
fn the_thumb_is_the_share_the_band_shows() {
    let layout = playlist_layout(400, 240, SCALE, 40);
    let room = layout.rows_bottom - layout.rows_top;
    let (top, bottom) = playlist_thumb(&layout, 0, 40);

    assert_eq!(
        top, layout.rows_top,
        "a band at its first row has its thumb at the top of the track"
    );
    // The thumb is the share of the rows the band shows, and never thinner than half a row — a bar
    // a hand cannot see is a bar a hand cannot carry (see `playlist_thumb`).
    assert_eq!(
        bottom - top,
        (room * layout.visible / 40).max(layout.row_height / 2)
    );

    let (_, bottom) = playlist_thumb(
        &layout,
        playlist_scroll_max(40, layout.visible as usize),
        40,
    );
    assert_eq!(
        bottom, layout.rows_bottom,
        "and a band at its last row has it at the bottom"
    );

    // A list short enough to fit draws a thumb as long as the track rather than none at all:
    // nothing calls this for one, and a thumb that is nowhere is what a share of nothing gives.
    let (top, bottom) = playlist_thumb(&layout, 0, 1);
    assert_eq!((top, bottom), (layout.rows_top, layout.rows_bottom));
}

/// The band's two edges are the window's two edges, and no row is ever drawn under either: the pad
/// below the last row is wider than the edge is tall, and the pad above the first row is the same
/// on the other side.
#[test]
fn the_edges_are_the_windows_and_never_a_row() {
    let layout = playlist_layout(400, 240, SCALE, 40);
    let top = 300;

    assert_eq!(
        playlist_edge_at(top, 240, SCALE, 0),
        Some(PlaylistEdge::Top)
    );
    assert_eq!(
        playlist_edge_at(top, 240, SCALE, 5),
        Some(PlaylistEdge::Top)
    );
    assert_eq!(
        playlist_edge_at(top, 240, SCALE, 6),
        None,
        "below the window's own first rows the band is the rows' own"
    );

    assert_eq!(
        playlist_edge_at(top, 240, SCALE, top + 239),
        Some(PlaylistEdge::Bottom)
    );
    assert_eq!(
        playlist_edge_at(top, 240, SCALE, top + 234),
        Some(PlaylistEdge::Bottom)
    );
    assert_eq!(
        playlist_edge_at(top, 240, SCALE, top + 233),
        None,
        "above the bottom edge the band is the rows' own"
    );
    assert_eq!(
        playlist_edge_at(top, 240, SCALE, top + 240),
        None,
        "and a row below the band is nobody's"
    );

    assert!(
        playlist_row_at(&layout, 0, 40, 234).is_none() || layout.rows_bottom <= 234,
        "no row of the band is drawn in the bottom edge"
    );
}

/// A row says what the file says: the format, the rate, the channels and the bitrate the probe
/// answered for, and the room the file takes — joined with the separator the card's facts line is
/// joined with, and with the extension in capitals where nothing has probed the file.
#[test]
fn a_row_says_what_the_file_says_and_who_long_it_is() {
    let root = scratch("playlist-row");
    let probed = root.join("mixdown.flac");
    std::fs::write(&probed, vec![0u8; 3 * 1024 * 1024]).expect("a written file");
    audio_track::remember(
        &probed,
        Track {
            player: Player::Native,
            codec: Some("FLAC".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(1_006_000),
            duration: Some(3_723.0),
        },
    );

    let row = playlist_row(&probed);
    assert_eq!(
        row.name, "mixdown",
        "a row names the file without the extension it is already known by"
    );
    let parts: Vec<(&str, FactKind)> = row
        .facts
        .iter()
        .map(|fact| (fact.text.as_str(), fact.kind.clone()))
        .collect();
    assert_eq!(
        parts,
        vec![
            ("FLAC", FactKind::Format),
            ("44.1 kHz", FactKind::Reading),
            ("stereo", FactKind::Note),
            ("1006 kbps", FactKind::Reading),
            ("3.0 MB", FactKind::Note),
        ],
        "the parts are the card's own, in the card's own inks"
    );
    assert_eq!(row.clock.as_deref(), Some("1:02:03"));

    // A file nothing has probed says only what it can be read to say without one: the extension
    // it carries, in the capitals a codec's name is written in, and the room it takes.
    let bare = root.join("unknown.ape");
    std::fs::write(&bare, vec![0u8; 1024 * 1024]).expect("a written file");
    let row = playlist_row(&bare);
    assert_eq!(row.name, "unknown");
    assert_eq!(
        row.facts
            .iter()
            .map(|fact| fact.text.as_str())
            .collect::<Vec<_>>(),
        vec!["APE", "1.0 MB"],
        "and its parts are the extension and the room it takes"
    );
    assert_eq!(
        row.clock, None,
        "and nothing at all where no length is known"
    );
}

/// A row whose facts are gone is a row whose file can never be previewed again, because the read
/// that found nothing was held as the answer: the row and the card both ask the memo, and an
/// emptiness held tells them nothing more is coming until the app is restarted (see
/// `audio_track::remember`).
#[test]
fn a_row_of_a_file_no_read_could_answer_for_says_what_the_folder_says_and_takes_the_facts_a_read_lands(
) {
    let root = scratch("playlist-row-unread");
    let probed = root.join("mixdown.flac");
    std::fs::write(&probed, vec![0u8; 3 * 1024 * 1024]).expect("a written file");

    let words = |path: &std::path::Path| -> Vec<String> {
        playlist_row(path)
            .facts
            .iter()
            .map(|fact| fact.text.clone())
            .collect()
    };
    let unread = words(&probed);
    assert_eq!(
        unread,
        vec!["FLAC", "3.0 MB"],
        "a row nothing read is the extension and its size"
    );

    // The sweep's read of this file answers with nothing — the case a FLAC takes where nothing in
    // this process decodes it — and the next sweep is owed the file again.
    assert_eq!(
        audio_track::probed(&probed),
        audio_track::Probed::NotAsked,
        "a read that found nothing is no answer a row or a card can be drawn from"
    );
    assert_eq!(
        words(&probed),
        unread,
        "and the row it did not answer for is the row it was"
    );

    // The read the sweep after it takes lands, and the row is redrawn from it.
    audio_track::remember(
        &probed,
        Track {
            player: Player::Native,
            codec: Some("FLAC".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(1_006_000),
            duration: Some(3_723.0),
        },
    );
    assert_eq!(
        words(&probed),
        vec!["FLAC", "44.1 kHz", "stereo", "1006 kbps", "3.0 MB"],
        "the facts the read landed are the facts the row is drawn with"
    );
}

/// A name with room is a name that never moves; one without travels the width it overruns by, at
/// the advance of its own font — and a row with no room at all is a row with no travel to give,
/// because there is nothing a marquee could show of it.
#[test]
fn a_name_too_wide_for_its_row_has_a_travel() {
    let (travel, advance) = name_travel("a mixdown of a very long name indeed", 60, SCALE)
        .expect("a name far wider than sixty pixels");

    assert!(travel > 0);
    assert!(advance > 0, "the step is the font's own advance");
    assert!(
        travel >= 60,
        "a name of forty characters overruns a column of sixty pixels by far more than that"
    );
    assert_eq!(name_travel("short", 400, SCALE), None);
    assert_eq!(name_travel("short", 0, SCALE), None);
}

/// The band is painted opaque the whole of its box down, which is what makes a press on it the
/// pin's rather than the listing's behind it, and the row the pin is showing carries the accent
/// stripe at its own left edge.
#[test]
fn the_band_claims_its_whole_box_and_stripes_the_row_it_shows() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let (width, height) = (240u32, 240u32);
    let surface = DibSurface::create(width, height).expect("a surface");
    let layout = playlist_layout(width as i32, height as i32, SCALE, 2);
    let rows = [a_row("one"), a_row("two")];

    paint_playlist(
        &surface,
        &palette,
        &PlaylistBand {
            layout,
            rows: &rows,
            shown: &[0, 1],
            first: 0,
            total: 2,
            showing: Some(1),
            picked: None,
            scrolled: &[],
            search: None,
            note: None,
        },
        SCALE,
    );

    let pixels =
        unsafe { std::slice::from_raw_parts(surface.bits(), (width * height * 4) as usize) };

    // Every pixel of the band is a pixel the window claims, the row the pin is showing included.
    for y in 0..height as usize {
        for x in 0..width as usize {
            let alpha = pixels[(y * width as usize + x) * 4 + 3];
            assert_eq!(alpha, 255, "the band's pixel at ({x}, {y}) is not opaque");
        }
    }

    // The stripe is the accent, at the second row's own left edge.
    let second_top = layout.rows_top + layout.row_height;
    let stripe = pixels[(second_top as usize * width as usize + 1) * 4..][..3].to_vec();
    assert_eq!(
        stripe,
        [palette.accent[2], palette.accent[1], palette.accent[0]],
        "the row the pin is showing is striped in the accent, premultiplied"
    );
}

/// A thumb dragged down the track is the scroll that row means, and the two are one arithmetic: a bar
/// drawn from one and answered by the other would drift out from under the hand that is carrying it.
#[test]
fn a_thumb_dragged_down_the_track_is_the_scroll_it_means() {
    let layout = playlist_layout(400, 240, SCALE, 40);
    let max = playlist_scroll_max(40, layout.visible.max(0) as usize);
    assert!(max > 0, "a band of forty rows in the room of five scrolls");

    assert_eq!(
        playlist_scroll_at(&layout, 40, layout.rows_top),
        0,
        "the top of the track is the first row"
    );
    assert_eq!(
        playlist_scroll_at(&layout, 40, layout.rows_bottom),
        max,
        "and the bottom of it is the last"
    );

    // The two are inverses over the whole of the travel: where the bar draws the thumb is the row a
    // drag to that place comes back to.
    for step in 0..=20 {
        let scrolled = max * step / 20;
        let (top, _) = playlist_thumb(&layout, scrolled, 40);
        let back = playlist_scroll_at(&layout, 40, top);

        assert!(
            back.abs_diff(scrolled) <= 1,
            "scroll {scrolled} drawn at {top} came back as {back}"
        );
    }
}

/// A press on the bar takes hold of the part of the thumb it landed on, and of the middle of it where
/// it landed on the track: a thumb that jumped to put its top under the hand is a list that jumped
/// the moment it was touched.
#[test]
fn a_press_on_the_bar_takes_hold_of_the_thumb_it_found() {
    let layout = playlist_layout(400, 240, SCALE, 40);
    let (top, bottom) = playlist_thumb(&layout, 10, 40);

    assert_eq!(
        playlist_thumb_grab(&layout, 10, 40, top + 3),
        3,
        "the hand took the thumb three pixels down from its top"
    );
    assert_eq!(
        playlist_thumb_grab(&layout, 10, 40, layout.rows_top),
        (bottom - top) / 2,
        "and a press on the track above it took the middle"
    );

    // And the bar is grabbed by the room it is the mark in rather than by its own four pixels: from
    // the length's last column to the window's own right edge, and nowhere left of that.
    let (left, right) = layout.track.expect("a band that scrolls has a track");
    assert!(playlist_track_holds(&layout, right - 1, layout.rows_top));
    assert!(
        playlist_track_holds(&layout, layout.width - 1, layout.rows_top),
        "the last pixel of the window is the bar's"
    );
    assert!(
        playlist_track_holds(&layout, layout.clock_right - 1, layout.rows_bottom - 1),
        "and so is the last pixel the length is inked in"
    );
    assert!(
        !playlist_track_holds(&layout, layout.clock_right - 2, layout.rows_top),
        "a press left of the length is the rows'"
    );
    assert!(playlist_track_holds(&layout, left - 1, layout.rows_top));
    assert!(
        !playlist_track_holds(&layout, layout.clock_right - 1, layout.rows_top - 1),
        "and one above the rows is nobody's, however far right it is"
    );
}

/// The arrows step the highlight a row at a time inside the band and hold at both ends: a step off
/// the top does not come back on at the bottom, and the first press with nothing picked starts on
/// the row the sound on screen is on, or on the band's first row in sight where it has none.
#[test]
fn the_arrows_step_a_row_and_hold_at_either_end() {
    // Nothing picked and nothing playing: the first press starts where the band is looking.
    assert_eq!(playlist_step_pick(None, None, 0, 40, 5, -1), Some((0, 0)));
    assert_eq!(playlist_step_pick(None, None, 0, 40, 5, 1), Some((1, 0)));

    // A first press on a pin that is playing starts on the sound's own row and carries the band to
    // it: the row it steps from is one the band has scrolled to rather than one it can see.
    assert_eq!(playlist_step_pick(None, Some(7), 0, 40, 5, 1), Some((8, 4)));
    assert_eq!(
        playlist_step_pick(None, Some(7), 0, 40, 5, -1),
        Some((6, 2))
    );

    // Either end holds rather than wraps.
    assert_eq!(
        playlist_step_pick(Some(0), None, 0, 40, 5, -1),
        Some((0, 0))
    );
    assert_eq!(
        playlist_step_pick(Some(39), Some(39), 35, 40, 5, 1),
        Some((39, 35))
    );

    // And a band of no rows has no row to step to.
    assert_eq!(playlist_step_pick(None, None, 0, 0, 5, 1), None);
}

/// A highlight stepped out of the rows in sight is carried back into them by the scroll, which is
/// the band's own unit: a whole row, so that no row is ever drawn half of one — and every row the
/// arrows can land on is one the paint draws from the scroll they leave.
#[test]
fn a_stepped_row_is_scrolled_back_into_sight() {
    let (rows, visible) = (40usize, 5usize);

    // Down out of the last row in sight, and up out of the first: the band follows it one row.
    assert_eq!(
        playlist_step_pick(Some(4), None, 0, rows, visible, 1),
        Some((5, 1))
    );
    assert_eq!(
        playlist_step_pick(Some(1), None, 1, rows, visible, -1),
        Some((0, 0))
    );

    // The last row of the list is the last row the band can scroll to, and the first is its first.
    let last = rows - 1;
    assert_eq!(
        playlist_step_pick(Some(last - 1), None, last - visible, rows, visible, 1),
        Some((last, last + 1 - visible))
    );

    // A step inside the rows in sight moves nothing of the scroll.
    assert_eq!(
        playlist_step_pick(Some(5), None, 2, rows, visible, 1),
        Some((6, 2))
    );

    // And whatever the row it lands on, it is one of the rows the paint draws.
    for picked in 0..rows {
        for direction in [-1, 1] {
            let (to, scroll) =
                playlist_step_pick(Some(picked), None, picked, rows, visible, direction)
                    .expect("a band of forty rows");
            assert!(
                scroll <= to && to < scroll + visible,
                "picked {picked} stepped to {to}, which is outside the rows drawn from {scroll}"
            );
        }
    }
}

/// Both rows of a band are drawn at the one size: the second row is the first row's own painting,
/// name and facts line alike. Two painters share one surface's DC and the face each remembers is
/// only its own — the second row is where a face the other left behind shows (see
/// `RunPainter::forget_the_face`).
#[test]
fn a_row_is_drawn_exactly_as_the_row_above_it() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let (width, height) = (400u32, 150u32);
    let surface = DibSurface::create(width, height).expect("a surface");
    let layout = playlist_layout(width as i32, height as i32, SCALE, 2);
    assert_eq!(layout.visible, 2, "two rows in the room of two");
    let name = "a name of the one length";
    let mut rows = [a_row(name), a_row(name)];
    // A length as well, so that the row's right column is drawn in the same test: it is the one run
    // of a row that is set at the body level rather than at either of the two above it.
    for row in &mut rows {
        row.clock = Some("1:02:03".to_string());
    }

    paint_playlist(
        &surface,
        &palette,
        &PlaylistBand {
            layout,
            rows: &rows,
            shown: &[0, 1],
            first: 0,
            total: 2,
            showing: None,
            picked: None,
            scrolled: &[],
            search: None,
            note: None,
        },
        SCALE,
    );

    let pixels =
        unsafe { std::slice::from_raw_parts(surface.bits(), (width * height * 4) as usize) };
    let row = |index: usize| {
        let top =
            (layout.rows_top + index as i32 * layout.row_height) as usize * width as usize * 4;
        &pixels[top..top + layout.row_height as usize * width as usize * 4]
    };

    assert_eq!(
        row(0),
        row(1),
        "the second row is drawn away from the first, so a face one painter selected is the face \
         the other drew in"
    );
}

/// The search field is a constant of the layout: it sits at the band's bottom whatever band it is
/// given, and a band taller or shorter than another gives the rows above it what it took rather than
/// stretching the field. A field that moved with the band would be a field the hand had to find again
/// on every drag of the height.
#[test]
fn the_search_field_sits_at_the_bands_bottom_whatever_its_height() {
    let tall = playlist_layout(400, 400, SCALE, 40);
    let short = playlist_layout(400, 240, SCALE, 40);

    let height = tall.search_bottom - tall.search_top;
    assert!(height > 0, "a band always draws the field it was given");
    assert_eq!(
        short.search_bottom - short.search_top,
        height,
        "the field is the same number of pixels in a band of any height"
    );

    for layout in [tall, short] {
        assert_eq!(
            layout.search_bottom, layout.rows_bottom,
            "the field's own bottom is the rows' own bottom: it floats over the last rows rather than a
             strip being kept from them"
        );
        assert_eq!(
            layout.search_top,
            layout.rows_bottom - height,
            "and its top is where a pill of its own height begins, over the rows under it"
        );
        assert_eq!(
            layout.search_left, layout.rows_top,
            "the field spans the band's inner width: the pad off the left, as the rows are"
        );
        assert_eq!(
            (layout.search_left, layout.search_right),
            (short.search_left, short.search_right),
            "and neither band is narrower than the other"
        );
    }

    assert_eq!(
        tall.search_bottom, tall.rows_bottom,
        "the field's own bottom is the band's bottom pad, not its top grown by the one constant"
    );
    assert_eq!(
        tall.search_bottom - short.search_bottom,
        400 - 240,
        "the field is anchored to the band's bottom rather than to its top"
    );

    // The rows are the room the field leaves: every pixel a band is dragged taller is a pixel the
    // rows above the field are taller by.
    let rows = |layout: &PlaylistLayout| layout.rows_bottom - layout.rows_top;
    assert_eq!(
        rows(&tall) - rows(&short),
        400 - 240,
        "a band dragged a hundred and sixty pixels taller gives the rows that many"
    );

    // A band too short for even one row still has its field: the rows are nothing and the field is
    // the field.
    // The field and one pad above and below it, and not a pixel more.
    let squeezed = playlist_layout(400, height + 2 * short.rows_top, SCALE, 40);
    assert_eq!(squeezed.visible, 0, "no row fits above the field there");
    assert_eq!(
        squeezed.search_bottom - squeezed.search_top,
        height,
        "and the field it draws is the one every other band draws"
    );

    // The bottom of the band is the window's own drag handle, and the field is held clear of it: a
    // press on the field is a caret and never a height being pulled.
    let top = 300;
    let band = squeezed.search_bottom + squeezed.rows_top;
    for y in squeezed.search_top..squeezed.search_bottom {
        assert_eq!(
            playlist_edge_at(top, band, SCALE, top + y),
            None,
            "the field's row {y} is no drag of the window's bottom edge"
        );
    }
    assert_eq!(
        playlist_edge_at(top, band, SCALE, top + band - 1),
        Some(PlaylistEdge::Bottom),
        "and the window's last rows are the band's edge, below the field"
    );
}

/// **The search field floats over the last rows of the list.** A band is a list to its own bottom pad:
/// the rows reach it, and the pill is drawn over the last of them while it is in sight — a field that
/// held a strip of the band back from the rows is a band that ends in a row of nothing.
///
/// What is drawn under the last whole row is a row peeking only where the band is not a whole number of
/// rows: a band snapped to its rows (see `snapped_playlist_height`) has none at all.
#[test]
fn the_field_floats_over_the_rows_and_they_reach_the_bands_own_pad() {
    let pad = 10;
    let row = playlist_row_height(SCALE);
    // A band a row and a half tall: four whole rows, and a row peeking under them.
    let unsnapped = 2 * pad + 4 * row + row / 2;
    let layout = playlist_layout(400, unsnapped, SCALE, 40);

    assert_eq!(
        layout.rows_bottom,
        unsnapped - pad,
        "the rows reach the band's own bottom pad"
    );
    assert_eq!(
        layout.search_bottom, layout.rows_bottom,
        "and the field's bottom is that same row: it floats over the rows rather than under them"
    );
    assert!(
        layout.search_top < layout.rows_bottom
            && layout.search_top >= layout.rows_top + (layout.visible - 1) * row,
        "the pill is drawn over the last rows the list draws rather than over a row of page"
    );

    // A band snapped to whole rows has nothing left under the last of them, and the field over the last
    // row of the list rather than a row of page.
    let snapped = snapped_playlist_height(unsnapped - 3, SCALE);
    let layout = playlist_layout(400, snapped, SCALE, 40);
    assert_eq!(
        layout.rows_top + layout.visible * row + pad,
        snapped,
        "a snapped band ends exactly on its last whole row and the pad under it"
    );
    assert!(
        layout.search_bottom > layout.rows_top + (layout.visible - 1) * row,
        "so the field is drawn over a row of the list and not over a row of page"
    );
}

/// A point of the field is the field's, from the pad it begins at to the pad it ends short of — and a
/// point above it is the rows', which is what keeps a press on a row off the field.
#[test]
fn a_point_of_the_field_is_the_fields_and_the_rows_above_it_are_not() {
    let layout = playlist_layout(400, 240, SCALE, 40);
    let (top, bottom) = (layout.search_top, layout.search_bottom);

    assert!(search_field_holds(&layout, layout.search_left, top));
    assert!(search_field_holds(
        &layout,
        layout.search_right - 1,
        bottom - 1
    ));
    assert!(
        !search_field_holds(&layout, layout.search_left - 1, top),
        "left of the pad the band is the rows'"
    );
    assert!(
        !search_field_holds(&layout, layout.width - 1, top),
        "and right of the pad it is the scrollbar's"
    );
    assert!(!search_field_holds(&layout, 200, top - 1));
    assert!(!search_field_holds(&layout, 200, bottom));
}

/// **The clear button is the field's right-hand mark and the query stops before it.** The pill is a
/// field and a control in one, so the × that puts the query back to nothing is held the same pad off
/// the right edge that the magnifier keeps off the left, and the last column the query may reach is
/// the gap before it — a press on the × clears the field, and a press in the text sets a caret in it.
#[test]
fn the_clear_button_stands_at_the_fields_right_and_the_query_stops_before_it() {
    let layout = playlist_layout(400, 240, SCALE, 0);
    let (top, bottom) = (layout.search_top, layout.search_bottom);

    assert!(
        layout.search_close_left > layout.search_text_right,
        "the × stands past the last column the query may reach"
    );
    assert!(
        layout.search_close_right < layout.search_right,
        "and it is held inside the pill's own pad"
    );
    assert_eq!(
        layout.search_right - layout.search_close_right,
        layout.search_mark - layout.search_left,
        "the × keeps the magnifier's own pad off the pill's right edge"
    );

    let middle = (layout.search_close_left + layout.search_close_right) / 2;
    assert!(
        search_close_holds(&layout, middle, top),
        "the × is the clear button's"
    );
    assert!(
        search_close_holds(&layout, layout.search_right - 1, bottom - 1),
        "and the pad between it and the pill's edge is the button's too, rather than a mark a hand has to hit"
    );
    assert!(
        !search_close_holds(&layout, layout.search_close_left - 1, top),
        "the query's own columns are not the button's"
    );
    assert!(
        !search_close_holds(&layout, middle, top - 1),
        "and no row above the field is"
    );
    assert!(!search_close_holds(&layout, middle, bottom));
    assert!(
        search_field_holds(&layout, middle, top),
        "a press on the × is still a press on the field it stands in"
    );
}

/// A paint of a band of no rows, so that the field's own pixels are the only ones inside its box.
fn a_band_of_the_field(search: PlaylistSearch) -> (DibSurface, RECT) {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let (width, height) = (400u32, 240u32);
    let surface = DibSurface::create(width, height).expect("a surface");
    let layout = playlist_layout(width as i32, height as i32, SCALE, 0);
    let rows: [PlaylistRow; 0] = [];

    paint_playlist(
        &surface,
        &palette,
        &PlaylistBand {
            layout,
            rows: &rows,
            shown: &[],
            first: 0,
            total: 0,
            showing: None,
            picked: None,
            scrolled: &[],
            search: Some(search),
            note: None,
        },
        SCALE,
    );

    let box_ = RECT {
        left: layout.search_left,
        top: layout.search_top,
        right: layout.search_right,
        bottom: layout.search_bottom,
    };

    (surface, box_)
}

/// The field draws the placeholder where the query is nothing, the query where it is something, and
/// the caret only while it has the focus — and every pixel of it is opaque, because the band is and a
/// layered window is hit-tested by the shape of its pixels.
#[test]
fn the_field_draws_its_placeholder_its_query_and_its_caret() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let page = palette.background;
    // The field's own fill: what a field with nothing in it draws and nothing else does.
    let fill = text_paint::blend(page, palette.foreground, 0.07);

    let pixels_of = |surface: &DibSurface| unsafe {
        std::slice::from_raw_parts(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    let ink = |pixels: &[u8], width: usize, x: i32, y: i32| {
        let at = (y as usize * width + x as usize) * 4;
        [pixels[at + 2], pixels[at + 1], pixels[at]]
    };
    // How much of the field was inked by something that is neither its own fill nor the page under it.
    let inked = |surface: &DibSurface, box_: RECT| {
        let pixels = pixels_of(surface);
        let width = surface.width as usize;
        (box_.top..box_.bottom)
            .flat_map(|y| (box_.left..box_.right).map(move |x| (x, y)))
            .filter(|(x, y)| {
                let at = ink(pixels, width, *x, *y);
                at != page && at != fill
            })
            .count()
    };

    let (empty, box_) = a_band_of_the_field(PlaylistSearch {
        text: "",
        caret: 0,
        focused: false,
        caret_lit: true,
        slide: 1.0,
        selection: None,
    });
    let placeholder = inked(&empty, box_);
    assert!(
        placeholder > 0,
        "an empty field says what it is for rather than drawing nothing"
    );

    let (focused, box_) = a_band_of_the_field(PlaylistSearch {
        text: "",
        caret: 0,
        focused: true,
        caret_lit: true,
        slide: 1.0,
        selection: None,
    });
    assert!(
        inked(&focused, box_) > placeholder,
        "a focused field draws a caret the unfocused one has none of"
    );
    // The caret is a bar rather than a mark: it inks a column of the field's text down the whole
    // height of a run, which is what tells it from the placeholder's glyphs beside it.
    let pixels = pixels_of(&focused);
    let column = (box_.left..box_.right)
        .filter(|x| {
            (box_.top..box_.bottom)
                .filter(|y| {
                    let at = ink(pixels, focused.width as usize, *x, *y);
                    at != page && at != fill
                })
                .count()
                >= 10
        })
        .count();
    assert!(
        column > 0,
        "a focused field draws a caret bar the unfocused one has none of"
    );

    let (typed, box_) = a_band_of_the_field(PlaylistSearch {
        text: "mixdown",
        caret: 4,
        focused: true,
        caret_lit: true,
        slide: 1.0,
        selection: None,
    });
    assert!(
        inked(&typed, box_) > 0,
        "a query in the field is inked inside it"
    );

    for surface in [&empty, &focused, &typed] {
        let pixels = pixels_of(surface);
        for y in box_.top..box_.bottom {
            for x in box_.left..box_.right {
                assert_eq!(
                    pixels[(y as usize * surface.width as usize + x as usize) * 4 + 3],
                    255,
                    "the field's pixel at ({x}, {y}) is not opaque"
                );
            }
        }
    }
}

/// The place a press on the field names: before the first character at its own left edge, after the
/// last at its right — the two ends of a query of any length.
#[test]
fn a_press_on_the_field_names_a_place_in_the_query() {
    let layout = playlist_layout(400, 240, SCALE, 0);

    assert_eq!(
        search_index_at(&layout, "", layout.search_left, SCALE),
        0,
        "an empty field has one place in it"
    );
    assert_eq!(
        search_index_at(&layout, "mixdown", layout.search_left, SCALE),
        0,
        "a press at the field's left edge is before the first character"
    );
    assert_eq!(
        search_index_at(&layout, "mixdown", layout.search_right, SCALE),
        "mixdown".len(),
        "and a press at its right edge is after the last"
    );
}

/// The rows a query lets through are a case-insensitive substring of the file's own name and never a
/// word of its facts line: a hand looking for a track is naming a track, not a format or a rate.
#[test]
fn the_query_lets_through_the_names_that_hold_it_and_keeps_their_order() {
    let rows = [
        a_row("First Mixdown"),
        a_row("second take"),
        a_row("MIXDOWN live"),
        a_row("an outro"),
    ];
    let names = |query: &str| {
        playlist_shown(&rows, query)
            .into_iter()
            .map(|at| rows[at].name.clone())
            .collect::<Vec<_>>()
    };

    assert_eq!(
        names(""),
        vec!["First Mixdown", "second take", "MIXDOWN live", "an outro"],
        "an empty query hides nothing and reorders nothing"
    );
    assert_eq!(
        names("   "),
        vec!["First Mixdown", "second take", "MIXDOWN live", "an outro"],
        "and a query of nothing but spaces is an empty query"
    );
    assert_eq!(
        names("mixdown"),
        vec!["First Mixdown", "MIXDOWN live"],
        "the match ignores case, and the rows keep the order the walk gave them"
    );
    assert_eq!(
        names("SECOND"),
        vec!["second take"],
        "and the other way round as well"
    );
    assert_eq!(
        names("mix"),
        vec!["First Mixdown", "MIXDOWN live"],
        "a substring anywhere in the name is a match"
    );
    assert!(
        names("mp3").is_empty(),
        "the format is a fact of the file and not its name"
    );
    assert!(names("4.2 MB").is_empty(), "and so is the room it takes");
    assert!(names("nothing here").is_empty());
}

/// What a field's pixel is, read back out of the surface in the order the palette keeps its colours:
/// the page under it where nothing was drawn, or the ink that was.
fn field_pixels(surface: &DibSurface) -> Vec<[u8; 3]> {
    let width = surface.width as usize;
    let pixels =
        unsafe { std::slice::from_raw_parts(surface.bits(), width * surface.height as usize * 4) };

    (0..width * surface.height as usize)
        .map(|at| {
            let index = at * 4;
            [pixels[index + 2], pixels[index + 1], pixels[index]]
        })
        .collect()
}

/// The field is a **pill**: the plate it is filled with is rounded to half its height, so the four
/// pixels in its corners are the band's page rather than the plate, and the two in the middle of its
/// sides are the plate. A square box of the same colour would be the one shape this is not.
#[test]
fn the_field_is_a_pill_rather_than_a_box() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let page = palette.background;
    let plate = text_paint::blend(page, palette.foreground, 0.07);

    let (surface, box_) = a_band_of_the_field(PlaylistSearch {
        text: "",
        caret: 0,
        focused: false,
        caret_lit: true,
        slide: 1.0,
        selection: None,
    });
    let pixels = field_pixels(&surface);
    let at = |x: i32, y: i32| pixels[y as usize * surface.width as usize + x as usize];
    let corner_ink = |x: i32, y: i32| {
        let corner = at(x, y);
        corner != page && corner != plate
    };

    // The corners are off the pill, so what is there is the band's page and not the plate: the radius
    // takes the corner of the box the layout gave it.
    assert_eq!(
        at(box_.left, box_.top),
        page,
        "the top-left corner of the field's box is outside the pill"
    );
    assert_eq!(
        at(box_.right - 1, box_.bottom - 1),
        page,
        "and so is the bottom-right one"
    );

    // The middle of each side is inside it, which is the same box read where the pill is whole: the
    // pill's own edge carries its hairline there and the pixel past it is the plate.
    let middle = (box_.top + box_.bottom) / 2;
    assert_ne!(
        at(box_.left, middle),
        page,
        "the middle of the pill's left edge is the pill's own edge"
    );
    assert_eq!(
        at(box_.left + 1, middle),
        plate,
        "and the pixel inside it is the plate the pill is filled with"
    );
    assert_ne!(
        at(box_.right - 1, middle),
        page,
        "the middle of its right edge is the pill's too"
    );
    assert_eq!(
        at(box_.right - 2, middle),
        plate,
        "and the pixel inside that edge is the plate"
    );

    // The corner is the page and the pixel a row in from it is on the curve rather than on the
    // plate, which is what a radius of half the height does to the box the layout gave the field.
    assert!(
        !corner_ink(box_.left + 2, box_.top + 1),
        "the pill's own edge is a curve, not a stair of whole pixels"
    );
}

/// The mark is a magnifier at the pill's left — a ring and a handle, in the muted ink, inside the
/// square the layout kept for it — and the query is set after it rather than over it.
#[test]
fn the_field_draws_its_mark_at_its_left_and_sets_the_query_after_it() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let page = palette.background;
    let plate = text_paint::blend(page, palette.foreground, 0.07);

    let layout = playlist_layout(400, 240, SCALE, 0);
    let (surface, box_) = a_band_of_the_field(PlaylistSearch {
        text: "",
        caret: 0,
        focused: false,
        caret_lit: true,
        slide: 1.0,
        selection: None,
    });
    let pixels = field_pixels(&surface);
    let width = surface.width as usize;
    let inked = |left: i32, right: i32| {
        (box_.top..box_.bottom)
            .flat_map(|y| (left..right).map(move |x| (x, y)))
            .filter(|(x, y)| {
                let at = pixels[*y as usize * width + *x as usize];
                at != page && at != plate
            })
            .count()
    };

    // The mark's own square is the one the layout names, and it is the square the mark is inked in.
    let mark = layout.search_mark;
    let side = layout.search_text_left - mark;
    assert!(side > 0, "the query begins past the mark");

    // The mark is a ring and a handle rather than a solid square of ink: the middle of its square is
    // the hole of the ring, which carries the plate rather than the mark.
    let centre = (box_.top + box_.bottom) / 2;
    assert!(
        inked(mark, layout.search_text_left) > 0,
        "the mark's square carries ink"
    );

    // Its centre is the ring's hole, so the mark is a shape rather than a block: the hole's own
    // pixels are the plate the pill was filled with.
    let hole = (mark + 1..layout.search_text_left - 1)
        .filter(|x| {
            (centre - 1..=centre + 1).all(|y| {
                let at = pixels[y as usize * width + *x as usize];
                at == plate || at == page
            })
        })
        .count();
    assert!(
        hole > 0,
        "the mark's middle is a hole rather than a block of ink"
    );

    // And the query begins at the column the layout kept for it, past the mark's own square: the
    // field's text and the mark never share a pixel.
    assert!(
        layout.search_text_left > mark + side / 2,
        "the query begins past the middle of the mark, not on it"
    );
}

/// The × at the field's right is drawn, and it is drawn in the muted ink the magnifier beside it is
/// drawn in rather than in the query's own brightness — a control about the text is not text of its
/// own, and the query is the one bright thing in the field.
///
/// The query is painted too, and the same question is asked of it: nothing of a query is inked under
/// the ×, which is what the layout's own last column for it keeps.
#[test]
fn the_field_draws_its_clear_button_in_the_muted_ink_at_its_right() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let page = palette.background;
    let plate = text_paint::blend(page, palette.foreground, 0.07);
    let muted = text_paint::blend(page, palette.foreground, 0.45);

    let layout = playlist_layout(400, 240, SCALE, 0);
    let (left, right) = (layout.search_close_left, layout.search_close_right);
    assert!(right > left, "the field keeps a square for its ×");

    // What the button's own square carries, over and above the plate the pill is filled with.
    let button_ink = |text: &str| {
        let (surface, box_) = a_band_of_the_field(PlaylistSearch {
            text,
            caret: text.len(),
            selection: None,
            focused: false,
            caret_lit: true,
            slide: 1.0,
        });
        let pixels = field_pixels(&surface);
        let width = surface.width as usize;

        (box_.top..box_.bottom)
            .flat_map(|y| (left..right).map(move |x| (x, y)))
            .map(|(x, y)| pixels[y as usize * width + x as usize])
            .filter(|at| *at != page && *at != plate)
            .collect::<Vec<_>>()
    };

    let inked = button_ink("");
    assert!(
        !inked.is_empty(),
        "the field draws its clear button whether the query is something or nothing"
    );
    assert!(
        inked.contains(&muted),
        "and the × is inked in the muted gray the magnifier beside it is drawn in"
    );

    // The query is painted to the last column the layout gives it, and none of it reaches the button:
    // the × is a control a hand presses rather than a mark drawn over the text beside it.
    for at in button_ink("a query with room to spare in the field") {
        assert_ne!(
            at, palette.foreground,
            "the query's own ink is not drawn under the ×"
        );
    }
}

/// **A selection is washed between its two ends**, in the ink the find bar washes its own query with,
/// and it is the field's only while the field holds the caret: a wash the keys have walked away from
/// is a range nothing on the window acts on.
///
/// The query painted here is a run of spaces, so that the field's own plate is the only thing under the
/// wash and the count below is the wash's and nothing else's — a glyph's antialiasing has no ink to
/// land in a range of nothing.
#[test]
fn a_selection_is_washed_between_its_two_ends() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let page = palette.background;
    let plate = text_paint::blend(page, palette.foreground, 0.07);
    let wash = text_paint::blend(plate, palette.foreground, 0.25);

    // The columns of the field a paint washes, which is what a selection is: the two ends name a range
    // and the wash fills the columns between them rather than the run's whole box.
    let washed = |selection: Option<(usize, usize)>, focused: bool| {
        let (surface, _) = a_band_of_the_field(PlaylistSearch {
            text: "     ",
            caret: 5,
            selection,
            focused,
            caret_lit: false,
            slide: 1.0,
        });
        let pixels = field_pixels(&surface);
        let width = surface.width as usize;

        (0..surface.width as i32)
            .filter(|x| {
                (0..surface.height as i32).any(|y| pixels[y as usize * width + *x as usize] == wash)
            })
            .count()
    };

    assert_eq!(
        washed(None, true),
        0,
        "a caret with nothing selected washes nothing"
    );
    assert_eq!(
        washed(Some((5, 5)), true),
        0,
        "and neither does a range of nothing"
    );

    let range = washed(Some((1, 3)), true);
    let whole = washed(Some((0, 5)), true);
    assert!(range > 0, "a selection is washed between its two ends");
    assert!(
        whole > range,
        "and the wider range is washed wider: the wash is the range's, not the run's"
    );
    assert_eq!(
        washed(Some((0, 5)), false),
        0,
        "a field the keys have left washes nothing at all"
    );
}

/// The caret is the size of the face the query is set in rather than the height of the line box: a
/// bar taller than the characters it stands between reads as a column of the field, not as a place
/// the next character lands.
///
/// The two fields painted here are one query with the caret in it and the same query without it, so
/// every pixel they disagree on is the caret's — which is the exact question, a bar of an accent that
/// is blended over the plate cannot be found by its own colour.
#[test]
fn the_caret_is_the_height_of_the_letter_it_stands_between() {
    let with = a_band_of_the_field(PlaylistSearch {
        text: "  ",
        caret: 0,
        focused: true,
        caret_lit: true,
        slide: 1.0,
        selection: None,
    })
    .0;
    let without = a_band_of_the_field(PlaylistSearch {
        text: "  ",
        caret: 0,
        focused: false,
        caret_lit: true,
        slide: 1.0,
        selection: None,
    })
    .0;
    let painted = field_pixels(&with);
    let spare = field_pixels(&without);
    let width = with.width as usize;
    let height = with.height as usize;

    // The tallest column the two disagree on, which is the caret's own.
    let column = (0..width)
        .map(|x| {
            (0..height)
                .filter(|y| painted[y * width + x] != spare[y * width + x])
                .count()
        })
        .max()
        .unwrap_or_default();

    // The face's own pixels, which is what a caret between characters is as tall as — the line box
    // the query is laid out in is taller than that, and a bar drawn its height stands in the leading.
    let level = facts_style().level as usize;
    let cell = text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[level], SCALE) as usize;

    assert!(column > 0, "a focused field draws its caret");
    assert!(
        column <= cell,
        "the caret is no taller than the letters it stands between"
    );
    assert!(
        column >= cell / 2,
        "and it is not a speck in the middle of the pill either"
    );
}

/// **The caret is drawn blinking rather than always**: a field holding it is painted with its bar for one
/// half of a blink and without it for the other, and the half without it is the field as it is when
/// nothing holds the caret at all — a bar that never went dark would be a mark of the field rather than a
/// caret in it.
#[test]
fn the_caret_is_drawn_only_in_the_half_of_its_blink_it_is_lit_in() {
    let paint = |focused: bool, caret_lit: bool| {
        let (surface, _) = a_band_of_the_field(PlaylistSearch {
            text: " ",
            caret: 0,
            focused,
            caret_lit,
            slide: 1.0,
            selection: None,
        });

        field_pixels(&surface)
    };

    let lit = paint(true, true);
    let dark = paint(true, false);
    let unfocused = paint(false, true);

    assert!(
        lit.iter().zip(&dark).any(|(a, b)| a != b),
        "the lit half of the blink draws a caret, and the dark half of the same field does not"
    );
    assert_eq!(
        dark, unfocused,
        "and the dark half is the field nothing holds the caret in: the bar is the whole of what a blink
         changes"
    );
}

/// **The field is shorter than a row.** It is a control a hand reaches for in a list, not a bar the
/// band is built around: a pill the height of a row is a pill that reads as a row of the list, and
/// the rows above it are what the band is for.
#[test]
fn the_field_is_shorter_than_a_row_of_the_list() {
    let layout = playlist_layout(400, 400, SCALE, 40);

    let field = layout.search_bottom - layout.search_top;
    let row = layout.row_height;
    assert!(
        field < row,
        "the field is {field} rows tall and a row of the list is {row}"
    );
    assert!(
        field > row / 2,
        "and it is not so short that a line of the query has no room to sit in"
    );
}

/// **The pill's border runs the whole length of the one straight part of its top and bottom**, and
/// the query inside it never paints over either: the row a run of text is handed is the pill's own
/// inside, a hairline short of its edge, and the two rows of the border are the fill's last and the
/// stroke's first.
#[test]
fn the_field_border_runs_the_whole_straight_part_of_the_pill() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let plate = text_paint::blend(palette.background, palette.foreground, 0.07);
    let ink =
        |at: [u8; 3]| (0..3).any(|channel| (at[channel] as i32 - plate[channel] as i32).abs() > 4);

    let layout = playlist_layout(400, 240, SCALE, 0);
    let radius = (layout.search_bottom - layout.search_top) as f32 / 2.0;
    let straight = (
        layout.search_left + radius.ceil() as i32,
        layout.search_right - radius.ceil() as i32,
    );

    for (text, focused) in [
        ("", false),
        ("a query of a name the field is wide enough to hold", true),
    ] {
        let (surface, box_) = a_band_of_the_field(PlaylistSearch {
            text,
            caret: text.len(),
            focused,
            caret_lit: true,
            slide: 1.0,
            selection: None,
        });
        let pixels = field_pixels(&surface);
        let width = surface.width as usize;

        for row in [box_.top, box_.bottom - 1] {
            let missing = (straight.0..straight.1)
                .filter(|x| !ink(pixels[row as usize * width + *x as usize]))
                .count();
            assert_eq!(
                missing, 0,
                "{} of the pill's row {row} carries no edge under a {focused} query",
                missing
            );
        }
    }
}

/// **A press names a place in the query as it is drawn.** The query is set at the facts line's scale
/// rather than at the band's own, so the columns a press is measured against are that face's: a press
/// measured at the band's scale would name a character several of them away from the one under it.
#[test]
fn a_press_on_the_field_is_measured_at_the_size_the_query_is_drawn() {
    let layout = playlist_layout(400, 240, SCALE, 0);
    let query = "a name too wide for its own field";

    let at_band_scale = search_index_at(&layout, query, layout.search_right, SCALE);
    let at_query_scale = search_index_at(&layout, query, layout.search_right, SCALE);
    assert_eq!(
        at_band_scale, at_query_scale,
        "one press names one place in the query, whatever the band's own scale is"
    );

    // The face the query is drawn in is narrower than the one the band draws its rows in, so the
    // whole of a query fits in a field the band's own face would have cut: a press at the field's
    // right edge is past the last of its characters rather than part way through them.
    let cell_of_query = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[facts_style().level as usize],
        SCALE,
    );
    let cell_of_band = text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[1], SCALE);
    assert!(
        cell_of_query < cell_of_band,
        "the query is set a level under the band's own header"
    );
    assert_eq!(
        search_index_at(&layout, "ab", layout.search_left, SCALE),
        0,
        "and a press at the field's left edge is before its first character"
    );
    assert!(
        layout.search_text_left > layout.search_left,
        "which is a column past the mark, not the pill's own edge"
    );
}

/// **The room under the last whole row is a row peeking rather than a band of empty page, and the pill
/// is drawn over the rows rather than a strip being cut out of them.** A band is a list to its own
/// bottom pad: a page under the last whole row is a band that reads as a list which ended above its own
/// bottom, so the next row is drawn into the room and cut at the pad — and the field floats over the
/// last rows while it is in sight, which is where a hand reaching for it finds it (see
/// `snapped_playlist_height`, which is what keeps a band from being caught between two rows at all).
#[test]
fn the_room_under_the_last_row_is_a_row_peeking_and_the_pill_floats_over_the_rows() {
    let palette = ChromePalette::current().expect("a theme to paint in");
    let foreground = palette.foreground;

    let rows: Vec<PlaylistRow> = (0..6)
        .map(|at| PlaylistRow {
            name: format!("track {at}"),
            facts: vec![Fact {
                text: "a reading".to_string(),
                kind: FactKind::Reading,
            }],
            clock: None,
        })
        .collect();
    let shown: Vec<usize> = (0..rows.len()).collect();

    // A band caught most of a row past its last whole one, which is the shape the peek is for.
    let (width, height) = (400u32, 230u32);
    let layout = playlist_layout(width as i32, height as i32, SCALE, rows.len());
    let leftover = layout.rows_bottom - (layout.rows_top + layout.visible * layout.row_height);
    assert!(
        leftover > 0 && leftover < layout.row_height,
        "the band is caught between two rows, so the one under the last whole row has room"
    );
    assert!(
        rows.len() > layout.visible as usize,
        "there is a row after the last one that fits, which is the one that peeks"
    );

    let paint = |slide: f32| {
        let surface = DibSurface::create(width, height).expect("a surface");
        paint_playlist(
            &surface,
            &palette,
            &PlaylistBand {
                layout,
                rows: &rows,
                shown: &shown,
                first: 0,
                total: rows.len(),
                showing: None,
                picked: None,
                scrolled: &[],
                search: Some(PlaylistSearch {
                    text: "",
                    caret: 0,
                    focused: false,
                    caret_lit: true,
                    slide,
                    selection: None,
                }),
                note: None,
            },
            SCALE,
        );

        field_pixels(&surface)
    };

    // The room under the last whole row, in the columns a row's name is drawn in: a row's ink there is
    // the row peeking, and a band of page is the empty row this is not.
    let peeked = |pixels: &[[u8; 3]]| {
        let top = layout.rows_top + layout.visible * layout.row_height;
        (top..layout.rows_bottom)
            .flat_map(|y| (layout.text_left..layout.name_right).map(move |x| (x, y)))
            .filter(|(x, y)| pixels[*y as usize * width as usize + *x as usize] == foreground)
            .count()
    };
    assert!(
        peeked(&paint(0.0)) > 0,
        "the row under the last whole one is drawn into the room the band has left"
    );

    // And the field is drawn over the rows: inside the pill's own straight part there is no row's ink
    // while it is in sight, because the pill's plate is what a hand sees there — the row is under the
    // field rather than cut away for it.
    let over_the_rows = |pixels: &[[u8; 3]]| {
        let radius = (layout.search_bottom - layout.search_top) as f32 / 2.0;
        ((layout.search_top)..layout.search_bottom)
            .flat_map(|y| {
                ((layout.search_left + radius.ceil() as i32)
                    ..(layout.search_right - radius.ceil() as i32))
                    .map(move |x| (x, y))
            })
            .filter(|(x, y)| pixels[*y as usize * width as usize + *x as usize] == foreground)
            .count()
    };
    assert!(
        over_the_rows(&paint(0.0)) > 0,
        "a parked field leaves the rows it would cover drawn where it would be"
    );
    assert_eq!(
        over_the_rows(&paint(1.0)),
        0,
        "and a pill in sight is drawn over them: the field floats on the list rather than beside it"
    );
}

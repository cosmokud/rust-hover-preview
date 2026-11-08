use super::super::playlist::{playlist_edge_at, playlist_row, playlist_thumb, PlaylistEdge};
use super::*;
use crate::readers::audio_track::{self, Player, Probed, Track};

/// A band of a display's own scale at 96 dpi, which is what every number below is counted in.
const SCALE: f32 = 1.0;

fn a_row(name: &str) -> PlaylistRow {
    PlaylistRow {
        name: name.to_string(),
        facts: "MP3 · 4.2 MB".to_string(),
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
        layout.name_right < layout.clock_left,
        "the name column stops short of the clock's own"
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
    // rows are narrowed by exactly that.
    let scrolling = playlist_layout(400, 240, SCALE, 40);
    let (left, right) = scrolling
        .track
        .expect("a band of forty rows in the room of five scrolls");
    assert_eq!(right, 390, "the track sits in the band's own margin");
    assert!(left < right);
    assert!(scrolling.clock_right < layout.clock_right);
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
    assert_eq!(bottom - top, room * layout.visible / 40);

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
        Probed::Track(Track {
            player: Player::Native,
            codec: Some("FLAC".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(1_006_000),
            duration: Some(3_723.0),
        }),
    );

    let row = playlist_row(&probed);
    assert_eq!(
        row.name, "mixdown",
        "a row names the file without the extension it is already known by"
    );
    assert_eq!(
        row.facts, "FLAC · 44.1 kHz · stereo · 1006 kbps · 3.0 MB",
        "the parts are the card's own, joined with the card's own separator"
    );
    assert_eq!(row.clock.as_deref(), Some("1:02:03"));

    // A file nothing has probed says only what it can be read to say without one: the extension
    // it carries, in the capitals a codec's name is written in, and the room it takes.
    let bare = root.join("unknown.ape");
    std::fs::write(&bare, vec![0u8; 1024 * 1024]).expect("a written file");
    let row = playlist_row(&bare);
    assert_eq!(row.name, "unknown");
    assert_eq!(row.facts, "APE · 1.0 MB");
    assert_eq!(
        row.clock, None,
        "and nothing at all where no length is known"
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
    let (width, height) = (240u32, 140u32);
    let surface = DibSurface::create(width, height).expect("a surface");
    let layout = playlist_layout(width as i32, height as i32, SCALE, 2);
    let rows = [a_row("one"), a_row("two")];

    paint_playlist(
        &surface,
        &palette,
        &PlaylistBand {
            layout,
            rows: &rows,
            first: 0,
            total: 2,
            showing: Some(1),
            hovered: None,
            picked: None,
            name_offset: 0,
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

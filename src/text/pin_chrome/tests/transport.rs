use super::*;

/// The popup the transport bar's own button opens, and the same panel hung from that button
/// given directly, are one panel and not two: a sound's card opens the very same one from a
/// button drawn on its own row, so a level that was placed one way there and another way on
/// the bar would be a level that moves as the file it is playing changes.
#[test]
fn the_popup_hung_from_a_button_is_the_one_the_transport_bar_opens() {
    for (width, height, top) in [(600, 30, 0), (600, 40, 200), (320, 30, 55)] {
        for dpi in [96, 144] {
            let strip = transport_layout(width, height, dpi, true).volume;
            let button = RECT {
                top: strip.top + top,
                bottom: strip.bottom + top,
                ..strip
            };

            let panel = volume_popup_layout(width, top, height, dpi);
            let from_button = volume_popup_from_button(button, width, dpi);
            let (a, b) = (
                RECT {
                    left: panel.panel.left,
                    top: panel.panel.top,
                    right: panel.panel.right,
                    bottom: panel.panel.bottom,
                },
                RECT {
                    left: from_button.panel.left,
                    top: from_button.panel.top,
                    right: from_button.panel.right,
                    bottom: from_button.panel.bottom,
                },
            );
            assert_eq!(a, b, "{width}x{height} at dpi {dpi}");

            let groove = RECT {
                left: panel.track.left,
                top: panel.track.top,
                right: panel.track.right,
                bottom: panel.track.bottom,
            };
            let hung = RECT {
                left: from_button.track.left,
                top: from_button.track.top,
                right: from_button.track.right,
                bottom: from_button.track.bottom,
            };
            assert_eq!(groove, hung, "{width}x{height} at dpi {dpi}");

            // And the panel is still where it belongs: centered on the button it
            // belongs to, above it, and inside the window it belongs to.
            assert!(panel.panel.right <= width && panel.panel.left >= 0);
            // And it floats above the button rather than over it — unless the window is too
            // short for that, which is what the panel's own floor is for: a level that cannot
            // be aimed at is worse than one drawn a little high.
            let panel_height = panel.panel.bottom - panel.panel.top;
            assert!(
                panel.panel.bottom <= button.top || panel.panel.top == 0,
                "the panel is not hung over the button it came out of"
            );
            assert!(
                panel.panel.bottom - panel.panel.top == panel_height,
                "{width}x{height} at dpi {dpi}"
            );
        }
    }
}

#[test]
fn a_clock_is_written_the_way_a_player_writes_one() {
    assert_eq!(clock_text(Some(7.4)), "0:07");
    assert_eq!(clock_text(Some(62.0)), "1:02");
    assert_eq!(clock_text(Some(3723.0)), "1:02:03");

    // A file whose container says nothing, and one whose length is nonsense, are the same
    // answer: a bar with no length to draw a playhead against.
    assert_eq!(clock_text(None), "--:--");
    assert_eq!(clock_text(Some(f64::NAN)), "--:--");
    assert_eq!(clock_text(Some(-1.0)), "--:--");
}

#[test]
fn the_buttons_sit_against_the_right_edge_in_the_order_windows_has_them() {
    let buttons = button_boxes(600, 30, 96, true);

    // The walk comes first in the run and the window's three after it, because the
    // window's are the ones against the right edge and the walk hangs off them.
    assert_eq!(buttons.len(), 8);
    assert_eq!(buttons[0].kind, CaptionButton::Previous);
    assert_eq!(buttons[1].kind, CaptionButton::Next);
    assert_eq!(buttons[2].kind, CaptionButton::OpenWith);
    assert_eq!(buttons[3].kind, CaptionButton::OpenWithList);
    assert_eq!(buttons[4].kind, CaptionButton::AlwaysOnTop);
    assert_eq!(buttons[5].kind, CaptionButton::Minimize);
    assert_eq!(buttons[6].kind, CaptionButton::Maximize);
    assert_eq!(buttons[7].kind, CaptionButton::Close);
    for pair in buttons.windows(2) {
        assert!(pair[0].rect.left < pair[1].rect.left);
    }
    assert_eq!(buttons[7].rect.right, 600);

    // The window's four keep the widths and the places they have always had: 46 pixels
    // each at 100%, packed to 600, which is what a hand has been reaching for on every
    // window it has ever had.
    for (index, left) in [416, 462, 508, 554].into_iter().enumerate() {
        assert_eq!(
            buttons[index + 4].rect,
            RECT {
                left,
                top: 0,
                right: left + 46,
                bottom: 30,
            },
            "the window's button at {left} has moved"
        );
    }

    // And the walk is the same width as the buttons beside it, laid end to end and
    // touching the group rather than leaving a gap in it.
    assert_eq!(buttons[3].rect.right, buttons[4].rect.left);
    for button in &buttons {
        assert_eq!(
            button.rect.right - button.rect.left,
            46,
            "a caption's buttons are one size"
        );
    }

    // And a point is on the button it looks like it is on, or on none of them.
    let close = buttons[7].rect;
    assert_eq!(
        button_at(close.left + 1, 5, 600, 30, 96, true),
        Some(CaptionButton::Close)
    );
    assert_eq!(button_at(1, 5, 600, 30, 96, true), None);
    assert_eq!(button_at(close.left + 1, 40, 600, 30, 96, true), None);
}

#[test]
fn a_transport_bar_is_taken_hold_of_where_it_was_pressed() {
    let layout = transport_layout(800, 30, 96, true);
    let middle = (layout.bar.left + layout.bar.right) / 2;
    let share = transport_share_at(middle, 800, 30, 96, true);
    assert!(
        (share - 0.5).abs() < 0.05,
        "the middle of the bar is half of it"
    );

    // A press past either end is the end it is past, which is what keeps a drag from asking
    // for a second of a file that is not there.
    assert_eq!(transport_share_at(0, 800, 30, 96, true), 0.0);
    assert_eq!(transport_share_at(800, 800, 30, 96, true), 1.0);
}

/// A thumb is drawn where the bar *draws* the playhead, and a press is read against the bar the
/// *bar itself* is: the two layouts are the same arithmetic, so the share a press is turned into
/// and the pixel the thumb lands on are one number read twice. A layout built from a different
/// height than the one the bar is drawn at is not that, and it fails at the two ends and nowhere
/// else — the middle of a wrongly-laid-out bar is the middle of the right one by accident, which
/// is exactly why a bar that is right in the middle reads as right.
#[test]
fn a_press_puts_the_thumb_under_the_hand_wherever_it_landed() {
    for (width, height, dpi) in [(800, 30, 96), (640, 40, 144), (480, 30, 96)] {
        let layout = transport_layout(width, height, dpi, true);
        let span = (layout.bar.right - layout.bar.left) as f64;

        for x in layout.bar.left..layout.bar.right {
            let share = transport_share_at(x, width, height, dpi, true);
            // Where `paint_transport` draws the thumb for that share.
            let thumb = layout.bar.left + (span * share).round() as i32;

            assert_eq!(
                thumb, x,
                "at {width}x{height} dpi {dpi}, a press at {x} of [{}, {}) drew the thumb at {thumb}",
                layout.bar.left, layout.bar.right,
            );
        }
    }
}

/// The volume button is the last thing on the bar and it answers a press whatever the player
/// behind the bar can be told: a level is this app's own, and FFmpeg's player takes one as well
/// as the media engine does — by being started again at it.
#[test]
fn the_volume_button_answers_on_a_bar_with_no_other_controls() {
    let live = transport_layout(800, 30, 96, true);
    let readout = transport_layout(800, 30, 96, false);

    // It keeps its own box at the right edge of the strip, and it is the same box whether or
    // not the bar carries a play button: a control that moved because the engine changed would
    // be one a hand has to look for.
    assert_eq!(live.volume, readout.volume);
    assert_eq!(
        live.volume.right,
        800 - 10,
        "against the strip's own padding"
    );
    assert!(
        live.volume.left > live.total.right,
        "the clocks end before it"
    );

    let x = (live.volume.left + live.volume.right) / 2;
    let y = (live.volume.top + live.volume.bottom) / 2;
    assert_eq!(
        transport_part_at(x, y, 800, 30, 96, true),
        Some(TransportPart::Volume)
    );
    assert_eq!(
        transport_part_at(x, y, 800, 30, 96, false),
        Some(TransportPart::Volume),
        "and on a bar whose player cannot be told anything"
    );

    // What the button did not take is not the track's: the room it occupies comes off the end
    // the clocks were drawn at.
    assert!(readout.total.right <= readout.volume.left);
    assert!(readout.bar.right <= readout.total.left);
}

/// The popup hangs over the button that opened it, inside the window it belongs to, and its
/// groove is what the level is measured against: the bottom of it is nothing and the top of it
/// is everything.
#[test]
fn a_volume_popup_is_placed_over_its_button_and_read_bottom_up() {
    let (width, strip_top, strip_height) = (800, 400, 30);
    let popup = volume_popup_layout(width, strip_top, strip_height, 96);
    let button = transport_layout(width, strip_height, 96, true).volume;

    // Above the bar, centered on the button that opened it, and clear of the button itself.
    assert_eq!(
        (popup.panel.left + popup.panel.right) / 2,
        (button.left + button.right) / 2,
        "the panel is centered on the button it was opened from"
    );
    assert_eq!(
        popup.panel.bottom,
        strip_top + button.top - VOLUME_PANEL_GAP
    );
    assert!(popup.panel.top >= 0);
    assert_eq!(
        popup.panel.right - popup.panel.left,
        VOLUME_PANEL_WIDTH,
        "the panel is the width a level is drawn in"
    );

    // The groove is inside the panel and centered in it, with the room the knob takes at either
    // end: a level of everything is a knob that is still whole. The room is the knob's own
    // radius, and it is asked for with the collar the knob is drawn with rather than tightly.
    assert!(popup.track.left > popup.panel.left && popup.track.right < popup.panel.right);
    assert_eq!(
        (popup.track.left + popup.track.right) / 2,
        (popup.panel.left + popup.panel.right) / 2
    );

    let collar = (VOLUME_THUMB_RADIUS + VOLUME_COLLAR_PIXELS).ceil() as i32;
    let ends = popup.track.top - popup.panel.top;
    assert!(
        ends >= collar && popup.panel.bottom - popup.track.bottom >= collar,
        "the groove is held off each end of the panel by the knob and its collar"
    );

    // And the panel is kept close around the knob: it is a strip of glass over somebody else's
    // picture, so what it is wider than the knob and its collar by is a couple of pixels and no
    // more, at either side and at both ends.
    let beside = (popup.panel.right - popup.panel.left - collar * 2) / 2;
    assert!(
        (1..=4).contains(&beside) && ends <= collar + 8,
        "the panel hugs the knob: {beside} beside it, {ends} past its ends"
    );

    // A point on the groove is the share of the level it is at, and a point past either end is
    // the end it is past.
    assert_eq!(volume_share_at(popup.track.bottom, popup.track), 0.0);
    assert_eq!(volume_share_at(popup.track.top, popup.track), 1.0);
    let middle = (popup.track.top + popup.track.bottom) / 2;
    assert!((volume_share_at(middle, popup.track) - 0.5).abs() < 0.02);
    assert_eq!(volume_share_at(0, popup.track), 1.0);
    assert_eq!(volume_share_at(strip_top + strip_height, popup.track), 0.0);

    // And what is drawn at a level is drawn where it was read from: the knob of 100% is at the
    // top of the groove, the knob of nothing is at the bottom, and the two are inside it.
    assert_eq!(volume_thumb_row(popup.track, 100), popup.track.top);
    assert_eq!(volume_thumb_row(popup.track, 0), popup.track.bottom);
    assert!(volume_thumb_row(popup.track, 50) < volume_thumb_row(popup.track, 25));
}

/// The panel a volume button opens is centered on that button's own middle, at every scale
/// the display runs at and every width the window is given: a hand aims at the button, and
/// the level answers from the middle of it rather than off one edge of it.
#[test]
fn a_volume_popup_is_centered_on_its_button() {
    for (width, height, strip_top, dpi) in [
        (800, 30, 400, 96),
        (600, 40, 200, 144),
        (320, 30, 55, 96),
        (1920, 40, 100, 144),
    ] {
        let strip = transport_layout(width, height, dpi, true).volume;
        let button = RECT {
            top: strip.top + strip_top,
            bottom: strip.bottom + strip_top,
            ..strip
        };

        let popup = volume_popup_layout(width, strip_top, height, dpi);

        // The panel's middle is the button's middle, and the whole of the panel is
        // inside the window it belongs to.
        assert_eq!(
            (popup.panel.left + popup.panel.right) / 2,
            (button.left + button.right) / 2,
            "{width}x{height} at dpi {dpi}"
        );
        assert!(
            popup.panel.left >= 0 && popup.panel.right <= width,
            "the panel is inside the window at {width}x{height} dpi {dpi}"
        );
    }

    // And the same centering for a button drawn somewhere else — a sound's card is
    // one such place — not only the transport bar's.
    for (button, width, dpi) in [
        (
            RECT {
                left: 100,
                top: 300,
                right: 130,
                bottom: 330,
            },
            800,
            96,
        ),
        (
            RECT {
                left: 40,
                top: 10,
                right: 88,
                bottom: 58,
            },
            200,
            144,
        ),
    ] {
        let popup = volume_popup_from_button(button, width, dpi);

        assert_eq!(
            (popup.panel.left + popup.panel.right) / 2,
            (button.left + button.right) / 2,
            "{width} wide at dpi {dpi}"
        );
        assert!(popup.panel.left >= 0 && popup.panel.right <= width);
    }
}

/// A window narrow beside the panel is not a popup drawn off the side of it: the panel is
/// kept inside the window it belongs to, centered on the button when the window is wide
/// enough and held at its edge when it is not, where the hand that opened it can still
/// reach it.
#[test]
fn a_volume_popup_is_kept_inside_a_narrow_window() {
    let width = 30;
    let popup = volume_popup_layout(width, 100, 30, 96);

    assert!(popup.panel.left >= 0 && popup.panel.right <= width);
    assert!(popup.panel.right > popup.panel.left);
    assert!(popup.track.left >= popup.panel.left);
    assert!(popup.track.bottom > popup.track.top);
}

#[test]
fn the_parts_of_a_transport_bar_are_hit_where_they_are_drawn() {
    let layout = transport_layout(800, 30, 96, true);

    let play = (layout.play.left + layout.play.right) / 2;
    assert_eq!(
        transport_part_at(play, 15, 800, 30, 96, true),
        Some(TransportPart::Play)
    );

    let bar = (layout.bar.left + layout.bar.right) / 2;
    assert_eq!(
        transport_part_at(bar, 15, 800, 30, 96, true),
        Some(TransportPart::Seek)
    );

    // Nothing is hit where nothing is drawn.
    assert_eq!(transport_part_at(2, 2, 800, 30, 96, true), None);
}

/// A bar whose player cannot be told anything is a read-out: FFmpeg's player reports no
/// position, takes no pause, and can only be "seeked" by being ended and begun again at a
/// second — so its bar carries no button and nothing to drag, and the two clocks begin where
/// the strip does rather than after a control that is not there.
#[test]
fn a_bar_with_no_controls_has_none_to_press() {
    let dead = transport_layout(800, 30, 96, false);
    let live = transport_layout(800, 30, 96, true);

    assert_eq!(dead.elapsed.left, 10, "the clock begins at the padding");
    assert_eq!(
        live.elapsed.left,
        dead.elapsed.left + (live.play.right - live.play.left) + 5,
        "a bar with a button begins after it"
    );

    // The track is longer by the room the button does not take, so a playhead drawn against
    // it is drawn against the same file rather than against a bar with a hole in it.
    assert!(dead.bar.left < live.bar.left);
    assert!(
        dead.bar.right - dead.bar.left > live.bar.right - live.bar.left,
        "the space a button does not take is the track's"
    );

    // And neither the button's place nor the track answers a press.
    let play = (live.play.left + live.play.right) / 2;
    let bar = (dead.bar.left + dead.bar.right) / 2;
    assert_eq!(transport_part_at(play, 15, 800, 30, 96, false), None);
    assert_eq!(transport_part_at(bar, 15, 800, 30, 96, false), None);
    assert_eq!(
        transport_part_at(play, 15, 800, 30, 96, true),
        Some(TransportPart::Play)
    );
}

/// The play mark and the two step marks are the marks a bar draws with an edge that runs across
/// the pixel grid: the pixels their slanted edges cross are some of the ink and some of what is
/// behind it, and a painter that gives each of them the one or the other is a staircase — which
/// is the one thing a triangle nine pixels across is good at showing (see `draw_play_glyph` and
/// `draw_track_step`).
#[test]
fn the_marks_with_a_slanted_edge_carry_a_share_of_the_ink() {
    const W: i32 = 32;
    const H: i32 = 24;
    // The ink is black and the button under it is white, so the byte a pixel keeps says which
    // share of the mark is in it: 0 is all of it, 255 is none of it, and a share is neither.
    let ink = [0u8, 0, 0];

    let mut play = vec![255u8; (W * H * 4) as usize];
    draw_play_glyph(&mut play, W, 16, 12, 9, ink);
    let mut on = vec![255u8; (W * H * 4) as usize];
    draw_track_step(&mut on, W, 16, 12, 14, ink, true);
    let mut back = vec![255u8; (W * H * 4) as usize];
    draw_track_step(&mut back, W, 16, 12, 14, ink, false);

    for (name, buffer) in [("play", play), ("the on step", on), ("the back step", back)] {
        let partial = buffer
            .as_chunks::<4>()
            .0
            .iter()
            .filter(|pixel| pixel[0] > 0 && pixel[0] < 255)
            .count();

        assert!(
            partial > 10,
            "{name} carries part of the ink in {partial} pixels only: a pixel an edge crosses \
             holds the area of the mark that is in it"
        );
    }
}

/// The four marks a bar's own button and a card's row carry — a play, a pause, a step each way —
/// at both scales, on the bar's own colour: the design under review.
///
/// A row of marks is judged as a row: the play mark is read against the pause mark it alternates
/// with and against the two steps beside it, and those three are the ones with an edge that runs
/// across the pixel grid (see `draw_play_glyph`, `draw_track_step`). One of them drawn a whole
/// column at a time, or drawn a size that is not its own, is what a row shows and a test on the
/// ink of one mark at a time cannot.
#[test]
fn draws_the_marks_of_a_bar() {
    let dir = scratch("transport-marks");
    // The strip's own colour and ink, as a dark theme's are: the marks are the theme's foreground
    // on the strip the bar paints (`paint_transport`).
    let background = [24u8, 26, 32];
    let ink = [240u8, 240, 240];

    for (dpi, tag) in [(96u32, "96dpi"), (192u32, "192dpi")] {
        let scale = dpi as f32 / 96.0;
        // The room the bar gives a button, which is the size the marks are measured in: a mark
        // drawn at a size of its own would be a mark that moves beside the ones here. The strip is
        // the one the bar is given at this scale, so that the marks hold the room they hold on the
        // bar rather than all of it.
        let side = {
            let layout = transport_layout(
                text_paint::scaled(600, scale),
                text_paint::scaled(30, scale),
                dpi,
                true,
            );
            layout.play.bottom - layout.play.top
        };

        let marks = ["play", "pause", "back", "on"];
        let width = side * marks.len() as i32;
        let mut strip = vec![0u8; (width * side * 4) as usize];
        for pixel in strip.as_chunks_mut::<4>().0 {
            pixel[0] = background[2];
            pixel[1] = background[1];
            pixel[2] = background[0];
            pixel[3] = 255;
        }

        for (cell, mark) in marks.into_iter().enumerate() {
            let center = cell as i32 * side + side / 2;
            let middle = side / 2;
            match mark {
                "play" => draw_play_glyph(
                    &mut strip,
                    width,
                    center,
                    middle,
                    size_in(scale, side, 9),
                    ink,
                ),
                "pause" => draw_pause_glyph(
                    &mut strip,
                    width,
                    center,
                    middle,
                    size_in(scale, side, 9),
                    ink,
                ),
                "back" => draw_track_step(
                    &mut strip,
                    width,
                    center,
                    middle,
                    size_in(scale, side, 12),
                    ink,
                    false,
                ),
                _ => draw_track_step(
                    &mut strip,
                    width,
                    center,
                    middle,
                    size_in(scale, side, 12),
                    ink,
                    true,
                ),
            }
        }

        let path = dir.join(format!("marks-{tag}.png"));
        write_png(path.clone(), &strip, width as u32, side as u32);
        println!("{}", path.display());

        if dpi == 96 {
            let (large, wide, tall) = zoom(&strip, width as u32, side as u32, 4);
            let path = dir.join(format!("marks-{tag}-x4.png"));
            write_png(path.clone(), &large, wide, tall);
            println!("{}", path.display());
        }
    }
}

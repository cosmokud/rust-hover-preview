use super::*;

/// The popup the transport bar's own button opens, and the same panel hung from that button
/// given directly, are one panel and not two: a sound's card opens the very same one from a
/// button drawn on its own row, so a level that was placed one way there and another way on
/// the bar would be a level that moves as the file it is playing changes.
#[test]
fn the_popup_hung_from_a_button_is_the_one_the_transport_bar_opens() {
    for (width, height, top) in [(600, 30, 0), (600, 40, 200), (320, 30, 55)] {
        for dpi in [96, 144] {
            let strip = transport_layout(width, height, dpi, true, false).volume;
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
    // window's are the ones against the right edge and the walk hangs off them. The detach
    // stands at the seam: the road out of being a pin, packed with the window's own.
    assert_eq!(buttons.len(), 9);
    assert_eq!(buttons[0].kind, CaptionButton::Previous);
    assert_eq!(buttons[1].kind, CaptionButton::Next);
    assert_eq!(buttons[2].kind, CaptionButton::OpenWith);
    assert_eq!(buttons[3].kind, CaptionButton::OpenWithList);
    assert_eq!(buttons[4].kind, CaptionButton::AlwaysOnTop);
    assert_eq!(buttons[5].kind, CaptionButton::Detach);
    assert_eq!(buttons[6].kind, CaptionButton::Minimize);
    assert_eq!(buttons[7].kind, CaptionButton::Maximize);
    assert_eq!(buttons[8].kind, CaptionButton::Close);
    for pair in buttons.windows(2) {
        assert!(pair[0].rect.left < pair[1].rect.left);
    }
    assert_eq!(buttons[8].rect.right, 600);

    // The window's three keep the widths and the places they have always had: 46 pixels
    // each at 100%, packed to 600, which is what a hand has been reaching for on every
    // window it has ever had — and the detach takes the slot the always-on-top gave up when
    // it stopped being the leftmost of the window's own.
    for (index, left) in [416, 462, 508, 554].into_iter().enumerate() {
        assert_eq!(
            buttons[index + 5].rect,
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
    let close = buttons[8].rect;
    assert_eq!(
        button_at(close.left + 1, 5, 600, 30, 96, true),
        Some(CaptionButton::Close)
    );
    assert_eq!(button_at(1, 5, 600, 30, 96, true), None);
    assert_eq!(button_at(close.left + 1, 40, 600, 30, 96, true), None);
}

#[test]
fn a_transport_bar_is_taken_hold_of_where_it_was_pressed() {
    let layout = transport_layout(800, 30, 96, true, false);
    let middle = (layout.bar.left + layout.bar.right) / 2;
    let share = transport_share_at(middle, 800, 30, 96, true, false);
    assert!(
        (share - 0.5).abs() < 0.05,
        "the middle of the bar is half of it"
    );

    // A press past either end is the end it is past, which is what keeps a drag from asking
    // for a second of a file that is not there.
    assert_eq!(transport_share_at(0, 800, 30, 96, true, false), 0.0);
    assert_eq!(transport_share_at(800, 800, 30, 96, true, false), 1.0);
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
        let layout = transport_layout(width, height, dpi, true, false);
        let span = (layout.bar.right - layout.bar.left - 1).max(1);

        for x in layout.bar.left..layout.bar.right {
            let share = transport_share_at(x, width, height, dpi, true, false);
            // Where `paint_transport` draws the thumb for that share.
            let thumb = layout.bar.left + (f64::from(span) * share).round() as i32;

            assert_eq!(
                thumb, x,
                "at {width}x{height} dpi {dpi}, a press at {x} of [{}, {}) drew the thumb at {thumb}",
                layout.bar.left, layout.bar.right,
            );
        }

        // And the last column of the track is the end of the file rather than a step short of it:
        // a press there is the one place a hand means the file's last frame.
        assert_eq!(
            transport_share_at(layout.bar.right - 1, width, height, dpi, true, false),
            1.0,
            "the track's own last column is the end at {width}x{height} dpi {dpi}"
        );
    }
}

/// The volume button is the last thing on the bar and it answers a press whatever the player
/// behind the bar can be told: a level is this app's own, and FFmpeg's player takes one as well
/// as the media engine does — by being started again at it.
#[test]
fn the_volume_button_answers_on_a_bar_with_no_other_controls() {
    let live = transport_layout(800, 30, 96, true, false);
    let readout = transport_layout(800, 30, 96, false, false);

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
        transport_part_at(x, y, 800, 30, 96, true, false),
        Some(TransportPart::Volume)
    );
    assert_eq!(
        transport_part_at(x, y, 800, 30, 96, false, false),
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
    let button = transport_layout(width, strip_height, 96, true, false).volume;

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
        let strip = transport_layout(width, height, dpi, true, false).volume;
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
    let layout = transport_layout(800, 30, 96, true, false);

    let play = (layout.play.left + layout.play.right) / 2;
    assert_eq!(
        transport_part_at(play, 15, 800, 30, 96, true, false),
        Some(TransportPart::Play)
    );

    let bar = (layout.bar.left + layout.bar.right) / 2;
    assert_eq!(
        transport_part_at(bar, 15, 800, 30, 96, true, false),
        Some(TransportPart::Seek)
    );

    // Nothing is hit where nothing is drawn.
    assert_eq!(transport_part_at(2, 2, 800, 30, 96, true, false), None);
}

/// A bar whose player cannot be told anything is a read-out: FFmpeg's player reports no
/// position, takes no pause, and can only be "seeked" by being ended and begun again at a
/// second — so its bar carries no button and nothing to drag, and the two clocks begin where
/// the strip does rather than after a control that is not there.
#[test]
fn a_bar_with_no_controls_has_none_to_press() {
    let dead = transport_layout(800, 30, 96, false, false);
    let live = transport_layout(800, 30, 96, true, false);

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
    assert_eq!(transport_part_at(play, 15, 800, 30, 96, false, false), None);
    assert_eq!(transport_part_at(bar, 15, 800, 30, 96, false, false), None);
    assert_eq!(
        transport_part_at(play, 15, 800, 30, 96, true, false),
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
                false,
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

/// The clock beside the track is written against it, and the room a picture that plays keeps for
/// the count of frames under a hand is left of the clock rather than between the clock and the
/// track: a clock laid out from the strip's own padding is a clock a bar's worth of nothing away
/// from the read-out it belongs to, which is the gap this app used to open there.
#[test]
fn the_clock_is_written_against_the_track_and_the_count_beside_it() {
    let picture = transport_layout(800, 30, 96, false, true);

    // What stands against the track is the clock's own box: the count's room is outside it, and
    // one gap is all that is held between the clock and the bar it belongs to.
    assert_eq!(picture.elapsed.right + strip_gap(1.0), picture.bar.left);
    assert_eq!(picture.frames.right, picture.elapsed.left);
    assert!(picture.frames.left < picture.frames.right);

    // And the room holds the widest count a strip is asked to write - a file of a thousand frames
    // and more, under a hand near its end - at every scale a strip is drawn at: a count written
    // into a room too narrow for it is a count clipped to `[10...`.
    for dpi in [96, 144, 192] {
        let scale = dpi as f32 / 96.0;
        let layout = transport_layout(text_paint::scaled(800, scale), 30, dpi, false, true);
        let widest = measure_text(
            &DibSurface::create(320, 40).expect("a surface to measure on"),
            &caption_style([255, 255, 255]),
            "[1021/1021]",
            scale,
        );
        assert!(
            widest > 0 && widest <= layout.frames.right - layout.frames.left,
            "at dpi {dpi} the count takes {widest} of the {} columns it is given",
            layout.frames.right - layout.frames.left
        );
    }

    // And a strip too narrow to hold both the count and a track worth dragging keeps the track:
    // a count is a read-out of a hand on the bar, and a bar of nothing is not a bar.
    let narrow = transport_layout(320, 30, 96, false, true);
    assert_eq!(narrow.frames.left, narrow.frames.right);
    assert_eq!(narrow.elapsed.right + strip_gap(1.0), narrow.bar.left);

    // And a film's bar keeps none of it, and writes its clock against the track the same way.
    let film = transport_layout(800, 30, 96, true, false);
    assert_eq!(film.frames.left, film.frames.right);
    assert_eq!(film.elapsed.right + strip_gap(1.0), film.bar.left);
}

/// The pace's button is the far end of the bar a picture that plays carries, and the level's is the
/// far end of every other bar: a file of frames has no sound for a level to be a level of, so the
/// two are never on one strip — a second control beside a picture that plays would be a button that
/// does nothing (see `TransportPart::Volume`).
#[test]
fn a_picture_that_plays_carries_one_control_at_the_far_end_and_never_two() {
    for (width, height) in [(320, 30), (600, 30), (600, 40), (1920, 48), (220, 26)] {
        for dpi in [96, 144, 192] {
            let picture = transport_layout(width, height, dpi, true, true);
            let film = transport_layout(width, height, dpi, true, false);

            // One square, at the strip's own right edge, and the same square the level's button is
            // on a film's bar: a hand finds the far control where it always was.
            assert_eq!(
                picture.fps, film.volume,
                "at {width}x{height} dpi {dpi} the pace is not where the level was"
            );
            assert_eq!(
                picture.volume.left, picture.volume.right,
                "at {width}x{height} dpi {dpi} a picture that plays drew a level's button"
            );
            assert_eq!(
                film.fps.left, film.fps.right,
                "at {width}x{height} dpi {dpi} a film drew a pace's button"
            );
            assert!(
                picture.fps.left >= 0 && picture.fps.right <= width,
                "at {width}x{height} dpi {dpi} the pace's button is off the bar"
            );

            // And nothing on a picture's bar is the level: every column of it, and every row wide
            // enough to hold the square, is answered by something else or by nothing at all.
            for x in 0..width {
                for y in [0, height / 2, height - 1] {
                    assert_ne!(
                        transport_part_at(x, y, width, height, dpi, true, true),
                        Some(TransportPart::Volume),
                        "at {width}x{height} dpi {dpi} a column of a picture's bar is a level"
                    );
                }
            }
        }
    }
}

/// A point on the FPS button is answered before anyone else's, and a bar drawn without one has
/// nothing of the pace's at that point at all.
#[test]
fn the_fps_button_is_hit_where_it_is_drawn() {
    let layout = transport_layout(800, 30, 96, true, true);
    let middle = (layout.fps.left + layout.fps.right) / 2;

    assert_eq!(
        transport_part_at(middle, 15, 800, 30, 96, true, true),
        Some(TransportPart::Fps)
    );
    // A film's bar has no pace to set, and the same column of it is the level's button - one
    // control at the far end, whichever bar it is on.
    assert_eq!(
        transport_part_at(middle, 15, 800, 30, 96, true, false),
        Some(TransportPart::Volume),
        "a film's bar has a level where a picture's has a pace"
    );

    // And the track stops short of the button, so a press near the end of the strip is a control
    // rather than a seek.
    assert_ne!(
        transport_part_at(layout.fps.left, 15, 800, 30, 96, true, true),
        Some(TransportPart::Seek)
    );
}

/// The pace's slider reads one frame a second at the bottom of its groove and the ceiling at the
/// top, and every whole frame in between: a groove whose ends were not those two would be a slider
/// a hand cannot reach either end of.
#[test]
fn the_pace_slider_reads_one_frame_a_second_up_to_the_ceiling() {
    let layout = transport_layout(800, 30, 96, true, true);
    let popup = fps_popup_from_button(
        RECT {
            top: layout.fps.top + 400,
            bottom: layout.fps.bottom + 400,
            ..layout.fps
        },
        800,
        96,
    );
    let track = popup.track;
    // The top this app's own setting names, which is the ceiling the slider is read against
    // rather than a second one of its own (see `animated_fps_max`).
    let max = DEFAULT_ANIMATED_FPS_MAX;

    assert_eq!(max, 180, "the pace this app opens at is a setting");
    assert_eq!(fps_at(track.bottom, track, max), 1, "the bottom is one");
    assert_eq!(
        fps_at(track.top, track, max),
        max,
        "and the top is all of it"
    );
    let quarter = fps_at(track.bottom - (track.bottom - track.top) / 4, track, max);
    assert!(
        (43..=49).contains(&quarter),
        "a quarter of the way up is a quarter of the pace above one, and not all of it: {quarter}"
    );

    // And a point off the groove is held at its ends rather than read past them.
    assert_eq!(fps_at(track.bottom + 500, track, max), 1);
    assert_eq!(fps_at(track.top - 500, track, max), max);

    // The knob is drawn where the pace it holds is: the two ends of the groove again, and a pace
    // in between the row between them.
    assert_eq!(fps_thumb_row(track, 1, max), track.bottom);
    assert_eq!(fps_thumb_row(track, max, max), track.top);
    assert!(fps_thumb_row(track, max / 2, max) < track.bottom);
    assert!(fps_thumb_row(track, max / 2, max) > track.top);

    // The panel is the level's own shape widened by what it has to hold: a number written across
    // its top is a figure read in a panel of its own width, and the two are one panel otherwise
    // (see `FPS_PANEL_WIDTH` and `fps_popup_text_rect`).
    let volume = volume_popup_from_button(
        RECT {
            top: layout.fps.top + 400,
            bottom: layout.fps.bottom + 400,
            ..layout.fps
        },
        800,
        96,
    );
    assert!(popup.panel.right - popup.panel.left > volume.panel.right - volume.panel.left);
    let text = fps_popup_text_rect(&popup, 96);
    assert!(
        text.right - text.left
            >= measure_text(
                &DibSurface::create(40, 20).expect("a surface to measure on"),
                &caption_style([255, 255, 255]),
                "180",
                1.0,
            ),
        "the room the pace is written in holds the figure the slider can be taken to"
    );
}

/// The pace is written over its own panel in blue and with nothing of the page behind it: the
/// figures are ink on a flat panel, and the panel's own alpha byte is handed back under them (see
/// `paint_fps_popup_text`).
///
/// It is judged on the strip a layered window carries rather than on the figure alone, because the
/// two faults it is written against are faults of what is *under* the figure: a run whose
/// background came from the caption's page instead of the panel is a patch of another colour around
/// a number, and a run that left GDI's zero alpha behind it is a number cut out of the panel it is
/// read on.
#[test]
fn the_pace_is_written_in_blue_with_no_panel_of_its_own_behind_it() {
    let shift = |rect: RECT, by: (i32, i32)| RECT {
        left: rect.left - by.0,
        top: rect.top - by.1,
        right: rect.right - by.0,
        bottom: rect.bottom - by.1,
    };
    let button = RECT {
        left: 600,
        top: 500,
        right: 622,
        bottom: 522,
    };
    let (width, height, text, strip_width, popup) = {
        let popup = fps_popup_from_button(button, 1200, 96);
        let bounds = volume_popup_bounds(&popup, 1200, 700);
        let by = (bounds.left, bounds.top);
        (
            bounds.right - bounds.left,
            bounds.bottom - bounds.top,
            shift(fps_popup_text_rect(&popup, 96), by),
            (bounds.right - bounds.left).max(1),
            VolumePopup {
                panel: shift(popup.panel, by),
                track: shift(popup.track, by),
            },
        )
    };

    // The strip a pace's panel is painted into, with or without the figures written over it, so
    // that what the figures did to the panel can be told apart from the panel itself.
    let painted = |figures: bool| {
        let palette = ChromePalette {
            background: [24, 26, 32],
            foreground: [240, 240, 240],
            accent: [86, 182, 194],
            blue: [97, 175, 239],
            number: [209, 154, 102],
            dark: true,
        };
        let surface = DibSurface::create(width as u32, height as u32)
            .expect("a strip to paint a pace's panel into");
        let len = (width.max(1) * height.max(1) * 4) as usize;
        {
            let pixels = unsafe { std::slice::from_raw_parts_mut(surface.bits(), len) };
            pixels.fill(0);
            paint_fps_popup(pixels, strip_width, &palette, &popup, 180, 180, false);
        }

        // The panel's own alpha byte under the figures, before they are written over it.
        let under = alphas(surface.bits(), strip_width, text, height);
        if figures {
            paint_fps_popup_text(&surface, &palette, &popup, 180, 96);
        }
        let pixels = unsafe { std::slice::from_raw_parts(surface.bits(), len) }.to_vec();

        (pixels, under)
    };

    let (bare, _) = painted(false);
    let (page, under) = painted(true);

    // **Nothing of the run is a panel of its own.** What it writes is the figure and the figure
    // only: the pixels it leaves alone are still the panel's, and none of them is outside the room
    // the pace is written in — a run whose background came from the page would be a rectangle of
    // another colour around a number, and every pixel of its box would have changed.
    let box_ = |x: i32, y: i32| ((y * strip_width + x) * 4) as usize;
    let mut changed = 0;
    for y in 0..height {
        for x in 0..strip_width {
            let at = box_(x, y);
            if page[at..at + 3] == bare[at..at + 3] {
                continue;
            }
            assert!(
                x >= text.left && x < text.right && y >= text.top && y < text.bottom,
                "the run painted outside the room the pace is written in, at ({x}, {y})"
            );
            changed += 1;
        }
    }
    let room = ((text.right - text.left).max(0) * (text.bottom - text.top).max(0)) as usize;
    assert!(changed > 0, "the pace was not written at all");
    assert!(
        changed < room,
        "the run filled its box with a background of its own: {changed} of {room} pixels changed"
    );

    // And the panel's alpha is handed back under the figures rather than left at the zero GDI draws
    // through: a strip carries per-pixel alpha, so a run that sealed it would be a hole in the panel.
    assert_eq!(
        under,
        alphas(page.as_ptr(), strip_width, text, height),
        "the run hands the panel's own alpha back"
    );

    // And the figure is the theme's blue: the bluest pixel of the room leans to it by a wide margin
    // over the panel it is written on, which no run in the panel's own colour could be.
    let bluest = (text.top.max(0)..text.bottom.min(height))
        .flat_map(|y| (text.left.max(0)..text.right.min(strip_width)).map(move |x| (x, y)))
        .map(|(x, y)| page[box_(x, y)..][..4].to_vec())
        .max_by_key(|pixel| i32::from(pixel[0]) - i32::from(pixel[2]))
        .expect("the room the pace is written in");
    assert!(
        i32::from(bluest[0]) - i32::from(bluest[2]) > 60,
        "the bluest pixel of the pace is not blue: {bluest:?}"
    );
}

/// The alpha byte of every pixel in `rect`, row by row, off a strip's own premultiplied BGRA.
fn alphas(bits: *const u8, width: i32, rect: RECT, height: i32) -> Vec<u8> {
    let pixels =
        unsafe { std::slice::from_raw_parts(bits, (width.max(1) * height.max(1) * 4) as usize) };
    let mut found = Vec::new();
    for y in rect.top.max(0)..rect.bottom.min(height) {
        for x in rect.left.max(0)..rect.right.min(width) {
            found.push(pixels[((y * width + x) * 4) as usize + 3]);
        }
    }
    found
}

/// The pace's panel is the volume panel, hung from a button of its own: inside the window, over the
/// media rather than over the button that opened it, and with a groove to drag.
#[test]
fn the_pace_popup_is_inside_the_window_and_floats_above_its_button() {
    for (width, height) in [(600, 30), (320, 30), (1920, 48)] {
        for dpi in [96, 144] {
            let layout = transport_layout(width, height, dpi, true, true);
            let button = RECT {
                top: layout.fps.top + height,
                bottom: layout.fps.bottom + height,
                ..layout.fps
            };
            let popup = fps_popup_from_button(button, width, dpi);

            assert!(popup.panel.left >= 0 && popup.panel.right <= width);
            // The panel floats over the media rather than over the button that opened it — unless
            // the window is too short to hold it above the button, which is where the panel's own
            // floor puts it at the window's top edge instead (see `volume_popup_from_button`).
            assert!(
                popup.panel.bottom <= button.top || popup.panel.top == 0,
                "the panel covers the button that opened it"
            );
            assert!(popup.track.bottom > popup.track.top, "a groove to drag");
        }
    }
}

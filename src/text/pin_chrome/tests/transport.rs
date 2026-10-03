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

            // And the panel is still where it belongs: hung from the button's own right edge,
            // above it, and inside the window it belongs to.
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
    assert_eq!(buttons.len(), 7);
    assert_eq!(buttons[0].kind, CaptionButton::Previous);
    assert_eq!(buttons[1].kind, CaptionButton::Next);
    assert_eq!(buttons[2].kind, CaptionButton::OpenWith);
    assert_eq!(buttons[3].kind, CaptionButton::OpenWithList);
    assert_eq!(buttons[4].kind, CaptionButton::Minimize);
    assert_eq!(buttons[5].kind, CaptionButton::Maximize);
    assert_eq!(buttons[6].kind, CaptionButton::Close);
    for pair in buttons.windows(2) {
        assert!(pair[0].rect.left < pair[1].rect.left);
    }
    assert_eq!(buttons[6].rect.right, 600);

    // The window's three keep the widths and the places they have always had: 46 pixels
    // each at 100%, packed to 600, which is what a hand has been reaching for on every
    // window it has ever had.
    for (index, left) in [462, 508, 554].into_iter().enumerate() {
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
    let close = buttons[6].rect;
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
    let share = transport_share_at(middle, 800, 96, true);
    assert!(
        (share - 0.5).abs() < 0.05,
        "the middle of the bar is half of it"
    );

    // A press past either end is the end it is past, which is what keeps a drag from asking
    // for a second of a file that is not there.
    assert_eq!(transport_share_at(0, 800, 96, true), 0.0);
    assert_eq!(transport_share_at(800, 800, 96, true), 1.0);
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

    // Above the bar, hung from the button's own right edge, and clear of the button itself.
    assert_eq!(popup.panel.right, button.right);
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

/// A window too narrow for the panel is not a popup drawn off the side of it: it is a panel
/// that starts at the window's own edge, where the hand that opened it can still reach it.
#[test]
fn a_volume_popup_is_kept_inside_a_narrow_window() {
    let popup = volume_popup_layout(30, 100, 30, 96);

    assert!(popup.panel.left >= 0);
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

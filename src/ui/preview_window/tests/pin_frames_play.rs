//! A pinned picture that plays: the bar it draws, the frames its track names, and the pace its
//! own button sets.

use super::*;

/// The kinds a bar with a track of frames is drawn for, and the ones it is not: a picture that
/// plays carries one, and a held sequence this app hands to the engine — a HEIF or an AVIF, or
/// an animated JXL, none of which is decoded frame by frame here — does not.
#[test]
fn only_the_pictures_this_app_decodes_itself_carry_a_bar_of_frames() {
    for kind in [
        MediaType::AnimatedGif,
        MediaType::AnimatedApng,
        MediaType::AnimatedWebP,
    ] {
        assert!(pin_transport_kind(Some(kind)), "{kind:?} carries a bar");
        assert!(
            is_played_picture(Some(kind)),
            "{kind:?} is the app's own frames"
        );
        assert!(
            !pin_transport_live(Some(kind)),
            "{kind:?} has no player behind it, so its button is not drawn"
        );
    }

    for kind in [MediaType::AnimatedHeif, MediaType::AnimatedJxl] {
        assert!(!is_played_picture(Some(kind)), "{kind:?} is the engine's");
    }

    // A film keeps its own bar and its own button, and a picture, a page and a sound carry none.
    assert!(pin_transport_live(Some(MediaType::Video)));
    assert!(pin_transport_live(Some(MediaType::NativeVideo)));
    assert!(!pin_transport_kind(Some(MediaType::Audio)));
    assert!(!pin_transport_kind(Some(MediaType::StaticImage)));
    assert!(!pin_transport_kind(None));
}

/// A picture that plays is pinned with the two things its bar is drawn from read off its own
/// frames: how long its run is, and the pace it keeps.
#[test]
fn a_pinned_picture_takes_its_length_and_its_pace_from_its_own_frames() {
    let mut pin = overlay_pin((100, 100, 500, 400), PinChrome::always());
    pin.transport_bar = true;
    pin.transport_animated = true;
    pin.fps = PinFps {
        detected: 12.5,
        wanted: None,
        ..PinFps::default()
    };
    pin.transport = PinTransport {
        duration: Some(2.4),
        ..PinTransport::default()
    };

    assert_eq!(
        pin.fps.pace(),
        13,
        "the pace the slider opens at is the file's own, in whole frames a second"
    );
    assert_eq!(
        pin_duration(&pin.transport),
        Some(2.4),
        "and its length is the run of its own frames"
    );

    // A hand that moved the slider is the pace the picture plays at, and the file's own is what is
    // under it — a picture put back to its own pace on every scrub would be a slider that snapped.
    pin.fps.wanted = Some(30);
    assert_eq!(pin.fps.pace(), 30);
}

/// The pace a hand asks for scales every frame's own delay rather than replacing it: a picture paced
/// at twice its own rate halves every delay, and one paced at half of it doubles them, so the
/// relationship between the frames stays where the file's author put it.
#[test]
fn the_pace_a_hand_asks_for_scales_every_delay_the_file_gives() {
    let mut playback = AnimationPlayback {
        detected: 10.0,
        wanted: None,
    };

    assert_eq!(
        playback.delay_scale(),
        1.0,
        "a picture nobody has touched plays at its own pace"
    );
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, playback),
        100
    );

    playback.wanted = Some(20);
    assert!(
        (playback.delay_scale() - 0.5).abs() < 0.001,
        "twice the file's own pace is half of every wait"
    );
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, playback),
        50
    );
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 200, playback),
        100,
        "and a frame authored to wait twice as long still waits twice as long"
    );

    playback.wanted = Some(5);
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, playback),
        200,
        "half the pace is twice the wait"
    );

    // And a picture whose own pace was never read — a file whose container could not be walked —
    // plays at the delays it was given, whatever a hand asked for: a pace has no base to scale
    // against, and a guessed base would be a picture played at a rate nobody chose.
    let unknown = AnimationPlayback {
        detected: 0.0,
        wanted: Some(30),
    };
    assert_eq!(unknown.delay_scale(), 1.0);
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, unknown),
        100
    );
}

/// The frame a share of the bar names is a place in the file rather than a second of a player.
#[test]
fn a_share_of_the_track_names_a_frame_of_the_file() {
    let two_frames = crate::formats::animation_frames::FrameTimeline::of_delays(&[100, 100]);

    assert_eq!(two_frames.frames(), 2);
    assert_eq!(
        two_frames.index_at_share(0.0),
        0,
        "the bar's start is the first"
    );
    assert_eq!(two_frames.index_at_share(0.5), 1, "half way is the second");
    assert_eq!(two_frames.index_at_share(1.0), 1, "and its end is the last");

    // A run of a single frame is not one a track can be dragged along, which is what the press's
    // own gate refuses (see `pin_frame_share`).
    let one_frame = crate::formats::animation_frames::FrameTimeline::of_delays(&[100]);
    assert_eq!(one_frame.frames(), 1);
    assert_eq!(one_frame.index_at_share(1.0), 0);
}

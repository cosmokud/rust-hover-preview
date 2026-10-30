use super::*;

/// What a bubble carries for each kind, read as the mark it is drawn with.
///
/// One question with two halves and both of them are asked here: which kinds carry a playback
/// glyph, and which of the two glyphs that is. A mark that answered only the first would be a
/// bubble over a paused film drawn with a play triangle, which is a bubble saying the opposite of
/// what the file is doing (see `MediaType::bubble_mark`).
#[test]
fn a_kinds_bubble_mark_is_the_glyph_of_what_that_kind_is_doing() {
    // The two that play: the glyph follows the playback.
    for kind in [MediaType::Video, MediaType::NativeVideo, MediaType::Audio] {
        assert_eq!(kind.bubble_mark(true), pin_chrome::BubbleMark::Play);
        assert_eq!(kind.bubble_mark(false), pin_chrome::BubbleMark::Pause);
        assert!(kind.bubble_mark(true).is_playback());
    }

    // A page: nothing about a page plays, so nothing about it changes with a playback it does not
    // have.
    for kind in [MediaType::Text, MediaType::Archive, MediaType::Peazip] {
        assert_eq!(kind.bubble_mark(true), pin_chrome::BubbleMark::Page);
        assert_eq!(kind.bubble_mark(false), pin_chrome::BubbleMark::Page);
    }

    // And everything else is a picture — including a file the pin was shown and could not draw,
    // whose bubble shows the failure frame itself (see `has_bubble_picture`).
    for kind in [
        MediaType::StaticImage,
        MediaType::Pdf,
        MediaType::Office,
        MediaType::Unplayable,
    ] {
        assert_eq!(kind.bubble_mark(true), pin_chrome::BubbleMark::Picture);
        assert_eq!(kind.bubble_mark(false), pin_chrome::BubbleMark::Picture);
        assert!(!kind.bubble_mark(true).is_playback());
    }
}

/// Which bubbles are painted on the clock, and which are left where they were painted.
///
/// The bars over a playing film's frame are read out of this app's own sound exactly as a sound's
/// are, so a film that is making sound is a bubble whose bars move — and a film the app plays at a
/// level of nothing is one making no sound at all, begun with no audio track, whose bars are a row
/// of nothing to paint again for (see `bubble_bars_are_moving`).
#[test]
fn a_bubbles_bars_move_for_a_sound_and_for_a_film_that_is_making_one() {
    let _room = the_room_to_itself();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    stand_pin(Some(overlay_pin((100, 100, 420, 340), PinChrome::always())));

    /// Stand a media of `kind` in the slot the question is asked of.
    fn stand_media(kind: MediaType) {
        let mut media = create_loading_media(320, 240);
        media.media_type = kind;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }
    }

    // A sound is on the clock for as long as it is playing, at any level: its bands are the
    // sound's own, and a sound at nothing is a flat row rather than no row at all.
    stand_media(MediaType::Audio);
    with_pin(|pin| {
        pin.volume = PinVolume {
            audio: true,
            level: 0,
            ..Default::default()
        }
    });
    assert!(bubble_bars_are_moving(), "a sound's bars are its own bands");

    // A film's bars are those same bands drawn over its own frame, and they are moving only while
    // the film is making sound: the level is what says so, since a film at nothing is played with
    // no audio track at all.
    for kind in [MediaType::Video, MediaType::NativeVideo] {
        stand_media(kind);
        with_pin(|pin| {
            pin.volume = PinVolume {
                audio: false,
                level: 40,
                ..Default::default()
            }
        });
        assert!(
            bubble_bars_are_moving(),
            "a film at 40% is a film making sound"
        );

        with_pin(|pin| {
            pin.volume = PinVolume {
                audio: false,
                level: 0,
                ..Default::default()
            }
        });
        assert!(
            !bubble_bars_are_moving(),
            "a film at a level of nothing makes no sound to move a bar"
        );
    }

    // And a picture is not on the clock however loud the level beside it stands: no bars of its
    // own are read from anything.
    stand_media(MediaType::StaticImage);
    with_pin(|pin| {
        pin.volume = PinVolume {
            audio: false,
            level: 100,
            ..Default::default()
        }
    });
    assert!(!bubble_bars_are_moving(), "a picture's bubble never moves");

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **The still a bubble goes up with is the film at the second the pin's bar is showing.** It is the
/// picture a bubble shows before any frame has been read beside the film's player, and the colour its
/// ring is drawn in for as long as the bubble stands — so it is taken at the *pin's own clock* and at
/// the circle's own side, and the frame of another film in the slot goes with the ask rather than
/// being left as this bubble's picture (see `grab_bubble_frame`).
#[test]
fn a_collapse_asks_for_a_still_of_the_film_at_the_second_the_bar_is_showing() {
    let _room = the_room_to_itself();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_still = BUBBLE_FRAME.lock().ok().and_then(|mut frame| frame.take());
    bubble_stream::stand_no_stream();

    let film = PathBuf::from("a-film-a-bubble-shows-playing.mkv");
    let mut pin = overlay_pin((100, 100, 420, 340), PinChrome::always());
    pin.path = film.clone();
    pin.transport.begun(92.5, true, false);
    stand_pin(Some(pin));

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    if let Ok(mut frame) = BUBBLE_FRAME.lock() {
        *frame = Some(BubbleFrame {
            pixels: vec![1, 2, 3, 4],
            side: 44,
        });
    }

    let box_ = (300, 300, 364, 364);
    grab_bubble_frame(box_);

    let asked = take_noted_grabs();
    assert_eq!(asked.len(), 1, "a collapse asks for one still and one only");
    let (asked_film, at, side) = &asked[0];
    assert_eq!(
        asked_film, &film,
        "of the film the pin is showing, which is the one the bubble is a bubble of"
    );
    assert!(
        (92.5..93.0).contains(at),
        "at the second the pin's own bar is showing ({at}): the still is the frame the bubble goes up \
         with, and it is the picture the film is at rather than the top of the file"
    );
    assert_eq!(
        *side, 44,
        "and cut to the circle's own side on this display rather than to the square the window is \
         painted in"
    );
    assert!(
        BUBBLE_FRAME.lock().ok().and_then(|frame| frame.clone()).is_none(),
        "while the frame that was in the slot goes with the ask: a grab that never answers leaves a \
         bubble with nothing in it rather than one showing the film before this one"
    );

    bubble_stream::stand_no_stream();
    if let Ok(mut frame) = BUBBLE_FRAME.lock() {
        *frame = previous_still;
    }
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// What a film a player of somebody else's is drawing is drawn with: the frames read beside that
/// player, and the colour the *still* was taken from around them.
///
/// The two are deliberately not the same frame. A ring averaged from every live frame that passes is a
/// hairline that changes twenty-five times a second — a reading of the film rather than its colour —
/// so the still the collapse grabbed is the ring's, and a bubble over a live picture is ringed in one
/// colour from the moment it goes up to the moment it comes down (see `bubble_art`).
#[test]
fn the_frames_read_beside_a_player_are_ringed_in_the_stills_own_colour() {
    let _room = the_room_to_itself();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_still = BUBBLE_FRAME.lock().ok().and_then(|mut frame| frame.take());
    bubble_stream::stand_no_stream();

    stand_pin(Some(overlay_pin((100, 100, 420, 340), PinChrome::always())));
    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    // The still as a collapse leaves it: a frame of one colour, so that the ring's colour is that
    // colour and nothing else, of the size a bubble is drawn at on this display.
    let side = 44u32;
    let still = [40u8, 90, 160, 255]
        .iter()
        .cycle()
        .take(side as usize * side as usize * 4)
        .copied()
        .collect::<Vec<u8>>();
    if let Ok(mut frame) = BUBBLE_FRAME.lock() {
        *frame = Some(BubbleFrame {
            pixels: still,
            side,
        });
    }

    // And the frames a stream would have read, of a different colour from the still: which of the two
    // is the picture and which is the ring is the whole of what this asks.
    let live = [7u8, 200, 30, 255]
        .iter()
        .cycle()
        .take(side as usize * side as usize * 4)
        .copied()
        .collect::<Vec<u8>>();
    bubble_stream::stand_live_frame(Some((live.clone(), side)));

    match bubble_art(side).0 {
        BubbleArtOwned::Picture(pixels, picture_side, _, ring) => {
            assert_eq!(
                pixels, live,
                "what is inside the circle is the frame the stream read, not the still it began from"
            );
            assert_eq!(picture_side, side);
            assert_eq!(
                ring,
                Some([160, 90, 40]),
                "and the ring is the still's own colour, in a palette's order — one colour for as \
                 long as the bubble stands rather than the average of every frame that passes"
            );
        }
        _ => panic!("a film being streamed is a picture in the bubble"),
    }

    // And with no frames read, the still is the picture — and answers the ring itself, which is the
    // same colour the live case was given (see `BubbleArt::Picture`).
    bubble_stream::stand_live_frame(None);
    match bubble_art(side).0 {
        BubbleArtOwned::Picture(_, _, _, ring) => assert_eq!(ring, None),
        _ => panic!("a film with no stream and a still in hand shows the still"),
    }

    bubble_stream::stand_no_stream();
    if let Ok(mut frame) = BUBBLE_FRAME.lock() {
        *frame = previous_still;
    }
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

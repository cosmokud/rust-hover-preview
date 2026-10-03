use super::*;

/// The three things a pin's end settles that are not the pin, from every road out of it.
///
/// The walk the planner is working on, the bubble's drag latch and the box a drag had left
/// the window at are each somebody else's state, so each has its own owner and its own lock —
/// but they are settled from the one exit rather than by each road remembering to, which is
/// the same drift the pin's own state had: a road that forgot one left it standing for a pin
/// that was gone.
#[test]
fn no_road_out_of_a_pin_leaves_what_it_left_behind_standing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    for reason in [
        Reason::Closed,
        Reason::Asked,
        Reason::SwitchedOff,
        Reason::MediaGone,
        Reason::Hung,
    ] {
        install(PinnedPreview::for_test());
        PIN_BUBBLE_MOVED.store(true, Ordering::Release);
        if let Ok(mut drag) = PIN_BUBBLE_DRAG.lock() {
            *drag = Some(((0, 0), (10, 10)));
        }
        if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
            *request = Some((0, 0, 10, 10));
        }
        assert!(
            queue_pin_job(PinJob::OpenWith {
                path: PathBuf::from("x")
            }),
            "a walk is queued for the pin going down"
        );

        end_pin(reason, &RecordedPinWindow::new(0x1000));

        assert!(
            PIN_JOBS.0.lock().ok().is_some_and(|jobs| jobs.is_none()),
            "{reason:?}: the queued walk is dropped with the pin it was worked out for"
        );
        assert!(
            !PIN_BUBBLE_MOVED.load(Ordering::Acquire),
            "{reason:?}: the bubble's drag latch does not outlive the pin"
        );
        assert!(
            PIN_BUBBLE_DRAG
                .lock()
                .ok()
                .is_some_and(|drag| drag.is_none()),
            "{reason:?}: nor the drag it was latched for"
        );
        assert!(
            PIN_BOX_REQUEST
                .lock()
                .ok()
                .is_some_and(|box_| box_.is_none()),
            "{reason:?}: nor the box a drag had left a window that is going away at"
        );
    }
}

#[test]
fn a_pinned_windows_box_is_its_media_plus_the_room_its_chrome_needs() {
    let content = (100, 130, 500, 600);

    // The box the window stands in is the media's, with a caption taken above it and, for a
    // kind that plays, a transport bar below it — and the same box read back the other way
    // round is the media again.
    assert_eq!(
        pinned_window_box_of(content, 96, false, false, pinned_caption_height(96, None)),
        (100, 100, 500, 600)
    );
    assert_eq!(
        pinned_window_box_of(content, 96, true, false, pinned_caption_height(96, None)),
        (100, 100, 500, 630)
    );
    assert_eq!(
        content_box_of(
            pinned_window_box_of(content, 96, true, false, pinned_caption_height(96, None)),
            96,
            true,
            false,
            pinned_caption_height(96, None)
        ),
        content
    );

    // A kind whose chrome is drawn over its media is its own box: there is no room to take
    // beside a picture for something painted on top of it.
    assert_eq!(
        pinned_window_box_of(content, 96, true, true, pinned_caption_height(96, None)),
        content
    );

    // And where each part of the window is drawn: the media between the two bands, or under
    // the chrome the whole of the window down, with the strips over its first and last rows.
    assert_eq!(pinned_band_rows(500, 30, 30, false), (30, 440));
    assert_eq!(pinned_band_rows(500, 30, 30, true), (0, 500));

    // What the room of a display is follows from the same answer: a pin that keeps its chrome
    // inside its own box has the whole work area to be placed in.
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1000,
        bottom: 800,
    };
    assert_eq!(
        pinned_room(bounds, 96, true, false, pinned_caption_height(96, None)).room(),
        (1000, 800 - 30 - 30)
    );
    assert_eq!(
        pinned_room(bounds, 96, true, true, pinned_caption_height(96, None)).room(),
        (1000, 800)
    );
}

/// A sound is given no caption at all, which is what makes its card the whole of its window.
///
/// The caption a banded kind carries is a band of the window above the media: room for a name
/// and a row of buttons, and the hand has to be on it to move the window. A sound's card is its
/// own size, carries its own name at its own top and its own controls on its own row, so a
/// caption above it would be a second name over the first with a close button on it — and there
/// would be nothing left of the window under it but the card, which is the whole of what the
/// window was for.
#[test]
fn a_sound_is_given_no_caption_above_its_card() {
    assert_eq!(
        pinned_caption_height(96, Some(MediaType::Audio)),
        0,
        "the card is the whole of a pinned sound's window"
    );

    // Every other kind, and nothing pinned, keep the caption where it has always been.
    for kind in [
        Some(MediaType::StaticImage),
        Some(MediaType::Text),
        Some(MediaType::Archive),
        Some(MediaType::NativeVideo),
        Some(MediaType::Video),
        None,
    ] {
        assert_eq!(
            pinned_caption_height(96, kind),
            pinned_caption_height(96, None),
            "{kind:?} keeps its caption where it was"
        );
    }

    // Which is a band of the window and not a change of what is drawn in it: a sound's window
    // is the card's box, read back the other way round it is the card again, and the band the
    // card is drawn in is the whole of the window rather than the rows below a caption.
    let window = pinned_window_box_of((100, 100, 500, 600), 96, false, false, 0);
    assert_eq!(window, (100, 100, 500, 600));
    assert_eq!(content_box_of(window, 96, false, false, 0), window);
    assert_eq!(pinned_band_rows(600, 0, 0, false), (0, 600));

    // A picture's band begins below the caption it has always had.
    let caption = pinned_caption_height(96, None);
    assert_eq!(
        pinned_band_rows(600, caption, 0, false),
        (caption, 600 - caption)
    );
}

/// There is no band above a sound's card, and so there is no title bar for a press there to be
/// answered by: what a hand lands on is the card's own row, which is the whole of the window.
#[test]
fn a_sound_with_no_caption_above_it_has_no_band_to_press_in() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();

    let mut pin = PinnedPreview {
        overlay: false,
        caption: pinned_caption_height(96, Some(MediaType::Audio)),
        content: (100, 100, 500, 300),
        ..PinnedPreview::for_test()
    };
    pin.chrome = PinChrome::always();
    stand_pin(Some(pin));

    let caption = pinned_caption_geometry().expect("a caption for a pin that is up");
    assert_eq!(caption.height, 0, "and nothing of one to be in");

    // Every press above the card's own row is the card's own, because there is no row above it
    // that is anything else (see `pinned_press`).
    assert!((0..caption.height).is_empty());

    stand_pin(previous_pin);
}

/// The middle of the card's volume button, in *screen* coordinates: what the cursor has to be
/// at for the card to answer with that button, and it is read of the card's own arithmetic
/// rather than counted out here — which is the whole of what the two tests below are about.
fn sound_pin_volume_button(pin: &PinnedPreview) -> (i32, i32) {
    let button = audio_preview::control_box(
        CardControl::Volume,
        (pin.content.2 - pin.content.0) as u32,
        pin.dpi,
        current_audio_options(),
        true,
    )
    .expect("a box on a card that carries its controls");
    let window = pin.window_box();
    let (_, height) = pin.window_size();
    let (top, _) = pinned_band_rows(height, pin.caption, 0, pin.overlay);

    (
        window.0 + (button.left + button.right) / 2,
        window.1 + button.top + top,
    )
}

/// A pointer that has walked off a sound's card leaves no button lit: the card is a media
/// frame rather than chrome, so what shows the wash is the card's own paint, and the tick that
/// notices the pointer has gone has to ask for that card rather than for the window.
#[test]
fn a_hover_off_the_card_leaves_its_buttons_unlit() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let pin = sound_pin();
    let (x, y) = sound_pin_volume_button(&pin);
    stand_pin(Some(pin));
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);

    let now = Instant::now();

    // A pointer on the button lights it, and the caption stays up: for a banded kind the band
    // is the whole of the window, so a hand anywhere on it has asked for the caption (see
    // `pin_chrome_near`).
    assert!(
        pin_state().and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            Some(refresh_pin_chrome(pin, now, Some((x, y))))
        }) == Some(true),
        "a pointer arriving on a button is a change worth a repaint"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_hovered)),
        Some(Some(CardControl::Volume)),
        "and the button it is on is what the card is painted from"
    );

    // And away from the whole of it, which is where the question is actually answered: the
    // button is put out, and the card is asked to be painted again.
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);
    assert!(
        pin_state().and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            Some(refresh_pin_chrome(pin, now, Some((-1000, -1000))))
        }) == Some(true),
        "which is a change worth a repaint"
    );
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and the card's own paint is what has to be redone: it is a media frame, not chrome, so \
             `render_layered_preview` alone would redraw the old card"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_hovered)),
        Some(None),
        "and the button is left unlit under a pointer that has walked away"
    );

    stand_pin(previous_pin);
}

/// A pin of a sound keeps the box its hover had: the card a hover is measured with and the card a
/// pin shows are the same size, because the row of buttons on the pinned one is the bar's own row
/// and the bar is what they are carved out of.
#[test]
fn a_pin_of_a_sound_is_given_the_box_its_controls_fit_inside() {
    let folder = std::env::temp_dir().join("rust-hover-preview-pin-sound-box");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("song.mp3");
    std::fs::write(&path, b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00").expect("a file");
    audio_track::remember(
        &path,
        audio_track::Probed::Track(audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        }),
    );

    let options = current_audio_options();
    let hover = audio_card(&path, None, None, 0, None).expect("a card");
    let pinned = audio_card(
        &path,
        None,
        None,
        0,
        Some(CardChrome {
            playing: false,
            volume: current_audio_volume(),
            hovered: None,
            pressed: None,
        }),
    )
    .expect("a card with its controls on it");
    let (hover_width, hover_height) =
        audio_preview::measure(&hover, 4096, 2160, 96, options).expect("a measured card");
    let (pinned_width, pinned_height) =
        audio_preview::measure(&pinned, 4096, 2160, 96, options).expect("a measured card");

    // The card a hover is shown and the card a pin shows are the same card, and the controls are
    // carved out of the bar rather than added to it — so the pin is the hover's box.
    assert_eq!(
        (pinned_width, pinned_height),
        (hover_width, hover_height),
        "a pinned card is the card a hover's is: {pinned_width}x{pinned_height} against \
             {hover_width}x{hover_height}"
    );

    // Which is the whole of what the take-up re-measures for: the box a pin is given holds
    // every control its card carries.
    let hover_box = (
        100,
        100,
        100 + hover_width as i32,
        100 + hover_height as i32,
    );
    let pin_box = pinned_audio_card_box(hover_box, &path, 96);
    assert_eq!(
        pin_box, hover_box,
        "a pin is given the very box its hover was given: {pin_box:?}"
    );

    let _ = std::fs::remove_file(&path);
}

/// The card a take-up is measured for is drawn again at the box the measurement came to: a pin is
/// given the box the hover it came from had, and a frame left at some other size is a card
/// drawn smaller than the window that is showing it.
#[test]
fn a_taken_up_sound_lays_its_card_out_again_for_the_box_it_is_given() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-sound-relayout");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("song.mp3");
    std::fs::write(&path, b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00").expect("a file");
    audio_track::remember(
        &path,
        audio_track::Probed::Track(audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        }),
    );

    // The card the hover drew, which is the card a pin inherits until it lays its own out.
    let hover = audio_card(&path, None, None, 0, None).expect("a card");
    let (width, height) = audio_preview::measure(&hover, 4096, 2160, 96, current_audio_options())
        .expect("a measured card");
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, width, height, 96);
    }

    // And the box a pin of it is measured for, which is the hover's own box: the card a pin shows is
    // the card a hover drew, controls and all.
    let mut pin = sound_pin();
    pin.content = (100, 100, 100 + width as i32, 100 + height as i32);
    let content = pin.content;
    assert_eq!(
        (content.3 - content.1) as u32,
        height,
        "and the box a pin is given is as tall as a hover's: {} against {height}",
        content.3 - content.1
    );
    assert_eq!(
        (content.2 - content.0) as u32,
        width,
        "and as wide: {} against {width}",
        content.2 - content.0
    );

    // The clock a take-up hands over: where the player is, how far a name has been scrolled,
    // and what the row of controls is saying. What is under the last of them is written here
    // rather than read of a standing pin, because a test that stands one shares a value the
    // whole process is reading (see `pin_window::PIN_TESTS_ONE_AT_A_TIME`).
    relayout_pinned_media(
        &path,
        content,
        96,
        Some(AudioCardClock {
            started: None,
            from: 0.0,
            paused: None,
            name_offset: 0,
            dpi: 96,
            chrome: Some(CardChrome {
                playing: false,
                volume: 40,
                hovered: None,
                pressed: None,
            }),
        }),
    );

    let frame = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref()?.frames.first().map(|frame| frame.width))
        .expect("a frame in the media slot");
    assert_eq!(
        frame,
        (content.2 - content.0) as u32,
        "the card is drawn into the pin's own box rather than the hover's"
    );

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    let _ = std::fs::remove_file(&path);
}

/// A pinned sound's window is the card's own box and the card is drawn from its first row, so
/// every control on the card is inside the window rather than under its edge — which is what a
/// press on the bar needs and what a window a band taller than its card would have taken from
/// it.
///
/// The three are one answer rather than three: a caption above a sound would put the card's
/// own rows that much further down (see `pinned_caption_height`), and a window shorter than the
/// card would put its last rows outside itself (see `pinned_audio_card_box`). Either one leaves
/// the bar at the bottom of the card answering nothing, and a bar that answers nothing is a
/// sound nobody can move through.
#[test]
fn a_pinned_sounds_bar_is_answered_from_the_card_and_not_from_below_it() {
    let pin = sound_pin();

    // The window is the card's box: no caption above it, no transport strip below it.
    assert_eq!(pin.window_box(), pin.content);
    let (width, height) = pin.window_size();
    let (top, _) = pinned_band_rows(height, pin.caption, 0, pin.overlay);
    assert_eq!(
        top, 0,
        "and the card is drawn from the window's own first row"
    );

    // Every control is where the card's own arithmetic puts it, in the card's own coordinates,
    // which are the window's because there is no band above them.
    for control in [
        CardControl::Previous,
        CardControl::Play,
        CardControl::Next,
        CardControl::Seek,
        CardControl::Volume,
    ] {
        let rect = audio_preview::control_box(
            control,
            (pin.content.2 - pin.content.0) as u32,
            pin.dpi,
            current_audio_options(),
            true,
        )
        .expect("a box on a card that carries its controls");

        assert!(
            rect.bottom <= height && rect.right <= width,
            "{control:?} at {rect:?} against a window of {width}x{height}"
        );
        assert_eq!(
            pin_audio_control_at(
                &pin,
                (rect.left + rect.right) / 2,
                (rect.top + rect.bottom) / 2
            ),
            Some(control),
            "and the middle of {control:?} is on {control:?}"
        );
    }

    // The bar in particular is answered with a share of the file rather than with nothing at
    // all: this is the one control a hand uses to move through a sound.
    let bar = audio_preview::control_box(
        CardControl::Seek,
        (pin.content.2 - pin.content.0) as u32,
        pin.dpi,
        current_audio_options(),
        true,
    )
    .expect("a bar");
    let share = audio_preview::bar_share_at(
        (bar.left + bar.right) / 2,
        (bar.top + bar.bottom) / 2,
        (pin.content.2 - pin.content.0) as u32,
        pin.dpi,
        current_audio_options(),
        true,
    )
    .expect("a share of the file under a press on the bar");

    assert!(
        (0.0..=1.0).contains(&share),
        "the middle of the bar is the middle of the file, and not {share}"
    );
}

/// The card's own paint is what shows the wash under a pointer, and the hover path is where
/// that is asked for — `AUDIO_CARD_DIRTY` is set rather than the window merely repainted, and
/// the distinction is the whole of why a sound's buttons light up at all.
#[test]
fn the_card_asks_to_be_repainted_when_its_hover_changes() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let pin = sound_pin();
    let (x, y) = sound_pin_volume_button(&pin);
    stand_pin(Some(pin));

    let now = Instant::now();

    AUDIO_CARD_DIRTY.store(false, Ordering::Release);
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(pin_audio_hover_refresh(pin, Some((x, y))))
    });
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "a change of the card's hover is a change of the card"
    );
    assert_eq!(
        changed,
        Some(true),
        "and the tick owes the window a repaint as well"
    );

    // Nothing changed, nothing is asked for: a pointer resting on the same button is not a
    // repaint a quarter of a second.
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(pin_audio_hover_refresh(pin, Some((x, y))))
    });
    assert_eq!(
        changed,
        Some(false),
        "the same answer twice is not a change"
    );
    assert!(
        !AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and nothing is asked of the card"
    );

    // And a pin brought up by a key has no pointer anywhere near it, which is why the arrival
    // window is the one every other kind gets (see `pin_hides_chrome`).
    let mut arrived = PinnedPreview::for_test();
    arrived.chrome = PinChrome::on_arrival(now);
    arrived.hides_chrome = true;
    arrived.overlay = false;
    assert!(!pin_audio_hover_refresh(&mut arrived, None));

    stand_pin(previous_pin);
}

/// A sound has no transport strip and no caption, and that is the whole of the design in one
/// line: its controls are the card's own, drawn on the card's own row, and the window is the
/// card and nothing else around it.
///
/// These three answers are a set, and the set is what keeps a sound from growing a second set
/// of controls on a strip of its own — a strip that came and went with the pointer would be a
/// caption hiding the very buttons it existed to reach.
#[test]
fn a_sound_has_no_transport_strip_and_no_caption() {
    // The video the media engine plays: a strip of chrome over a picture, both of which the
    // pointer asks for. A video FFmpeg plays keeps its bands and shows them always, because the
    // band is where the player's own window stands.
    assert!(pin_transport_kind(Some(MediaType::NativeVideo)));
    assert!(pin_hides_chrome(Some(MediaType::NativeVideo)));
    assert!(pin_transport_kind(Some(MediaType::Video)));
    assert!(!pin_hides_chrome(Some(MediaType::Video)));
    assert!(!pin_hides_chrome(None));

    assert!(
        !pin_transport_kind(Some(MediaType::Audio)),
        "a sound's controls are the card's, drawn on the card's own row"
    );
    assert!(
        !pin_hides_chrome(Some(MediaType::Audio)),
        "and it is given no caption to hide, so its chrome never comes or goes"
    );

    // The card is its own size and is not framed into one either way (see `pin_frame`).
    assert_eq!(pin_frame(Some(MediaType::Audio)), PinFrame::None);
    assert!(!pin_overlay_chrome(Some(MediaType::Audio)));

    // And those three facts together are what says "this pin is showing a sound's card",
    // which is the gate every question about the card's controls is behind (see
    // `pin_shows_an_audio_card`).
    assert!(pin_shows_an_audio_card(&sound_pin()));
    assert!(!pin_shows_an_audio_card(&overlay_pin(
        (100, 100, 500, 300),
        PinChrome::always()
    )));
}

#[test]
fn the_chrome_is_drawn_over_the_media_of_the_kinds_this_app_draws_itself() {
    // A picture, a texture, an animation, a drawing, a page of a document or a book, and the
    // video the media engine decodes: this app holds the band of every one of them, so a strip
    // of chrome can be painted over it.
    for kind in [
        MediaType::StaticImage,
        MediaType::Dds,
        MediaType::AnimatedGif,
        MediaType::Design,
        MediaType::Vector,
        MediaType::Pdf,
        MediaType::NativeVideo,
    ] {
        assert!(pin_overlay_chrome(Some(kind)), "{kind:?}");
    }

    // The kinds that keep their chrome in bands around the media: the two windows this app
    // does not own (the player's video, and the page an SVG or a font is drawn on), a page that
    // is laid out to whatever box it is given, and a sound's card.
    for kind in [
        MediaType::Text,
        MediaType::Archive,
        MediaType::Audio,
        MediaType::Video,
        MediaType::EngineSvg,
        MediaType::EngineFont,
    ] {
        assert!(!pin_overlay_chrome(Some(kind)), "{kind:?}");
    }

    assert!(!pin_overlay_chrome(None));
}

#[test]
fn a_pins_chrome_is_asked_for_one_strip_at_a_time() {
    let caption = pinned_caption_height(96, None);
    let pin = overlay_pin((100, 100, 500, 400), PinChrome::always());

    // The caption's strip and the room beside it: a pointer over the picture's first rows, or
    // just above the window, is a pointer that has come for the title bar.
    assert_eq!(pin_chrome_near(&pin, Some((300, 100))), (true, false));
    assert_eq!(pin_chrome_near(&pin, Some((300, 90))), (true, false));
    assert_eq!(
        pin_chrome_near(&pin, Some((300, 100 + caption))),
        (true, false)
    );
    assert_eq!(
        pin_chrome_near(&pin, Some((300, 100 + caption + 24))),
        (true, false)
    );
    assert_eq!(
        pin_chrome_near(&pin, Some((300, 100 + caption + 25))),
        (false, false)
    );

    // And out in the picture, which is where the chrome is out of the way: a hand there is
    // reading the file rather than looking for its buttons. A hand out beside the window is not
    // near anything of it either, whichever end it is level with.
    assert_eq!(pin_chrome_near(&pin, Some((300, 300))), (false, false));
    assert_eq!(pin_chrome_near(&pin, Some((500 + 25, 100))), (false, false));
    assert_eq!(pin_chrome_near(&pin, None), (false, false));

    // A kind that plays carries the bar across the bottom, and *that* is what a hand down there
    // has asked for: the title bar is the other end of the window, and it stays where it is.
    let mut playing = pin;
    playing.transport_bar = true;
    assert_eq!(
        pin_chrome_near(&playing, Some((300, 400 - 12))),
        (false, true)
    );
    assert_eq!(
        pin_chrome_near(&playing, Some((300, 400 + 24))),
        (false, true)
    );
    assert_eq!(
        pin_chrome_near(&playing, Some((300, 400 - 30 - 25))),
        (false, false)
    );
    assert_eq!(pin_chrome_near(&playing, Some((300, 300))), (false, false));
    assert_eq!(pin_chrome_near(&playing, Some((300, 100))), (true, false));
}

#[test]
fn a_pins_chrome_is_shown_and_hidden_by_where_the_pointer_is() {
    let now = Instant::now();
    let content = (100, 100, 500, 400);
    let far = Some((-1000, -1000));
    let near = Some((300, 100 + pinned_caption_height(96, None) / 2));

    // A pin comes up with the whole of its chrome showing, and it stays that way for a moment
    // whether or not the pointer is anywhere near it — the moment a hand looks for the buttons
    // in — and nothing about that costs a repaint: it is already drawn.
    let mut pin = overlay_pin(content, PinChrome::on_arrival(now));
    pin.transport_bar = true;
    assert!(pin.chrome.caption && pin.chrome.bar);
    assert!(!refresh_pin_chrome(
        &mut pin,
        now + Duration::from_millis(1400),
        far
    ));
    assert!(pin.chrome.caption && pin.chrome.bar);

    // Then the strips left alone are gone, in the tick that notices rather than over a fade of
    // them, and the answer having changed is the repaint that shows the picture in their place.
    let leaving = now + Duration::from_millis(1500);
    assert!(refresh_pin_chrome(&mut pin, leaving, far));
    assert!(!pin.chrome.caption && !pin.chrome.bar);

    // And asking again is not a change: a strip that has gone costs nothing to keep gone.
    assert!(!refresh_pin_chrome(
        &mut pin,
        leaving + Duration::from_secs(1),
        far
    ));
    assert!(!pin.chrome.caption && !pin.chrome.bar);

    // A hand coming for the title bar gets the title bar — and not the bar at the other end of
    // the window, which is what it did not ask for. A hand at the bottom is answered the same
    // way round.
    assert!(refresh_pin_chrome(
        &mut pin,
        leaving + Duration::from_secs(2),
        near
    ));
    assert!(pin.chrome.caption && !pin.chrome.bar);

    let at_the_bar = Some((300, 400 - 12));
    assert!(refresh_pin_chrome(
        &mut pin,
        leaving + Duration::from_secs(3),
        at_the_bar
    ));
    assert!(!pin.chrome.caption && pin.chrome.bar);

    // A strip that is already showing is not a change, and neither is one that leaves.
    assert!(!refresh_pin_chrome(
        &mut pin,
        leaving + Duration::from_secs(4),
        at_the_bar
    ));
    assert!(refresh_pin_chrome(
        &mut pin,
        leaving + Duration::from_secs(5),
        far
    ));
    assert!(!pin.chrome.caption && !pin.chrome.bar);

    // A press or a drag is a hand on the picture rather than a hand asking for a title bar, so
    // the chrome is not brought out by one — what is asked is where the pointer is and nothing
    // else. What holds a button is a pointer that is on the strip the button is in, which is
    // where the pointer has to be for the press to have landed on it at all.
    let mut pressed = overlay_pin(content, PinChrome::always());
    pressed.chrome = PinChrome {
        caption: false,
        bar: false,
        until: None,
    };
    pressed.pressed = Some(pin_chrome::CaptionButton::Minimize);
    pressed.dragging = Some(PinDrag {
        from: (0, 0),
        window: content,
        action: PinDragAction::Move,
        delivered: true,
        carried: (i32::MIN, i32::MIN),
    });
    assert!(!refresh_pin_chrome(&mut pressed, now, far));
    assert!(!pressed.chrome.caption && !pressed.chrome.bar);

    // A pin whose chrome is not drawn over its media has nothing to show or hide and nothing it
    // is asked: a text preview's caption and bar are where they have always been, always there.
    let mut text = overlay_pin(content, PinChrome::always());
    text.overlay = false;
    text.hides_chrome = false;
    assert!(!refresh_pin_chrome(&mut text, now, far));
    assert!(text.chrome.caption && text.chrome.bar);
}

/// A caption's two hand-off buttons say what they are, and say it only after the pointer has
/// settled on one.
///
/// The two are the pair a glyph cannot tell apart — one icon for "open this" and one for
/// "open this with something else" is a question asked twice with the same picture — so
/// each names itself, and the name of the default is the program the machine would really
/// use, which is the one fact a hand cannot work out from the icon.
///
/// The wait is the other half of it: a name that appeared as the pointer crossed the strip
/// would be a word left behind on every button the pointer passed, and a caption is where
/// the pointer is already moving through on its way to somewhere else.
#[test]
fn the_two_hand_off_buttons_name_themselves_once_the_pointer_settles() {
    let now = Instant::now();
    let content = (100, 100, 500, 400);

    let mut pin = overlay_pin(content, PinChrome::always());
    pin.tooltip = PinTooltip {
        default_app: "Adobe Photoshop".to_string(),
        ..Default::default()
    };

    // The default's own name, which is the one a hand cannot get from the icon.
    assert_eq!(
        pin.tooltip
            .text_for(pin_chrome::CaptionButton::OpenWith)
            .as_deref(),
        Some("Open With Adobe Photoshop")
    );
    // And the list, which says what it is rather than what it would open with.
    assert_eq!(
        pin.tooltip
            .text_for(pin_chrome::CaptionButton::OpenWithList)
            .as_deref(),
        Some("Open With...")
    );
    // A button whose glyph already says what it is does not say it again in words.
    assert_eq!(pin.tooltip.text_for(pin_chrome::CaptionButton::Close), None);
    assert_eq!(pin.tooltip.text_for(pin_chrome::CaptionButton::Next), None);

    // Arriving is not saying: the pointer has to rest before the name is written, and every
    // tick before that is the same answer and costs no repaint.
    pin.hovered = Some(pin_chrome::CaptionButton::OpenWith);
    assert!(!pin.tooltip.refresh(pin.hovered, now));
    assert!(!pin
        .tooltip
        .refresh(pin.hovered, now + PIN_TOOLTIP_DELAY / 2));
    assert_eq!(pin.tooltip.shown, None);

    // And then it is said, once, and the tick after that says nothing new.
    let said = now + PIN_TOOLTIP_DELAY;
    assert!(pin.tooltip.refresh(pin.hovered, said));
    assert_eq!(pin.tooltip.shown, Some(pin_chrome::CaptionButton::OpenWith));
    assert!(!pin
        .tooltip
        .refresh(pin.hovered, said + Duration::from_secs(1)));

    // Moving on puts it away at once rather than waiting out the delay a second time, and
    // a different button starts its own wait from scratch.
    pin.hovered = Some(pin_chrome::CaptionButton::OpenWithList);
    assert!(pin
        .tooltip
        .refresh(pin.hovered, said + Duration::from_secs(1)));
    assert_eq!(pin.tooltip.shown, None);
    assert!(!pin
        .tooltip
        .refresh(pin.hovered, said + Duration::from_millis(1)));
    assert!(pin.tooltip.refresh(
        pin.hovered,
        said + Duration::from_secs(1) + PIN_TOOLTIP_DELAY
    ));
    assert_eq!(
        pin.tooltip.shown,
        Some(pin_chrome::CaptionButton::OpenWithList)
    );

    // A caption that has gone takes the name with it: a button that is not drawn is not a
    // button a hand is on, and a name left over would be a caption naming nothing.
    assert!(pin.tooltip.refresh(None, said + Duration::from_secs(3)));
    assert_eq!(pin.tooltip.shown, None);

    // And a machine with nothing filed against the format has no program to name, so the
    // default's button says nothing rather than saying "Open With" on its own.
    let unnamed = PinTooltip::default();
    assert_eq!(unnamed.text_for(pin_chrome::CaptionButton::OpenWith), None);
    assert_eq!(
        unnamed
            .text_for(pin_chrome::CaptionButton::OpenWithList)
            .as_deref(),
        Some("Open With..."),
        "the list is there whatever the machine has filed the format under"
    );
}

/// Every kind of pin says the name of its two hand-off buttons — not only the kinds whose
/// caption the pointer has to ask for.
///
/// The two kinds of pin are different in one way only: whether the pointer can bring the
/// caption out. A kind whose caption is in bands around its media has it always, which is
/// no reason for it to be a caption that says nothing — and it carries the same two buttons
/// and answers a press on both the same way, so a hand on a text preview's "Open With..."
/// has as much to learn from a name as a hand on a picture's.
#[test]
fn a_caption_says_its_name_whatever_kind_of_pin_it_is_on() {
    let now = Instant::now();
    let content = (100, 100, 500, 400);

    for (name, overlay) in [("a picture", true), ("a page of text", false)] {
        let mut pin = overlay_pin(content, PinChrome::always());
        pin.overlay = overlay;
        pin.hovered = Some(pin_chrome::CaptionButton::OpenWithList);

        // The caption is in a different place on the two kinds: over the picture's own first
        // rows, and in a strip of its own above them. The pointer is put where each kind
        // actually carries it, because the question a name is asked on is the same one for
        // both and it is about where the pointer is.
        let window = pin.window_box();
        let on_the_caption = Some((
            window.0 + 300,
            window.1 + pinned_caption_height(pin.dpi, None) / 2,
        ));

        // The first tick also settles the chrome itself, which is a repaint of its own; the
        // name is measured from where the pointer was noticed, so it is asked about on the
        // next one rather than this.
        let _ = refresh_pin_chrome(&mut pin, now, on_the_caption);
        assert_eq!(pin.tooltip.shown, None, "{name}: arriving is not saying");

        // And once the wait is over, the name is written on both kinds alike.
        let said_at = now + PIN_TOOLTIP_DELAY;
        assert!(
            refresh_pin_chrome(&mut pin, said_at, on_the_caption),
            "{name}: a name is a repaint the window owes"
        );
        assert_eq!(
            pin.tooltip.shown,
            Some(pin_chrome::CaptionButton::OpenWithList),
            "{name}: the button under the pointer names itself"
        );

        // And asking again while nothing has moved is not a change, which is what keeps a
        // name up without the tick repainting the window sixty times a second for it.
        assert!(!refresh_pin_chrome(
            &mut pin,
            said_at + Duration::from_secs(1),
            on_the_caption
        ));
    }
}

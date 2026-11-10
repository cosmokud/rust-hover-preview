use super::*;

// The stand-in for the machine's own `ReleaseCapture` re-entrancy, shared with
// the tests about the roads a release ends (see `pin_input`).
use super::pin_input::CapturingWindow;

/// The two settings a sound's card's font is read from — the Audio
/// Scaling share and the text font size — stood where a test asks and
/// put back when the guard is dropped, so a test beside this one is
/// answered with the machine's own whatever this one asserts, even one
/// that fails on the way (the house pattern: `PinLevels` in
/// `pin_volume`).
struct CardFontSettings {
    audio_scale: PreviewScale,
    text_font_scale_percent: u32,
}

impl CardFontSettings {
    /// Stands the share and the text font size where a test wants them,
    /// holding what the machine's own settings were to put back.
    fn stood_at(audio_scale: PreviewScale, text_font_scale_percent: u32) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = CardFontSettings {
            audio_scale: config.audio_scale,
            text_font_scale_percent: config.text_font_scale_percent,
        };
        config.audio_scale = audio_scale;
        config.text_font_scale_percent = text_font_scale_percent;

        was
    }

    /// Stands the share where a test wants it, the text font size left
    /// alone, holding what the machine's own settings were to put back.
    fn stood_the_share_at(audio_scale: PreviewScale) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = CardFontSettings {
            audio_scale: config.audio_scale,
            text_font_scale_percent: config.text_font_scale_percent,
        };
        config.audio_scale = audio_scale;

        was
    }
}

impl Drop for CardFontSettings {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.audio_scale = self.audio_scale;
            config.text_font_scale_percent = self.text_font_scale_percent;
        }
    }
}

/// The theme a sound's card is painted with, stood where a test wants it
/// and put back when the guard is dropped, the way `CardFontSettings`
/// stands the share and the text font size. The theme is the one setting
/// a repaint of a card on screen still answers to, so the test that
/// exercises that stands it where it is not, and a test that fails on the
/// way still leaves the machine's own theme behind it.
struct CardThemeSettings {
    theme: TextTheme,
}

impl CardThemeSettings {
    /// Stands the theme where a test wants it, holding what the machine's
    /// own theme was to put back.
    fn stood_at(theme: TextTheme) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = CardThemeSettings {
            theme: config.theme,
        };
        config.theme = theme;

        was
    }
}

impl Drop for CardThemeSettings {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.theme = self.theme;
        }
    }
}

/// A sound on disk, probed and remembered the way a hover finds one: the
/// card a test paints is the card of a file the audio track cache knows,
/// with a length and four facts a card is worth showing.
///
/// The folder is named per test, so a test's file is its own, and the
/// file is the test's to take away when it is done with it.
fn a_remembered_sound(folder: &str) -> PathBuf {
    let folder = std::env::temp_dir().join(folder);
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("song.mp3");
    std::fs::write(&path, b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00").expect("a file");
    audio_track::remember(
        &path,
        audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        },
    );

    path
}

/// The three things a pin's end settles that are not the pin, from every road out of it.
///
/// The walk the planner is working on, the bubble's drag latch and the box a drag had left
/// the window at are each somebody else's state, so each has its own owner and its own lock —
/// but they are settled from the one exit rather than by each road remembering to, which is
/// the same drift the pin's own state had: a road that forgot one left it standing for a pin
/// that was gone.
#[test]
fn no_road_out_of_a_pin_leaves_what_it_left_behind_standing() {
    let _room = the_room_to_itself();

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
            pinned_caption_height(96, None),
            0
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
    assert_eq!(pinned_band_rows(500, 30, 30, false, 0), (30, 440));
    assert_eq!(pinned_band_rows(500, 30, 30, true, 0), (0, 500));

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
    assert_eq!(content_box_of(window, 96, false, false, 0, 0), window);
    assert_eq!(pinned_band_rows(600, 0, 0, false, 0), (0, 600));

    // A picture's band begins below the caption it has always had.
    let caption = pinned_caption_height(96, None);
    assert_eq!(
        pinned_band_rows(600, caption, 0, false, 0),
        (caption, 600 - caption)
    );
}

/// There is no band above a sound's card, and so there is no title bar for a press there to be
/// answered by: what a hand lands on is the card's own row, which is the whole of the window.
#[test]
fn a_sound_with_no_caption_above_it_has_no_band_to_press_in() {
    let _room = the_room_to_itself();

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
    let (top, _) = pinned_media_band_rows(pin);

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
    let _room = the_room_to_itself();

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

/// The facts line of a card travels only under a pointer resting on it: the row it is drawn in is
/// the card's own — stored in the card's pixels, and taken to the card by the media band's top row,
/// the same subtraction a press on one of its buttons is answered by (see `pin_audio_facts_at`).
#[test]
fn the_facts_line_is_engaged_by_a_pointer_on_its_own_row() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let mut pin = sound_pin();
    // The row the take-up writes down, in the card's own pixels (see `note_audio_facts_rect`).
    pin.audio_facts_rect = Some((8, 24, 380, 40));
    let (top, _) = pinned_media_band_rows(&pin);
    let window = pin.window_box();
    let on_the_line = (window.0 + 100, window.1 + top + 30);

    assert!(
        pin_audio_facts_at(&pin, 100, top + 30),
        "a point on the row is taken to the card through the media band's own top"
    );
    assert!(
        !pin_audio_facts_at(&pin, 100, top + 60),
        "and the bar's own row below it is not the facts line"
    );

    stand_pin(Some(pin));
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);

    let now = Instant::now();
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(pin, now, Some(on_the_line)))
    });
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_facts_hovered)),
        Some(true),
        "a pointer on the line is what engages the line's scroll"
    );
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and the card is asked to be painted again, as much as a lit button does"
    );

    // And off it, which lets the line go: what the loop's own step winds back to the first fact.
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(pin, now, Some((-1000, -1000))))
    });
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_facts_hovered)),
        Some(false),
        "a pointer that has walked away leaves the line to wind home"
    );

    stand_pin(previous_pin);
}

/// A pin of a sound keeps the box its hover had: the card a hover is measured with and the card a
/// pin shows are the same size, because the row of buttons on the pinned one is the bar's own row
/// and the bar is what they are carved out of.
#[test]
fn a_pin_of_a_sound_is_given_the_box_its_controls_fit_inside() {
    let _room = the_room_to_itself();

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-sound-box");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("song.mp3");
    std::fs::write(&path, b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00").expect("a file");
    audio_track::remember(
        &path,
        audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        },
    );

    let options = current_audio_options();
    let hover = audio_card(&path, None, None, 0, 0, None).expect("a card");
    let pinned = audio_card(
        &path,
        None,
        None,
        0,
        0,
        Some(CardChrome {
            detached: false,
            playing: false,
            volume: current_audio_volume(),
            hovered: None,
            pressed: None,
            window_buttons: false,
            always_on_top: true,
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

    // Which is the whole of what the take-up re-measures for: the box a pin is
    // given holds every control its card carries. The box is the audio room —
    // the room the Audio Scaling setting gives the display the hover stands
    // on — because that is the room the hover's own card is measured at (see
    // `audio_box`), and a pin is given the very box its hover was given.
    let room = audio_box_room(work_area_at(100, 100), current_audio_scale(), 96);
    let (room_width, room_height) =
        audio_preview::measure(&hover, room.width, room.height, 96, options)
            .expect("a measured card");
    let hover_box = (100, 100, 100 + room_width as i32, 100 + room_height as i32);
    let pin_box = pinned_audio_card_box(hover_box, &path, 96, current_audio_scale());
    assert_eq!(
        pin_box, hover_box,
        "a pin is given the very box its hover was given: {pin_box:?}"
    );

    let _ = std::fs::remove_file(&path);
}

/// The card's font is the share's own — the default text size at the
/// 10% anchor, scaled by the share's fraction of it — and not the text
/// font size, which sizes every other painted preview: a change to
/// Text Preview → Font Size leaves a sound's card where it is at every
/// share.
#[test]
fn a_card_is_built_at_the_share_s_font_not_the_text_s() {
    let _room = the_room_to_itself();
    let _settings = CardFontSettings::stood_at(PreviewScale::Percent(5), 200);

    assert_eq!(
        current_audio_options().font_scale_percent,
        63,
        "a 5% card is built at 63%, whatever the text font size is"
    );

    if let Ok(mut config) = crate::CONFIG.lock() {
        config.audio_scale = PreviewScale::Percent(25);
    }
    assert_eq!(
        current_audio_options().font_scale_percent,
        313,
        "a 25% card is built at 313%"
    );
}

/// The card a take-up is measured for is drawn again at the box the measurement came to: a pin is
/// given the box the hover it came from had, and a frame left at some other size is a card
/// drawn smaller than the window that is showing it.
#[test]
fn a_taken_up_sound_lays_its_card_out_again_for_the_box_it_is_given() {
    let _room = the_room_to_itself();

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-sound-relayout");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("song.mp3");
    std::fs::write(&path, b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00").expect("a file");
    audio_track::remember(
        &path,
        audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        },
    );

    // The card the hover drew, which is the card a pin inherits until it lays its own out.
    let hover = audio_card(&path, None, None, 0, 0, None).expect("a card");
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
            facts_offset: 0,
            dpi: 96,
            chrome: Some(CardChrome {
                detached: false,
                playing: false,
                volume: 40,
                hovered: None,
                pressed: None,
                window_buttons: false,
                always_on_top: true,
            }),
        }),
        PinRelayoutRoad::TakeUp,
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

/// A repaint of a card on screen keeps the frame it was painted
/// at: the options a sound's card is drawn with are part of the
/// drawing, so the clock-and-bar repaint paints with the options
/// the frame was first painted with, and a change to Audio
/// Scaling — the configuration moved on to another share — reaches
/// the next hover and the next pin only, the bargain every other
/// Scaling setting already makes. The card's clock and its creeping
/// bar keep their own cadence; only the size they are painted at
/// stays where it was.
#[test]
fn a_repaint_keeps_the_frame_it_was_painted_at() {
    let _room = the_room_to_itself();
    let _settings = CardFontSettings::stood_the_share_at(DEFAULT_AUDIO_SCALE);

    let path = a_remembered_sound("rust-hover-preview-pin-sound-repaint");

    // The card a hover shows, loaded into the media slot: the
    // frame's first painting, at the anchor's share.
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let painted_with = current_audio_scale();
    let hover = audio_card(&path, None, None, 0, 0, None).expect("a card");
    let (width, height) =
        audio_preview::measure(&hover, 4096, 2160, 96, audio_options_of(painted_with))
            .expect("a measured card");
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, width, height, 96);
    }
    let mut media = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|mut media| media.take())
        .expect("a card in the media slot");

    // The configuration moves on to another share.
    drop(_settings);
    let _moved = CardFontSettings::stood_the_share_at(PreviewScale::Percent(25));

    // The clock-and-bar repaint, handed a clock, a chrome and a
    // name offset the way the loop hands them over.
    let chrome = Some(CardChrome {
        detached: false,
        playing: false,
        volume: 40,
        hovered: None,
        pressed: None,
        window_buttons: true,
        always_on_top: true,
    });
    assert!(
        media.refresh_audio_card(&path, Some(67.0), Some(180.0), 96, 0, 0, chrome, 0.0),
        "the card is repainted"
    );

    // The frame it leaves is the frame the anchor's font paints:
    // the same box, and the same pixels as a render of the same
    // card into the same box at the anchor's options — with the
    // theme the configuration now answers, which is the one
    // setting a repaint still follows. A render is deterministic,
    // so the pixel equality holds, and it is the assertion that
    // the repaint did not follow the configuration.
    let frame = media.frames.first().expect("a frame").clone();
    assert_eq!(
        (frame.width, frame.height),
        (width, height),
        "the repainted frame keeps the box it was painted at"
    );
    let card = audio_card(&path, Some(67.0), Some(180.0), 0, 0, chrome).expect("a card");
    let (pixels, ..) =
        audio_preview::render(&card, width, height, 96, audio_options_of(painted_with))
            .expect("a painted card");
    assert_eq!(
        frame.pixels, pixels,
        "the repaint painted at the share the frame was painted at"
    );

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    let _ = std::fs::remove_file(&path);
}

/// A pin answers its controls at the share it was taken up at:
/// the options the card was taken up with are the options the
/// card is drawn at for as long as the pin is up, so the card's
/// own arithmetic — the control a press is answered with, the
/// window buttons' band, the volume popup's geometry — is
/// answered against the layout the card is drawn in after the
/// configuration has moved on to another share, so a press lands
/// on the control as drawn. A walk is a fresh take-up — its media
/// is freshly loaded — so it takes up the new share.
#[test]
fn a_pin_answers_its_controls_at_the_share_it_was_taken_up_at() {
    let _room = the_room_to_itself();
    let _settings = CardFontSettings::stood_the_share_at(DEFAULT_AUDIO_SCALE);

    let path = a_remembered_sound("rust-hover-preview-pin-sound-take-up");
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    // A card in the media slot at the anchor, which is what a
    // take-up of a sound takes its options from: the frame's
    // first painting, at the share the card is to be drawn at.
    let hover = audio_card(&path, None, None, 0, 0, None).expect("a card");
    let (width, height) = audio_preview::measure(&hover, 4096, 2160, 96, current_audio_options())
        .expect("a measured card");
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, width, height, 96);
    }
    let pin = take_up_pinned_window(&path, (100, 100, 100 + width as i32, 100 + height as i32));
    let anchor = current_audio_scale();

    // The configuration moves on to another share.
    drop(_settings);
    let _moved = CardFontSettings::stood_the_share_at(PreviewScale::Percent(25));

    // The control a press is answered with, at the centre of the
    // anchor's volume button: the button the card's own
    // arithmetic puts there at the share the card is drawn at,
    // which is the share the pin remembers rather than the share
    // the configuration now answers.
    let button = audio_preview::control_box(
        CardControl::Volume,
        (pin.content.2 - pin.content.0).max(1) as u32,
        pin.dpi,
        audio_options_of(anchor),
        true,
    )
    .expect("a box on a card that carries its controls");
    let (top, _) = pinned_media_band_rows(&pin);
    assert_eq!(
        pin_audio_control_at(
            &pin,
            (button.left + button.right) / 2,
            (button.top + button.bottom) / 2 + top,
        ),
        Some(CardControl::Volume),
        "the press is answered against the layout the card is drawn in"
    );

    // The window buttons' band is the anchor's: a hand in the
    // band the anchor's buttons stand in and across their columns
    // is a hand near the buttons, and a hand one row past it is
    // not.
    let band = audio_preview::window_button_band(pin.dpi, audio_options_of(anchor));
    let on_a_button = audio_preview::control_box(
        CardControl::AlwaysOnTop,
        (pin.content.2 - pin.content.0).max(1) as u32,
        pin.dpi,
        audio_options_of(anchor),
        true,
    )
    .expect("a box on a card that carries its controls");
    let x = (on_a_button.left + on_a_button.right) / 2;
    assert!(
        pin_audio_window_buttons_at(&pin, x, band - 1),
        "a hand at a button's own column and in the anchor's band is a hand near the buttons"
    );
    assert!(
        !pin_audio_window_buttons_at(&pin, x, band),
        "while a hand one row below it is a hand past the buttons"
    );

    // And the volume popup hangs off the anchor's button box —
    // the panel floating over the very button it was opened from,
    // not over where the moved share's button would have been.
    let popup = pinned_audio_volume_popup(&pin).expect("a popup over the card's own button");
    let expected = pin_chrome::volume_popup_from_button(
        RECT {
            top: button.top + top,
            ..button
        },
        pin.window_size().0,
        pin.dpi,
    );
    assert_eq!(
        popup.panel, expected.panel,
        "the popup's panel is hung off the button as the card drew it"
    );
    assert_eq!(popup.track, expected.track, "and its groove with it");

    // Then walk: the media slot is given a card freshly loaded at
    // the moved share, and the pin taken up over it is a fresh
    // take-up, which asks the configuration — the next pin gets
    // the new share, and its controls are answered at the new
    // share's layout.
    stand_pin(Some(pin));
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, width, height, 96);
    }
    let walked = take_up_pinned_window(&path, (100, 100, 100 + width as i32, 100 + height as i32));
    let moved = current_audio_scale();
    let walked_button = audio_preview::control_box(
        CardControl::Volume,
        (walked.content.2 - walked.content.0).max(1) as u32,
        walked.dpi,
        audio_options_of(moved),
        true,
    )
    .expect("a box on a card that carries its controls");
    let (walked_top, _) = pinned_media_band_rows(&walked);
    assert_eq!(
        pin_audio_control_at(
            &walked,
            (walked_button.left + walked_button.right) / 2,
            (walked_button.top + walked_button.bottom) / 2 + walked_top,
        ),
        Some(CardControl::Volume),
        "a fresh take-up asks the configuration, so the walk gets the new share"
    );

    stand_pin(previous_pin);
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    let _ = std::fs::remove_file(&path);
}

/// A repaint still follows the Theme setting: what the remembered
/// options freeze is the size the card is drawn at, not the theme,
/// so a change to the theme re-styles a card on screen — the
/// repaint composes the configuration's theme over the remembered
/// font size — even as a change to the share leaves the card
/// where it is.
#[test]
fn a_repaint_still_follows_the_theme_setting() {
    let _room = the_room_to_itself();
    let _settings = CardFontSettings::stood_the_share_at(DEFAULT_AUDIO_SCALE);

    let path = a_remembered_sound("rust-hover-preview-pin-sound-theme");

    // The frame's first painting, at the anchor's share.
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let painted_with = current_audio_scale();
    let painted_theme = current_audio_theme();
    let hover = audio_card(&path, None, None, 0, 0, None).expect("a card");
    let (width, height) =
        audio_preview::measure(&hover, 4096, 2160, 96, audio_options_of(painted_with))
            .expect("a measured card");
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, width, height, 96);
    }
    let mut media = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|mut media| media.take())
        .expect("a card in the media slot");

    // The share moved on, and the theme stood where it is not: the
    // configuration's opposite of the painted one, so the repaint
    // has a theme to follow that the frame it is repainting did
    // not paint with.
    drop(_settings);
    let _moved = CardFontSettings::stood_the_share_at(PreviewScale::Percent(25));
    let opposite = match painted_theme {
        TextTheme::Dark => TextTheme::Light,
        TextTheme::Light | TextTheme::Custom(_) => TextTheme::Dark,
    };
    let _theme = CardThemeSettings::stood_at(opposite);
    assert_ne!(
        current_audio_theme(),
        painted_theme,
        "the theme is stood where the card was not painted with"
    );

    let chrome = Some(CardChrome {
        detached: false,
        playing: false,
        volume: 40,
        hovered: None,
        pressed: None,
        window_buttons: true,
        always_on_top: true,
    });
    assert!(
        media.refresh_audio_card(&path, Some(67.0), Some(180.0), 96, 0, 0, chrome, 0.0),
        "the card is repainted"
    );

    // The pixels are a render at the *configuration's* theme over
    // the *remembered* font size — the composition, externally. A
    // render at the moved share's font, or at the painted theme,
    // is a different painting, so the equality is the assertion
    // that the repaint followed the theme and nothing else.
    let frame = media.frames.first().expect("a frame").clone();
    let card = audio_card(&path, Some(67.0), Some(180.0), 0, 0, chrome).expect("a card");
    let (pixels, ..) =
        audio_preview::render(&card, width, height, 96, audio_options_of(painted_with))
            .expect("a painted card");
    assert_eq!(
        frame.pixels, pixels,
        "the repaint painted the configuration's theme over the remembered font"
    );

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    let _ = std::fs::remove_file(&path);
}

/// A pin taken up after a change to Audio Scaling keeps the
/// hover's size: the box a pin of a sound is given is the box
/// of the card as it is drawn — at the share its frame was
/// painted at — and not the box the share the configuration
/// now answers names, so a hover whose card is on screen when
/// the setting moves is pinned at the size it was showing at,
/// and the window pinning it is not resized by the pinning.
/// The pin's controls answer at the anchor's layout, the
/// layout the card is drawn in.
#[test]
fn a_pin_taken_up_after_a_change_keeps_the_hovers_size() {
    let _room = the_room_to_itself();
    let _settings = CardFontSettings::stood_the_share_at(DEFAULT_AUDIO_SCALE);

    let path = a_remembered_sound("rust-hover-preview-pin-sound-take-up-size");
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    // The hover's card, loaded into the media slot at the
    // anchor, and the hover's box measured at the anchor: the
    // box the hover was given.
    let hover = audio_card(&path, None, None, 0, 0, None).expect("a card");
    let anchor = current_audio_scale();
    let room = audio_box_room(work_area_at(100, 100), anchor, 96);
    let (hover_width, hover_height) = audio_preview::measure(
        &hover,
        room.width,
        room.height,
        96,
        audio_options_of(anchor),
    )
    .expect("a measured card");
    let hover_box = (
        100,
        100,
        100 + hover_width as i32,
        100 + hover_height as i32,
    );
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, hover_width, hover_height, 96);
    }

    // The configuration moves on to another share, and the
    // hover — whose card is on screen at the anchor — is
    // pinned over the box it was showing at.
    drop(_settings);
    let _moved = CardFontSettings::stood_the_share_at(PreviewScale::Percent(25));
    let pin = take_up_pinned_window(&path, hover_box);

    // The pin's box is the hover's box — the box the hover
    // was given — and not the box the moved share names,
    // which is a bigger window around the same card.
    let moved_room = audio_box_room(work_area_at(100, 100), current_audio_scale(), 96);
    let (moved_width, moved_height) = audio_preview::measure(
        &hover,
        moved_room.width,
        moved_room.height,
        96,
        audio_options_of(current_audio_scale()),
    )
    .expect("a measured card");
    let moved_box = (
        100,
        100,
        100 + moved_width as i32,
        100 + moved_height as i32,
    );
    let pin_box = pin.content;
    assert_ne!(
        moved_box, hover_box,
        "the moved share names a different box: {moved_box:?} against {hover_box:?}"
    );
    assert_eq!(
        pin_box, hover_box,
        "a pin is given the very box its hover was given: {pin_box:?}"
    );
    assert_ne!(
        pin.content, moved_box,
        "the pin's window is not resized to the box the moved share names"
    );

    // The pin's controls still answer at the anchor's layout:
    // the volume button the anchor's font puts there is the
    // control a press at its centre lands on.
    let button = audio_preview::control_box(
        CardControl::Volume,
        (pin.content.2 - pin.content.0).max(1) as u32,
        pin.dpi,
        audio_options_of(anchor),
        true,
    )
    .expect("a box on a card that carries its controls");
    let (top, _) = pinned_media_band_rows(&pin);
    assert_eq!(
        pin_audio_control_at(
            &pin,
            (button.left + button.right) / 2,
            (button.top + button.bottom) / 2 + top,
        ),
        Some(CardControl::Volume),
        "the press is answered against the layout the card is drawn in"
    );

    stand_pin(previous_pin);
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    let _ = std::fs::remove_file(&path);
}

/// A relayout of a card on screen keeps the share it was
/// painted at: the layout a maximize or a resize asks for is
/// asked of the frame as it stands, so the card is laid out
/// again into the new box at the share its frame was first
/// painted at — at the anchor's font — and not at the share
/// the configuration has moved on to. The window the card is
/// in changes size; the card stays one card at its own share.
#[test]
fn a_relayout_of_a_card_on_screen_keeps_the_share_it_was_painted_at() {
    let _room = the_room_to_itself();
    let _settings = CardFontSettings::stood_the_share_at(DEFAULT_AUDIO_SCALE);

    let path = a_remembered_sound("rust-hover-preview-pin-sound-relayout-share");

    // The frame's first painting, at the anchor's share.
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let anchor = current_audio_scale();
    let hover = audio_card(&path, None, None, 0, 0, None).expect("a card");
    let (width, height) = audio_preview::measure(&hover, 4096, 2160, 96, audio_options_of(anchor))
        .expect("a measured card");
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, width, height, 96);
    }
    let mut media = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|mut media| media.take())
        .expect("a card in the media slot");

    // The configuration moves on to another share.
    drop(_settings);
    let _moved = CardFontSettings::stood_the_share_at(PreviewScale::Percent(25));

    // The relayout a maximize or a resize asks for, at a box
    // of a different size than the frame's, with the clock
    // the loop hands over.
    let chrome = Some(CardChrome {
        detached: false,
        playing: false,
        volume: 40,
        hovered: None,
        pressed: None,
        window_buttons: true,
        always_on_top: true,
    });
    let size = (width + 120, height + 60);
    assert!(
        media.relayout_audio_card(
            &path,
            AudioCardClock {
                started: None,
                from: 0.0,
                paused: None,
                name_offset: 0,
                facts_offset: 0,
                dpi: 96,
                chrome,
            },
            Some(67.0),
            Some(180.0),
            size,
        ),
        "the card is laid out again"
    );

    // The frame it leaves is a render of the same card into
    // that same box at the anchor's options — the theme the
    // configuration now answers over the font the anchor's
    // share names — and not a render at the moved share's
    // font. A render is deterministic, so the pixel equality
    // holds, and it is the assertion that the relayout did
    // not follow the configuration.
    let frame = media.frames.first().expect("a frame").clone();
    assert_eq!(
        (frame.width, frame.height),
        size,
        "the relaid-out frame fills the box it was given"
    );
    let card = audio_card(&path, Some(67.0), Some(180.0), 0, 0, chrome).expect("a card");
    let (pixels, ..) = audio_preview::render(&card, size.0, size.1, 96, audio_options_of(anchor))
        .expect("a painted card");
    assert_eq!(
        frame.pixels, pixels,
        "the relayout painted at the share the frame was painted at"
    );

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    let _ = std::fs::remove_file(&path);
}

/// A band under a pinned sound's card grows the window downward and leaves the card exactly where
/// it was: the same box, the same first row, and the same media band — the rows the band takes are
/// window the card never had, not rows of the card given away.
///
/// The band is the session's rather than the pin's (it is remembered so that the next pin comes up
/// with it), so the test puts it back down before it ends.
#[test]
fn a_band_under_a_card_grows_the_window_and_leaves_the_card_alone() {
    let _room = the_room_to_itself();

    let pin = sound_pin();
    assert_eq!(
        pin_playlist::playlist_rows(&pin),
        0,
        "a band is down until a hand opens one"
    );
    assert_eq!(pin.window_box(), pin.content);

    pin_playlist::toggle(&pin);
    let (band_top, band_height) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");

    assert!(
        pin_playlist::the_tick_is_owed_a_repaint(),
        "a band opened is a box the window has to be given, so the tick owes the paint that gives it"
    );

    assert!(
        band_height > 0,
        "a band opened on a display is a band with a height"
    );
    assert_eq!(
        band_top,
        pin.content.3 - pin.content.1,
        "it begins at the card's own bottom edge, counted in the window's own rows, so nothing of \
         the card moves for it"
    );
    assert_eq!(
        pin.window_box(),
        (
            pin.content.0,
            pin.content.1,
            pin.content.2,
            pin.content.3 + band_height
        ),
        "the window is the card's box grown downward and no other way"
    );
    assert_eq!(
        (pin.content.2 - pin.content.0, pin.content.3 - pin.content.1),
        (400, 200),
        "and the card's own box is the one it was taken up with"
    );

    // The media band is the card's rows and no more of them: what the band took is window the card
    // was never drawn in, which is what keeps `compose_media_into_band` from stretching it.
    let (top, rows) = pinned_media_band_rows(&pin);
    assert_eq!(
        top, 0,
        "the card still begins at the window's own first row"
    );
    assert_eq!(rows, 200, "and it is still as tall as the card is");

    pin_playlist::toggle(&pin);
    assert_eq!(
        pin.window_box(),
        pin.content,
        "and closing it puts the window back to the card's box"
    );
    assert!(
        pin_playlist::the_tick_is_owed_a_repaint(),
        "and a box given back is owed a paint as much as one given: the band is down, and the \
         window it grew is the size it grew to until a paint says otherwise"
    );
}

/// A window carried is a place and not a size, and the band under a card is the window's: the box a
/// move reads back out of the window it left is the card's own.
///
/// A move that read the band back as content would hand the card the rows of its own playlist, and
/// then measure the window as that taller card plus a band again — a card stretched down a window
/// that walks off the screen one drag at a time (see `content_box_of`).
#[test]
fn a_move_over_a_band_leaves_the_card_the_box_it_had() {
    let _room = the_room_to_itself();

    let pin = sound_pin();
    pin_playlist::toggle(&pin);
    let (_, band) = pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");

    // The window carried a hundred pixels right and two hundred down.
    let window = pin.window_box();
    let carried = (
        window.0 + 100,
        window.1 + 200,
        window.2 + 100,
        window.3 + 200,
    );
    let content = content_box_of(
        carried,
        pin.dpi,
        pin.transport_bar,
        pin.overlay,
        pin.caption,
        band,
    );

    assert_eq!(
        (content.2 - content.0, content.3 - content.1),
        (pin.content.2 - pin.content.0, pin.content.3 - pin.content.1),
        "the card was given a size by a drag that only changed its place"
    );
    assert_eq!(
        (content.0 - pin.content.0, content.1 - pin.content.1),
        (100, 200),
        "and it was carried where the hand carried it"
    );

    // And the two roads are one: the window that content stands in is the window the drag left.
    let measured = pinned_window_box_of(
        content,
        pin.dpi,
        pin.transport_bar,
        pin.overlay,
        pin.caption,
    );
    assert_eq!(
        (measured.0, measured.1, measured.2, measured.3 + band),
        carried,
        "the window measured from the card it was read back from is not the window the drag left"
    );

    pin_playlist::toggle(&pin);
}

/// A press the band took is a press the band gives back, and only that one: the flag a release is
/// answered by is spent where it is read.
///
/// A flag left standing is the release of every drag, every caption button and every seek the card
/// begins afterwards, taken for a press the band is still holding — a window left carrying the
/// pointer for the rest of the run, and a bar that never lets go of a seek.
#[test]
fn a_press_the_band_took_is_given_back_by_the_release_that_follows_it() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let pin = sound_pin();
    pin_playlist::toggle(&pin);
    let (band_top, band_height) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    stand_pin(Some(pin));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the band owes a paint for the rows it just built"
    );

    // A row of the band, and a press on it: the band's own, and not the window's to carry. The row
    // is taken past the band's own pad, which is the part of it a press picks nothing in.
    let row = band_top + pin_chrome::playlist_row_height(1.0) / 2 + 12;
    assert!(row < band_top + band_height, "a row inside the band");
    assert!(
        unsafe { pin_playlist::press(HWND::default(), 100, row, false) },
        "a press on a row is the band's"
    );

    // The release that follows it is the band's too — once, and only once: the press it answers is
    // spent by it, so the next release the window is sent is the window's again.
    assert!(
        unsafe { pin_playlist::release(HWND::default()) },
        "the release of a press the band took is the band's"
    );
    assert!(
        !unsafe { pin_playlist::release(HWND::default()) },
        "and a release with no press behind it is the window's"
    );

    // And the band is put back down, so the next test does not come up with one.
    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// A pointer in the band is the row of the band it is over, read off the screen row the tick reads
/// the cursor at: the band begins at the window's own row `band_top`, and a row counted in the
/// window's rows instead is one that far down the list — a band no hand is ever on, and a name
/// nothing carries (see `pin_playlist::pin_playlist_hover_refresh`).
#[test]
fn a_pointer_in_the_band_is_a_row_of_the_band() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let mut pin = sound_pin();
    // A name the band's column has no room for, so that the row the hand is on is a name that
    // marquees rather than one drawn once and held.
    pin.path = PathBuf::from(
        "a mixdown of a very long name indeed, and one no column of a band holds.mp3",
    );
    pin_playlist::toggle(&pin);
    let (band_top, _) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    let window = pin.window_box();
    let row = pin_chrome::playlist_row_height(pin.dpi as f32 / 96.0) / 2;
    // The window a pointer's own place is read against, standing where the pin's box says it does:
    // a screen row is a row of a window only the window itself can answer (see
    // `pin_input::the_point_in_the_window`).
    let hwnd = unsafe {
        a_window_of_the_suite(WINDOW_EX_STYLE::default(), w!("STATIC"), WS_POPUP, window)
    };
    stand_pin(Some(pin));
    pin_playlist::refresh_the_rows();

    // The pointer inside the band's own first row, as the screen row the tick reads it at.
    let cursor = Some((window.0 + 40, window.1 + band_top + row));
    let asked = |cursor| {
        pin_state()
            .and_then(|pinned| {
                pinned
                    .pin()
                    .map(|pin| pin_playlist::pin_playlist_hover_refresh(pin, hwnd, cursor))
            })
            .expect("the pin the test stood up")
    };

    assert!(asked(cursor), "a row the pointer has come to is a change");
    let (hovered, _, scroll) = pin_playlist::the_band_in_hand().expect("the band's own rows");
    assert_eq!(
        hovered,
        Some(scroll),
        "the pointer in the band's first row is on the row the band draws first"
    );
    // The name the hand is on is one the band carries: a marquee is only ever set up for a name the
    // column has no room for, so a row that has one is a row wider than the band (see
    // `pin_playlist::set_the_marquee`).
    assert!(
        pin_playlist::the_marquees_in_hand()
            .iter()
            .any(|(row, _)| *row == scroll),
        "and the name the hand is on is one of the names the band carries"
    );
    assert!(!asked(cursor), "the same row again is no change");

    // SAFETY: the window is the suite's own, and this is the road one is put away down.
    unsafe { put_away_a_window_of_the_suite(hwnd) };
    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// **A row the hand is on is the row the tick is on**, whichever of the two roads reads the pointer:
/// the move the window is sent carries its own rows, and the tick reads the cursor off the screen.
/// Two readings of the band's geometry were two answers, and the tick's — counted from the window's
/// own top rather than from the band's — cleared the row the move had just named, so a name
/// travelling across a row the hand was merely *on* stopped while one on a row it had *picked* went
/// on (see `pin_playlist::hover` and `pin_playlist::pin_playlist_hover_refresh`).
#[test]
fn a_row_the_hand_is_on_survives_the_tick_that_reads_the_pointer() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let mut pin = sound_pin();
    // A name the band's column has no room for, and one no other test of a band names: the row a
    // hand is on is only a name that marquees for a row the band had to build afresh.
    pin.path =
        PathBuf::from("a mixdown long enough that no column of a band holds the whole of it.mp3");
    pin_playlist::toggle(&pin);
    let (band_top, _) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    let window = pin.window_box();
    let row = pin_chrome::playlist_row_height(pin.dpi as f32 / 96.0) / 2;
    let hwnd = unsafe {
        a_window_of_the_suite(WINDOW_EX_STYLE::default(), w!("STATIC"), WS_POPUP, window)
    };
    stand_pin(Some(pin));
    pin_playlist::refresh_the_rows();

    assert!(
        unsafe { pin_playlist::hover(40, band_top + row) },
        "a row a hand has come to is a change"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(Some(0)),
        "the move says the hand is on the band's own first row"
    );
    assert!(
        pin_playlist::the_marquees_in_hand()
            .iter()
            .any(|(row, offset)| *row == 0 && *offset == 0),
        "and that is the row the name is set up for, held at its resting place"
    );

    // The tick that reads the same pointer off the screen: the row is the row, so there is
    // nothing to change and the name is still the name.
    let at = |at| {
        pin_state()
            .and_then(|pinned| {
                pinned
                    .pin()
                    .map(|pin| pin_playlist::pin_playlist_hover_refresh(pin, hwnd, at))
            })
            .expect("the pin the test stood up")
    };
    assert!(
        !at(Some(client_to_screen(hwnd, 40, band_top + row))),
        "the tick's pointer is the row the move already named"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(Some(0)),
        "so the row the hand is on is still the row it was"
    );
    assert!(
        pin_playlist::the_marquees_in_hand()
            .iter()
            .any(|(row, _)| *row == 0),
        "and the name the tick has to carry is still on it"
    );

    // And carried on: a tick past the row's hold moves the name along it, which is the whole of
    // what a tick that cleared the hover would have stopped.
    std::thread::sleep(audio_preview::NAME_HOLD + std::time::Duration::from_millis(60));
    pin_playlist::the_tick_is_owed_a_repaint();
    let offset = pin_playlist::the_marquees_in_hand()
        .into_iter()
        .find(|(row, _)| *row == 0)
        .map(|(_, offset)| offset)
        .expect("the name still marqueed on the row the hand is on");
    assert!(offset > 0, "and it has moved along it");

    // A pointer that has left the window is on no row at all — and the name it was carrying is not
    // stopped where it stands, nor given up: it is on its way back to its resting place, and the row
    // the pin is showing is one the band carries whatever the hand does (see
    // `pin_playlist::set_the_marquee`).
    assert!(
        at(Some((window.0 - 10, window.1 + band_top + row))),
        "a pointer beside the window is on nothing of it"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(None),
        "the row the hand has left is no row at all"
    );
    assert!(
        pin_playlist::the_marquees_in_hand()
            .iter()
            .any(|(row, _)| *row == 0),
        "and the name it left behind is still the band's to carry"
    );

    // SAFETY: the window is the suite's own, and this is the road one is put away down.
    unsafe { put_away_a_window_of_the_suite(hwnd) };
    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// **A row is a row of the window, at the scale the window is drawn at**: a display at 144% is where
/// the box the pin's content projects and the rect the window really rests in stop being the same
/// numbers, and the tick that read the row off the box instead cleared the one the move had named —
/// a name on a row the hand was merely *on* stopped, while one on a row it had *picked* went on (see
/// `pin_playlist::pin_playlist_hover_refresh` and `pin_input::the_point_in_the_window`).
#[test]
fn a_row_the_hand_is_on_is_a_row_of_the_window_at_any_scale() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let mut pin = sound_pin();
    pin.dpi = 144;
    pin.caption = pinned_caption_height(144, Some(MediaType::Audio));
    pin.path = PathBuf::from(
        "a third mixdown, whose name no column of a band holds and no other test names.mp3",
    );
    pin_playlist::toggle(&pin);
    let (band_top, _) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    let row = pin_chrome::playlist_row_height(pin.dpi as f32 / 96.0) / 2;

    // The window stands a window's own distance away from the box the pin is laid out in, and is the
    // size the pin asks its window to be: what the tick reads back is the window's rows, and a box
    // it cannot be asked about.
    let window = pin.window_box();
    let (width, height) = pin.window_size();
    let stands_at = (
        window.0 + 213,
        window.1 + 307,
        window.0 + 213 + width,
        window.1 + 307 + height,
    );
    assert_ne!(
        stands_at.1, window.1,
        "the window is not where the box says"
    );
    let hwnd = unsafe {
        a_window_of_the_suite(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            WS_POPUP,
            stands_at,
        )
    };
    stand_pin(Some(pin));
    pin_playlist::refresh_the_rows();

    assert!(
        unsafe { pin_playlist::hover(40, band_top + row) },
        "a row a hand has come to is a change"
    );
    let carried = pin_playlist::the_marquees_in_hand();
    assert_eq!(
        carried.iter().map(|(row, _)| *row).collect::<Vec<_>>(),
        vec![0],
        "the name is set up for the row the move named, at 144%"
    );

    // The same pointer read off the screen, the way the tick reads it: the window's own client point
    // is that same row, so nothing of the band's changes and the name is still the name.
    let at = |at| {
        pin_state()
            .and_then(|pinned| {
                pinned
                    .pin()
                    .map(|pin| pin_playlist::pin_playlist_hover_refresh(pin, hwnd, at))
            })
            .expect("the pin the test stood up")
    };
    assert!(
        !at(Some(client_to_screen(hwnd, 40, band_top + row))),
        "the tick's pointer is the row the move already named"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(Some(0)),
        "so the row the hand is on is still the row it was, at 144%"
    );
    assert_eq!(
        pin_playlist::the_marquees_in_hand()
            .iter()
            .map(|(row, _)| *row)
            .collect::<Vec<_>>(),
        vec![0],
        "and the name the tick has to carry is still on it"
    );

    std::thread::sleep(audio_preview::NAME_HOLD + std::time::Duration::from_millis(60));
    pin_playlist::the_tick_is_owed_a_repaint();
    let offset = pin_playlist::the_marquees_in_hand()
        .into_iter()
        .find(|(row, _)| *row == 0)
        .map(|(_, offset)| offset)
        .expect("the name still marqueed on the row the hand is on");
    assert!(offset > 0, "and it has moved along it");

    // SAFETY: the window is the suite's own, and this is the road one is put away down.
    unsafe { put_away_a_window_of_the_suite(hwnd) };
    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// **A name the band lets go of parks rather than snapping home.** A row the hand has left is a name
/// on its way back to its resting place — the card's own rule for the line beside its name — so what
/// the eye sees is the name walking back under its column rather than jumping to the front of it (see
/// `pin_playlist::set_the_marquee` and `audio_preview::NameScroll::set_engaged`).
///
/// The row travelled here is one the pin is **not** showing, because a highlighted row and a playing
/// one travel whatever the hand does: only a row the hand alone was carrying is ever let go of.
#[test]
fn a_name_the_band_lets_go_of_parks_before_it_stops() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let folder = PathBuf::from("a folder of three sounds for a name that parks");
    let sounds = [
        folder.join("a first mixdown.mp3"),
        folder.join("a second mixdown of a name so long that no column of any band in this suite holds the whole of it.mp3"),
        folder.join("a third mixdown.mp3"),
    ];

    // The pin stands on the first of the three, so the row a hand comes to is neither the row the pin
    // is showing nor one any pick has made its own.
    let mut pin = sound_pin();
    pin.path = sounds[0].clone();
    pin_playlist::toggle(&pin);
    let (band_top, _) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    let row_height = pin_chrome::playlist_row_height(pin.dpi as f32 / 96.0);
    stand_pin(Some(pin));
    stand_pin_walk(Some(PinStep {
        at: sounds[0].clone(),
        from: sounds[0].clone(),
        step: 1,
        order: sounds.to_vec(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the band builds the rows of the walk"
    );

    // The band's second row: the one name of the three the column has no room for.
    assert!(
        unsafe { pin_playlist::hover(40, band_top + row_height + row_height / 2) },
        "a row a hand has come to is a change"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(Some(1)),
        "the hand is on the band's second row"
    );
    assert_eq!(
        pin_playlist::the_marquees_in_hand()
            .iter()
            .map(|(row, _)| *row)
            .collect::<Vec<_>>(),
        vec![1],
        "and its name is the one name the band carries"
    );

    // A tick past the row's hold carries the name along it, so that there is somewhere to park from.
    std::thread::sleep(audio_preview::NAME_HOLD + std::time::Duration::from_millis(60));
    pin_playlist::the_tick_is_owed_a_repaint();
    let carried = pin_playlist::the_marquees_in_hand()
        .into_iter()
        .find(|(row, _)| *row == 1)
        .map(|(_, offset)| offset)
        .expect("the name still carried on the row the hand is on");
    assert!(carried > 0, "and it has moved along the row");

    // The hand leaves the band, and the name it was carrying is not stopped where it stands: it is
    // still the band's to carry, and it is on its way back to its resting place.
    assert!(
        unsafe { pin_playlist::hover(40, band_top - 1) },
        "a hand above the band is on no row of it"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(None),
        "so the band has no row the hand is on"
    );
    assert_eq!(
        pin_playlist::the_marquees_in_hand(),
        vec![(1, carried)],
        "and the name it left behind is still where it stood"
    );

    std::thread::sleep(std::time::Duration::from_millis(30));
    pin_playlist::the_tick_is_owed_a_repaint();
    let homeward = pin_playlist::the_marquees_in_hand()
        .into_iter()
        .find(|(row, _)| *row == 1)
        .map(|(_, offset)| offset)
        .expect("the name is on its way home rather than given up where it stood");
    assert!(homeward < carried, "and it has walked back towards home");
    assert!(homeward > 0, "and it is not home yet");

    // And the band is put back down, so the next test does not come up with one.
    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// A `Return` in a pinned sound's card plays the row the band has picked, and asks for nothing at all
/// where there is no band up or no row picked: the file is the hand's own pick, and a key that guessed
/// one would start a file nobody asked for (see `pin_playlist::play_the_pick`).
#[test]
fn a_return_plays_the_row_a_sounds_band_has_picked() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let folder = PathBuf::from("a folder of two sounds for the return key");
    let sounds = [folder.join("a.mp3"), folder.join("b.mp3")];

    // The pin's walk is two files of which it is showing the first, so the row a hand picks is one
    // that names a file the pin is not already on: a pick of the file on screen has nothing to ask.
    let pin = sound_pin();
    let mut standing = sound_pin();
    standing.path = sounds[0].clone();
    stand_pin(Some(standing));
    stand_pin_walk(Some(PinStep {
        at: sounds[0].clone(),
        from: sounds[0].clone(),
        step: 1,
        order: sounds.to_vec(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the band builds the rows of the walk"
    );

    // A pin is a window onto something, and what the drain below asks before it answers a key: the
    // media slot is stood with a card so that the pin is not read as one that came apart.
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(40, 24);
        media.media_type = MediaType::Audio;
        *slot = Some(media);
    }

    // A channel of the preview's own, stood in the slot a pick is sent on, so that what the key asks
    // for is what this test reads and a later test's channel is not read here (see
    // `show_pinned_file`).
    let (sender, answers) = channel();
    let previous_sender = PREVIEW_SENDER
        .lock()
        .ok()
        .and_then(|mut held| held.replace(sender));
    note_the_pin_file_in_flight(None);

    // One `Return`, through the loop's own drain, which is where a key is answered.
    let a_return = || {
        ask_pin(PinCommand::PlayRow);
        let mut request = None;
        unsafe {
            pin_command_request(
                HWND(0x1000 as *mut _),
                &mut request,
                &mut None,
                &mut None,
                &mut 0.0,
                &mut None,
                false,
                None,
            );
        }
        assert!(
            request.is_none(),
            "a row played is not a box asked for and not a pin taken down"
        );
        answers.try_recv()
    };

    // A band that is down is no list for a key to be in, whatever row it was left on.
    assert!(
        a_return().is_err(),
        "a `Return` with no band up asks for no file"
    );

    // And a band that is up with nothing picked is a band no row of the hand's own.
    pin_playlist::toggle(&pin);
    // The band's own turn over asks the walk behind it again, and that is not the key's answer
    // (see `PreviewMessage::PinWalkAgain`).
    let _ = answers.try_recv();
    assert!(
        a_return().is_err(),
        "and one with no row picked asks for none either"
    );

    // The row a `Down` arrow picks is the row the key plays, and it is the second of the walk. The
    // hand is put on the band's own first row first, which is where a hand reading the list is. The
    // band's own row is read and the pin's lock let go of before the move is answered: the pin is
    // not asked twice on one thread (see `the_band_of_the_pin`).
    let band_top = pin_state()
        .and_then(|pinned| pinned.pin().and_then(pin_playlist::pin_playlist_band))
        .expect("the band")
        .0;
    let _ = unsafe { pin_playlist::hover(40, band_top + pin_chrome::playlist_row_height(1.0) / 2) };
    assert!(
        pin_playlist::step_the_pick(1),
        "the band takes the arrow over it"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(_, picked, _)| picked),
        Some(Some(1)),
        "and the row it picked is the second of the walk"
    );
    let asked = a_return();

    pin_playlist::toggle(&pin);
    if let Ok(mut held) = PREVIEW_SENDER.lock() {
        *held = previous_sender;
    }
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    stand_pin(previous_pin);

    assert!(
        matches!(asked, Ok(PreviewMessage::PinUpdate(ref path)) if *path == sounds[1]),
        "the key asks the pin for the file of the row it has picked"
    );
}

/// The band is the width of the walk behind it: opening it narrows a sound's playlist to the
/// folder's sounds and closing it puts the setting's own order back, and neither is a press —
/// the band turned over is the only thing that asks, and the file the pin is showing is the file
/// it keeps (see `pin_walk::ask_the_walk_again` and `pin_navigation::nav_mode_for`).
#[test]
fn a_band_turned_over_asks_the_walk_behind_it_again() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    // A folder of a test's own, because a walk of one is read off the disk: a sound with a picture
    // beside it is the whole question the band asks.
    let folder = std::env::temp_dir().join("rust-hover-preview-band-asks-the-walk");
    let _ = std::fs::remove_dir_all(&folder);
    std::fs::create_dir_all(&folder).expect("a folder a test can write to");
    for file in ["a.mp3", "b.mp3", "a.png"] {
        std::fs::write(folder.join(file), vec![0u8; 8]).expect("a file a test can write");
    }
    let sound = folder.join("a.mp3");
    let picture = folder.join("a.png");

    // `Audio -> All`: the pin's own walk is the folder end to end, the picture with the sounds.
    let was = {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = (
            config.pin_nav_audio,
            config.pin_nav_image,
            config.pin_mode_audio_shuffle,
        );
        config.pin_nav_audio = PinNavFileTypes::All;
        config.pin_nav_image = PinNavFileTypes::Category;
        config.pin_mode_audio_shuffle = false;
        was
    };
    let mut standing = sound_pin();
    standing.path = sound.clone();
    stand_pin(Some(standing));
    let full = match answer_pin_job(&PinJob::Walk {
        at: sound.clone(),
        step: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
        config: Box::new(CONFIG.lock().expect("the configuration").clone()),
        session: None,
    }) {
        Some(PinPlanned::Walk(walk)) => walk,
        _ => panic!("a walk is always answered"),
    };
    assert!(
        full.order.contains(&picture),
        "under `All` the walk is the folder end to end: {:?}",
        full.order
    );
    stand_pin_walk(Some(full.clone()));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the band builds the rows it has a walk for"
    );

    // A channel of the test's own, stood where the loop's is: what the band asks the loop for is
    // what this test reads (see `show_pinned_file`).
    let (sender, asked) = channel();
    let previous_sender = PREVIEW_SENDER
        .lock()
        .ok()
        .and_then(|mut held| held.replace(sender));

    // The band up: the toggle asks the walk again, and the loop asks the planner for it.
    pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(
        matches!(asked.try_recv(), Ok(PreviewMessage::PinWalkAgain)),
        "a band opened is the walk behind it asked again, with no press to ask it"
    );
    ask_the_walk_again();
    let job = PIN_JOBS
        .0
        .lock()
        .expect("the pin's job slot")
        .take()
        .expect("the walk the band asked for");
    assert!(
        matches!(&job, PinJob::Walk { step, .. } if *step == 0),
        "and it asks for the walk's order rather than for a step along it"
    );
    let Some(PinPlanned::Walk(audio_only)) = answer_pin_job(&job) else {
        panic!("a walk is always answered")
    };
    assert_eq!(
        audio_only.at, sound,
        "a walk with no step asked for stands where it asked from, so nothing the pin shows moves"
    );
    assert!(
        audio_only
            .order
            .iter()
            .all(|path| path.extension().is_some_and(|ext| ext == "mp3")),
        "the band's own walk is the folder's sounds: {:?}",
        audio_only.order
    );
    // What the loop does with that answer: its order goes onto the pin's walk, which is what the
    // rows are built from (see `PreviewMessage::PinAnswered`).
    pin_walk_synced_its_playlist(&audio_only);
    assert!(
        pin_playlist::refresh_the_rows(),
        "and the rows banded are the rows of a walk that moved"
    );
    assert_eq!(
        pin_state().and_then(|state| state.walk().map(|walk| walk.order.clone())),
        Some(audio_only.order.clone()),
        "which is the audio-only walk the band asked for"
    );

    // And the band down: the order the setting asks for comes back, by the same road.
    pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(
        matches!(asked.try_recv(), Ok(PreviewMessage::PinWalkAgain)),
        "a band put down asks the walk again too"
    );
    ask_the_walk_again();
    let job = PIN_JOBS
        .0
        .lock()
        .expect("the pin's job slot")
        .take()
        .expect("the walk the band asked for");
    let Some(PinPlanned::Walk(restored)) = answer_pin_job(&job) else {
        panic!("a walk is always answered")
    };
    assert_eq!(
        restored.order, full.order,
        "a band put down gives back the order the setting asks for"
    );
    assert_eq!(
        restored.at, sound,
        "and the pin is still standing on the file it was showing"
    );
    pin_walk_synced_its_playlist(&restored);
    assert!(
        pin_playlist::refresh_the_rows(),
        "and the rows the band is owed are the folder's again"
    );

    if let Ok(mut held) = PREVIEW_SENDER.lock() {
        *held = previous_sender;
    }
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_nav_audio = was.0;
        config.pin_nav_image = was.1;
        config.pin_mode_audio_shuffle = was.2;
    }
    let _ = std::fs::remove_dir_all(&folder);
    stand_pin(previous_pin);
}

/// The band's two edges are the window's two edges, and a press on either is the height being
/// dragged rather than a row being picked: the card is never resized, and the band is the only thing
/// a pinned sound's window has to be resized by.
///
/// And the card's own window buttons keep the rows they stand on, because a border that answered for
/// them would be a border that ate the close button.
#[test]
fn the_bands_edges_drag_the_height_and_the_cards_buttons_keep_their_rows() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let pin = sound_pin();
    pin_playlist::toggle(&pin);
    let (band_top, band_height) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    let width = pin.window_size().0;
    let card = pin.content;
    stand_pin(Some(pin));
    // The rows are the band's own and are built for whatever order the pin holds; a test that runs
    // after another one of this pin's is a band whose rows are already there (see `THE_PLAYLIST`).
    pin_playlist::refresh_the_rows();

    // The bottom edge is the window's own last rows, and the top edge its first ones.
    assert_eq!(
        pin_playlist::edge_at(width / 2, band_top + band_height - 1),
        Some(pin_chrome::PlaylistEdge::Bottom)
    );
    assert_eq!(
        pin_playlist::edge_at(4, 1),
        Some(pin_chrome::PlaylistEdge::Top),
        "the window's first rows are its top edge, away from the card's own buttons"
    );

    // And where the card's window buttons stand the edge is not the band's at all: they are four
    // buttons in the window's own first rows and they keep them.
    let buttons = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            (0..width)
                .find(|x| pin_audio_window_buttons_at(pin, *x, 1))
                .map(|x| (x, pin_audio_window_buttons_at(pin, x, 1)))
        })
        .expect("the card's own window buttons, which a card always carries");
    assert!(buttons.1, "the button column the test found");
    assert_eq!(
        pin_playlist::edge_at(buttons.0, 1),
        None,
        "the band's top edge took the rows the card's window buttons stand on"
    );

    // A press on the bottom edge is the height's, and dragging it down grows the band and leaves the
    // card exactly where it was.
    let edge = band_top + band_height - 1;
    assert!(
        unsafe { pin_playlist::press(HWND::default(), width / 2, edge, false) },
        "a press on the band's bottom edge is the band's"
    );
    assert!(
        unsafe { pin_playlist::drag(HWND::default(), width / 2, edge + 40) },
        "and the drag that follows it is the band's too"
    );
    assert_eq!(
        pin_playlist::pin_playlist_band(pin_state().expect("the pin").pin().expect("the pin"))
            .expect("the band")
            .1,
        band_height + 40,
        "dragging the bottom edge down did not grow the band by the travel"
    );
    assert_eq!(
        pin_state()
            .expect("the pin")
            .pin()
            .expect("the pin")
            .content,
        card,
        "and the card moved for a drag of the band's bottom edge"
    );
    assert!(unsafe { pin_playlist::release(HWND::default()) });

    // And the band is put back down, so the next test does not come up with one.
    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
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
    let _room = the_room_to_itself();

    // The window is a fixed box, and the box a 400-by-200 window holds is the
    // card at the anchor's font — the 10% share's own — so the share is stood
    // at the anchor for the test and put back when the guard is dropped: what
    // the test is about is where the card's controls sit inside the window
    // they are drawn in, which is the anchor's answer whatever share the
    // machine is set to.
    let _settings = CardFontSettings::stood_the_share_at(DEFAULT_AUDIO_SCALE);

    let pin = sound_pin();

    // The window is the card's box: no caption above it, no transport strip below it.
    assert_eq!(pin.window_box(), pin.content);
    let (width, height) = pin.window_size();
    let (top, _) = pinned_media_band_rows(&pin);
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
    let _room = the_room_to_itself();

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
    // pointer asks for. A video FFmpeg plays is asked for the same way: the band is the whole
    // of its window, and the pin's own window is the one on top for as long as any strip of
    // its chrome is showing (see `pin_overlay_chrome`).
    assert!(pin_transport_kind(Some(MediaType::NativeVideo)));
    assert!(pin_hides_chrome(Some(MediaType::NativeVideo), false));
    assert!(pin_transport_kind(Some(MediaType::Video)));
    assert!(pin_hides_chrome(Some(MediaType::Video), false));
    assert!(!pin_hides_chrome(None, false));

    assert!(
        !pin_transport_kind(Some(MediaType::Audio)),
        "a sound's controls are the card's, drawn on the card's own row"
    );
    assert!(
        !pin_hides_chrome(Some(MediaType::Audio), false),
        "and it is given no caption to hide, so its chrome never comes or goes"
    );

    // The card is its own size and is not framed into one either way (see `pin_frame`).
    assert_eq!(pin_frame(Some(MediaType::Audio)), PinFrame::None);
    assert!(!pin_overlay_chrome(Some(MediaType::Audio), false));

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
    // A picture, a texture, an animation, a drawing, a page of a document or a book, and
    // either video there is — the one the media engine decodes and the one FFmpeg's player
    // plays: the band of every one of them is this app's to paint a strip of chrome over
    // (the player's own window stands behind the pin's for as long as any strip is showing,
    // see `pin_overlay_chrome`).
    for kind in [
        MediaType::StaticImage,
        MediaType::Dds,
        MediaType::AnimatedGif,
        MediaType::Design,
        MediaType::Vector,
        MediaType::Pdf,
        MediaType::NativeVideo,
        MediaType::Video,
    ] {
        assert!(pin_overlay_chrome(Some(kind), false), "{kind:?}");
        assert!(pin_overlay_chrome(Some(kind), true), "{kind:?}");
    }

    // The kinds that keep their chrome in bands around the media: a page that is laid out to
    // whatever box it is given, a sound's card, and a page of HTML the browser runs — the last
    // of which is a window standing in front of the pin rather than a picture under its chrome.
    for kind in [MediaType::Text, MediaType::Archive, MediaType::Audio] {
        assert!(!pin_overlay_chrome(Some(kind), false), "{kind:?}");
        assert!(!pin_overlay_chrome(Some(kind), true), "{kind:?}");
    }
    for kind in [MediaType::EngineSvg, MediaType::EngineFont] {
        assert!(
            pin_overlay_chrome(Some(kind), false),
            "{kind:?} is a picture the pin draws through the Shell's thumbnail"
        );
        assert!(
            !pin_overlay_chrome(Some(kind), true),
            "{kind:?} is also a page of HTML, which runs and holds the band itself"
        );
    }

    assert!(!pin_overlay_chrome(None, false));
    assert!(!pin_overlay_chrome(None, true));
}

/// The caption's popup is the two kinds the Shell composes into the pin and no other, it joins the
/// pin's own band because an owned window inherits none, and a pin that is minimized or off the screen
/// cannot show it (see `the_caption_popup_at`, `the_popup_band` and `the_pin_can_show_it`).
#[test]
fn the_caption_popup_rides_with_the_pin_and_only_for_a_page_the_shell_composes() {
    // Band mirrored both ways, and never `HWND_TOP`.
    assert_eq!(pin_chrome_strips::the_popup_band(true), HWND_TOPMOST);
    assert_eq!(pin_chrome_strips::the_popup_band(false), HWND_NOTOPMOST);

    assert!(pin_chrome_strips::the_pin_can_show_it(false, true));
    assert!(
        !pin_chrome_strips::the_pin_can_show_it(true, true),
        "minimized"
    );
    assert!(
        !pin_chrome_strips::the_pin_can_show_it(false, false),
        "hidden"
    );

    // Every kind but the two the Shell composes keeps the child strip it always had, whatever the two
    // flags say.
    for kind in [
        MediaType::StaticImage,
        MediaType::Dds,
        MediaType::AnimatedGif,
        MediaType::AnimatedApng,
        MediaType::AnimatedWebP,
        MediaType::AnimatedHeif,
        MediaType::AnimatedJxl,
        MediaType::Video,
        MediaType::NativeVideo,
        MediaType::Audio,
        MediaType::Pdf,
        MediaType::Text,
        MediaType::Archive,
        MediaType::Peazip,
        MediaType::Office,
        MediaType::Design,
        MediaType::Vector,
        MediaType::Libre,
        MediaType::Magick,
        MediaType::Calibre,
        MediaType::Comic,
        MediaType::Unplayable,
        MediaType::Loading,
    ] {
        assert!(
            !pin_chrome_strips::the_caption_rides_in_a_popup(true, true, Some(kind)),
            "{kind:?}"
        );
    }
    assert!(!pin_chrome_strips::the_caption_rides_in_a_popup(
        true, true, None
    ));

    // And the two it is for need both flags: a page that runs, and a pin that draws its own picture.
    for kind in [MediaType::EngineSvg, MediaType::EngineFont] {
        assert!(pin_chrome_strips::the_caption_rides_in_a_popup(
            true,
            true,
            Some(kind)
        ));
        assert!(!pin_chrome_strips::the_caption_rides_in_a_popup(
            false,
            true,
            Some(kind)
        ));
        assert!(!pin_chrome_strips::the_caption_rides_in_a_popup(
            true,
            false,
            Some(kind)
        ));
    }
}

/// **A page the browser draws is the pin's picture until it runs**, and the whole of the difference
/// is visible in the window a pin is taken up with: a drawing's caption floats over its picture, in
/// a window that is the media's own box, and a page of HTML's caption is a band above one.
///
/// Both are `EngineSvg`, so nothing but the page itself can say which (see `pin_overlay_chrome`).
#[test]
fn a_pinned_page_floats_its_chrome_over_the_picture_unless_the_page_runs() {
    let _room = the_room_to_itself();

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let was_rendering = crate::CONFIG.lock().ok().map(|config| config.render_html);

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-page-overlay");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let drawing = folder.join("drawing.svg");
    std::fs::write(&drawing, br#"<svg xmlns="http://www.w3.org/2000/svg"/>"#)
        .expect("a written drawing");
    let specimen = folder.join("specimen.ttf");
    std::fs::write(&specimen, b"not really a font").expect("a written specimen");
    let page = folder.join("page.html");
    std::fs::write(&page, b"<!doctype html><title>a page</title>").expect("a written page");

    // An engine page in the media slot, which is the kind a drawing and a page of HTML both are.
    let stand_an_engine_page = || {
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            let mut media = create_loading_media(40, 24);
            media.media_type = MediaType::EngineSvg;
            *current = Some(media);
        }
    };
    let content = (100, 80, 500, 320);

    stand_pin(None);
    stand_an_engine_page();
    let svg = take_up_pinned_window(&drawing, content);
    assert!(
        svg.overlay,
        "a page that does not run is a picture under this app's chrome"
    );
    assert!(
        svg.hides_chrome,
        "so its chrome comes and goes with the pointer"
    );
    assert_eq!(svg.content, content, "the media is the box it was given");
    assert_eq!(
        svg.window_box(),
        content,
        "the caption is over the picture rather than above it"
    );
    let (_, height) = svg.window_size();
    assert_eq!(
        pinned_media_band_rows(&svg),
        (0, height),
        "and the page it frames is drawn the whole window down"
    );

    // A specimen is the same kind of thing, drawn by the same engine on the same road.
    stand_pin(None);
    stand_an_engine_page();
    let font = take_up_pinned_window(&specimen, content);
    assert!(font.overlay && font.hides_chrome, "a specimen is a picture");

    // And a page of HTML, with the switch that draws one on, is the other answer: a window in
    // front of the pin, whose band this app cannot paint a caption over.
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.render_html = true;
    }
    stand_pin(None);
    stand_an_engine_page();
    let running = take_up_pinned_window(&page, content);
    assert!(
        !running.overlay,
        "a page that runs holds the band in front of the pin"
    );
    assert!(!running.hides_chrome, "and its band is never asked away");
    assert_eq!(running.content, content, "the page is the box it was given");
    assert_eq!(
        running.window_box(),
        (content.0, content.1 - svg.caption, content.2, content.3),
        "with the caption in a band of its own above it"
    );

    // A step between the two is a step between the two answers, and the chrome arrives with the
    // answer rather than carried from the file just left (see `take_up_pinned_window`).
    stand_pin(Some(running));
    stand_an_engine_page();
    let back = take_up_pinned_window(&drawing, content);
    assert!(back.overlay && back.hides_chrome, "the drawing again");
    assert!(
        !back.chrome.caption && !back.chrome.bar,
        "and a chrome that is asked for is not one carried over from a band"
    );
    stand_pin(Some(svg));
    stand_an_engine_page();
    let across = take_up_pinned_window(&page, content);
    assert!(!across.overlay && !across.hides_chrome, "the page again");
    assert!(
        across.chrome.caption && across.chrome.bar,
        "and a band is always there, whatever the pin it replaced was hiding"
    );

    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
    if let Some(was_rendering) = was_rendering {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.render_html = was_rendering;
        }
    }
    let _ = std::fs::remove_dir_all(&folder);
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
    let _room = the_room_to_itself();
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
    let _room = the_room_to_itself();
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

/// The middle of one of the two window buttons a pinned sound's card
/// carries in its top margin, in the window's own coordinates: the box
/// is asked of the card's own arithmetic rather than reproduced here,
/// and the middle of the box a press is answered against is inside the
/// box the button is drawn in.
fn a_sound_window_button_point(pin: &PinnedPreview, control: CardControl) -> (i32, i32) {
    let (width, _) = pin.window_size();
    let box_ = audio_preview::control_box(
        control,
        width as u32,
        pin.dpi,
        current_audio_options(),
        true,
    )
    .expect("a box on a card that carries its controls");

    ((box_.left + box_.right) / 2, (box_.top + box_.bottom) / 2)
}

/// The two window buttons a pinned sound's card carries in its top
/// margin, and the cushion each is answered against: the box reaches
/// the window's own top edge above the drawn button and a little below
/// it, which is what makes a mark this small a thing a hand can hit.
#[test]
fn a_sounds_window_buttons_are_answered_over_the_margin_they_stand_in() {
    let _room = the_room_to_itself();

    let mut pin = sound_pin();
    pin.audio_window_buttons = true;
    let (width, _) = pin.window_size();

    for control in [
        CardControl::AlwaysOnTop,
        CardControl::Minimize,
        CardControl::Close,
    ] {
        let box_ = audio_preview::control_box(
            control,
            width as u32,
            pin.dpi,
            current_audio_options(),
            true,
        )
        .expect("a box on a card that carries its controls");
        let middle = (box_.left + box_.right) / 2;

        // The middle of the button, and the window's own top row
        // above it: the cushion the hit box reaches to is part of the
        // button.
        assert_eq!(
            pin_audio_control_at(&pin, middle, (box_.top + box_.bottom) / 2),
            Some(control),
            "the middle of {control:?}'s box is {control:?}"
        );
        assert_eq!(
            pin_audio_control_at(&pin, middle, 0),
            Some(control),
            "and so is the window's own top row, where the cushion reaches"
        );

        // The cushion's last row under the drawn button, and nothing
        // past it.
        assert_eq!(
            pin_audio_control_at(&pin, middle, box_.bottom - 1),
            Some(control),
            "the cushion's last row is still {control:?}"
        );
        assert_eq!(
            pin_audio_control_at(&pin, middle, box_.bottom),
            None,
            "while a row below it is the card's own margin"
        );
    }
}

/// What a press left armed on the pin's card, if anything.
fn the_pressed_control() -> Option<CardControl> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.audio_pressed))
}

/// A press on one of the two window buttons arms it, and a release on
/// it asks the pin's own command — the minimize that shrinks the pin
/// into its bubble and the close that ends the pin, the player and the
/// window together. A press on a button is not a hand carrying the
/// window: the buttons win over the drag, which is what the `dragging`
/// assertion below is about.
#[test]
fn a_sounds_window_buttons_ask_the_pin_for_its_two_ways_out() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    for (control, command) in [
        (CardControl::Minimize, PinCommand::Minimize),
        (CardControl::Close, PinCommand::Close),
    ] {
        // The window buttons are controls only while they are
        // showing, which is the state a hand near the buttons
        // themselves — the hand that is about to press one —
        // leaves the pin in (see `pin_audio_window_buttons_at`).
        let mut shown = sound_pin();
        shown.audio_window_buttons = true;
        stand_pin(Some(shown));
        take_pin_commands();

        let (x, y) = a_sound_window_button_point(&sound_pin(), control);
        assert!(
            unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
            "a hand on {control:?} is the card's to act on"
        );
        assert_eq!(
            the_pressed_control(),
            Some(control),
            "so the press armed it"
        );
        assert!(
            take_pin_commands().is_empty(),
            "and a press alone asks for nothing — the command is the release's to ask for"
        );
        assert!(
            pin_state()
                .and_then(|pinned| pinned.pin().map(|pin| pin.dragging.is_none()))
                .unwrap_or(false),
            "and a press on a button is not a hand carrying the window"
        );

        let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
        assert!(
            unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
            "and the pointer is still on it, so the release is the card's own to answer"
        );
        assert_eq!(
            take_pin_commands(),
            vec![command],
            "which asks for the command the button exists to ask for"
        );
        assert_eq!(the_pressed_control(), None, "and the button is let go of");
    }

    stand_pin(previous_pin);
}

#[test]
fn a_sounds_always_on_top_button_toggles_pin_always_on_top() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    let mut shown = sound_pin();
    shown.audio_window_buttons = true;
    stand_pin(Some(shown));

    let (x, y) = a_sound_window_button_point(&sound_pin(), CardControl::AlwaysOnTop);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on AlwaysOnTop is the card's to act on"
    );
    assert_eq!(
        the_pressed_control(),
        Some(CardControl::AlwaysOnTop),
        "so the press armed it"
    );

    let before = pin_always_on_top();
    let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "the release triggers AlwaysOnTop toggle"
    );
    // The release asks the tray side for the flip rather than making it, because the setting is a row of
    // that side's with a file behind it (see `crate::app::seam::TrayMessage::TogglePinAlwaysOnTop`).
    crate::shell::explorer_hook::answer_what_the_preview_side_asked();
    assert_eq!(
        pin_always_on_top(),
        !before,
        "AlwaysOnTop state was toggled"
    );
    assert_eq!(the_pressed_control(), None, "and the button is let go of");

    // Toggle back to restore initial state
    assert!(unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) });
    assert!(unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) });
    crate::shell::explorer_hook::answer_what_the_preview_side_asked();
    assert_eq!(pin_always_on_top(), before, "restored to original state");

    stand_pin(previous_pin);
}

/// The two window buttons come and go with the hand's every move, on both
/// axes: a pointer move onto them — in the top band, the band as tall as the
/// buttons' own hit boxes, and across the columns they stand in — brings them
/// up and asks for the card again, and a move off them puts them away and asks
/// for it again too.
#[test]
fn a_hand_near_the_top_border_brings_a_sounds_window_buttons_up() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();

    stand_pin(Some(sound_pin()));
    let pin = sound_pin();
    let band = audio_preview::window_button_band(96, current_audio_options());
    let (x, on_a_button) = a_sound_window_button_point(&pin, CardControl::AlwaysOnTop);

    // A hand in the top band but hard against the card's own left
    // border: the band is the whole width and the buttons are not, so
    // this is a hand near neither and nothing changes.
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);
    unsafe { pinned_mouse_move(HWND(0x1000 as *mut _), 1, 1) };
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_window_buttons)),
        Some(false),
        "a hand in the top band at the card's left is a hand near no button"
    );
    assert!(
        !AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and asks the card for nothing"
    );

    // A hand on the buttons themselves: they come up, and the card is
    // asked for again, which is what shows them.
    unsafe { pinned_mouse_move(HWND(0x1000 as *mut _), x, on_a_button) };
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_window_buttons)),
        Some(true),
        "a hand on a window button's own box is a hand asking for the two window buttons"
    );
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and the buttons coming up is a change of the card, which is what shows them"
    );

    // A hand one row below the band, at the buttons' own column: out of
    // it, so the buttons go away and the name's corner is the card's own
    // again.
    unsafe { pinned_mouse_move(HWND(0x1000 as *mut _), x, band) };
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_window_buttons)),
        Some(false),
        "a hand below the band is a hand asking for the corner back"
    );
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and their going away is a change of the card as well"
    );

    stand_pin(previous_pin);
}

/// The tick asks the same question of the pointer the move asks: a
/// screen cursor on the buttons brings a sound's two window buttons
/// up, and a cursor in the top band at the card's own left — or below
/// the band, or off the window — puts them away.
#[test]
fn the_tick_asks_for_the_window_buttons_where_a_hand_is_near_the_top_border() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();

    let pin = sound_pin();
    let window = pin.window_box();
    let band = audio_preview::window_button_band(96, current_audio_options());
    let (x, y) = a_sound_window_button_point(&pin, CardControl::AlwaysOnTop);
    let on_a_button = (window.0 + x, window.1 + y);
    stand_pin(Some(pin));

    // A screen cursor in the top band, in the window's own left
    // margin: in the band and across the card from the buttons, so
    // nothing changes.
    let in_the_band = (window.0 + 10, window.1 + 1);
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(pin_audio_hover_refresh(pin, Some(in_the_band)))
    });
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_window_buttons)),
        Some(false),
        "a screen cursor in the top band at the card's left is near no button"
    );
    assert_eq!(changed, Some(false), "and nothing is asked of the card");

    // A cursor on the buttons themselves is a hand near them: they come
    // up and the card is asked for.
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(pin_audio_hover_refresh(pin, Some(on_a_button)))
    });
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_window_buttons)),
        Some(true),
        "a screen cursor on a window button is a hand near the buttons"
    );
    assert_eq!(
        changed,
        Some(true),
        "and the tick owes the card a repaint, which is what shows them"
    );
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and the card is asked for by name"
    );

    // The same cursor resting there is not a change worth a repaint.
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(pin_audio_hover_refresh(pin, Some(on_a_button)))
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

    // A cursor one row below the band, at the buttons' own column: the
    // buttons go away.
    let below_the_band = (window.0 + x, window.1 + band);
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(pin_audio_hover_refresh(pin, Some(below_the_band)))
    });
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.audio_window_buttons)),
        Some(false),
        "a screen cursor below the band is a hand asking for the corner back"
    );
    assert_eq!(
        changed,
        Some(true),
        "and the tick owes the card a repaint again"
    );
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and the card is asked for by name"
    );

    // A cursor off the window is a hand nowhere near the border:
    // nothing changes, because the buttons are already away.
    let off_the_window = (window.0 - 10, window.1 + 1);
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(pin_audio_hover_refresh(pin, Some(off_the_window)))
    });
    assert_eq!(changed, Some(false), "a hand off the window is nothing new");

    stand_pin(previous_pin);
}

/// The two axes a hand asks for a sound's card window buttons on, at a display's own
/// scale: the band as tall as the buttons' own hit boxes, and the columns from the
/// leftmost button to the rightmost widened by a hand's margin. A hand in the band at
/// the card's own left is near neither; a hand on a button is near both; and a hand a
/// pixel past the margin is past the buttons, whatever its row.
#[test]
fn a_sounds_window_buttons_are_asked_for_on_both_axes() {
    // The options the span and the band are measured from are the configuration's, and this test
    // asks them twice — once for the test's own `right` and once inside `pin_audio_window_buttons_at`
    // — so a test that folded a card with another scale in between is a span the two disagree about.
    let _room = the_room_to_itself();
    for dpi in [96u32, 192] {
        let mut pin = sound_pin();
        pin.dpi = dpi;

        let width = (pin.content.2 - pin.content.0).max(1);
        let options = pinned_audio_options(&pin);
        let band = audio_preview::window_button_band(dpi, options);
        let (left, right) = audio_preview::window_button_span(dpi, options, width);
        let button =
            audio_preview::control_box(CardControl::AlwaysOnTop, width as u32, dpi, options, true)
                .expect("a box on a card that carries its controls");
        let x = (button.left + button.right) / 2;
        let y = (button.top + button.bottom) / 2;

        assert!(
            !pin_audio_window_buttons_at(&pin, 1, 1),
            "at {dpi} dpi: the card's top-left is in the band and across the card from the buttons"
        );
        assert!(
            pin_audio_window_buttons_at(&pin, x, y),
            "at {dpi} dpi: the always-on-top button's own box is near the buttons"
        );
        assert!(
            pin_audio_window_buttons_at(&pin, left, 1),
            "at {dpi} dpi: the margin's own edge is still near them"
        );
        assert!(
            !pin_audio_window_buttons_at(&pin, left - 1, 1),
            "at {dpi} dpi: a pixel past the margin is past the buttons"
        );
        assert!(
            pin_audio_window_buttons_at(&pin, right - 1, 1),
            "at {dpi} dpi: the other margin's own edge is still near them"
        );
        assert!(
            !pin_audio_window_buttons_at(&pin, right, 1),
            "at {dpi} dpi: and a pixel past it is past the buttons"
        );
        assert!(
            !pin_audio_window_buttons_at(&pin, x, band),
            "at {dpi} dpi: a hand at the buttons' column but below the band is past them"
        );
    }
}

/// A press on a window button that slides off it asks for nothing: the
/// release is answered where the pointer is rather than where the press
/// was, and no pressed state is left standing.
#[test]
fn a_press_that_slides_off_a_sounds_window_button_asks_for_nothing() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    // The window buttons are controls only while they are
    // showing, which is the state a hand near the buttons
    // themselves — the hand that is about to press one —
    // leaves the pin in (see `pin_audio_window_buttons_at`).
    let mut shown = sound_pin();
    shown.audio_window_buttons = true;
    stand_pin(Some(shown));
    let (x, y) = a_sound_window_button_point(&sound_pin(), CardControl::Minimize);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the minimize button is the card's to act on"
    );

    // The pointer walks off the button and off the card entirely.
    unsafe { pinned_mouse_move(HWND(0x1000 as *mut _), -1000, -1000) };
    let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), -1000, -1000, &window) },
        "the release is the card's to answer, wherever the pointer came to rest"
    );

    assert_eq!(
        take_pin_command(),
        None,
        "a press that slid off asks for nothing"
    );
    assert_eq!(the_pressed_control(), None, "and no button is left held");

    stand_pin(previous_pin);
}

/// The bubble a pinned sound's minimize leaves stands where the button
/// did: the minimize's own box at the top-right corner of the window,
/// which is the corner every other pin's bubble stands in.
#[test]
fn a_sounds_bubble_takes_the_place_of_the_minimize_button() {
    let _room = the_room_to_itself();

    let pin = sound_pin();
    let window = pin.window_box();

    let box_ = pinned_minimize_box(&pin);
    let (width, _) = pin.window_size();
    let button = audio_preview::control_box(
        CardControl::Minimize,
        width as u32,
        pin.dpi,
        current_audio_options(),
        true,
    )
    .expect("a minimize on a card that carries its controls");

    assert_eq!(
        box_,
        (
            window.0 + button.left,
            window.1 + button.top,
            window.0 + button.right,
            window.1 + button.bottom,
        ),
        "the bubble is centred on the minimize's own box"
    );

    // The box the bubble is centred on begins at the window's own top
    // edge — its hit box reaches it — so the bubble stands in the
    // window's top-right corner: its centre is in the top half and the
    // right half of the window, a button's own width from each edge.
    assert_eq!(
        box_.1, window.1,
        "the bubble's box begins at the window's own top edge"
    );
    let centre = ((box_.0 + box_.2) / 2, (box_.1 + box_.3) / 2);
    assert!(
        centre.0 > (window.0 + window.2) / 2,
        "and its centre is in the window's right half"
    );
    assert!(centre.1 < (window.1 + window.3) / 2, "and in its top half");

    // Every other kind keeps the bubble it has always had: one centred
    // on the caption's own minimize button, which is a button a kind
    // with a caption carries — a sound is the one kind whose window
    // has no caption, and whose bubble stands where the card's own
    // minimize did instead.
    let picture = overlay_pin((100, 100, 500, 300), PinChrome::always());
    let picture_window = picture.window_box();
    let caption_minimize = pin_chrome::button_boxes(
        picture_window.2 - picture_window.0,
        picture.caption,
        picture.dpi,
        picture.frame != PinFrame::None,
    )
    .into_iter()
    .find(|button| button.kind == pin_chrome::CaptionButton::Minimize)
    .expect("a caption carries a minimize");
    assert_eq!(
        pinned_minimize_box(&picture),
        (
            picture_window.0 + caption_minimize.rect.left,
            picture_window.1 + caption_minimize.rect.top,
            picture_window.0 + caption_minimize.rect.right,
            picture_window.1 + caption_minimize.rect.bottom,
        ),
        "a picture's bubble is the caption's own minimize, as it has always been"
    );
}

/// A play pressed on the bubble of a collapsed sound, with `Pin Mode → Pause on Minimize` on:
/// the park is given back the road a restore gives it back, and the tick that parked it does not
/// park the sound again over the hand that asked for it.
///
/// Its old road was a dead one, which was the whole of the fault: the park ends the player, so a
/// press found no player behind the card and let go of nothing at all.
#[test]
fn a_play_on_a_collapsed_pins_bubble_begins_the_parked_sound_again() {
    let _room = the_room_to_itself();
    let _switches = PausePreviewSettings::stood_at(true, true);

    let path = a_remembered_sound("rust-hover-preview-pin-bubble-play");

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);

    // The card the pin is showing, with no player behind it: the park ended the one the collapse
    // found playing (see `bubble_playback_to_park`).
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = load_audio_card(&path, 240, 80, 96);
    }

    let mut pin = sound_pin();
    pin.path = path.clone();
    pin.collapsed = true;
    pin.bubble_pause = Some(BubblePause::Player(12.5));
    stand_pin(Some(pin));

    let mut started = None;
    let mut offset = 0.0;
    toggle_pinned_by_key(&mut started, &mut offset, &mut None);

    assert!(
        started.is_some(),
        "a player of this app's is begun again: a park ended is a park a press gives back"
    );
    assert_eq!(
        offset, 12.5,
        "and it is begun at the second the park wrote down"
    );
    assert!(
        pin_bubble_pause().is_none(),
        "the bubble holds no park for the sound it has just given back"
    );
    let running = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.video_process.is_some()))
        .unwrap_or(false);
    assert!(running, "and it is behind the card");

    // And the tick that parked it stands down over it.
    settle_bubble_playback(&mut started, &mut offset);
    assert!(
        started.is_some(),
        "the switches do not park the sound back over the hand that asked for it"
    );
    assert_eq!(offset, 12.5, "nor take the clock the card is drawn from");
    assert!(
        pin_bubble_pause().is_none(),
        "and the bubble still holds no park"
    );

    // A second press is the ordinary hold, and it outlives the pin coming back up: a hold the hand
    // asked for is not a park the collapse made.
    toggle_pinned_by_key(&mut started, &mut offset, &mut None);
    assert!(
        started.is_none(),
        "a player of this app's is ended by the hold"
    );

    restore_pin(&a_restored_window(), PinRestore::Pressed);
    let collapsed = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.collapsed))
        .expect("a pin is up");
    assert!(!collapsed, "the pin is up again");
    assert!(
        started.is_none(),
        "and the hold is still there: what is not playing is not parked"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    let _ =
        std::fs::remove_dir_all(std::env::temp_dir().join("rust-hover-preview-pin-bubble-play"));
}

/// The card's command bar is asked out by a hand near the row its controls stand in, and by
/// nothing else: a hand on the artwork's own rows — the top-left of the card — leaves the row
/// unlit, and a hand one row past the band's margin is no longer near it. Asked at two display
/// scales because the row and its margin are read at the display's own scale.
#[test]
fn a_hand_on_the_artwork_does_not_ask_the_cards_command_bar_out() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    for dpi in [96u32, 192] {
        let mut pin = sound_pin();
        pin.dpi = dpi;
        let window = pin.window_box();
        let width = (pin.content.2 - pin.content.0).max(1) as u32;
        let play = audio_preview::control_box(
            CardControl::Play,
            width,
            dpi,
            current_audio_options(),
            true,
        )
        .expect("a button on the card's own row");
        let bar = audio_preview::control_box(
            CardControl::Seek,
            width,
            dpi,
            current_audio_options(),
            true,
        )
        .expect("the card's own bar");
        let (top, bottom) = pin_audio_command_bar_band(&pin).expect("a band over the row's rows");
        let margin = logical_px(dpi, AUDIO_COMMAND_BAR_REVEAL_PIXELS);
        assert_eq!(
            (top, bottom),
            (
                play.top.min(bar.top) - margin,
                bar.bottom.max(play.bottom) + margin
            ),
            "the band is the row widened by the margin at this display's scale"
        );

        stand_pin(Some(pin));

        let hovered_at = |x: i32, y: i32| {
            pin_state().and_then(|mut pinned| {
                let pin = pinned.pin_mut()?;
                let _ = pin_audio_hover_refresh(pin, Some((x, y)));
                Some(pin.audio_hovered)
            })
        };

        // The artwork's own top-left corner: a hand on the card, nowhere near its row.
        assert_eq!(
            hovered_at(window.0 + 1, window.1 + 1),
            Some(None),
            "a hand on the artwork asks for nothing at {dpi} dpi"
        );

        // The row itself, where the bar is drawn: the hand is near it and the row is lit.
        assert_eq!(
            hovered_at(
                window.0 + (bar.left + bar.right) / 2,
                window.1 + (bar.top + bar.bottom) / 2
            ),
            Some(Some(CardControl::Seek)),
            "a hand on the command bar's own row is a hand on its bar at {dpi} dpi"
        );

        // And the four window buttons in the card's own corner, the detach among them, are answered
        // wherever they are drawn: they stand in a band of their own at the top of the card, above
        // the bar's row, and a hand on one is a hand on that button (see
        // `pin_audio_control_revealed`).
        for control in [
            CardControl::AlwaysOnTop,
            CardControl::Detach,
            CardControl::Minimize,
            CardControl::Close,
        ] {
            let button =
                audio_preview::control_box(control, width, dpi, current_audio_options(), true)
                    .expect("a window button in the card's own corner");
            assert_eq!(
                hovered_at(
                    window.0 + (button.left + button.right) / 2,
                    window.1 + (button.top + button.bottom) / 2
                ),
                Some(Some(control)),
                "a hand on {control:?}'s own box is a hand on it at {dpi} dpi"
            );
        }

        // One row past the band's margin: near the row, and no longer near enough.
        assert_eq!(
            hovered_at(window.0 + (bar.left + bar.right) / 2, window.1 + top - 1),
            Some(None),
            "a hand past the margin is a hand on the artwork at {dpi} dpi"
        );
    }

    stand_pin(previous_pin);
}

/// The Explorer hook's own reading of a sound card's button: the screen point the always-on-top
/// button's press lands on is the pin's chrome, not a pick of the file the pin stands over — a
/// press on it must not arrive with the file under the card (see `pin_button_covers`).
///
/// A sound has no caption, so the controls it carries in its card's place are the only buttons it
/// has: without them asked here, the hook reads the press as a pick of the listing and the preview
/// switches to whatever file is under the card.
#[test]
fn a_sounds_always_on_top_button_is_the_pins_own_and_not_a_file_under_it() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    let pin = sound_pin();
    let window = pin.window_box();
    let (_, height) = pin.window_size();
    stand_pin(Some(pin));

    let (x, y) = a_sound_window_button_point(&sound_pin(), CardControl::AlwaysOnTop);
    publish_pin_button_boxes();
    assert!(
        pin_button_covers((window.0 + x, window.1 + y)),
        "the always-on-top button is the pin's own chrome"
    );

    // And the card's artwork, which is not a button of the pin's, keeps the pass-through a click
    // on a pinned window has.
    assert!(
        !pin_button_covers((window.0 + 1, window.1 + height / 2)),
        "while the card's artwork is the listing's, not the chrome's"
    );

    stand_pin(previous_pin);
}

/// A band's search field is the pin's own field, and the three things a hand does with it are one
/// answer each: a press inside it puts the caret there and picks no row, a character typed into it
/// hides the rows its query does not name, and `Esc` gives the whole list back and lets the caret go.
///
/// The pin is stood on a walk of three sounds rather than on a folder of the disk's own, because the
/// rows are what a query filters and a folder a test cannot see the whole of is a folder it cannot
/// assert on.
#[test]
fn a_query_in_the_bands_field_hides_the_rows_it_does_not_name() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let folder = PathBuf::from("a folder of three sounds for the bands own search");
    let sounds = [
        folder.join("a first mixdown.mp3"),
        folder.join("an interlude.mp3"),
        folder.join("a last mixdown.mp3"),
    ];

    let mut pin = sound_pin();
    pin.path = sounds[1].clone();
    pin_playlist::toggle(&pin);
    let (band_top, band_height) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    let window = pin.window_size();
    stand_pin(Some(pin));
    stand_pin_walk(Some(PinStep {
        at: sounds[1].clone(),
        from: sounds[1].clone(),
        step: 1,
        order: sounds.to_vec(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the band builds the rows of the walk"
    );
    assert_eq!(
        pin_playlist::the_rows_in_sight(),
        vec![0, 1, 2],
        "a band with no query draws the whole walk"
    );

    // The field's own box, in the window's rows: it is inside the band and nowhere else.
    let (field_top, field_bottom) =
        pin_playlist::the_field_rows().expect("a band with its field on screen");
    assert!(
        field_top >= band_top && field_bottom <= window.1,
        "the field is drawn inside the window the band is in"
    );
    let field = (field_top + field_bottom) / 2;

    // A press on the field is the caret's and picks nothing — and it is a press on a field a hand has
    // reached for, since a field parked under the edge is a field nothing can press.
    assert!(
        unsafe { pin_playlist::hover(40, field) },
        "the hand reaches for the field"
    );
    for _ in 0..14 {
        pin_playlist::slide_the_search_on(20);
    }
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(_, picked, _)| picked),
        Some(None),
        "the band holds no pick before the hand has pressed anything"
    );
    assert!(
        unsafe { pin_playlist::press(HWND::default(), 100, field, false) },
        "a press on the field is the band's rather than the card's"
    );
    assert_eq!(
        pin_playlist::the_query_in_hand(),
        Some((String::new(), true)),
        "the press put the caret in an empty field and focused it"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(_, picked, _)| picked),
        Some(None),
        "and a press on the field picks no row of the list"
    );

    // Typed through the road `WM_CHAR` answers the field down: the rows the query does not name are
    // hidden, and the row the pin is showing is one of those that go.
    pin_playlist::type_into_the_search('m');
    assert_eq!(
        pin_playlist::the_query_in_hand(),
        Some(("m".to_string(), true)),
        "the character is the query's"
    );
    assert_eq!(
        pin_playlist::the_rows_in_sight(),
        vec![0, 2],
        "only the names that hold the query are drawn"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(_, _, scroll)| scroll),
        Some(0),
        "and the band is scrolled to the first row it still draws"
    );

    // A row the query hides is a row the pick cannot be on, so the arrows step the positions the
    // query lets through: the row the pin is showing is one of those it hides, so the first press
    // starts on the band's first row in sight and lands one past it.
    assert!(
        pin_playlist::step_the_pick(1),
        "the band takes the arrow over it"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(_, picked, _)| picked),
        Some(Some(2)),
        "the walk's third row, which is the second of the two the query lets through"
    );
    assert!(
        pin_playlist::step_the_pick(1),
        "and a second one is still the band's"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(_, picked, _)| picked),
        Some(Some(2)),
        "which holds on the last row in sight rather than stepping onto one the query hides"
    );

    // `Esc` gives the whole list back and takes the caret out of the field.
    assert!(
        pin_playlist::key_while_the_search_has_focus(VK_ESCAPE.0 as i32),
        "the field has the caret, so `Esc` is its"
    );
    assert_eq!(
        pin_playlist::the_rows_in_sight(),
        vec![0, 1, 2],
        "the query was emptied and the whole walk is drawn again"
    );
    assert_eq!(
        pin_playlist::the_query_in_hand(),
        Some((String::new(), false)),
        "and the field let the caret go"
    );
    assert!(
        !pin_playlist::key_while_the_search_has_focus(VK_BACK.0 as i32),
        "so the keys are the pin's again"
    );

    // And the window's own box is the one it was: the field is inside the band rather than rows the
    // window grew for it.
    let (_, after) =
        pin_playlist::pin_playlist_band(pin_state().expect("the pin").pin().expect("the pin"))
            .expect("the band");
    assert_eq!(
        after, band_height,
        "the band's height is the height it was given, query or no query"
    );

    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// A press anywhere on the pin that is not the field is the hand leaving it: the query it typed stays
/// applied — the list under the press is the list it asked for — and the caret goes, so the keys are
/// the pin's again.
///
/// And the field itself does not move or change size for it: it is a constant of the layout rather than
/// a box the query grows.
#[test]
fn a_press_beside_the_field_blurs_it_and_leaves_the_query_it_typed() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let folder = PathBuf::from("a folder of two sounds for a blur of the bands own search");
    let sounds = [folder.join("a shown.mp3"), folder.join("an interlude.mp3")];

    let mut pin = sound_pin();
    pin.path = sounds[0].clone();
    pin_playlist::toggle(&pin);
    stand_pin(Some(pin));
    stand_pin_walk(Some(PinStep {
        at: sounds[0].clone(),
        from: sounds[0].clone(),
        step: 1,
        order: sounds.to_vec(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the band builds the rows of the walk"
    );

    let (field_top, field_bottom) =
        pin_playlist::the_field_rows().expect("a band with its field on screen");
    let field = (field_top + field_bottom) / 2;
    assert!(unsafe { pin_playlist::press(HWND::default(), 100, field, false) });

    pin_playlist::type_into_the_search('i');
    assert_eq!(
        pin_playlist::the_rows_in_sight(),
        vec![1],
        "the row the query does not name is hidden"
    );

    // A press on the card above the band is not the band's at all, and it is still the press that
    // takes the caret out of the field.
    let band_top = pin_state()
        .and_then(|pinned| pinned.pin().and_then(pin_playlist::pin_playlist_band))
        .map(|(top, _)| top)
        .expect("the band");
    assert!(
        !unsafe { pin_playlist::press(HWND::default(), 100, band_top - 1, false) },
        "a press above the band is the card's"
    );
    pin_playlist::blur_the_search();
    assert_eq!(
        pin_playlist::the_query_in_hand(),
        Some(("i".to_string(), false)),
        "the query is the one the hand typed and the caret has gone"
    );
    assert_eq!(
        pin_playlist::the_rows_in_sight(),
        vec![1],
        "and the list under the press is the list the query asked for"
    );

    // The field's box is the same box, whatever the band's own height has become.
    let field_rows = pin_playlist::the_field_rows().expect("the field");
    let edge = pin_state()
        .and_then(|pinned| pinned.pin().and_then(pin_playlist::pin_playlist_band))
        .map(|(top, height)| (top, top + height))
        .expect("the band");
    assert!(
        unsafe { pin_playlist::press(HWND::default(), 100, edge.1 - 1, false) },
        "a press on the band's bottom edge is the drag that sets its height"
    );
    assert!(
        unsafe { pin_playlist::drag(HWND::default(), 100, edge.1 + 40) },
        "and the drag that follows it is the band's"
    );
    assert!(unsafe { pin_playlist::release(HWND::default()) });

    // The drag was taken one row above the band's last and carried forty rows past it, so the band
    // grew by the whole of that travel and the field — its last rows — moved with it.
    let grown = pin_playlist::the_field_rows().expect("the field");
    assert_eq!(
        grown.1 - field_rows.1,
        41,
        "the field's bottom moved with the band's last rows"
    );
    assert_eq!(
        grown.1 - grown.0,
        field_rows.1 - field_rows.0,
        "and the field itself is the same height it was"
    );

    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// A band of twelve sounds with the pin stood on the first of them, which is the only list a test of
/// the band's own scroll can count rows on: a folder of the disk's own is a folder a test cannot see
/// the whole of.
fn a_walk_of_twelve_sounds(mut pin: PinnedPreview) -> Vec<PathBuf> {
    let folder = PathBuf::from("a folder of twelve sounds for the bands own slide");
    let sounds: Vec<PathBuf> = (0..40)
        .map(|at| folder.join(format!("a track {at:02} of the folder.mp3")))
        .collect();

    pin.path = sounds[0].clone();
    stand_pin(Some(pin));
    stand_pin_walk(Some(PinStep {
        at: sounds[0].clone(),
        from: sounds[0].clone(),
        step: 1,
        order: sounds.clone(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the band builds the rows of the walk"
    );

    sounds
}

/// **The field is the pointer's**: it comes into sight when a hand reaches into the band and it parks
/// when the hand leaves — and it parks by travelling rather than by jumping, which is the one thing a
/// slide is for.
#[test]
fn the_bands_field_slides_in_with_the_pointer_and_parks_when_it_goes() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let folder = PathBuf::from("a folder of three sounds for the bands own slide");
    let sounds = [
        folder.join("a first mixdown.mp3"),
        folder.join("a middle take.mp3"),
        folder.join("a last mixdown.mp3"),
    ];

    let mut pin = sound_pin();
    pin.path = sounds[0].clone();
    pin_playlist::toggle(&pin);
    let (band_top, _) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    stand_pin(Some(pin));
    stand_pin_walk(Some(PinStep {
        at: sounds[0].clone(),
        from: sounds[0].clone(),
        step: 1,
        order: sounds.to_vec(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(pin_playlist::refresh_the_rows(), "the band builds its rows");

    // **A band opens parked.** The field is a hand's rather than a band's, and a hand that has not
    // reached into it yet has not asked for it.
    assert_eq!(
        pin_playlist::the_search_slide().map(|(slide, _)| slide),
        Some(0.0),
        "the field is out of sight until the pointer comes into the band"
    );

    // **A row of the list is not the field's.** A hand reading the rows of a band is a hand on the list,
    // and a field that came up under every row it read is a field that covered the row it was reading:
    // what the pointer is asked for is the strip the field lives in, and nothing above it.
    let row = pin_chrome::playlist_row_height(1.0) / 2;
    let _ = unsafe { pin_playlist::hover(40, band_top + row) };
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(Some(0)),
        "a pointer on the band's first row is a hand on that row"
    );
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((0.0, false)),
        "and it asks for no field: the rows are what the hand is reading"
    );

    // **The field's own strip is what a hand reaches for.** A pointer at it is a hand asking for the
    // field: the target moves, and the paint that follows travels toward it rather than landing on it.
    let (field_top, field_bottom) =
        pin_playlist::the_field_rows().expect("a band with its field on screen");
    assert!(
        unsafe { pin_playlist::hover(40, (field_top + field_bottom) / 2) },
        "a pointer at the field's strip is the field's own"
    );
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((0.0, true)),
        "the field is wanted and has not travelled yet"
    );
    assert!(
        pin_playlist::slide_the_search_on(20),
        "and the tick after it moves it"
    );
    let (part_way, _) = pin_playlist::the_search_slide().expect("the field");
    assert!(
        part_way > 0.0 && part_way < 1.0,
        "a field on its way in is drawn part way rather than blinking, not at {part_way}"
    );

    // A pointer that has left the band lets it go, and it travels back the same way.
    assert!(
        unsafe { pin_playlist::hover(40, 0) },
        "a pointer over the card is the band's answer that it is in it no longer"
    );
    let (held, wanted) = pin_playlist::the_search_slide().expect("the field");
    assert!(!wanted, "the pointer gone is the field no longer wanted");
    assert!(
        held > 0.0,
        "and it is not taken away where it stands: it has a way to go back"
    );

    // **And it comes back to exactly nothing**, so that a parked field asks for no paint: twelve ticks
    // of a loop's own cadence is longer than the whole slide, so the last of them lands it.
    for _ in 0..14 {
        pin_playlist::slide_the_search_on(20);
    }
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((0.0, false)),
        "a parked field rests at nothing rather than a hair short of it"
    );

    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// **A caret belongs to the hand that reached for the field, and a query outlives them both.** A
/// pointer that has walked away from a field it pressed is a pointer no key is going to be typed
/// from: the caret goes with it, and a field with nothing in it parks — the one state that used to
/// hold a band open with no pointer near it and nothing in the field. What a query keeps is its own:
/// a list with rows missing and no field to explain it is a list the hand cannot get back.
#[test]
fn a_caret_goes_with_the_pointer_and_a_query_holds_the_field_open_alone() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let folder = PathBuf::from("a folder of three sounds for a field held open");
    let sounds = [
        folder.join("a first mixdown.mp3"),
        folder.join("a middle take.mp3"),
        folder.join("a last mixdown.mp3"),
    ];

    let mut pin = sound_pin();
    pin.path = sounds[0].clone();
    pin_playlist::toggle(&pin);
    stand_pin(Some(pin));
    stand_pin_walk(Some(PinStep {
        at: sounds[0].clone(),
        from: sounds[0].clone(),
        step: 1,
        order: sounds.to_vec(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(pin_playlist::refresh_the_rows(), "the band builds its rows");

    // The hand reaches for the field's strip, the field comes in to meet it, and a press takes it.
    let (field_top, field_bottom) =
        pin_playlist::the_field_rows().expect("a band with its field on screen");
    assert!(
        unsafe { pin_playlist::hover(40, (field_top + field_bottom) / 2) },
        "the hand reaches for the field"
    );
    for _ in 0..14 {
        pin_playlist::slide_the_search_on(20);
    }
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((1.0, true)),
        "the field is in sight whole before a hand presses it"
    );
    assert!(unsafe {
        pin_playlist::press(HWND::default(), 100, (field_top + field_bottom) / 2, false)
    });
    // **The pointer leaves the band**, which the tick that finds the cursor outside it is what says: a
    // move the window is sent and the cursor the tick reads are one answer (see `set_the_hover`). The
    // caret is the pointer's, so it goes with it.
    assert!(
        pin_playlist::the_pointer_has_left_the_band(),
        "a caret let go of is a field that changed"
    );
    assert_eq!(
        pin_playlist::the_query_in_hand().map(|(_, focused)| focused),
        Some(false),
        "the caret is the pointer's, and the pointer has gone"
    );
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((1.0, false)),
        "and with no query in it the field is wanted by nothing: it starts back"
    );

    // And it parks, which is the state a hand that has walked away comes back to.
    for _ in 0..14 {
        pin_playlist::slide_the_search_on(20);
    }
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((0.0, false)),
        "a field nothing is reaching for rests at nothing"
    );

    // **A query holds it open by itself**, with no pointer and no caret: it is what is hiding rows.
    pin_playlist::type_into_the_search('m');
    assert_eq!(
        pin_playlist::the_search_slide().map(|(_, wanted)| wanted),
        Some(true),
        "a query hiding rows is a field the hand has to be able to see and edit"
    );
    assert_eq!(
        pin_playlist::the_query_in_hand(),
        Some(("m".to_string(), false)),
        "and the caret it was typed with is gone, which is the state the field is held in"
    );

    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// **A hand that is only near the band is reaching for its field.** The field stands on the band's last
/// row, which is the window's own bottom edge, so the gesture a hand makes for it arrives from under
/// that edge or from beside it: those pixels are the field's too, and a field that waited for the
/// pointer to be inside the band is one that comes up a row too late.
///
/// What they are not is a row of the band. The row a name travels for, and the row a press picks, are
/// the pointer's inside the band's own rect and nowhere else — the reach is the field's alone.
#[test]
fn a_pointer_near_the_band_reaches_for_its_field_and_is_on_no_row_of_it() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let pin = sound_pin();
    pin_playlist::toggle(&pin);
    let (band_top, band_height) =
        pin_playlist::pin_playlist_band(&pin).expect("the band the hand just opened");
    let reach = pin_chrome::playlist_reach(pin.dpi as f32 / 96.0);
    stand_pin(Some(pin));
    pin_playlist::refresh_the_rows();

    for _ in 0..14 {
        pin_playlist::slide_the_search_on(20);
    }
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((0.0, false)),
        "the field is parked before the pointer comes near it"
    );

    // **A row of the list is not the field's.** The strip at the band's bottom is what a hand reaches
    // for, so a hand reading the band's own rows asks for nothing: a field that came up under every row
    // the hand read is a field that covered the row it was reading.
    let _ = unsafe { pin_playlist::hover(40, band_top + 4) };
    assert_eq!(
        pin_playlist::the_search_slide().map(|(_, wanted)| wanted),
        Some(false),
        "and a row of the list asks for no field"
    );

    // Four pixels under the window's own last row, which is the band's: a hand at the field rather
    // than in a row of the list.
    assert!(
        unsafe { pin_playlist::hover(40, band_top + band_height + 4) },
        "a pointer beside the band is a change the band owes a paint for"
    );
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((0.0, true)),
        "the field is wanted by a hand that is only reaching for it"
    );
    assert_eq!(
        pin_playlist::the_band_in_hand().map(|(hovered, _, _)| hovered),
        Some(None),
        "and the hand is on no row of the band, which is its own rect and nothing beside it"
    );

    // Past the reach the band is nobody's: a pointer that is not coming at the field leaves it.
    assert!(
        unsafe { pin_playlist::hover(40, band_top + band_height + reach + 1) },
        "a pointer past the reach is a change of its own"
    );
    assert_eq!(
        pin_playlist::the_search_slide().map(|(_, wanted)| wanted),
        Some(false),
        "and the field is wanted by nothing out there"
    );

    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// **A press where the field would be is not a press in the field** while it is parked: the strip is
/// reserved whether the pill is drawn in it or not, so the rows above it never jump — which means the
/// strip is a place a hand can press, and what it must not get is a caret it cannot see.
#[test]
fn a_press_where_the_field_would_be_focuses_nothing_while_it_is_parked() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let pin = sound_pin();
    pin_playlist::toggle(&pin);
    stand_pin(Some(pin));

    for _ in 0..14 {
        pin_playlist::slide_the_search_on(20);
    }
    assert_eq!(
        pin_playlist::the_search_slide(),
        Some((0.0, false)),
        "the field is parked before the press"
    );

    let (field_top, field_bottom) =
        pin_playlist::the_field_rows().expect("a band with its field on screen");
    assert!(
        unsafe { pin_playlist::press(HWND::default(), 100, (field_top + field_bottom) / 2, false) },
        "a press in the band is the band's"
    );
    assert_eq!(
        pin_playlist::the_query_in_hand(),
        Some((String::new(), false)),
        "and it puts no caret in a field that is not there"
    );

    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

/// **A pick is not a scroll home.** A press on a row, or a `Return` on one the hand picked, walks the
/// pin onto that row's file — which reconciles the band's rows — and the band used to reset its scroll
/// to the top on every reconciliation, so a hand that pressed row eleven was left reading row one.
///
/// What a reconciliation of rows costs the scroll is nothing: the hand is reading the row it picked.
#[test]
fn a_pick_leaves_the_band_scrolled_where_the_hand_left_it() {
    let _room = the_room_to_itself();

    let previous_pin = take_pin_for_a_test();
    let pin = sound_pin();
    pin_playlist::toggle(&pin);
    let sounds = a_walk_of_twelve_sounds(pin);

    // The band is scrolled down the way a wheel scrolls it — a notch forward is a notch down the list,
    // which is the wheel's own reading of it.
    assert!(
        pin_playlist::step_the_notches_for_a_test(-1200),
        "the wheel is the band's"
    );
    let (_, scrolled) = pin_playlist::the_rows_and_scroll();
    assert!(
        scrolled > 0,
        "the band is scrolled down before the pick, and it is at {scrolled}"
    );

    // The arrows step down the band a row at a time, and the last press of them is the row a hand reads
    // before it presses Return on it.
    assert!(
        pin_playlist::step_the_pick(1),
        "the band takes the arrow over it"
    );
    let picked = pin_playlist::the_band_in_hand()
        .and_then(|(_, picked, _)| picked)
        .expect("a row the arrow picked");

    // **The read of the band the pick leaves behind**, before a row is reconciled against it.
    let (_, picked_before, scroll_before) = pin_playlist::the_band_in_hand().expect("the band");
    assert_eq!(
        picked_before,
        Some(picked),
        "the band's pick is the arrow's"
    );

    // **The pick walks the pin onto that row's file**, which is what reconciles the rows of the band:
    // the pin is stood on the file the press named and the walk with it, exactly as a double press on
    // a row does (see `show_pinned_file` and `refresh_the_rows`).
    if let Some(mut state) = pin_state() {
        if let Some(pin) = state.pin_mut() {
            pin.path = sounds[picked].clone();
        }
    }
    stand_pin_walk(Some(PinStep {
        at: sounds[picked].clone(),
        from: sounds[picked].clone(),
        step: 1,
        order: sounds.clone(),
        missed: 0,
        shuffle: false,
        progress: 0,
        anchor: Some(crate::formats::routing::NavCategory::Audio),
    }));
    assert!(
        pin_playlist::refresh_the_rows(),
        "the pin walked, so the band reconciles the rows"
    );

    // **What the reconciliation costs the band is nothing.** A hand that picked row eleven was reading
    // row eleven, and the band it is left with is scrolled where it was rather than at its top.
    let (_, picked_after, scroll_after) = pin_playlist::the_band_in_hand().expect("the band");
    assert_ne!(
        scroll_after, 0,
        "a pick on a row down the list does not scroll the band back to its top"
    );
    assert_eq!(
        scroll_after, scroll_before,
        "and the band is scrolled where the hand left it"
    );
    assert_eq!(
        picked_after,
        Some(picked),
        "the pick survived the reconciliation of rows"
    );

    let opened = pin_state().and_then(|pinned| pinned.pin().map(pin_playlist::toggle));
    assert!(opened.is_some(), "the pin the test stood up");
    stand_pin(previous_pin);
}

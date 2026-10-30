//! Which window a taken-up pin stands in: every kind this app draws the whole way stands in the
//! pin's own window, and the three kinds something else draws in a window of its own keep the
//! preview's.

use super::*;

/// The kinds a pin draws itself are the ones Windows may activate, list in Alt+Tab and give a
/// taskbar button to. Only a film a player of its own draws and the browser's two keep the
/// preview's window, because what fills their band is somebody else's window.
#[test]
fn every_kind_the_pin_draws_whole_stands_in_the_pins_own_window() {
    for kind in [
        MediaType::StaticImage,
        MediaType::Dds,
        MediaType::AnimatedGif,
        MediaType::AnimatedApng,
        MediaType::AnimatedWebP,
        MediaType::AnimatedHeif,
        MediaType::AnimatedJxl,
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
    ] {
        assert!(
            pin_should_stand_in_its_own_window(Some(kind), Path::new("a-file")),
            "{kind:?} is drawn the whole way by this app, so its pin is the pin's own window"
        );
    }

    for kind in [MediaType::EngineSvg, MediaType::EngineFont] {
        assert!(
            !pin_should_stand_in_its_own_window(Some(kind), Path::new("a-file")),
            "{kind:?} is drawn in a window of somebody else's, so its pin keeps the preview's"
        );
    }
}

/// A film is drawn into a host this app owns under the pin's own window, whichever player draws
/// it: mpv through `--wid`, FFmpeg's player adopted as a disabled child after it makes its own
/// window. So a pin of one stands in the pin's own window. The kind on screen and the file's own
/// route answer the same way, placeholder or not (see `pin_film_host`).
#[test]
fn a_film_stands_in_the_pins_own_window_whichever_player_draws_it() {
    let _room = the_room_to_itself();

    for (mpv, ffplay_event, engine) in [
        (true, true, VideoEngine::Mpv),
        (false, true, VideoEngine::Ffmpeg),
    ] {
        stand_mpv_in(Some(mpv));
        stand_ffplay_in(Some(ffplay_event));
        stand_video_engine_in(Some((engine, true)));

        for kind in [
            Some(MediaType::Video),
            Some(MediaType::Loading),
            Some(MediaType::Unplayable),
            None,
        ] {
            assert!(
                pin_should_stand_in_its_own_window(kind, Path::new("a.mp4")),
                "{kind:?} of a film {engine:?} draws is drawn in a host of the pin's own, so the                  pin is the pin's own window"
            );
        }
    }

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
}

/// A drag parks the window behind the band, and a film drawn into the pin's own host leaves no
/// window of the desktop's behind it to park, whichever player draws it (see
/// `MediaType::draws_in_a_window_of_its_own`).
#[test]
fn a_hosted_film_leaves_nothing_for_a_drag_to_park() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let previous_own = pin_is_its_own_window();
    let path = Path::new("a-film.mkv");

    stand_pin(Some(overlay_pin((100, 80, 420, 320), PinChrome::always())));
    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(true));

    pin_stands_in_its_own_window(true);
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    assert!(
        !MediaType::Video.draws_in_a_window_of_its_own(path),
        "a hosted film's picture is a child of the pin's own window, not a window of the desktop's"
    );

    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    assert!(
        !MediaType::Video.draws_in_a_window_of_its_own(path),
        "and a film FFmpeg's player draws is adopted into the same host, so it is no window to park"
    );
    assert!(
        !MediaType::NativeVideo.draws_in_a_window_of_its_own(path),
        "the media engine's film is a frame this window draws, as it always was"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// A film's answer about a window of its own is the same whichever player the route names now: an
/// engine row can change under a pin that is already up, and both players draw into the pin's own
/// host (see `MediaType::draws_in_a_window_of_its_own`).
#[test]
fn a_pins_film_is_drawn_in_its_host_whichever_route_it_is_drawn_by_now() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let previous_own = pin_is_its_own_window();

    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(true));

    let path = Path::new("a-film.mkv");
    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = path.to_path_buf();
    stand_pin(Some(pin));

    pin_stands_in_its_own_window(true);
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    assert_eq!(video_route(path), VideoRoute::Mpv);
    assert!(!MediaType::Video.draws_in_a_window_of_its_own(path));
    assert!(pin_should_stand_in_its_own_window(
        Some(MediaType::Video),
        path
    ));

    // The engine row changes to FFmpeg's player under the same pin: the film is adopted into the
    // same host, not left as a window of the desktop's.
    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    assert_eq!(video_route(path), VideoRoute::Ffplay);
    assert!(
        pin_is_its_own_window(),
        "which the take-up decision has not heard about"
    );
    assert!(
        !MediaType::Video.draws_in_a_window_of_its_own(path),
        "the film FFmpeg's player draws stands in the pin's host too"
    );
    assert!(
        pin_should_stand_in_its_own_window(Some(MediaType::Video), path),
        "so a pin taken up again on it is the pin's own window"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// The kind on screen is the answer where there is one, and the file's name is not asked again:
/// a picture is the pin's own window whatever it is called, and a film FFmpeg's player draws is
/// hosted like every other film even under a picture's name (see
/// `pin_should_stand_in_its_own_window`).
#[test]
fn a_known_kind_is_not_re_read_from_the_path() {
    let _room = the_room_to_itself();

    assert!(
        pin_should_stand_in_its_own_window(Some(MediaType::StaticImage), Path::new("a.svg")),
        "the picture in hand is a picture, whatever the name beside it says"
    );

    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    assert!(
        pin_should_stand_in_its_own_window(Some(MediaType::Video), Path::new("a.png")),
        "and a film is still hosted under a picture's name"
    );
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
}

/// A pin whose media is a placeholder — the spinner before a file is read, or the cross a failed
/// file leaves — is still a pin of the file under it, so the file's own route says whose window
/// the band is: a picture, a page, a card and a film the media engine draws are this app's to
/// host, while a film a player of its own draws and the browser's three are not.
#[test]
fn a_placeholder_pin_stands_where_its_files_route_says() {
    for path in [
        "a.pdf", "a.docx", "a.cr2", "a.txt", "a.mp3", "a.cbz", "a.zip", "a.psd", "a.emf", "a.epub",
    ] {
        for kind in [None, Some(MediaType::Loading), Some(MediaType::Unplayable)] {
            assert!(
                pin_should_stand_in_its_own_window(kind, Path::new(path)),
                "{path} is drawn by this app, so a pin of it is the pin's own window even while \
                 its media is a placeholder"
            );
        }
    }

    // A page of HTML is the browser's only while the browser is drawing it: with `Render HTML`
    // off it is this app's own text preview, and a pin of it belongs in the pin's window with
    // every other app-drawn kind. So only the two the engine always draws are named here.
    for path in ["a.svg", "a.ttf"] {
        for kind in [None, Some(MediaType::Loading), Some(MediaType::Unplayable)] {
            assert!(
                !pin_should_stand_in_its_own_window(kind, Path::new(path)),
                "{path} is drawn in a window of somebody else's, so a placeholder pin of it keeps \
                 the preview's"
            );
        }
    }
}

/// A failed film is the cross, but it is a failed *film*: a player of its own is drawn in the pin's
/// host like any film, so its pin is still the pin's own window.
#[test]
fn a_failed_film_a_player_draws_stands_in_the_pins_own_window() {
    let _room = the_room_to_itself();

    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_mpv_in(Some(false));

    assert!(
        pin_should_stand_in_its_own_window(Some(MediaType::Unplayable), Path::new("a.mp4")),
        "the cross over a film a player draws is still a pin of a hosted film"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
}

/// A pinned FFmpeg video now stands in the pin's own window, while an audio pin is unchanged: its
/// card keeps the pin's window and never a film host, and its Space stays the card's own toggle.
#[test]
fn a_pinned_ffplay_film_stands_in_the_pins_own_window() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let previous_own = pin_is_its_own_window();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

    let video = Path::new("a-film.mkv");
    assert!(
        pin_should_stand_in_its_own_window(Some(MediaType::Video), video),
        "a pinned FFmpeg film is hosted in the pin's own window"
    );

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = video.to_path_buf();
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);
    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }
    assert!(
        pin_film_host::the_film_is_shown_in_a_host(video, None),
        "and the film-host road is open for it"
    );

    // A sound is the card this app paints: no host, and the film-host road refuses it even where
    // the file's route names the same FFmpeg player (see `the_film_is_shown_in_a_host`).
    let sound = Path::new("a-sound.mp3");
    assert!(
        pin_should_stand_in_its_own_window(Some(MediaType::Audio), sound),
        "an audio pin keeps the pin's own window, as it always did"
    );
    assert!(
        !MediaType::Audio.draws_in_a_window_of_its_own(sound),
        "the card is this app's own frame rather than a window of somebody else's"
    );
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.media_type = MediaType::Audio;
        }
    }
    assert!(
        !pin_film_host::the_film_is_shown_in_a_host(sound, None),
        "an audio pin is never given a film host"
    );
    unsafe { pin_film_host::take_up_the_pin_film_host(None) };
    assert_eq!(
        pin_film_host::film_host(),
        None,
        "and taking an audio pin up leaves no host standing"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// **A pin that arrives is never asked for the foreground.** The take-up and the show behind it are
/// the whole of the road a pin arrives by, and nothing on that road asks Windows for the front.
///
/// **What puts a pin in front is the band, and which band is the setting's to say.** A pin that is
/// `Always On Top` is placed in the topmost band by the show every pin gets; a pin that is not has the
/// topmost band taken and given straight back for the moment it appears, so the next click on another
/// window puts it behind exactly as any window behaves (see `band_place`,
/// `showing_takes_the_band_to_appear` and `take_the_band_and_give_it_back`). Nothing of the pin's
/// arrival is a request for the keyboard, and nothing of this road may become one.
///
/// It is read off the road's own recorder rather than off the desktop: a take-up that asked for the
/// front wrote that down, and the show's own line is what says this is the road that was read.
#[test]
fn a_pin_that_arrives_is_never_asked_for_the_foreground() {
    let _room = the_room_to_itself();

    let film = super::pin_film_host_window::StandingFilm::new();
    let before = pin_always_on_top();
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    // A pin that arrives comes up with no pin standing behind it, which is the take-up whose state is
    // written whole from the file the key asked for (see `take_up_pinned_window`).
    stand_pin(None);
    pin_trace::clear_noted();

    install(take_up_pinned_window(
        Path::new("a-pin-that-arrives.mkv"),
        film.box_,
    ));
    // Safety: the window is the stand's own, and every global written here is put back when it is
    // dropped.
    unsafe { show_pinned_window(film.pin_hwnd) };

    assert!(
        unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "the pin is shown: a take-up that asks nothing of the front is still a pin on the screen"
    );
    assert!(
        !unsafe { window_holds_the_topmost_band(film.pin_hwnd) },
        "and the band it appears in is given straight back, because `Always On Top` is off"
    );

    let trace = pin_trace::noted_lines();
    assert!(
        trace.iter().any(|line| line.starts_with("show hwnd")),
        "this is the take-up/show road that was read: {trace:?}"
    );
    assert!(
        !trace.iter().any(|line| line.contains("front")),
        "and no part of it asks for the foreground: a pin arrives without its window ever being put \
         in front by a request, which is the band's work (see `band_place`): {trace:?}"
    );

    // **And the ask was never what put a pin in front.** With `Always On Top` on, the same show
    // places the same window in the topmost band: the setting is the mechanism, and it is the only
    // one.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    // Safety: the window is the stand's own, as above.
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(
        unsafe { window_holds_the_topmost_band(film.pin_hwnd) },
        "with `Always On Top` on the pin is in the band above everything, which is the placement's \
         doing and not a request for the foreground"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
}

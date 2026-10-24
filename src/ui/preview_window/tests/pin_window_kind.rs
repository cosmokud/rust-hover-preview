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
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

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
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
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
fn a_pins_film_is_a_host_child_whichever_route_it_is_drawn_by_now() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
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
    assert!(pin_should_stand_in_its_own_window(Some(MediaType::Video), path));

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
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

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
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

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

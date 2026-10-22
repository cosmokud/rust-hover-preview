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

/// A film mpv plays is drawn into a host this app owns under the pin's own window, so a pin of
/// one stands in the pin's own window; a film FFmpeg's player draws is that player's window
/// behind the band, and its pin keeps the preview's until its own ticket. The kind on screen and
/// the file's own route answer the same way, placeholder or not (see `pin_film_host`).
#[test]
fn a_film_mpv_plays_stands_in_the_pins_own_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));

    for kind in [
        Some(MediaType::Video),
        Some(MediaType::Loading),
        Some(MediaType::Unplayable),
        None,
    ] {
        assert!(
            pin_should_stand_in_its_own_window(kind, Path::new("a.mp4")),
            "{kind:?} of a film mpv plays is drawn in a host of the pin's own, so the pin is the \
             pin's own window"
        );
    }

    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

    for kind in [
        Some(MediaType::Video),
        Some(MediaType::Loading),
        Some(MediaType::Unplayable),
        None,
    ] {
        assert!(
            !pin_should_stand_in_its_own_window(kind, Path::new("a.mp4")),
            "{kind:?} of a film FFmpeg's player draws in a window of its own keeps the preview's"
        );
    }

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
}

/// A drag parks the window behind the band, and a film mpv draws into the pin's own host leaves
/// no window of the desktop's behind it to park; a film FFmpeg's player draws still does (see
/// `MediaType::draws_in_a_window_of_its_own`).
#[test]
fn a_hosted_film_leaves_nothing_for_a_drag_to_park() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_own = pin_is_its_own_window();

    stand_pin(Some(overlay_pin((100, 80, 420, 320), PinChrome::always())));

    pin_stands_in_its_own_window(true);
    assert!(
        !MediaType::Video.draws_in_a_window_of_its_own(),
        "a hosted film's picture is a child of the pin's own window, not a window of the desktop's"
    );

    pin_stands_in_its_own_window(false);
    assert!(
        MediaType::Video.draws_in_a_window_of_its_own(),
        "and a film FFmpeg's player draws is still the window a drag parks"
    );
    assert!(
        !MediaType::NativeVideo.draws_in_a_window_of_its_own(),
        "the media engine's film is a frame this window draws, as it always was"
    );

    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// The kind on screen is the answer where there is one, and the file's name is not asked again:
/// a picture is the pin's own window whatever it is called, and a film FFmpeg's player draws keeps
/// the preview's even under a picture's name. A film's route is the one thing asked beside the
/// kind, because it is the player that decides where its picture is drawn (see
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
        !pin_should_stand_in_its_own_window(Some(MediaType::Video), Path::new("a.png")),
        "and a player's window is still a player's window under a picture's name"
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

/// A failed film is the cross, but it is a failed *film*: the route that says a player of its own
/// draws it still says so, so the pin stays in the preview's window rather than becoming a real
/// window over a player's band the user asked to leave alone.
#[test]
fn a_failed_film_a_player_draws_keeps_the_previews_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_mpv_in(Some(false));

    assert!(
        !pin_should_stand_in_its_own_window(Some(MediaType::Unplayable), Path::new("a.mp4")),
        "the cross over a film a player draws is still that player's band, not the pin's window"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
}

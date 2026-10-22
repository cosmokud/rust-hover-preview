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

    for kind in [
        MediaType::Video,
        MediaType::EngineSvg,
        MediaType::EngineFont,
    ] {
        assert!(
            !pin_should_stand_in_its_own_window(Some(kind), Path::new("a-file")),
            "{kind:?} is drawn in a window of somebody else's, so its pin keeps the preview's"
        );
    }
}

/// The kind on screen is the answer where there is one, and the file's name is not asked again:
/// a picture is the pin's own window whatever it is called, and a film a player draws keeps the
/// preview's even under a picture's name.
#[test]
fn a_known_kind_is_not_re_read_from_the_path() {
    assert!(
        pin_should_stand_in_its_own_window(Some(MediaType::StaticImage), Path::new("a.svg")),
        "the picture in hand is a picture, whatever the name beside it says"
    );
    assert!(
        !pin_should_stand_in_its_own_window(Some(MediaType::Video), Path::new("a.png")),
        "and a player's window is still a player's window under a picture's name"
    );
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

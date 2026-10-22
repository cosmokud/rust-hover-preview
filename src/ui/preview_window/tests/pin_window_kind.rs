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
        MediaType::Unplayable,
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

/// A pin taken up before its preview has answered has no media kind to ask, so the file's own
/// route is what says whose window the band is: a picture, a page, a card and a film the media
/// engine draws are this app's to host, while a film a player of its own draws and the browser's
/// three are not.
#[test]
fn a_pin_taken_up_while_loading_stands_where_its_files_route_says() {
    for path in [
        "a.pdf", "a.docx", "a.cr2", "a.txt", "a.mp3", "a.cbz", "a.zip", "a.psd", "a.emf", "a.epub",
    ] {
        for kind in [None, Some(MediaType::Loading)] {
            assert!(
                pin_should_stand_in_its_own_window(kind, Path::new(path)),
                "{path} is drawn by this app, so a pin of it is the pin's own window even while \
                 its preview is still loading"
            );
        }
    }

    for path in ["a.svg", "a.html", "a.ttf"] {
        for kind in [None, Some(MediaType::Loading)] {
            assert!(
                !pin_should_stand_in_its_own_window(kind, Path::new(path)),
                "{path} is drawn in a window of somebody else's, so a loading pin of it keeps \
                 the preview's"
            );
        }
    }
}

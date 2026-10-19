//! The engines the file on screen was built under, read where a preview begins and kept for
//! the whole of it.
//!
//! `Engine → Video Engine`, the `Fallback` switch above it and `Engine → Office Engine` are
//! written the moment a row is clicked, and that is the whole of what the click does: nothing
//! on screen is rebuilt for it. A film already playing belongs to the player that started it —
//! the rule the rows of that pair already stated (`set_mpv_path`,
//! `toggle_video_engine_fallback`) and which nothing kept, because the router and the page
//! engine were reading the setting again on every question they were asked. A resize, a
//! maximize, a seek, a transport key and a step are all questions, and each of them came back
//! with the engine of *now* rather than the engine of the file: a film ended and begun again
//! in another player in the middle of its pass, or a page asked of an engine that has never
//! seen the document.
//!
//! So the choices are read once, where a preview is built — [`latch_engines`] — and every look
//! after that is answered with what was read then ([`video_engine_in_effect`],
//! [`office_engine_in_effect`]). The preview loop makes that read at the two places a file is
//! built anew: a hover, and a pinned window being shown a file it is not already showing. A
//! pin taken up over a hover is not one of them — the window adopts the player the hover
//! started and the frame it is drawing, so the choices that file was built under are still the
//! ones in effect.

use std::sync::Mutex;

use super::defaults::{DEFAULT_OFFICE_ENGINE, DEFAULT_VIDEO_ENGINE, DEFAULT_VIDEO_ENGINE_FALLBACK};
use super::setting_types::{OfficeEngine, VideoEngine};
use crate::CONFIG;

/// The choices a preview was begun with, read in one look so that a file is never built under
/// one of them from one moment and another from the next.
#[derive(Clone, Copy)]
struct Engines {
    video: VideoEngine,
    /// Whether an explicit choice falls through to the engines under it. It is half of the
    /// video choice rather than a setting of its own — the router reads the two together, and
    /// a preview built under one was built under both (see `resolve_video_engine`).
    fallback: bool,
    office: OfficeEngine,
}

/// What [`latch_engines`] read last, and nothing at all before the first preview. One for the
/// whole process, which is what the app shows at a time: a hover is refused while a pin is up,
/// so the file on screen is one file (see `show_preview`).
static IN_EFFECT: Mutex<Option<Engines>> = Mutex::new(None);

/// Read the choices again, for a preview that is about to be built, and answer whether the
/// video choice moved.
///
/// The move is the one thing the preview side has to act on: a geometry probed under the
/// engine that has just been left behind is not the geometry the preview about to be built is
/// placed by, so the probe cache is given up with this read rather than with the tray row that
/// moved the setting — a preview on screen is still placed by what it was built under (see
/// `forget_video_geometry`). A read whose lock cannot be taken is not a read: the choices are
/// left where they are, which is the preview that is up kept rather than rebuilt on a guess.
pub fn latch_engines() -> bool {
    let configured = configured();

    match IN_EFFECT.lock() {
        Ok(mut held) => read(&mut held, configured),
        Err(_) => false,
    }
}

/// Which engine plays a video here and now, and whether an explicit choice falls through to
/// the engines under it — the two halves the router is handed together (see
/// `video_hw::engine_choice`).
pub fn video_engine_in_effect() -> (VideoEngine, bool) {
    let in_effect = looked_up(held(), configured());

    (in_effect.video, in_effect.fallback)
}

/// Which engine an Office document's page is asked of here and now (see
/// `office_formats::page_engine`).
pub fn office_engine_in_effect() -> OfficeEngine {
    looked_up(held(), configured()).office
}

/// The choices a look is answered with: the ones the preview on screen was begun with, and the
/// configuration's own where no preview has begun — which is every look the tray, a probe of
/// its own or a run that has shown nothing makes.
fn looked_up(held: Option<Engines>, configured: Engines) -> Engines {
    held.unwrap_or(configured)
}

/// What [`latch_engines`] does with the choices it read: they become the ones in effect, and
/// the answer is whether the video choice moved — the fall-back switch counted with it, since
/// the two are one choice (see [`Engines::fallback`]).
fn read(held: &mut Option<Engines>, chosen: Engines) -> bool {
    let moved = held.is_some_and(|previous| {
        previous.video != chosen.video || previous.fallback != chosen.fallback
    });

    *held = Some(chosen);
    moved
}

/// What [`latch_engines`] read last, where a preview has begun at all.
fn held() -> Option<Engines> {
    IN_EFFECT.lock().ok().and_then(|held| *held)
}

/// The choices the configuration names now, or the defaults where its lock cannot be read —
/// the answer each of the two readers used to make for itself.
fn configured() -> Engines {
    CONFIG
        .lock()
        .map(|config| Engines {
            video: config.video_engine,
            fallback: config.video_engine_fallback,
            office: config.office_engine,
        })
        .unwrap_or(Engines {
            video: DEFAULT_VIDEO_ENGINE,
            fallback: DEFAULT_VIDEO_ENGINE_FALLBACK,
            office: DEFAULT_OFFICE_ENGINE,
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eng(video: VideoEngine, fallback: bool, office: OfficeEngine) -> Engines {
        Engines {
            video,
            fallback,
            office,
        }
    }

    /// What a look is answered with, and what a read answers about the choice that moved: the
    /// whole of the protection is these two rules, and neither needs a configuration.
    #[test]
    fn a_preview_is_begun_under_the_choices_read_then_and_kept_under_them() {
        let begun = eng(VideoEngine::Native, true, OfficeEngine::MicrosoftOffice);
        let moved = eng(VideoEngine::Mpv, false, OfficeEngine::LibreOffice);
        let mut held = None;

        assert_eq!(
            looked_up(held, moved).video,
            VideoEngine::Mpv,
            "no preview has begun, so the configuration's own choice answers"
        );

        assert!(
            !read(&mut held, begun),
            "a first read has nothing behind it to have moved from"
        );
        assert_eq!(
            looked_up(held, moved).video,
            VideoEngine::Native,
            "a preview is up, so the choice it was begun with answers however the \
             configuration has moved since"
        );
        assert_eq!(looked_up(held, moved).office, OfficeEngine::MicrosoftOffice);

        assert!(
            !read(&mut held, begun),
            "the same choices read again are not a move, so no probed geometry is given up"
        );
        assert!(
            read(&mut held, eng(begun.video, !begun.fallback, begun.office)),
            "the fall-back switch is half of the video choice"
        );
        assert!(
            !read(
                &mut held,
                eng(begun.video, !begun.fallback, OfficeEngine::LibreOffice)
            ),
            "and the Office choice is nothing the film's side has to act on"
        );
        assert_eq!(
            looked_up(held, moved).office,
            OfficeEngine::LibreOffice,
            "a look is answered with the choices in effect, whatever they are"
        );
    }
}

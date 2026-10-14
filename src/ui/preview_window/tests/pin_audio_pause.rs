use super::*;

/// A press on a pinned sound's play/pause is a toggle, and the whole of it is answered by
/// whether the sound is being held — which is the engine's own answer for the engine, because
/// nothing this side writes down about an engine's hold.
///
/// The fault this is a regression test for: the question was asked of whether a *session* is
/// there, and a session that was asked to hold is a session that is still there. So a press on a
/// sound the engine was already holding asked it to hold again — the sound stayed paused, with
/// no way to let it go, and the card went on drawing the pause glyph of a sound that was going
/// (see `toggle_pinned_audio`).
#[test]
fn a_sound_the_engine_is_holding_is_held_whatever_is_written_down_for_it() {
    // The engine's own answer, in both directions, is the whole of what a press turns on.
    assert!(
        !a_sound_that_is_held(None, Player::Native, false),
        "a sound the engine is playing is not held, so a press asks it to hold"
    );
    assert!(
        a_sound_that_is_held(None, Player::Native, true),
        "and a sound the engine is holding is held: a paused session is a session that is still \
         there, and only the engine knows it is one"
    );

    // A player of this app's is the other way round: it was ended to hold it, so the hold is
    // what was written down and the engine has nothing to do with it.
    assert!(
        a_sound_that_is_held(Some(37.5), Player::Ffmpeg, false),
        "a player of this app's is held at the second it was ended at, and nothing else says so"
    );
    assert!(
        !a_sound_that_is_held(None, Player::Ffmpeg, true),
        "and the engine holding a film of its own is not a sound of this app's being held"
    );

    // The hold the end of a file writes down stands over both players: a sound whose pass has
    // ended has no session behind it, and the card sits at the beginning of the file with the
    // play button up (see `pinned_native_audio_end`).
    assert!(
        a_sound_that_is_held(Some(0.0), Player::Native, false),
        "a sound held at the end of its file is held, with no session for the engine to be asked \
         about"
    );
}

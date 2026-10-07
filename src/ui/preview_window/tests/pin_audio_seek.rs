use super::*;
use crate::config::config::{AudioSeek, DEFAULT_AUDIO_SEEK, DEFAULT_PIN_MODE_AUDIO_SEEK};

/// The two `Audio Seek` settings, put where a test asks and put
/// back when the guard is dropped, so a test beside this one is
/// answered with the machine's own.
struct SeekSettings {
    hover: AudioSeek,
    pin: AudioSeek,
}

impl SeekSettings {
    fn put(hover: AudioSeek, pin: AudioSeek) -> Self {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = SeekSettings {
            hover: config.audio_seek,
            pin: config.pin_mode_audio_seek,
        };
        config.audio_seek = hover;
        config.pin_mode_audio_seek = pin;

        was
    }
}

impl Drop for SeekSettings {
    fn drop(&mut self) {
        if let Ok(mut config) = CONFIG.lock() {
            config.audio_seek = self.hover;
            config.pin_mode_audio_seek = self.pin;
        }
    }
}

/// Where a pinned sound starts is the pin's own setting's answer,
/// and where a hovered one is the hover's: the two are two answers
/// to one question, so a configuration that names them differently
/// is answered differently by the two consults — the pin's player
/// starts where the pin says, and the hover's where the hover says.
#[test]
fn a_pinned_sound_starts_where_the_pin_says_a_hover_does_not_have_to() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = SeekSettings::put(AudioSeek::Random, AudioSeek::Middle);
    assert_eq!(
        current_audio_seek(),
        AudioSeek::Random,
        "a hover starts where the hover's setting says"
    );
    assert_eq!(
        current_pin_mode_audio_seek(),
        AudioSeek::Middle,
        "a pin starts where the pin's own setting says, not where the hover's does"
    );

    let _settings = SeekSettings::put(AudioSeek::Remember, AudioSeek::Start);
    assert_eq!(
        current_audio_seek(),
        AudioSeek::Remember,
        "a hover starts where the sound was left, where that is what the hover's setting says"
    );
    assert_eq!(
        current_pin_mode_audio_seek(),
        AudioSeek::Start,
        "a pin starts at the beginning, where that is what the pin's setting says"
    );
}

/// The two settings start at their own answers: a pin at the
/// beginning, a hover where the sound was left. Which is which is
/// what the defaults are, so a configuration that names neither —
/// one written before either setting existed — is answered with
/// the beginning by the pin's consult and with the memory by the
/// hover's.
#[test]
fn the_two_settings_start_at_their_own_answers() {
    assert_eq!(
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        AudioSeek::Start,
        "a pinned sound starts at the beginning unless the file says otherwise"
    );
    assert_eq!(
        DEFAULT_AUDIO_SEEK,
        AudioSeek::Remember,
        "a hovered sound starts where it was left unless the file says otherwise"
    );
}

/// Whether the position of the sound on screen is written down
/// as its card is repainted is the answer of the setting that
/// governs that sound: the pin's own setting for a pinned file,
/// the hover's for a hovered one. The two are two answers to one
/// question, so a configuration that names them differently is
/// answered differently by the two — a pin whose own setting is
/// `Remember` keeps its position even where the hover's is not,
/// and a hover keeps its even where the pin's is not.
#[test]
fn the_position_written_down_is_what_the_sound_s_own_setting_says() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // The pin's setting is the one that resumes a pinned sound
    // and the hover's the one that resumes a hovered one, so each
    // keeps its position under its own setting and neither under
    // the other's.
    let _settings = SeekSettings::put(AudioSeek::Start, AudioSeek::Remember);
    assert!(
        position_is_remembered(true),
        "a pinned file's position is written down where the pin's own setting says so, even where the hover's says otherwise"
    );
    assert!(
        !position_is_remembered(false),
        "a hovered file's position is not written down where the hover's setting says otherwise"
    );

    let _settings = SeekSettings::put(AudioSeek::Remember, AudioSeek::Start);
    assert!(
        position_is_remembered(false),
        "a hovered file's position is written down where the hover's own setting says so, even where the pin's says otherwise"
    );
    assert!(
        !position_is_remembered(true),
        "a pinned file's position is not written down where the pin's own setting says otherwise"
    );
}

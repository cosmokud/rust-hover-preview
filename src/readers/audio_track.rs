//! What a sound file is, and which of the two players this machine has for it.
//!
//! A sound is played by the media engine Windows has where its decoders reach the format,
//! and by FFmpeg's `ffplay` where they do not — and which of the two a file is, is a question
//! about the file and the machine rather than about the list its name is in. That answer is
//! what [`Track`] carries, and it is one answer per file and per version of it, measured once
//! and held for the hovers that follow (see [`probe`] and [`remember`]).
//!
//! The two players answer differently and the difference is the point. Windows' engine is
//! asked for *decoded* PCM on the file's first audio stream, which it can only give where a
//! decoder for the codec is registered: a yes is the whole of "this machine can play this".
//! FFmpeg is asked nothing at all beyond whether its player is installed — a player that is
//! there plays nearly everything, and a format it cannot read is a file the probe's own
//! `ffprobe` pass finds no audio stream in, which is the same answer as no player at all.
//! Which engine answers is settled by `preview_window`, which is where a file is measured off
//! the preview thread and a process may be started; this module keeps the answer.
//!
//! What is deliberately *not* here is the position of the playback: that is the session's
//! (see `video_player`) and the FFmpeg player's own wall clock, and neither is a fact about
//! the file.

use crate::formats::head;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::path::Path;
use std::sync::Mutex;

/// Which of the two engines plays a file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Player {
    /// The media engine Windows has, in this app's own process: the engine that plays a
    /// video's frames plays a sound's as well, with nothing to draw (see `video_player`).
    Native,
    /// FFmpeg's `ffplay`, with no window at all: `-nodisp` is a player that makes a sound
    /// and shows nothing, which is the whole of what a card with no picture needs.
    Ffmpeg,
}

/// What a sound file is, as far as a preview is concerned: which engine plays it and the
/// facts the card is drawn with.
///
/// Every fact but the player is optional, and each is left out rather than guessed at: a
/// container that does not say how long it is draws a card with no clock, and one whose codec
/// this app has no friendly name for is drawn with the name its own extension carries.
#[derive(Debug, Clone, PartialEq)]
pub struct Track {
    pub player: Player,
    /// The codec's own name — `FLAC`, `MP3`, `Vorbis` — where the engine named one.
    pub codec: Option<String>,
    /// Samples a second, as the file writes them.
    pub rate: Option<u32>,
    /// Channels the file carries: one, two, or however many.
    pub channels: Option<u16>,
    /// Bits a second, as an average, where the file says.
    pub bitrate: Option<u32>,
    /// How long it plays, where the file says.
    pub duration: Option<f64>,
}

/// What a probe has said about a file: nothing where it has not been asked, what the machine
/// plays where it has, and nothing-to-play where it has that answer instead.
///
/// It is three answers rather than two because a preview is waited for on one of them and
/// dropped for the other: a hover on a file no probe has looked at is a wait, and a hover on
/// one the machine cannot play is over (see `audio_box` and `Probed::Nothing`).
#[derive(Debug, Clone, PartialEq)]
pub enum Probed {
    /// No probe has been run for the file in this run.
    NotAsked,
    /// This machine plays the file, and this is what the file holds.
    Track(Track),
    /// The file is not a sound this machine can play — or holds no sound at all.
    Nothing,
}

/// What every probe has said, held per file and version: a track where the machine plays one,
/// and nothing where it does not — which is an answer too, and one worth holding (see
/// [`Probed`]).
static PROBED: Lazy<Mutex<HashMap<head::Key, Option<Track>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// How many files are held. The same bound and the same reasoning as every other memo of this
/// shape: a folder swept a file at a time.
const TRACKS_MAX_ENTRIES: usize = 512;

/// What a probe has said about `path`, or that there is nothing to say yet.
pub fn probed(path: &Path) -> Probed {
    let key = head::key(path);

    match PROBED.lock().ok().and_then(|held| held.get(&key).cloned()) {
        Some(Some(track)) => Probed::Track(track),
        Some(None) => Probed::Nothing,
        None => Probed::NotAsked,
    }
}

/// The track a player would be started from, where there is one to start.
pub fn playable(path: &Path) -> Option<Track> {
    match probed(path) {
        Probed::Track(track) => Some(track),
        Probed::NotAsked | Probed::Nothing => None,
    }
}

/// Hold what a probe found about `path`, the answer that there was nothing to find included.
pub fn remember(path: &Path, probed: Probed) {
    let value = match probed {
        Probed::Track(track) => Some(track),
        Probed::Nothing => None,
        Probed::NotAsked => return,
    };

    let key = head::key(path);

    if let Ok(mut held) = PROBED.lock() {
        if held.len() >= TRACKS_MAX_ENTRIES {
            held.clear();
        }
        held.insert(key, value);
    }
}

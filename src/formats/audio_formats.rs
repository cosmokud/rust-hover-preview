//! Which files are sounds, and the one thing about a sound a probe has to be asked.
//!
//! The list of names this answers for is a row of `crate::formats::lists` — the one table
//! every kind's list is a row of, and the one place a list is written down, the built-in
//! entries and the older lists this app shipped and then changed included.

use crate::formats::{head, text_formats};
use once_cell::sync::Lazy;
use std::collections::HashSet;
use std::path::Path;
use std::sync::Mutex;

/// Whether the configured list claims `path`.
pub fn matches_audio_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// The containers a probe found a sound in and no picture, held for the run.
///
/// It is the answer to the one question the lists cannot settle. An `.mp4`, a `.mka` and an
/// `.ogg` are containers that can hold a film or a song, and what they hold is a fact about
/// the streams inside them rather than about the box: the probe opens the file, finds no
/// video stream and a sound it can play, and says so here — after which the router answers
/// the file as the sound it is, which is what routes the replayed hover to the card rather
/// than to the black window a video player would put up for a file with nothing to draw.
///
/// What is remembered lives for the run and is keyed by the file *and* the version of it that
/// was read, so a file replaced by one that does hold a picture is asked about again — the
/// same key `content_type` holds its own answers under (see `formats::head::Key`).
static AUDIO_ONLY: Lazy<Mutex<HashSet<head::Key>>> = Lazy::new(|| Mutex::new(HashSet::new()));

/// How many files are remembered as sounds. A pointer swept across a folder is a file at a
/// time, so what this bounds is a session's worth of them, and what it costs when it is
/// reached is everything held, the way every other cache of this app's shape answers that.
const AUDIO_ONLY_MAX_ENTRIES: usize = 512;

/// Remember that a probe found a sound in `path` and no picture.
pub fn remember_audio_only(path: &Path) {
    let key = head::key(path);

    if let Ok(mut remembered) = AUDIO_ONLY.lock() {
        if remembered.len() >= AUDIO_ONLY_MAX_ENTRIES {
            remembered.clear();
        }
        remembered.insert(key);
    }
}

/// Whether a probe has already found a sound in `path` and no picture.
pub fn probed_audio_only(path: &Path) -> bool {
    let key = head::key(path);

    AUDIO_ONLY
        .lock()
        .is_ok_and(|remembered| remembered.contains(&key))
}

/// Whether a probe has already found a sound in this version of `path` and no picture.
///
/// The question is a `HashSet` lookup, and taking it used to cost a `fs::metadata` of the file
/// to build the key: the entry is read to answer a question about a table already in memory.
/// A caller that has read the entry for something else — the head, a cache key, the hook's own
/// gate — has the key already, and one hover asks this question from four places, so the
/// repeated `metadata` calls were the same read of the same directory entry several times over.
///
/// Nothing is asked of the file here: a file that is not there has a key like any other, and
/// the table cannot hold it because nothing ever remembered it.
pub fn probed_audio_only_in(key: &head::Key) -> bool {
    AUDIO_ONLY
        .lock()
        .is_ok_and(|remembered| remembered.contains(key))
}

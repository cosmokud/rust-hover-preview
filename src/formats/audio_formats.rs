//! Which files are sounds, and the one thing about a sound a probe has to be asked.
//!
//! The extension list lives in `config.ini`, written from the built-in list on first run
//! and read back from there, exactly as the video preview's list and the rest of them are —
//! so a user can add a format this list does not name, or take one out, without a rebuild.
//!
//! The question here is only what a file is *called*. What it *is*, and whether this machine
//! can play it, is settled by the probe in `audio_track`: a sound is read by the media engine
//! Windows has where its decoders reach it and by FFmpeg's player where they do not, and the
//! list has no opinion about either.
//!
//! Two names in this list are also a list's of another kind, and both are settled by the
//! bytes rather than by the name — see `content_type` for what a `.mpc` and an `.ogg` are
//! asked before the order of the lists decides, and `SHARED` in `routing` for the pair.
//!
//! What is deliberately *not* in the list is as worth saying. A **MIDI sequence** is the
//! loudest of them: the GS Wavetable synthesizer is a MIDI *out* device and not a decoder,
//! so Media Foundation has nothing to hand a `.mid` to, and FFmpeg's own player turns one
//! down as well — no sound of the format is playable by this app, and a hover onto one shows
//! nothing. Beside it are the **tracker modules** (`.mod`, `.xm`, `.it`), which only a build
//! of FFmpeg carrying libopenmpt plays; the **protected** files (`m4p`, `aa`, `aax`), whose
//! DRM neither engine will decrypt; and the **playlists** (`.m3u`, `.pls`), which are text
//! that names other files. Each is named in `TODO.md` with what would have to change for it.

use crate::formats::{head, text_formats};
use once_cell::sync::Lazy;
use std::collections::HashSet;
use std::path::Path;
use std::sync::Mutex;

/// The extensions written to `config.ini` on first run: every sound this app has an engine
/// to play, in the one list, because which engine plays a format is the machine's answer and
/// not the list's.
///
/// The native families are here — WAV, MP3, AAC in its two spellings, WMA, FLAC, ALAC, the
/// AMR pair, AIFF, DSD's two containers, AC-3 and DTS — and so is everything an installed
/// FFmpeg reads and Windows does not: Ogg Vorbis and Opus, Matroska's audio, Musepack,
/// WavPack, Monkey's Audio, True Audio, Shorten, TAK, OptimFROG, CAF, AU, VOC and
/// RealAudio. A name whose decoders are on neither engine is a name that shows nothing, and
/// the list is a list of the ones that do.
pub const DEFAULT_AUDIO_EXTENSIONS: &str = "aac,ac3,aif,aifc,aiff,amr,ape,au,awb,caf,dff,dsf,dts,dtshd,eac3,flac,m4a,m4b,mka,mp2,mp3,mpa,mpc,oga,ogg,ofr,ofs,opus,ra,shn,snd,spx,tak,tta,voc,wav,wave,wma,wv";

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

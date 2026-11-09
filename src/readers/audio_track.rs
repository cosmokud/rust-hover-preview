//! What a sound file is, and which of the two players this machine has for it.
//!
//! A sound is played by the media engine Windows has where its decoders reach the format,
//! and by FFmpeg's `ffplay` where they do not — and which of the two a file is, is a question
//! about the file and the machine rather than about the list its name is in. That answer is
//! what [`Track`] carries, and it is one answer per file and per version of it, measured once
//! and held for the hovers that follow — and for as long as one of them has a card on screen,
//! which is a longer life than a hover's and the reason a full memo gives up the file it was
//! asked about last (see [`remember`]).
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
//! What a file asks for in the way it is played is here as well — the gain that brings its
//! measured loudness to the level the tray's `Normalize` plays files at, where that measurement
//! was asked for (see [`gain`]) — because that is a fact about the file too, and one a hover
//! must not measure twice.
//!
//! What is deliberately *not* here is the position of the playback: that is the session's
//! (see `video_player`) and the FFmpeg player's own wall clock, and neither is a fact about
//! the file.

use crate::formats::head;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::path::Path;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;

mod disk;

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

/// What a probe has said about a file: nothing where it has not been asked, and what the machine
/// plays where it has.
///
/// There is no third answer, and the one that is missing is the whole point: a read that found
/// nothing to play is *not held*, because it is a fact about one read rather than about the file —
/// a player this machine does not have, a format its decoders do not reach, a prober that was not
/// on the `PATH` when a folder was swept. Held, it was terminal: the file could not be previewed
/// again for the rest of the run, and every row of the folder it sat in lost the facts it had just
/// been given (see [`remember`]).
///
/// So `NotAsked` covers two answers that are one: nothing has looked at the file, and the read that
/// did found nothing usable. Both mean ask again, and the gates that ask are where the asking is
/// bounded — one ask per hover, one per pin attempt, one per sweep item, and never a retry inside
/// one of them (see `audio_box`, `video_probe::probe_video_geometry` and
/// `pin_playlist::start_the_prober`).
#[derive(Debug, Clone, PartialEq)]
pub enum Probed {
    /// No probe has been run for the file, the one that was found nothing a preview can use, or the
    /// answer it gave has been given up to make room for another file (see `make_room`).
    ///
    /// The three are one answer and have to be: a hover that finds this one measures the file
    /// again, and a card that finds it is a card with nothing to lay out at all — which is the
    /// whole of why what a full memo gives up is the file nothing is asking about, and never the
    /// one a card is on screen for.
    NotAsked,
    /// This machine plays the file, and this is what the file holds.
    Track(Track),
}

/// An answer, and when it was last asked for: the two things both memos of this shape are for,
/// because what is held is only half of it and the other half is which of it was wanted longest
/// ago.
struct Held<T> {
    answer: T,
    used: u64,
}

/// Every ask either memo has been made, and so the stamp an answer is given whether it is
/// written or read.
///
/// A count of this run's own asks rather than a clock, and for the reason every such stamp in
/// this app is one: two files asked for inside the same millisecond are two files, and what
/// tells them apart is which of them was asked for rather than what the wall said.
static ASKS: AtomicU64 = AtomicU64::new(0);

/// The stamp for an ask being made now.
fn asked() -> u64 {
    ASKS.fetch_add(1, Ordering::Relaxed)
}

/// Give up what a memo at its bound has no room for: the file nothing has asked about for
/// longest, one file at a time.
///
/// The whole of the reason this is not a clear of everything is the file on screen. Every
/// question a preview asks about a sound is asked of a memo *while the sound is on screen* —
/// the card's own facts four times a second as its clock moves, the length its bar is a share of
/// when it is pressed, which of the two players a key acts on — so the file being looked at is
/// the most-asked-for file in the map and is the last of it to be given up. That is what the
/// stamp on [`Held`] is for, and the asking is what puts it there: an answer that is read is
/// stamped as well as one that is written.
///
/// Clearing instead took the whole memo down the moment a folder ran past the bound, and
/// `Probed::NotAsked` is not a wait once a card is up. It is a card that will not be laid out
/// again, so its clock stops with no clock drawn rather than waiting for the probe a hover
/// waits for; a bar whose share is no length and so answers no press; a Space that toggles
/// nothing; and a walk that steps over the file it landed on, because the player it was to
/// start could not be started for a file no probe has answered for. Every one of those follows
/// from the file being *forgotten* rather than from its never having been read, and a file being
/// forgotten is what a bound is for and what the file on screen must not be.
fn make_room<T>(held: &mut HashMap<head::Key, Held<T>>, bound: usize) {
    while held.len() > bound {
        let Some(oldest) = held
            .iter()
            .min_by_key(|(_, answer)| answer.used)
            .map(|(key, _)| key.clone())
        else {
            break;
        };

        held.remove(&oldest);
    }
}

/// What every probe has said, held per file and version: the track the machine plays the file with,
/// and nothing at all for a file no read has answered for (see [`Probed`]).
static PROBED: Lazy<Mutex<HashMap<head::Key, Held<Track>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// How many files are held: a large music folder swept whole, at a few hundred bytes a track and so
/// a few hundred kilobytes of them, because a sweep that evicts its own answers is a sweep whose
/// rows lose the facts they have just been given (see [`make_room`]).
const TRACKS_MAX_ENTRIES: usize = 4096;

/// What a probe has said about `path`, or that there is nothing to say yet.
pub fn probed(path: &Path) -> Probed {
    let key = head::key(path);

    let Ok(mut held) = PROBED.lock() else {
        return Probed::NotAsked;
    };

    // An answer that is found is asked about, and being asked about is what keeps it in a memo
    // that is full: the card on screen reads its own file between every other pair of hovers,
    // and the sweep that fills the memo is what this is measured against (see `make_room`).
    let Some(found) = held.get_mut(&key) else {
        return Probed::NotAsked;
    };
    found.used = asked();

    Probed::Track(found.answer.clone())
}

/// The track a player would be started from, where there is one to start.
pub fn playable(path: &Path) -> Option<Track> {
    match probed(path) {
        Probed::Track(track) => Some(track),
        Probed::NotAsked => None,
    }
}

/// What this file's answer is, asking for it only where neither memo has one: the memo of this run
/// first, the disk a run before this one wrote second, and `read` last — which is the probe, the
/// engine or the process, and the thing this exists to keep from being run again.
///
/// Every gate that would otherwise probe asks this rather than the probe (see `audio_box`,
/// `video_probe::probe_video_geometry` and `pin_playlist::start_the_prober`), so one file read in
/// an earlier run is one row filled here with nothing started at all.
pub fn ask(path: &Path, read: impl FnOnce() -> Option<Track>) -> Option<Track> {
    if let Probed::Track(track) = probed(path) {
        return Some(track);
    }

    // A hit is held for this run as well, and not written again: it is already on the disk it was
    // read from (see `disk::entry`).
    if let Some(track) = disk::entry(path) {
        hold(path, track.clone());
        return Some(track);
    }

    // A read that found nothing writes nothing, here or on the disk (see [`remember`]).
    let track = read()?;
    remember(path, track.clone());

    Some(track)
}

/// Hold what a probe found for `path`, and nothing for a probe that found nothing: a read that came
/// back empty leaves the file unanswered rather than answered with an emptiness, so the next ask
/// reads it again instead of being told for the rest of the run that the machine cannot play it.
///
/// Only a [`Track`] can be written, which is the same answer `audio_card` and the start of a pin's
/// player both read to decide a file can be previewed (see `audio_engine::audio_card` and
/// `start_audio_playback_at`) — so what is held is only ever what those two accept. Both levels are
/// written: the memo of this run, and the entry the next run reads (see [`ask`] and `disk::store`).
///
/// The disk write is the caller's own thread — the hover's measure thread, the film's load worker,
/// the playlist's prober — and never a tick's, which is the one thread that must not wait for a
/// file to be written (see `pin_playlist::start_the_prober`).
pub fn remember(path: &Path, track: Track) {
    hold(path, track.clone());
    disk::store(path, &track);
}

/// The first level on its own: what this run's memo holds, and nothing written anywhere.
fn hold(path: &Path, track: Track) {
    let key = head::key(path);

    if let Ok(mut held) = PROBED.lock() {
        // The room is made after the file is in rather than before, so that what is given up is
        // never the file that has just arrived: the answer being written is the freshest thing in
        // the map by a whole file's worth of asks, which is what a file the pointer has just
        // moved to is.
        held.insert(
            key,
            Held {
                answer: track,
                used: asked(),
            },
        );
        make_room(&mut held, TRACKS_MAX_ENTRIES);
    }
}

/// The lock every test that writes either memo holds for its whole run: both are process-wide values
/// shared with every other test in the binary, and two of these at once is one test's sweep giving up
/// the file another test's card is asking for (see `pin_window::tests::ONE_AT_A_TIME`).
#[cfg(test)]
static ONE_AT_A_TIME: Mutex<()> = Mutex::new(());

/// The gain that brings a file's measured loudness to the level files are normalized to, where
/// one has been measured for it: the number the tray's `Normalize` is played at, on top of the
/// level the sound is played at.
///
/// It is held per file and version like the track above it, and the two are measured at different
/// moments and by different questions: a track is what the machine has for the file, and this is
/// what the file asks for — which is why a file can have one and no other (see [`gain`]).
static GAINS: Lazy<Mutex<HashMap<head::Key, Held<f64>>>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// How many files are held. The same bound and the same reasoning as every other memo of this
/// shape: a folder swept a file at a time, and the file a player is being started with is the one
/// that sweep left behind last (see [`make_room`]).
const GAINS_MAX_ENTRIES: usize = 512;

/// The gain this file's loudness was measured to ask for, where one has been measured.
///
/// Nothing is an answer this tells apart from a gain of one: a file nothing has measured is one a
/// hover would have to wait for a decode of, while a measured file that already stands at the
/// target — or one that is silence rather than sound, which measures as nothing at all — is played
/// as it holds. Both are played through the filter rather than around it, so which of the two a
/// caller is asking about does not change what is heard (see `start_audio_playback`).
pub fn gain(path: &Path) -> Option<f64> {
    let key = head::key(path);

    // Stamped on the way out for the reason the track above is: a gain a player is being started
    // with is a gain that was asked about, and a memo that gave it up would start the next
    // player without the filter and measure the file again to get it back.
    let mut held = GAINS.lock().ok()?;
    let found = held.get_mut(&key)?;
    found.used = asked();

    Some(found.answer)
}

/// Hold the gain a file's measured loudness asked for.
pub fn remember_gain(path: &Path, gain: f64) {
    let key = head::key(path);

    if let Ok(mut held) = GAINS.lock() {
        held.insert(
            key,
            Held {
                answer: gain,
                used: asked(),
            },
        );
        make_room(&mut held, GAINS_MAX_ENTRIES);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A file no probe has run for, named so that the tests here cannot be each other's.
    ///
    /// A name is all a file needs to be held under: a file that is not on the machine is keyed
    /// by its path and no version, which is the same answer a file that is there and unreadable
    /// gives (see `head::key`).
    fn a_file(case: &str, nth: usize) -> std::path::PathBuf {
        std::env::temp_dir()
            .join("rust-hover-preview-audio-track")
            .join(case)
            .join(format!("{nth}.flac"))
    }

    /// A sound the machine plays, at a length its card's bar is a share of.
    fn a_track(duration: f64) -> Track {
        Track {
            player: Player::Ffmpeg,
            codec: Some("FLAC".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(900_000),
            duration: Some(duration),
        }
    }

    /// A card that is on screen keeps its file, and a folder walked past the bound is what fills
    /// this memo: a file is asked about here as often as the preview loop asks about the one its
    /// card is drawn for, which is the whole of what a full memo has to be able to tell apart
    /// from a file that was swept a long time ago.
    ///
    /// What clearing the memo instead of giving up one file cost is the file on screen losing its
    /// answer to whichever file happened to cross the bound: a card whose clock stops because it
    /// will not be laid out again, a bar that answers no press, and a walk that steps over the
    /// sound it landed on (see `make_room`).
    #[test]
    fn a_file_the_card_on_screen_is_drawn_for_is_the_one_a_full_memo_does_not_give_up() {
        let _one = ONE_AT_A_TIME.lock();
        let on_screen = a_file("on-screen", 0);

        remember(&on_screen, a_track(180.0));

        for nth in 0..=TRACKS_MAX_ENTRIES {
            remember(&a_file("swept", nth), a_track(180.0));

            // What the preview loop does to the file it is showing, four times a second, between
            // every other pair of hovers.
            assert!(probed(&on_screen) == Probed::Track(a_track(180.0)));
        }

        assert_eq!(
            probed(&on_screen),
            Probed::Track(a_track(180.0)),
            "the file a card is drawn for keeps its track while a folder is walked past the bound"
        );
        assert_eq!(
            probed(&a_file("swept", 0)),
            Probed::NotAsked,
            "and the file that sweep reached first is the one the bound gives up"
        );
    }

    /// What every gate that asks this memo does: one ask per event, and an ask again only where the
    /// answer was nothing usable. Nothing here retries the read twice inside one ask, which is what
    /// keeps a file nothing will ever play from being read on every frame (see `audio_box`,
    /// `video_probe::probe_video_geometry` and `pin_playlist::start_the_prober`).
    fn ask_once(
        path: &Path,
        read: impl FnOnce() -> Option<Track>,
        reads: &mut usize,
    ) -> Option<Track> {
        if matches!(probed(path), Probed::NotAsked) {
            *reads += 1;

            if let Some(track) = read() {
                remember(path, track);
            }
        }

        playable(path)
    }

    /// A read that found nothing is not an answer, and a memo that held it as one was the end of the
    /// file for the rest of the run: every gate above asks only what a `NotAsked` file has, so an
    /// emptiness stored here was never read again — no card, no player, no facts in a row, until the
    /// app was restarted (see [`remember`]).
    #[test]
    fn a_read_that_found_nothing_is_not_held_and_the_next_ask_reads_the_file_again() {
        let _one = ONE_AT_A_TIME.lock();
        let path = a_file("nothing-found", 0);
        let mut reads = 0;

        assert_eq!(ask_once(&path, || None, &mut reads), None);
        assert_eq!(
            probed(&path),
            Probed::NotAsked,
            "a file nothing could be read from is a file nothing has answered for"
        );

        // The same file asked about again — the next hover, the next sweep — is read again, and what
        // this read finds is what is held.
        assert_eq!(
            ask_once(&path, || Some(a_track(240.0)), &mut reads),
            Some(a_track(240.0))
        );
        assert_eq!(
            reads, 2,
            "the read that found nothing was not kept, so the file was read a second time"
        );

        // And a file with a track is a file that answers from the memo: one read, however many asks.
        assert_eq!(
            ask_once(
                &path,
                || panic!("a file with a track was read again"),
                &mut reads
            ),
            Some(a_track(240.0))
        );
        assert_eq!(
            reads, 2,
            "a track is held, and every ask after it is answered from the memo"
        );
    }

    /// A folder swept whole is a folder whose every row is a file this memo is asked about, and the
    /// bound has to hold the sweep that fills it: an answer given up while its own folder is still
    /// being read is a row that loses the facts it was just drawn with (see `playlist_row`).
    #[test]
    fn a_folder_swept_past_the_old_bound_keeps_every_answer_it_landed() {
        let _one = ONE_AT_A_TIME.lock();
        let swept = TRACKS_MAX_ENTRIES.min(1_500);

        for nth in 0..swept {
            remember(&a_file("wide-sweep", nth), a_track(200.0));
        }

        assert!(
            swept > 512,
            "this folder is wider than the bound this memo had: {swept}"
        );

        for nth in 0..swept {
            assert_eq!(
                probed(&a_file("wide-sweep", nth)),
                Probed::Track(a_track(200.0)),
                "row {nth} of a folder swept whole keeps its own answer"
            );
        }
    }

    /// The same for the gain beside the track, which is read where a player is started rather than
    /// where a card is drawn: a memo that gave it up starts the next player without the filter
    /// the file's peak asked for, and the file is measured again to get it back.
    #[test]
    fn a_gain_a_player_is_being_started_with_is_the_one_a_full_memo_does_not_give_up() {
        let _one = ONE_AT_A_TIME.lock();
        let on_screen = a_file("normalized", 0);

        remember_gain(&on_screen, 4.0);

        for nth in 0..=GAINS_MAX_ENTRIES {
            remember_gain(&a_file("normalized-swept", nth), 1.0);

            // What `start_audio_playback` reads of the file it is about to play.
            assert_eq!(gain(&on_screen), Some(4.0));
        }

        assert_eq!(
            gain(&on_screen),
            Some(4.0),
            "the gain a sound is played at survives a folder walked past the bound"
        );
    }

    /// A bound is a bound: the memos are there so that a run that sweeps a whole drive does not
    /// grow without end, and giving up one file at a time rather than all of them is not a
    /// licence to hold more of them.
    #[test]
    fn a_memo_asked_to_hold_more_than_its_bound_holds_no_more_than_its_bound() {
        let _one = ONE_AT_A_TIME.lock();

        for nth in 0..(TRACKS_MAX_ENTRIES + 8) {
            remember(&a_file("overflow", nth), a_track(10.0));
            remember_gain(&a_file("overflow", nth), 1.0);
        }

        assert!(
            PROBED.lock().expect("the memo's own lock").len() <= TRACKS_MAX_ENTRIES,
            "a memo given twice its bound holds no more than its bound: {}",
            PROBED.lock().expect("the memo's own lock").len()
        );
        assert!(
            GAINS.lock().expect("the memo's own lock").len() <= GAINS_MAX_ENTRIES,
            "a memo given twice its bound holds no more than its bound: {}",
            GAINS.lock().expect("the memo's own lock").len()
        );
    }
}

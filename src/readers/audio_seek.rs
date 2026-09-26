//! Where in a file a sound starts, and where each sound was last left.
//!
//! Two halves of one question, and the question belongs to a sound alone: a hover on a picture
//! or a document is answered in a moment and the preview goes away, while a hover on a sound is
//! a few seconds of the file *being listened to* — and a file being listened to again is
//! usually wanted from where it was left rather than from the top. What the tray's
//! `Volume → Audio Seek` names is how a hover answers that, and [`AudioSeek`] holds the four
//! answers: where the file was left, at its beginning, half way in, or anywhere at all.
//!
//! The first of those is the only one this module remembers anything for, and the memory is
//! deliberately a piece of disk rather than a piece of the run: a hover that was resumed after
//! a restart is the whole point of it, and a map of a path to a number of seconds is a few
//! dozen bytes — cheap enough that keeping it under the temp folder costs less than deciding
//! what is worth keeping. It is held in memory once read and written back when it changes,
//! which is what a run pays: one read at the first sound hovered, and one small write per five
//! seconds of listening and per hover that ends (see [`remember`] and [`flush`]).
//!
//! What the file is bounded by is a ceiling of one gigabyte, in the code rather than in
//! `config.ini` or the tray, and there is deliberately no setting beside it: what a user would
//! be answering with such a setting is a question about a file of a few dozen bytes per sound
//! ever hovered, and the ceiling is not a budget anyone tunes — it is what keeps the one case
//! that could grow it from growing it without end (see [`MEMORY_LIMIT_BYTES`]).
//!
//! The position itself is the player's: the engine's own clock where Windows plays a sound and
//! this app's clock over the player's start where FFmpeg does (see `audio_clock`). Nothing here
//! measures anything, and nothing here knows what a sound is — what it is handed is where the
//! sound on screen had got to, and what it hands back is where the next hover of that file
//! starts.

use crate::config::config::AudioSeek;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

/// How much of the temp folder the memory may take.
///
/// A ceiling rather than a budget, and one no honest use of the file can reach: an entry is
/// thirty-odd bytes — a hash of a path, the seconds into the file, and when it was touched —
/// so a gigabyte of them is some thirty million sounds hovered, which is a machine that has
/// listened to every file it owns and every file it ever deleted as well. What it is here for
/// is the other end of that: a sound hovered is a line written, nothing ever prunes a line
/// whose file is gone, and what this says is that even a machine left running over a lifetime
/// of hovers has a file of a known, finite size rather than one that grows with the years. The
/// least recently left-off sounds are the ones given up first (see `render`), so what a memory
/// at the ceiling has forgotten is what nothing has listened to in the longest time.
const MEMORY_LIMIT_BYTES: u64 = 1024 * 1024 * 1024;

/// How long a position may stand unwritten while a sound plays.
///
/// The memory is written as it changes rather than per tick — the card is repainted four times
/// a second and a write per repaint would be four writes a second for a number that moved by a
/// quarter of one — and a hover that ends writes what it has whatever this says (see [`flush`]),
/// which is the write that matters: a file the pointer has left is a file whose position is
/// settled. What is left between the two is a crash or a kill, which costs at most this much of
/// the sound's position rather than the whole of it.
const FLUSH_INTERVAL: Duration = Duration::from_secs(5);

/// Where the memory is kept: a folder of its own under the temp folder, because the loose files
/// directly under that one are what a run which ended mid-render left behind and are swept on
/// the way up (see `document_cache::discard_leftovers`), and a sound hovers were remembered
/// across runs are not that.
fn memory_path() -> PathBuf {
    #[cfg(test)]
    let root = std::env::temp_dir().join("rust-hover-preview-audio-seek-tests");
    #[cfg(not(test))]
    let root = std::env::temp_dir().join("rust-hover-preview").join("audio");

    root.join("seek.txt")
}

/// Where each sound was left, as one line a file: the key of the file it is about, the seconds
/// into it the sound had played, and the moment it was last left at.
///
/// The last of those is what the ceiling gives up first and nothing else reads it, so it is a
/// plain count of seconds since the epoch rather than anything this run has to keep. The key is
/// a hash rather than the path itself, for the reason every other file this app writes under
/// the temp folder is named the way it is: a line is a fixed size whatever the file is called,
/// and nothing about a path can be read as the shape of the line it is on.
#[derive(Debug, Clone, Copy, PartialEq)]
struct Entry {
    /// How far into the file the sound had played, in seconds.
    position: f64,
    /// When it was last left off, in seconds since the epoch.
    used: u64,
}

/// Everything remembered, and what the run knows about the file it is written to.
struct Memory {
    /// Where each sound was left, by the key of the file it is about.
    entries: HashMap<String, Entry>,
    /// Whether anything has been remembered since the file was last written.
    dirty: bool,
    /// Whether the file has been read in yet. Nothing is read until a sound asks: the launch
    /// has a tray icon to put up, and a memory nothing has hovered for is not part of it.
    loaded: bool,
    /// When the file was last written, for the interval above.
    written: Option<Instant>,
}

/// The memory, held as the process's own copy of the file.
static MEMORY: Lazy<Mutex<Memory>> = Lazy::new(|| {
    Mutex::new(Memory {
        entries: HashMap::new(),
        dirty: false,
        loaded: false,
        written: None,
    })
});

/// Where a sound should start playing, for the way the tray has asked for.
///
/// Every answer but the first is arithmetic: the beginning is no seconds in, the middle is half
/// of the length the probe read, and anywhere at all is a fraction of it drawn from the clock
/// (see `anywhere`). A file whose length is not known is a file none of those three can be
/// asked of, and one is answered with the beginning — which is also what a hover did before
/// there was a setting, and the only answer that needs nothing known about the file.
///
/// `Remember` is the memory: a file this app has left off somewhere starts there, and one it
/// has never hovered — or one whose remembered position is past a length the file no longer
/// has — starts at the beginning.
pub fn start_position(path: &Path, seek: AudioSeek, duration: Option<f64>) -> f64 {
    let remembered = match seek {
        AudioSeek::Remember => remembered(path),
        _ => None,
    };

    planned(seek, remembered, duration, anywhere())
}

/// The answer the four ways of starting a sound come to, from what is remembered about the file
/// and a fraction of it to land at.
///
/// Split out of [`start_position`] so that the rule can be read and tested without a file, a
/// clock or a random number: what the three are is the file's own length, the one number the
/// answers that are not the beginning all lean on.
fn planned(seek: AudioSeek, remembered: Option<f64>, duration: Option<f64>, pick: f64) -> f64 {
    let length = duration.filter(|seconds| seconds.is_finite() && *seconds > 0.0);

    match seek {
        AudioSeek::Remember => remembered
            .filter(|position| position.is_finite() && *position >= 0.0)
            .map(|position| match length {
                // A file that has been edited since it was listened to — shortened, or
                // re-encoded to something else — is one the old position may not even be
                // inside of, so it is brought into the file rather than handed to a player
                // that would refuse it or play it from wherever it liked.
                Some(length) => position.min(length * 0.99),
                None => position,
            })
            .unwrap_or(0.0),
        AudioSeek::Start => 0.0,
        AudioSeek::Middle => length.map(|length| length / 2.0).unwrap_or(0.0),
        AudioSeek::Random => length.map(|length| length * pick).unwrap_or(0.0),
    }
}

/// How far into a file `path` was left off, where anything is remembered about it.
fn remembered(path: &Path) -> Option<f64> {
    let key = key(path);

    with_memory(|memory| {
        memory
            .entries
            .get(&key)
            .map(|entry| entry.position)
            .filter(|position| position.is_finite() && *position >= 0.0)
    })
}

/// Remember that the sound of `path` had played to `position` seconds, writing the memory out
/// when enough has changed since the last write for a write to be worth it.
///
/// What this costs a hover is nothing at all unless the mode is the one that reads it back:
/// the other three ways of starting a sound are rules rather than memories, and a run that is
/// on one of them neither reads a file nor writes one (see `start_position`).
pub fn remember(path: &Path, position: f64) {
    if !position.is_finite() || position < 0.0 {
        return;
    }

    let key = key(path);

    with_memory(|memory| {
        memory
            .entries
            .insert(key, Entry { position, used: now() });
        memory.dirty = true;

        if memory.written.is_none_or(|at| at.elapsed() >= FLUSH_INTERVAL) {
            write(memory);
        }
    });
}

/// Write out what has been remembered since the last write, however recently that was.
///
/// It is what a hover that ends calls: the position of a sound is settled the moment its file
/// is left, and a memory waiting out the interval above would be a position the next hover of
/// the file was answered from the position before — which is a sound dropped back a few seconds
/// between two hovers of it, and exactly the thing this module exists to avoid.
pub fn flush() {
    with_memory(|memory| {
        if memory.dirty {
            write(memory);
        }
    });
}

/// Ask the memory a question, reading the file in the first time one is asked.
fn with_memory<T>(ask: impl FnOnce(&mut Memory) -> T) -> T {
    let Ok(mut memory) = MEMORY.lock() else {
        // A poisoned lock is another thread's panic in a map of positions, which is not a
        // reason to answer a hover with a player that will not start: what is remembered is
        // read as nothing and written over.
        return ask(&mut Memory {
            entries: HashMap::new(),
            dirty: false,
            loaded: true,
            written: None,
        });
    };

    if !memory.loaded {
        memory.entries = read();
        memory.loaded = true;
    }

    ask(&mut memory)
}

/// The file as the map it holds, with anything a line does not say dropped rather than guessed
/// at: a memory is a convenience, and a file something else has written into is answered with
/// the sounds that are still readable in it.
fn read() -> HashMap<String, Entry> {
    let Ok(text) = std::fs::read_to_string(memory_path()) else {
        return HashMap::new();
    };

    parse(&text)
}

/// One line a file, as the file it came from: `key\tposition\tlast left off`.
fn parse(text: &str) -> HashMap<String, Entry> {
    let mut entries = HashMap::new();

    for line in text.lines() {
        let line = line.trim();
        if line.is_empty() || line.starts_with(';') {
            continue;
        }

        let mut fields = line.split('\t');

        let Some(key) = fields.next().filter(|key| !key.is_empty()) else {
            continue;
        };
        let Some(position) = fields.next().and_then(|value| value.parse::<f64>().ok()) else {
            continue;
        };
        let Some(used) = fields.next().and_then(|value| value.parse::<u64>().ok()) else {
            continue;
        };
        if !position.is_finite() || position < 0.0 {
            continue;
        }

        entries.insert(
            key.to_string(),
            Entry {
                position,
                used,
            },
        );
    }

    entries
}

/// Write the memory out, giving up what the ceiling leaves no room for.
///
/// Nothing is written where nothing has changed, and nothing is left half-written where the
/// write fails: the text goes to a file beside the real one and is moved over it, so a run that
/// is killed mid-write leaves the memory it had rather than one that cannot be read.
fn write(memory: &mut Memory) {
    let text = render(&mut memory.entries, MEMORY_LIMIT_BYTES);
    let path = memory_path();

    if let Some(folder) = path.parent() {
        if std::fs::create_dir_all(folder).is_err() {
            return;
        }
    }

    let scratch = path.with_extension("writing");
    if std::fs::write(&scratch, text.as_bytes()).is_err() {
        return;
    }
    if std::fs::rename(&scratch, &path).is_err() {
        let _ = std::fs::remove_file(&scratch);
        return;
    }

    memory.dirty = false;
    memory.written = Some(Instant::now());
}

/// The memory as the text the file holds, with the entries `limit` leaves no room for given up
/// — the ones left off longest ago first — so that the file is never larger than the ceiling
/// however many sounds have been hovered.
///
/// The lines are written in the order of their keys rather than in the order the map hands them
/// back, for the reason `config` writes its own file that way: a file that is the same map
/// written twice is the same bytes twice, which is what lets a change in it mean something.
fn render(entries: &mut HashMap<String, Entry>, limit: u64) -> String {
    if entries.is_empty() {
        return String::new();
    }

    // What each line weighs, in the bytes it is written as — the key, the tab-separated
    // position and moment, and the newline — so that the ceiling is measured on the file the
    // memory would be rather than on the map it is. The one line that says what the file is
    // is not weighed: it does not grow with what is remembered, and a ceiling a few bytes
    // under a gigabyte would be a strange thing to trim a memory for.
    let weighed = |key: &str, entry: &Entry| key.len() as u64 + line_rest(entry).len() as u64;

    let mut order: Vec<(&str, u64)> = entries
        .iter()
        .map(|(key, entry)| (key.as_str(), entry.used))
        .collect();
    order.sort_unstable_by_key(|(_, used)| *used);

    // What is given up first is what was left off longest ago, and only as much of it as the
    // ceiling asks for: a memory with room to spare is a memory that forgets nothing.
    let dropped: Vec<String> = {
        let mut over: u64 = entries
            .iter()
            .map(|(key, entry)| weighed(key, entry))
            .sum::<u64>()
            .saturating_sub(limit);
        let mut dropped = Vec::new();

        for (key, _) in order {
            if over == 0 {
                break;
            }
            if let Some(entry) = entries.get(key) {
                over = over.saturating_sub(weighed(key, entry));
                dropped.push(key.to_string());
            }
        }

        dropped
    };
    for key in &dropped {
        entries.remove(key);
    }

    let mut keys: Vec<&String> = entries.keys().collect();
    keys.sort_unstable();

    let mut text =
        String::from("; Where each sound was left off, for `Volume → Audio Seek → Remember`.\n");
    for key in keys {
        let entry = &entries[key.as_str()];
        text.push_str(key);
        text.push_str(&line_rest(entry));
    }

    text
}

/// The half of a line that is not the key: the position, the moment, and the newline.
fn line_rest(entry: &Entry) -> String {
    format!("\t{:.3}\t{}\n", entry.position, entry.used)
}

/// What a sound file is remembered by: its path, folded the way Windows folds one, so that a
/// file is one sound however it was spelled when the pointer found it.
///
/// A hash of the path rather than the path, because a line has to be a known size — see
/// [`Entry`] — and nothing about the file's own name is in it, so a sound listened to and then
/// re-encoded where it stands is the same sound with a position in a file that may now be a
/// different length (which `planned` is what brings back into the file).
fn key(path: &Path) -> String {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    path.to_string_lossy().to_lowercase().hash(&mut hasher);

    format!("{:016x}", hasher.finish())
}

/// A fraction of a file to start a sound at, for the way of starting one that is anywhere at
/// all.
///
/// Nothing here needs a generator anyone could reproduce, and nothing is seeded from the file:
/// two hovers of the same sound should not land in the same place, which is the whole of what
/// `Random` is for. What it is drawn from is the clock, mixed with a count of how many of these
/// have been drawn in this run — the clock alone is a value that barely moves between two
/// hovers a moment apart, and the count alone would be a sequence the same every run.
fn anywhere() -> f64 {
    static DRAWN: AtomicU64 = AtomicU64::new(0);

    let clock = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|since| since.as_nanos() as u64)
        .unwrap_or(0);
    let roll = clock ^ DRAWN.fetch_add(1, Ordering::Relaxed).wrapping_mul(0x9E37_79B9_7F4A_7C15);

    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    roll.hash(&mut hasher);

    // The top fifty-three bits of the hash are a number that fits a `double` exactly, which is
    // what makes the division land in `0.0..1.0` without a value ever coming back as one.
    (hasher.finish() >> 11) as f64 / (1u64 << 53) as f64
}

/// How many seconds have passed since the epoch, as the moment a sound was last left off.
fn now() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|since| since.as_secs())
        .unwrap_or(0)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Each of the four ways of starting a sound is the position it names: where the file was
    /// left, the beginning, half of it, and the fraction of it a roll lands at.
    #[test]
    fn every_way_of_starting_a_sound_is_the_position_it_names() {
        let length = Some(200.0);

        assert_eq!(
            planned(AudioSeek::Remember, Some(42.0), length, 0.5),
            42.0,
            "a sound that was left somewhere starts where it was left"
        );
        assert_eq!(
            planned(AudioSeek::Remember, None, length, 0.5),
            0.0,
            "a sound nothing is remembered about starts at the beginning"
        );
        assert_eq!(
            planned(AudioSeek::Start, Some(42.0), length, 0.5),
            0.0,
            "and so does every sound, where the beginning is what was asked for"
        );
        assert_eq!(
            planned(AudioSeek::Middle, None, length, 0.5),
            100.0,
            "the middle of the file is half of what the probe read"
        );
        assert_eq!(
            planned(AudioSeek::Random, None, length, 0.25),
            50.0,
            "anywhere at all is the fraction of the file the roll landed at"
        );
    }

    /// A file whose length is not known is a file only one of the four can be asked of: the
    /// beginning. The rest are answered with it rather than with a player handed a position it
    /// has no whole to measure against.
    #[test]
    fn a_sound_of_unknown_length_starts_at_the_beginning() {
        for seek in [AudioSeek::Middle, AudioSeek::Random] {
            assert_eq!(
                planned(seek, None, None, 0.5),
                0.0,
                "{seek:?} is a share of a length, and there is no length"
            );
            assert_eq!(
                planned(seek, None, Some(0.0), 0.5),
                0.0,
                "and a file that says it lasts no time at all is one with no middle either"
            );
        }

        assert_eq!(
            planned(AudioSeek::Remember, Some(30.0), None, 0.5),
            30.0,
            "what is remembered about a file is a position rather than a share of one"
        );
    }

    /// A remembered position is brought into the file it is about, because the file may have
    /// been edited — shortened, or replaced by something else — since it was listened to.
    #[test]
    fn a_remembered_position_is_brought_into_the_file_it_is_about() {
        assert_eq!(
            planned(AudioSeek::Remember, Some(900.0), Some(120.0), 0.5),
            118.8,
            "a position past the end of a shortened file is brought back inside it"
        );
        assert_eq!(
            planned(AudioSeek::Remember, Some(-4.0), Some(120.0), 0.5),
            0.0,
            "and a position that is not one is not one to hand a player"
        );
    }

    /// A roll is a fraction of a file and always the same kind of number: inside `0.0..1.0`,
    /// which is what keeps `Random` a place *in* the sound — a value of one would be the end of
    /// it and a value past that no place at all.
    #[test]
    fn a_roll_lands_inside_the_file() {
        for _ in 0..64 {
            let roll = anywhere();
            assert!(
                (0.0..1.0).contains(&roll),
                "a roll is a fraction of the file, and this one is {roll}"
            );
        }
    }

    /// The memory is the file it writes, read back: what a run remembers is what the run after
    /// it starts from, which is the whole of what the file is for.
    #[test]
    fn the_memory_is_written_as_the_file_it_is_read_back_from() {
        let held = parse("6f1c0a4b9d2e3f40\t12.500\t1700000000\n");
        let entry = held["6f1c0a4b9d2e3f40"];
        assert_eq!(entry.position, 12.5);
        assert_eq!(entry.used, 1_700_000_000);

        let mut entries = held.clone();
        let text = render(&mut entries, MEMORY_LIMIT_BYTES);
        assert_eq!(
            parse(&text),
            held,
            "a file written out by this app is one it reads back as itself"
        );
        assert!(
            text.starts_with(';'),
            "and it says what it is, since it is a file a person may find: {text}"
        );
    }

    /// A line the memory cannot make sense of is left where it is rather than answered with a
    /// position nothing asked for: the file is a convenience, and a file half-written by
    /// something else still gives up the sounds that can be read in it.
    #[test]
    fn a_line_the_memory_cannot_read_is_dropped() {
        let held = parse(
            "; a comment\n\
             \n\
             nonsense\n\
             6f1c0a4b9d2e3f40\tnothing\t1700000000\n\
             6f1c0a4b9d2e3f41\t12.500\n\
             6f1c0a4b9d2e3f42\t-3.0\t1700000000\n\
             6f1c0a4b9d2e3f43\t8.000\t1700000000\n",
        );

        assert_eq!(
            held.len(),
            1,
            "only the line that says all three of its fields is one: {held:?}"
        );
        assert!(held.contains_key("6f1c0a4b9d2e3f43"));
    }

    /// A remembered position is written where the run after this one reads it: the file under
    /// the temp folder, which is the whole of what remembering is across runs — a hover on a
    /// file this machine listened to yesterday starts where it was left off, and nothing else
    /// in the app keeps anything that says so.
    #[test]
    fn a_remembered_position_is_written_where_the_next_run_reads_it() {
        let path = Path::new(r"C:\Music\Kind of Blue - So What (Remastered).flac");
        let _ = std::fs::remove_file(memory_path());

        assert_eq!(
            remembered(path),
            None,
            "a file nothing has been listened to is one nothing is remembered about"
        );

        remember(path, 42.5);
        flush();

        let written = std::fs::read_to_string(memory_path()).expect("the memory, as written");
        assert_eq!(
            parse(&written).get(&key(path)).map(|entry| entry.position),
            Some(42.5),
            "what the memory is written as is what it is read back from: {written}"
        );
        assert_eq!(
            remembered(path),
            Some(42.5),
            "and what it answers with is the position the file was left at"
        );

        // A position that is not one — a player that never started, a deck that reports
        // nonsense — is not one anything is remembered about either.
        let nonsense = Path::new(r"C:\Music\nothing playing.mp3");
        remember(nonsense, f64::NAN);
        assert_eq!(remembered(nonsense), None);

        let _ = std::fs::remove_file(memory_path());
    }

    /// The ceiling is what bounds the file, and what it gives up is what has been left off
    /// longest ago: a memory at the ceiling forgets the sounds nothing has listened to in the
    /// longest time rather than whichever ones the map happened to hand back first.
    ///
    /// The ceiling the app ships with is a gigabyte, which no test can fill; what is tested is
    /// the rule it is applied by, at a ceiling of a few lines.
    #[test]
    fn the_ceiling_gives_up_what_was_left_off_longest_ago() {
        let sound = |index: u64| {
            (
                format!("{index:016x}"),
                Entry {
                    position: index as f64,
                    used: 1_700_000_000 + index,
                },
            )
        };
        let mut entries: HashMap<String, Entry> = (0..4).map(sound).collect();

        // A memory with room for all four keeps all four.
        let text = render(&mut entries, 4096);
        assert_eq!(entries.len(), 4, "nothing is given up where everything fits");
        assert_eq!(text.lines().count(), 5, "a header and one line per sound");

        // And one with room for two keeps the two that were left off most recently.
        let mut entries: HashMap<String, Entry> = (0..4).map(sound).collect();
        let line = line_rest(&entries["0000000000000000"]).len() as u64 + 16;
        let text = render(&mut entries, line * 2);

        let mut kept: Vec<&String> = entries.keys().collect();
        kept.sort();
        assert_eq!(
            kept,
            vec!["0000000000000002", "0000000000000003"],
            "the sounds left off longest ago are the ones a memory at the ceiling forgets"
        );
        assert_eq!(
            text.lines().filter(|line| !line.starts_with(';')).count(),
            2,
            "and what is written is only what the ceiling has room for: {text}"
        );
    }
}

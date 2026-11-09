//! The second level of the track memo: what a probe found, kept on disk so that a run after this
//! one opens a folder and fills its rows without probing a file (see [`super::ask`]).
//!
//! One file per answer, `<general>/audio-track-<hash>`, beside the subtitle folders and inside the
//! same `general_disk_cache_mb` budget. A file rather than an index because the budget gives up the
//! least recently written thing first, and a file's own modified time is that answer without
//! anything having to be kept in step with it: one entry given up is one file removed, no category
//! goes at once, and no reader has to read-modify-write a shared index that three threads write to.
//!
//! What is held is only ever a [`Track`]: a read that found nothing is not stored here either (see
//! `super::remember`), because an emptiness on the disk would outlive the run that found it — a
//! player installed afterwards, a format a later version learns to read — and be the answer for the
//! rest of the machine's life with nothing to ever revisit it.

use super::{Player, Track};
use crate::formats::head;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

/// The meaning of the bytes below. An entry written by another version is a miss rather than a
/// guess: the fields a track is made of are the fields a card is drawn from, and a misread one is a
/// card of wrong facts where a miss costs one probe.
const VERSION: u8 = 1;

/// The prefix every entry's name carries, which is what tells this folder's files from the folders
/// its films' subtitles are kept in.
const PREFIX: &str = "audio-track-";

/// What was written when. Every entry counts against the one budget, so the folder is trimmed now
/// and then rather than after each file: a sweep of a folder's worth of songs writes one file per
/// song, and a walk of the whole folder after each of them reads the folder a folder's worth of
/// times over.
///
/// shortcut: an entry that takes the folder over the budget can stand until the 64th write or the
/// tray's own trim; upgrade to a running byte count if the budget is ever raised far enough for it
/// to matter.
const TRIM_EVERY_WRITES: u64 = 64;

static WRITES: AtomicU64 = AtomicU64::new(0);

/// The folder the entries are kept in: the app's `general` folder, shared with the subtitles (see
/// `subtitle_files::general_folder`).
fn folder() -> PathBuf {
    crate::ui::preview_window::general_folder()
}

/// The name this version of this file's answer is kept under.
///
/// The same recipe every other cache here names a file by — the lowercased path, its length and its
/// modification time, hashed (see `subtitle_files::key` and `document_cache::key`) — so a file saved
/// again is a name nothing has written and the entry it had is left to the trim.
fn entry_path(key: &head::Key) -> PathBuf {
    let (path, modified, len) = key.stamp();
    let mut hasher = DefaultHasher::new();

    path.to_string_lossy().to_lowercase().hash(&mut hasher);
    // `None` is a file whose entry could not be read, and it hashes like the answer it is: a
    // stamp nothing can be compared against is not one an entry is kept under (`store`).
    nanos(modified).hash(&mut hasher);
    len.hash(&mut hasher);

    folder().join(format!("{PREFIX}{:016x}", hasher.finish()))
}

/// Nanoseconds since the epoch, which is the unit the name above hashes and the entry states.
fn nanos(modified: Option<SystemTime>) -> Option<u64> {
    modified
        .and_then(|modified| modified.duration_since(UNIX_EPOCH).ok())
        .map(|since| u64::try_from(since.as_nanos()).unwrap_or(u64::MAX))
}

/// The track kept for this version of this file, where one was written and it is whole.
///
/// Anything that does not answer exactly — another version's bytes, a stamp that is not this
/// file's, a file cut short by a run that ended inside the write, an entry of no bytes at all — is
/// a miss *and* is given up here rather than left to be read again for the rest of the machine's
/// life. Nothing panics: an entry is derived data, and one that cannot be read is one probe.
pub(super) fn entry(path: &Path) -> Option<Track> {
    let key = head::key(path);
    let file = entry_path(&key);
    let track = decode(&std::fs::read(&file).ok()?, &key);

    if track.is_none() {
        let _ = std::fs::remove_file(&file);
    }

    track
}

/// Write what a probe found for this file, and give up who it displaced where the folder has
/// filled.
///
/// Silent on failure: a folder that cannot be written, or a file that cannot be replaced, costs the
/// next run a probe rather than a preview. A file whose own entry cannot be read has no stamp to
/// keep an answer under, and nothing is written for it.
pub(super) fn store(path: &Path, track: &Track) {
    let key = head::key(path);

    if nanos(key.stamp().1).is_none() {
        return;
    }

    let file = entry_path(&key);

    if std::fs::create_dir_all(folder()).is_err() {
        return;
    }

    if write_whole(&file, &encode(&key, track)).is_none() {
        return;
    }

    // Trimming is the general folder's own business — the film folders beside these entries are
    // counted against the same budget, and the oldest of either kind goes first (see
    // `subtitle_files::trim_folder`).
    if WRITES
        .fetch_add(1, Ordering::Relaxed)
        .is_multiple_of(TRIM_EVERY_WRITES)
    {
        crate::ui::preview_window::trim_subtitle_cache();
    }
}

/// Write a file whole, under a name nothing reads until it is complete: a run ended inside the
/// write leaves half an entry under a name of its own rather than an entry a reader would trust
/// (the same discipline the document cache writes a page with).
fn write_whole(path: &Path, bytes: &[u8]) -> Option<()> {
    let writing = path.with_extension("writing");
    std::fs::write(&writing, bytes).ok()?;
    std::fs::rename(&writing, path).ok()?;

    Some(())
}

fn encode(key: &head::Key, track: &Track) -> Vec<u8> {
    let (path, modified, len) = key.stamp();
    let mut bytes = vec![
        VERSION,
        match track.player {
            Player::Native => 0,
            Player::Ffmpeg => 1,
        },
    ];

    bytes.extend_from_slice(&len.to_le_bytes());
    bytes.push(u8::from(nanos(modified).is_some()));
    bytes.extend_from_slice(&nanos(modified).unwrap_or(0).to_le_bytes());

    let name = path.to_string_lossy().to_lowercase();
    bytes.extend_from_slice(&(name.len() as u32).to_le_bytes());
    bytes.extend_from_slice(name.as_bytes());

    let codec = track.codec.as_deref().unwrap_or_default();
    bytes.extend_from_slice(&(codec.len() as u16).to_le_bytes());
    bytes.extend_from_slice(codec.as_bytes());
    bytes.push(u8::from(track.rate.is_some()));
    bytes.extend_from_slice(&track.rate.unwrap_or(0).to_le_bytes());
    bytes.push(u8::from(track.channels.is_some()));
    bytes.extend_from_slice(&track.channels.unwrap_or(0).to_le_bytes());
    bytes.push(u8::from(track.bitrate.is_some()));
    bytes.extend_from_slice(&track.bitrate.unwrap_or(0).to_le_bytes());
    bytes.push(u8::from(track.duration.is_some()));
    bytes.extend_from_slice(&track.duration.unwrap_or(0.0).to_le_bytes());

    bytes
}

/// What an entry says, where it says it about *this* version of *this* file: the stamp written into
/// it is compared rather than trusted, so an entry copied, half-replaced or written by a name that
/// hashed the same is a miss.
fn decode(bytes: &[u8], key: &head::Key) -> Option<Track> {
    let mut entry = Reading { bytes, at: 0 };

    if entry.byte()? != VERSION {
        return None;
    }

    let player = match entry.byte()? {
        0 => Player::Native,
        1 => Player::Ffmpeg,
        _ => return None,
    };

    let len = entry.number(8)?;
    let modified = match entry.byte()? {
        0 => None,
        1 => Some(entry.number(8)?),
        _ => return None,
    };
    let name_len = entry.number(4)? as usize;
    let name = String::from_utf8(entry.take(name_len)?.to_vec()).ok()?;

    let codec_len = entry.number(2)? as usize;
    let codec = String::from_utf8(entry.take(codec_len)?.to_vec()).ok()?;
    let rate = entry
        .optional(|entry| entry.number(4))?
        .map(|rate| rate as u32);
    let channels = entry
        .optional(|entry| entry.number(2))?
        .map(|channels| channels as u16);
    let bitrate = entry
        .optional(|entry| entry.number(4))?
        .map(|bitrate| bitrate as u32);
    let duration = entry.optional(|entry| entry.number(8))?.map(f64::from_bits);

    let (path, written, written_len) = key.stamp();

    if name != path.to_string_lossy().to_lowercase()
        || len != written_len
        || modified != nanos(written)
    {
        return None;
    }

    Some(Track {
        player,
        codec: (!codec.is_empty()).then_some(codec),
        rate,
        channels,
        bitrate,
        duration,
    })
}

/// The bytes of one entry, read a field at a time, and nothing at all where a field runs off the end
/// of what is there — which is what an entry cut short answers.
struct Reading<'a> {
    bytes: &'a [u8],
    at: usize,
}

impl Reading<'_> {
    fn take(&mut self, count: usize) -> Option<&[u8]> {
        let end = self.at.checked_add(count)?;
        let taken = self.bytes.get(self.at..end)?;
        self.at = end;

        Some(taken)
    }

    fn byte(&mut self) -> Option<u8> {
        Some(self.take(1)?[0])
    }

    fn number(&mut self, width: usize) -> Option<u64> {
        Some(
            self.take(width)?
                .iter()
                .enumerate()
                .fold(0u64, |value, (shift, byte)| {
                    value | (u64::from(*byte) << (shift * 8))
                }),
        )
    }

    /// The field under an optional prefix: nothing for a field that is not there, and nothing at all
    /// for one whose prefix is not a mark this version writes.
    fn optional(&mut self, read: impl FnOnce(&mut Self) -> Option<u64>) -> Option<Option<u64>> {
        match self.byte()? {
            0 => Some(None),
            1 => read(self).map(Some),
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests;

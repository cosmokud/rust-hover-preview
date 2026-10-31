//! The pipe a player of somebody else's writes its log lines into, and the thread that empties it.
//!
//! **Why a pipe at all.** A film that is damaged plays — silently, or stuttering, or in a loop of two
//! frames — and the only thing on the machine that says so is the player's own log. Nothing this app
//! decodes can be asked: ffplay and mpv are separate processes, and what they saw is in what they
//! printed (see `engines::corruption_log`, which is what the lines are matched against).
//!
//! **It is `stdout` and not `stderr`, and that is a measurement rather than a preference.** On
//! Windows, mpv 0.41 writes its terminal output to the standard output — with `--terminal=yes` and
//! with the default `auto` alike — and to the standard error not at all: a missing file launched
//! with both redirected printed `[file] Cannot open file '…': Invalid argument` into `stdout` (134
//! bytes) and nothing into `stderr` (0 bytes), twice over. The same pipe is live rather than
//! flushed at exit: a player writing its status line grew the piped file from 297 bytes to 365
//! between the second and the fifth second of a run that was still going. ffplay's own errors do
//! reach `stderr`, but its launch ignores what it decodes (`-err_detect ignore_err`, `+discardcorrupt`)
//! and a film it recovered from leaves nothing at warning level (measured: five bytes, one escape
//! sequence), so it is not wired at all — see `video_playback`.
//!
//! **The thread is the whole of the contract with the player.** A pipe nobody reads fills, and a
//! player blocked on a full pipe stops playing; so the drain starts with the player, reads in chunks,
//! keeps no more than one line, and ends when the pipe closes with the process. Nothing here is on
//! the UI thread, and nothing here holds a lock the paint wants.
//!
//! **And nothing is remembered forever.** The watch of a player is kept under its id and the file it
//! was begun for, and the table is a short one: a pin is taken up within a second of the player that
//! shows its file, so a handful of entries covers every road (see `watch_of`, which is only ever
//! asked about the player a pin was just taken up on).

use super::*;
use crate::engines::corruption_log::CorruptionWatch;

/// The longest line kept: a player's log lines are one sentence, and a run of bytes with no newline
/// in it is not something to grow a buffer for.
const LONGEST_LINE: usize = 512;

/// How many players' watches are kept. A pin is taken up on the player that was begun for its file
/// moments before, so a table this short answers every road and cannot grow with a session's hovers.
const WATCHES_KEPT: usize = 8;

/// What a player's log lines are matched against: one latch, begun before the process is begun.
pub(super) fn a_watch() -> Arc<CorruptionWatch> {
    Arc::new(CorruptionWatch::new())
}

/// The stream a launch gives its player's log lines to, which the drain below empties.
///
/// It is `piped` rather than inherited for the reason every other stream a player is given is
/// `null`: a preview must not write into a console it does not own, and nothing of the player's is
/// to reach a window of this app's.
pub(super) fn a_stream_to_drain() -> Stdio {
    Stdio::piped()
}

/// The latch of the player `pid`, where it was begun for `path`.
///
/// **Both halves are asked.** The id alone would be enough while one player per file stands, and the
/// path is what makes the answer a fact about the file rather than about a moment: a pin taken up on
/// a file it has stepped to asks about a player that has only just been begun, and an id read a
/// moment too early — the outgoing player's — is refused by the path it was begun for rather than
/// answering a card the new file has not earned (see `PinnedPreview::corruption`).
pub(super) fn watch_of(pid: u32, path: &Path) -> Option<Arc<CorruptionWatch>> {
    let watches = WATCHES.lock().ok()?;
    watches
        .iter()
        .find(|watched| watched.pid == pid && watched.path == path)
        .map(|watched| Arc::clone(&watched.watch))
}

/// Remember which player is showing which file, and what its log lines are read as.
pub(super) fn remember(pid: u32, path: &Path, watch: &Arc<CorruptionWatch>) {
    let Ok(mut watches) = WATCHES.lock() else {
        return;
    };

    watches.retain(|watched| watched.pid != pid && watched.path != path);
    if watches.len() >= WATCHES_KEPT {
        watches.remove(0);
    }
    watches.push(WatchedPlayer {
        pid,
        path: path.to_path_buf(),
        watch: Arc::clone(watch),
    });
}

/// Read a player's log to its end on a thread of its own, matching every line and keeping none.
///
/// The thread is the pipe's only reader from the moment the player exists, which is what says a
/// player can never be blocked by its own logging; it ends at the pipe's end of file, which is the
/// process ending.
pub(super) fn drain_a_player_log(stream: std::process::ChildStdout, watch: Arc<CorruptionWatch>) {
    let _ = std::thread::Builder::new()
        .name("player-log".to_string())
        .spawn(move || follow(stream, &watch));
}

/// Match every complete line of `reader` against `watch` until it ends.
///
/// The lines are handed over where they lie: no lowercased copy, no `String` per line, and a line
/// longer than `LONGEST_LINE` is dropped whole rather than grown. A line ending in the `\r\n` a
/// Windows player writes is trimmed of its `\r` first, so a signature at the end of a line is not
/// lost to it.
fn follow(mut reader: impl std::io::Read, watch: &CorruptionWatch) {
    let mut chunk = [0u8; 4096];
    let mut line: Vec<u8> = Vec::with_capacity(LONGEST_LINE);
    let mut over_long = false;

    loop {
        let count = match reader.read(&mut chunk) {
            Ok(0) | Err(_) => return,
            Ok(count) => count,
        };

        for &byte in &chunk[..count] {
            if byte != b'\n' {
                if !over_long {
                    if line.len() < LONGEST_LINE {
                        line.push(byte);
                    } else {
                        // Nothing this long is one of the lines the table names, and keeping it
                        // in pieces would only risk reading half a signature as a whole one.
                        over_long = true;
                    }
                }
                continue;
            }

            if !over_long {
                if line.last() == Some(&b'\r') {
                    line.pop();
                }
                if let Ok(line) = std::str::from_utf8(&line) {
                    watch.observe(line);
                }
            }

            line.clear();
            over_long = false;
        }
    }
}

/// One player whose log lines are being read: the id they are read from, the file it was begun
/// for, and the latch its lines are matched against.
struct WatchedPlayer {
    pid: u32,
    path: PathBuf,
    watch: Arc<CorruptionWatch>,
}

/// The players whose log lines are being read, under the id they are read from and the file they
/// were begun for.
static WATCHES: Lazy<Mutex<Vec<WatchedPlayer>>> = Lazy::new(|| Mutex::new(Vec::new()));

#[cfg(test)]
mod tests {
    use super::*;
    use crate::engines::corruption_log::Weight;

    /// A folder no other test looks at, named for the test that asked for it.
    fn a_folder(name: &str) -> PathBuf {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-player-logs")
            .join(name);
        let _ = std::fs::remove_dir_all(&folder);
        let _ = std::fs::create_dir_all(&folder);
        folder
    }

    #[test]
    fn a_line_is_read_off_the_pipe_and_matched_where_it_lies() {
        let watch = CorruptionWatch::new();
        follow(
            std::io::Cursor::new(
                b"[ffmpeg/video] h264: error while decoding MB 53 19\r\n\
                  [cplayer] Starting playback...\r\n\
                  [mkv] Invalid EBML length at position 20511866\r\n"
                    .to_vec(),
            ),
            &watch,
        );

        assert!(watch.tripped());
    }

    #[test]
    fn the_last_line_of_a_player_that_was_still_writing_it_is_read_too() {
        let watch = CorruptionWatch::new();
        // No newline at the end: the line the player was in the middle of when the pipe closed.
        follow(
            std::io::Cursor::new(b"Failed to recognize file format.".to_vec()),
            &watch,
        );

        assert!(!watch.tripped(), "a line without its end is not a line yet");
    }

    #[test]
    fn a_line_with_no_newline_in_it_never_grows_a_buffer_and_never_trips_anything() {
        let watch = CorruptionWatch::new();
        let mut fired = Vec::new();
        fired.extend(std::iter::repeat_n(b'x', LONGEST_LINE * 4));
        fired.extend_from_slice(b"\n");
        fired.extend_from_slice(b"[mkv] Invalid EBML length\n");

        follow(std::io::Cursor::new(fired), &watch);

        // The long run was dropped, and the line after it was read as the line it is.
        assert!(watch.tripped());
    }

    #[test]
    fn a_watch_is_found_by_the_player_and_the_file_it_was_begun_for() {
        let folder = a_folder("found");
        let first = folder.join("a.mkv");
        let second = folder.join("b.mkv");
        let watch = a_watch();
        remember(4, &first, &watch);

        let found = watch_of(4, &first).expect("the player and its file name their watch");
        assert!(Arc::ptr_eq(&found, &watch));

        // The same player asked about another file, and another player asked about this one: the
        // answer is nothing either way, which is what keeps a card off a file that has not earned
        // it.
        assert!(watch_of(4, &second).is_none());
        assert!(watch_of(8, &first).is_none());

        let _ = std::fs::remove_dir_all(&folder);
    }

    #[test]
    fn the_table_keeps_a_handful_of_players_and_forgets_the_oldest() {
        let folder = a_folder("kept");
        let path = folder.join("a.mkv");

        for pid in 0..(WATCHES_KEPT as u32 * 3) {
            remember(pid * 4, &path, &a_watch());
        }

        // Every player but the oldest few is still answered, and the table itself has not grown.
        assert!(watch_of((WATCHES_KEPT as u32 * 3 - 1) * 4, &path).is_some());
        assert!(!watch_of(0, &path).is_some());
        let kept = WATCHES.lock().map(|watches| watches.len()).unwrap_or(0);
        assert!(kept <= WATCHES_KEPT);

        let _ = std::fs::remove_dir_all(&folder);
    }

    #[test]
    fn a_players_own_line_trips_the_latch_the_pin_is_taken_up_with() {
        let folder = a_folder("pin");
        let path = folder.join("damaged.mkv");
        let watch = a_watch();
        remember(12, &path, &watch);

        follow(
            std::io::Cursor::new(
                b"[mkv] Corrupt file detected. Trying to resync starting from position 20511866...\r\n"
                    .to_vec(),
            ),
            &watch,
        );

        let pin_watch = watch_of(12, &path).expect("the pin reads the watch of its own player");
        assert!(pin_watch.tripped());
        assert_eq!(
            pin_watch.observe("[ffmpeg/video] h264: error while decoding MB 1 2"),
            Some(Weight::Soft)
        );

        let _ = std::fs::remove_dir_all(&folder);
    }
}

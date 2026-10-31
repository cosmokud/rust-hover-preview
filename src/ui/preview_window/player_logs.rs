//! The pipe a player of somebody else's writes its own lines into, and the thread that empties it.
//!
//! **Why a pipe at all.** A film that plays with no sound and no entry in Windows' Volume Mixer has
//! two causes, and only one of them is this app's: the launch may have given the player no track, or
//! the player may have had a track and found no device to open. The second is the player's own to
//! report — mpv says `Could not open/initialize audio device -> no sound` in its log and plays on
//! without it — and nothing this app decodes can be asked: ffplay and mpv are separate processes
//! (see `playback_trace`, which is where the lines are written).
//!
//! **It is `stdout` and not `stderr`, and that is a measurement rather than a preference.** On
//! Windows, mpv 0.41 writes its terminal output to the standard output — with `--terminal=yes` and
//! with the default `auto` alike — and to the standard error not at all: a missing file launched
//! with both redirected printed `[file] Cannot open file '…': Invalid argument` into `stdout` (134
//! bytes) and nothing into `stderr` (0 bytes), twice over. The same pipe is live rather than
//! flushed at exit: a player writing its status line grew the piped file from 297 bytes to 365
//! between the second and the fifth second of a run that was still going. ffplay's own errors do
//! reach `stderr`, which is why a launch of it that keeps a line is given that stream to write on and
//! a thread to empty it (see `drain_a_players_errors`).
//!
//! **A stream is piped only on a launch that has a use for a line, and the thread that empties it is
//! begun on the same launches and no others.** mpv's log is that for a traced launch alone: an
//! ordinary launch of that player writes nothing anywhere and no thread is spawned for it. **FFmpeg's
//! player is the one engine whose lines are kept without the trace**, and for a reason rather than an
//! oversight: its sound is the one thing it can only say on that stream, so a launch the guard stands
//! behind is piped and drained too (see `audio_guard`), while an ordinary hover keeps the `null`
//! stream it always had. Either way the thread is what the pipe costs, and what it buys is a player
//! that can never be blocked by its own writing.
//!
//! **The thread is the whole of the contract with the player.** A pipe nobody reads fills, and a
//! player blocked on a full pipe stops playing; so the drain starts with the player, reads in chunks,
//! keeps no more than one line, and ends when the pipe closes with the process. Nothing here is on
//! the UI thread, and nothing here holds a lock the paint wants.
//!
//! **FFmpeg's player has a drain of its own, and it is the same shape.** Its errors were thrown away
//! entirely — `-loglevel quiet` into a null stream — so a start that found no audio device was a start
//! nothing anywhere said anything about. Every launch that keeps a line is drained now, and the level
//! it is started at is what decides whether there is anything to read (see `drain_a_players_errors`,
//! `playback_trace` and `audio_guard`).

use super::*;

/// The longest line kept: a player's log lines are one sentence, and a run of bytes with no newline
/// in it is not something to grow a buffer for.
const LONGEST_LINE: usize = 512;

/// The stream a launch that keeps its player's own lines writes them into, which the drain below
/// empties — piped only on those launches, so a launch with no reader is never given one.
///
/// It is `piped` rather than inherited for the reason every other stream a player is given is
/// `null`: a preview must not write into a console it does not own, and nothing of the player's is
/// to reach a window of this app's.
pub(super) fn a_stream_to_drain() -> Stdio {
    Stdio::piped()
}

/// Read a player's log to its end on a thread of its own, writing every line into the playback
/// trace and keeping none of it here.
///
/// The thread is the pipe's only reader from the moment the player exists, which is what says a
/// player can never be blocked by its own logging; it ends at the pipe's end of file, which is the
/// process ending.
pub(super) fn drain_a_player_log(stream: std::process::ChildStdout, pid: u32) {
    let _ = std::thread::Builder::new()
        .name("player-log".to_string())
        .spawn(move || {
            lines_of(stream, |line| {
                playback_trace::a_players_line(pid, "mpv log", line);
            });
            // The end of the stream is the end of the player, which is the end of anything watching
            // it (see `audio_guard::a_player_ended`).
            audio_guard::a_player_ended(pid);
        });
}

/// Read a player's own error stream to its end on a thread of its own, writing every line into the
/// playback trace and keeping none of it here — except the lines that are a sound device's failure,
/// which are what a pin's film is begun again for (see `audio_guard::a_players_line`).
///
/// **FFmpeg's player says everything it has to say on the standard error, and this app used to throw
/// it away** — the launch runs it at `-loglevel quiet` with the stream `null`, so a start that found
/// no audio device left no trace anywhere at all (see `playback_trace`). **Every launch that keeps the
/// line pipes it**, and the level a launch is given is what decides whether the player writes anything
/// into it: a film the playback guard stands behind is played at a level that prints, because a failure
/// the player reports on a stream nobody reads is the same silent film with one more reason nobody
/// knows. The thread is the same contract the log's drain keeps: the pipe has one reader from the moment
/// the player exists, so the player can never be blocked by its own writing — and that contract is why
/// the launch that pipes a stream is the launch that empties it, traced or not. A launch with no reader
/// is given `null` and begins no thread. The child stays windowless:
/// `engine_processes::hidden_command` starts it with `CREATE_NO_WINDOW`, so a redirected stream is a
/// pipe and never a console.
pub(super) fn drain_a_players_errors(stream: std::process::ChildStderr, pid: u32) {
    let _ = std::thread::Builder::new()
        .name("player-errors".to_string())
        .spawn(move || {
            lines_of(stream, |line| {
                playback_trace::a_players_line(pid, "ffplay stderr", line);
                audio_guard::a_players_line(pid, line);
            });
            // The end of the stream is the end of the player, which is the end of anything watching
            // it (see `audio_guard::a_player_ended`).
            audio_guard::a_player_ended(pid);
        });
}

/// Hand every complete line of `reader` to `on_line` until the stream ends.
///
/// **The one shape every reader of a player's own writing has**, which is why it is one function: the
/// two streams are read into the same trace, and what a line *is* must not differ between them. The
/// lines are handed over where they lie: no lowercased copy, no `String` per line, and a line longer
/// than `LONGEST_LINE` is dropped whole rather than grown. A line ending in the `\r\n` a Windows
/// player writes is trimmed of its `\r` first, so what reaches the trace is the line the player wrote.
fn lines_of(mut reader: impl std::io::Read, mut on_line: impl FnMut(&str)) {
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
                        // Nothing this long is a line of a player's own prose, and keeping it in
                        // pieces would only risk writing half of it as a whole line.
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
                    on_line(line);
                }
            }

            line.clear();
            over_long = false;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The lines `lines_of` hands over, read off a stream that ends where the bytes do.
    fn the_lines_of(bytes: &[u8]) -> Vec<String> {
        let mut lines = Vec::new();
        lines_of(std::io::Cursor::new(bytes.to_vec()), |line| {
            lines.push(line.to_string())
        });
        lines
    }

    #[test]
    fn a_line_is_read_off_the_pipe_where_it_lies() {
        assert_eq!(
            the_lines_of(
                b"[cplayer] Starting playback...\r\n\
                  [mkv] Invalid EBML length at position 20511866\r\n"
            ),
            vec![
                "[cplayer] Starting playback...",
                "[mkv] Invalid EBML length at position 20511866",
            ],
            "the `\\r` a Windows player ends a line with is trimmed, and the line reaches the \
             trace as it was written"
        );
    }

    #[test]
    fn the_last_line_of_a_player_that_was_still_writing_it_is_not_a_line_yet() {
        assert_eq!(
            the_lines_of(b"Failed to recognize file format."),
            Vec::<String>::new(),
            "a line without its end is not a line yet"
        );
    }

    #[test]
    fn a_line_with_no_newline_in_it_never_grows_a_buffer() {
        let mut stream = Vec::new();
        stream.extend(std::iter::repeat_n(b'x', LONGEST_LINE * 4));
        stream.extend_from_slice(b"\n[cplayer] Starting playback...\n");

        assert_eq!(
            the_lines_of(&stream),
            vec!["[cplayer] Starting playback..."],
            "the long run is dropped whole, and the line after it is read as the line it is"
        );
    }
}

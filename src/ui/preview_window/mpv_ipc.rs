//! Telling a running mpv what to do, over its own JSON IPC.
//!
//! mpv is the one player here that can be spoken to while it plays, and that is the whole reason
//! the mpv route exists as something other than FFmpeg's player (see `video_route`): a volume, a
//! seek, a pause and a track step are lines written down a named pipe rather than a player ended
//! and begun again.
//!
//! **Write-only by decision.** The pipe is opened, one JSON command line is written, and the
//! handle is dropped: there is no reader thread, no request id, no reply parsed and no state kept
//! between commands, so there is nothing left behind to clean up when a player is killed. The app
//! knows what it needs without asking — the sidecar from the filesystem, the level from the pin —
//! so nothing is read back. A command that wants no answer says so with `"async": true` rather
//! than leaving a reply nobody reads in the pipe.
//!
//! The pipe the launch names is remembered against the player's own process id, and forgotten when
//! that player is killed or parked, so a command can never reach a player this app has stopped
//! holding (see `kill_player_process` and `restart_pinned_player`). The name is unique per launch
//! (a counter) so two previews could never share one.

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;

/// The pipe a launch names, unique to that launch.
///
/// A counter rather than the process id because the name is written into the argument list before
/// there is a process to name: `mpv_args` builds the command line, and the player only exists once
/// it has been spawned. It is one apart from the last, so two previews never share a pipe.
pub(super) fn next_pipe_name() -> String {
    static NEXT: AtomicU64 = AtomicU64::new(1);
    format!("rust-hover-preview-mpv-{}", NEXT.fetch_add(1, Ordering::Relaxed))
}

/// The pipe the player behind `pid` listens on, while it is the one this app holds.
static CURRENT_PIPE: Mutex<Option<(u32, String)>> = Mutex::new(None);

/// Remember the pipe a launch named for the player it has just begun.
pub(super) fn remember_pipe(pid: u32, name: &str) {
    if let Ok(mut slot) = CURRENT_PIPE.lock() {
        *slot = Some((pid, name.to_string()));
    }
}

/// Forget the pipe of `pid`, where it is that player's — a kill drops it too.
pub(super) fn forget_pipe(pid: u32) {
    if let Ok(mut slot) = CURRENT_PIPE.lock() {
        if slot.as_ref().is_some_and(|(held, _)| *held == pid) {
            *slot = None;
        }
    }
}

/// The pipe the player behind `pid` listens on, or nothing where that player is not the one this
/// app holds.
fn pipe_for(pid: u32) -> Option<String> {
    CURRENT_PIPE
        .lock()
        .ok()
        .and_then(|slot| {
            slot.as_ref()
                .filter(|(held, _)| *held == pid)
                .map(|(_, name)| name.clone())
        })
}

/// Where a command actually goes: the pipe of the player behind `pid`, or nothing where there is
/// no such player.
fn write(pid: u32, line: String) -> bool {
    let Some(name) = pipe_for(pid) else {
        return false;
    };
    deliver(&name, line)
}

/// Write one command line down the pipe and close it. MPV holds the other end; a pipe that cannot
/// be opened is a player that is not there, and a command for one is not an error.
///
/// **A test run notes the line instead and writes nothing**, which is what keeps the suite from
/// opening a pipe to a player it never began (see `NOTED`); the one test that has a real player to
/// hand a line to says so first (see `write_to_a_real_pipe`). It is one function rather than two so
/// that the bytes a real mpv is handed are the bytes that test writes: the missing newline these
/// lines are read by was a fault in exactly this code, and the copy an ordinary test writes its
/// asserted string from is the copy that could not see it.
fn deliver(name: &str, line: String) -> bool {
    #[cfg(test)]
    if !WRITE_TO_A_REAL_PIPE.load(Ordering::Acquire) {
        if let Ok(mut noted) = NOTED.lock() {
            noted.push(line);
        }
        return true;
    }

    use std::io::Write;
    let path = format!(r"\\.\pipe\{name}");
    let Ok(mut pipe) = std::fs::OpenOptions::new().write(true).open(path) else {
        return false;
    };
    pipe.write_all(line.as_bytes()).is_ok()
}

/// Whether a line is written down a real pipe rather than noted.
///
/// A test run notes every line, so that the suite opens no pipe and starts no player (see
/// `deliver`); the one test that runs against a real mpv sets this first, and the same function
/// then writes the bytes. A flag rather than a second delivery so that the write under test is the
/// write the app does.
#[cfg(test)]
static WRITE_TO_A_REAL_PIPE: std::sync::atomic::AtomicBool =
    std::sync::atomic::AtomicBool::new(false);

/// Write what is told from here on down a real pipe, for the one test that has a player to write to
/// (see `WRITE_TO_A_REAL_PIPE`).
#[cfg(test)]
pub(super) fn write_to_a_real_pipe(real: bool) {
    WRITE_TO_A_REAL_PIPE.store(real, Ordering::SeqCst);
}

/// Set the level mpv is playing at, in the 0–100 the pin's knob is drawn in.
pub(super) fn set_volume(pid: u32, level: u32) -> bool {
    write(
        pid,
        // mpv's JSON IPC reads a message up to its newline, and a line with none is a command it
        // never sees — closing the pipe does not deliver it either — so the newline is part of
        // the line rather than of the write (see `write`).
        format!(r#"{{"command":["set_property","volume",{level}],"async":true}}"#) + "\n",
    )
}

/// Take mpv to a second of its file.
pub(super) fn seek_absolute(pid: u32, seconds: f64) -> bool {
    write(
        pid,
        format!(r#"{{"command":["seek",{seconds:.3},"absolute"],"async":true}}"#) + "\n",
    )
}

/// Hold mpv where it is, or set it going again.
pub(super) fn set_paused(pid: u32, paused: bool) -> bool {
    write(
        pid,
        format!(r#"{{"command":["set_property","pause",{paused}],"async":true}}"#) + "\n",
    )
}

/// Walk mpv's own subtitle tracks, off and round again — mpv's own word for the `T` key's step.
pub(super) fn cycle_subtitles(pid: u32) -> bool {
    write(pid, r#"{"command":["cycle","sub"],"async":true}"#.to_string() + "\n")
}

/// The JSON lines this app would have written, in order, for a test build.
///
/// The pipe is the one seam under `#[cfg(test)]` in the same spirit as `spawn_player`'s noted
/// player: a test run opens no pipe and starts no player, and what a command would have said is
/// noted instead (see `video_playback::spawn_player`).
#[cfg(test)]
static NOTED: Mutex<Vec<String>> = Mutex::new(Vec::new());

/// The lines noted so far (see `NOTED`).
#[cfg(test)]
pub(super) fn noted_lines() -> Vec<String> {
    NOTED.lock().map(|noted| noted.clone()).unwrap_or_default()
}

/// Forget the lines noted so far, so a test reads its own commands.
#[cfg(test)]
pub(super) fn clear_noted() {
    if let Ok(mut noted) = NOTED.lock() {
        noted.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The name a launch names is one apart from the last, so no two previews can share a pipe —
    /// the last thing a write-only client can afford is a command reaching the wrong player.
    #[test]
    fn the_pipe_a_launch_names_is_never_one_another_launch_named() {
        assert_ne!(
            next_pipe_name(),
            next_pipe_name(),
            "two launches must never be told the same pipe, or a command aimed at one player \
             reaches whichever process happened to open that name first"
        );
    }

    /// A command is written for the player this app holds and for no other, which is what the
    /// pid guard is for: the pipe outlives a kill by nothing, so a stale pid names no pipe.
    #[test]
    fn a_command_reaches_the_player_it_names_and_no_other() {
        clear_noted();
        remember_pipe(41, "a-pipe");

        assert!(
            set_volume(41, 40),
            "the player this app holds is the one a command goes to"
        );
        assert_eq!(
            noted_lines(),
            vec![format!(
                "{}\n",
                r#"{"command":["set_property","volume",40],"async":true}"#
            )],
            "the volume is set on the player's own property, and the line wants no answer"
        );

        assert!(
            !set_volume(43, 40),
            "a pid this app holds no pipe for is nobody's player, so nothing is written"
        );
        assert_eq!(
            noted_lines().len(),
            1,
            "and nothing is written for it"
        );

        forget_pipe(41);
        forget_pipe(43);
        clear_noted();
    }

    /// Every command the pin sends is a JSON command line with the one property or the one seek
    /// the manual names, and each says it wants no reply.
    #[test]
    fn each_command_is_the_line_mpv_is_told_with() {
        clear_noted();
        remember_pipe(7, "a-pipe");

        assert!(seek_absolute(7, 90.0));
        assert!(set_paused(7, true));
        assert!(set_paused(7, false));
        assert!(cycle_subtitles(7));

        assert_eq!(
            noted_lines(),
            vec![
                format!("{}\n", r#"{"command":["seek",90.000,"absolute"],"async":true}"#),
                format!("{}\n", r#"{"command":["set_property","pause",true],"async":true}"#),
                format!("{}\n", r#"{"command":["set_property","pause",false],"async":true}"#),
                format!("{}\n", r#"{"command":["cycle","sub"],"async":true}"#),
            ],
            "each line names the command and its argument, and every one is async"
        );

        forget_pipe(7);
        clear_noted();
    }

    /// The pipe, end to end: one command line handed to a real mpv down its own pipe, and the player
    /// doing what the line says.
    ///
    /// **Ignored, and it has to be.** It starts a real mpv — found the way the app finds it, in the
    /// installs this machine has, by the app's own list and its own rule for choosing between them —
    /// and it needs one to be installed at all, while the suite proper must start no player and open
    /// no pipe (see `write_to_a_real_pipe`). A machine with no mpv, or one where mpv will not start,
    /// is answered with a printed reason and no failure: this machine cannot ask is not the same
    /// answer as the line is wrong.
    ///
    /// **What a run of it proves is the half nothing else in this module can.** Every other test here
    /// asserts the *string* this app meant to send, and a string is not what a player reads: the
    /// delimiter the JSON IPC reads a message up to was missing from every command line this module
    /// wrote, and every test in it stayed green, because the string they asserted was the very string
    /// written here. This one asserts what the bytes do to a real player, and `quit` is the one
    /// command whose answer is visible from outside it — no film, no ffmpeg, no window and no reply
    /// to read: mpv is gone, or it is not.
    ///
    /// It writes for real, so run it alone, the way the ignore asks for — `cargo test
    /// a_real_mpv_quits_on_the_line_this_app_writes -- --ignored` — and not beside the tests that
    /// read the noted lines. Nor beside the other test here that writes for real: the flag that
    /// turns the writing on and the pipe this app holds are one pair for the whole process, so the
    /// two of them run one after the other rather than together (`cargo test a_real_mpv --
    /// --ignored --test-threads=1`).
    #[test]
    #[ignore = "starts a real mpv, which a test run must never do; needs one installed"]
    fn a_real_mpv_quits_on_the_line_this_app_writes() {
        use crate::app::engine_processes;
        use crate::formats::codecs;
        use std::process::Stdio;
        use std::time::{Duration, Instant};

        let Some(program) = codecs::resolve_mpv(None, &codecs::mpv_installs()) else {
            println!("this machine has no mpv, so there is no pipe to hand a line to");
            return;
        };

        // The app's own flag and the app's own name, one apart from the last launch's.
        let name = next_pipe_name();
        let ipc = format!(r#"--input-ipc-server=\\.\pipe\{name}"#);

        // Headless and idle: no window to open, no sound to open, no configuration of the user's to
        // read (the app's own launches pass `--no-config` for the same reason), and no console of
        // its own. What is left is a player waiting for a line.
        let started = engine_processes::hidden_command(&program)
            .args([
                "--no-config",
                "--idle=yes",
                "--vo=null",
                "--ao=null",
                ipc.as_str(),
            ])
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn();

        let mut player = match started {
            Ok(player) => player,
            Err(error) => {
                println!(
                    "mpv is installed but could not be started ({}), so there is no pipe to hand \
                     a line to",
                    error.kind()
                );
                return;
            }
        };

        remember_pipe(player.id(), &name);
        write_to_a_real_pipe(true);

        // The player puts its end of the pipe up while it starts, and a line written before it is
        // there is a line no player received — so the handover is tried until it lands.
        let quit = r#"{"command":["quit"]}"#.to_string() + "\n";
        let patience = Instant::now() + Duration::from_secs(5);
        let look = Duration::from_millis(25);
        let mut handed = false;
        while !handed && Instant::now() < patience {
            handed = write(player.id(), quit.clone());
            if !handed {
                std::thread::sleep(look);
            }
        }

        // **The assertion this test exists for.** A line mpv cannot read is one it never acts on,
        // and the player is then still up when it is asked about — which is exactly what the missing
        // newline did, and what no assertion about the note could ever have shown.
        let gone_by = Instant::now() + Duration::from_secs(5);
        let mut gone = false;
        while !gone && Instant::now() < gone_by {
            gone = matches!(player.try_wait(), Ok(Some(_)));
            if !gone {
                std::thread::sleep(look);
            }
        }

        if !gone {
            let _ = player.kill();
        }
        let _ = player.wait();
        forget_pipe(player.id());
        write_to_a_real_pipe(false);

        assert!(
            handed,
            "the line has to reach the player's pipe: a pipe that cannot be opened is not a player \
             this app can tell anything, and the one just started is up"
        );
        assert!(
            gone,
            "mpv has to quit on the line this app writes it, and a player still up five seconds \
             later is a line mpv never read — which is what a command line without the newline the \
             JSON IPC delimits on is, however exactly the string of it is asserted"
        );
    }

    /// **Two lines written straight after one another both reach the player.**
    ///
    /// The tick writes a pair — the rewind that takes a film back to its first frame and the line
    /// that settles which state it is left in, played or held (see `video_loop_tick`) — and every
    /// line is a connection of its own: the pipe is opened, one line written, the handle dropped
    /// (`deliver`). The second line is the one this test is about, and it is read from outside the
    /// pipe the only way a write-only client can read anything: `quit`, whose answer is a player
    /// that is gone.
    ///
    /// Run it the way the ignore asks for — `cargo test
    /// a_real_mpv_reads_two_lines_written_straight_after_one_another -- --ignored` — and not beside
    /// the tests that read the noted lines: the flag that turns the writing on and the pipe this app
    /// holds are one pair for the whole process, so this test and the one above it run one after the
    /// other rather than together (`cargo test a_real_mpv -- --ignored --test-threads=1`).
    #[test]
    #[ignore = "starts a real mpv, which a test run must never do; needs one installed"]
    fn a_real_mpv_reads_two_lines_written_straight_after_one_another() {
        use crate::app::engine_processes;
        use crate::formats::codecs;
        use std::process::Stdio;
        use std::time::{Duration, Instant};

        let Some(program) = codecs::resolve_mpv(None, &codecs::mpv_installs()) else {
            println!("this machine has no mpv, so there is no pipe to hand two lines to");
            return;
        };

        let name = next_pipe_name();
        let ipc = format!(r#"--input-ipc-server=\\.\pipe\{name}"#);

        let started = engine_processes::hidden_command(&program)
            .args([
                "--no-config",
                "--idle=yes",
                "--vo=null",
                "--ao=null",
                ipc.as_str(),
            ])
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn();

        let mut player = match started {
            Ok(player) => player,
            Err(error) => {
                println!(
                    "mpv is installed but could not be started ({}), so there is no pipe to hand \
                     two lines to",
                    error.kind()
                );
                return;
            }
        };

        remember_pipe(player.id(), &name);
        write_to_a_real_pipe(true);

        // The first line is tried until it lands, the way the test above tries its own: the player
        // puts its end of the pipe up while it starts, and it changes nothing about the player.
        let first = r#"{"command":["set_property","volume",50],"async":true}"#.to_string() + "\n";
        let patience = Instant::now() + Duration::from_secs(5);
        let look = Duration::from_millis(25);
        let mut handed = false;
        while !handed && Instant::now() < patience {
            handed = write(player.id(), first.clone());
            if !handed {
                std::thread::sleep(look);
            }
        }
        assert!(
            handed,
            "the first line has to reach the player's pipe: a pipe that cannot be opened is not a \
             player this app can tell anything"
        );

        // **And the second in the same breath as the first**, which is the shape of the pair a press
        // writes: the first line's connection is closed, and the next one is opened at once.
        let second = r#"{"command":["quit"]}"#.to_string() + "\n";
        let sent = write(player.id(), second);

        let gone_by = Instant::now() + Duration::from_secs(5);
        let mut gone = false;
        while !gone && Instant::now() < gone_by {
            gone = matches!(player.try_wait(), Ok(Some(_)));
            if !gone {
                std::thread::sleep(look);
            }
        }

        if !gone {
            let _ = player.kill();
        }
        let _ = player.wait();
        forget_pipe(player.id());
        write_to_a_real_pipe(false);

        assert!(
            sent,
            "the second line has to be written as well: a pair is two commands, and one that \
             cannot be written is a command the player never reads"
        );
        assert!(
            gone,
            "mpv has to quit on the second line, and a player still up five seconds later is a \
             second line that was written and never reached it"
        );
    }
}

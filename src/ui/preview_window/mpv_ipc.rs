//! Telling a running mpv what to do, over its own JSON IPC.
//!
//! mpv is the one player here that can be spoken to while it plays, and that is the whole reason
//! the mpv route exists as something other than FFmpeg's player (see `video_route`): a volume, a
//! seek, a pause and a track step are lines written down a named pipe rather than a player ended
//! and begun again.
//!
//! **Write-only by decision.** The pipe is opened, one JSON command line is written, and the
//! handle is dropped: there is no reader thread, no request id, no reply parsed and nothing read
//! back — the app knows what it needs without asking, the sidecar from the filesystem and the
//! level from the pin. A command that wants no answer says so with `"async": true` rather than
//! leaving a reply nobody reads in the pipe.
//!
//! The pipe the launch names is remembered against the player's own process id, and forgotten when
//! that player is killed or parked, so a command can never reach a player this app has stopped
//! holding (see `kill_player_process` and `restart_pinned_player`). What is remembered beside the
//! name is what the launch itself said of its player — whether it was given any audio — because a
//! level moved on a player begun with none has a track to ask back for (see `HeldPlayer` and
//! `set_volume`). The name is unique per launch (a counter) so two previews could never share one.

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;

/// The pipe a launch names, unique to that launch and to this run of the app.
///
/// A counter rather than the process id because the name is written into the argument list before
/// there is a process to name: `mpv_args` builds the command line, and the player only exists once
/// it has been spawned. It is one apart from the last, so two previews never share a pipe. **The
/// app's own process id is what the counter alone cannot do**: a counter begins at one every run,
/// and a player left listening by a run that was killed holds the very first name the next run
/// asks for, answers its lines as though they were meant for it, and the app's `deliver` reports
/// success for a command the player on screen never heard (see `CURRENT_PIPE`, which is a single
/// slot rather than a map of them).
pub(super) fn next_pipe_name() -> String {
    static NEXT: AtomicU64 = AtomicU64::new(1);
    format!(
        "rust-hover-preview-mpv-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    )
}

/// A player this app holds while it runs: the pipe it listens on, and whether the launch gave it
/// no audio at all.
///
/// **The audio is a fact of the launch and the pipe is where it is needed.** A level of nothing is
/// begun with `--aid=no` rather than a level of nothing (see `video_launch::mpv_args`), and a level
/// moved on to such a player has no track to be heard at until one is asked for (see `set_volume`).
/// The two are one player's facts, so they are kept together and forgotten together: a kill or a
/// park drops the record, and a stale launch can never leave an answer behind for the player that
/// replaces it.
#[derive(Clone)]
struct HeldPlayer {
    pid: u32,
    /// The pipe name, without the `\\.\pipe\` the delivery adds.
    pipe: String,
    /// Whether the player has no audio track selected, which is what the first level above nothing
    /// asks back for (see `set_volume`).
    audio_off: bool,
}

/// The player behind a pid this app holds while it runs.
static CURRENT_PIPE: Mutex<Option<HeldPlayer>> = Mutex::new(None);

/// Remember the pipe a launch named for the player it has just begun, and whether that launch gave
/// the player any audio at all (see `HeldPlayer`).
pub(super) fn remember_pipe(pid: u32, name: &str, audio_off: bool) {
    if let Ok(mut slot) = CURRENT_PIPE.lock() {
        *slot = Some(HeldPlayer {
            pid,
            pipe: name.to_string(),
            audio_off,
        });
    }
}

/// Forget the pipe of `pid`, where it is that player's — a kill drops it too.
pub(super) fn forget_pipe(pid: u32) {
    if let Ok(mut slot) = CURRENT_PIPE.lock() {
        if slot.as_ref().is_some_and(|held| held.pid == pid) {
            *slot = None;
        }
    }
}

/// The player this app holds behind `pid`, or nothing where `pid` names no such player.
fn held_player(pid: u32) -> Option<HeldPlayer> {
    CURRENT_PIPE
        .lock()
        .ok()
        .and_then(|slot| slot.as_ref().filter(|held| held.pid == pid).cloned())
}

/// Where a command actually goes: the pipe of the player behind `pid`, or nothing where there is
/// no such player.
fn write(pid: u32, line: String) -> bool {
    let Some(held) = held_player(pid) else {
        return false;
    };
    deliver(&held.pipe, line)
}

/// How many times a line is offered to the pipe, and how long the wait between two offers is.
///
/// mpv serves one client at a time and frees the next instance of its pipe a moment after the
/// client before it has been let go of, so two lines written in one breath — the pair the tick's
/// rewind is (see `video_loop_tick`) — can have the *second* refused with `ERROR_PIPE_BUSY`.
/// Measured against mpv 0.41.0 with a film playing: 60 such pairs written back to back the way
/// `deliver` writes them had the second line refused 3 times, and none at all with a millisecond
/// between the two. A few milliseconds of patience is many times the moment it takes, and only a
/// refused line ever pays it.
const PIPE_OFFERS: usize = 6;
const PIPE_OFFER_WAIT: std::time::Duration = std::time::Duration::from_millis(2);

/// Windows' own `ERROR_PIPE_BUSY`, which is what an open that finds every instance of the pipe
/// taken answers — the one failure worth waiting for rather than reporting.
const ERROR_PIPE_BUSY: i32 = 231;

/// Write one command line down the pipe and close it. MPV holds the other end; a pipe that cannot
/// be opened is a player that is not there, and a command for one is not an error.
///
/// **And a pipe with every instance taken is tried again rather than answered `false`**, because
/// that is a line the player is ready to hear: what a line refused there costs is not the waiting
/// but the command — the rewind's settle is the half that holds a film at its top, and one lost to
/// the pipe is a film that plays on from the first frame instead (see `video_loop_tick`).
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
    let mut offers = PIPE_OFFERS;
    loop {
        offers -= 1;
        match std::fs::OpenOptions::new().write(true).open(&path) {
            Ok(mut pipe) => return pipe.write_all(line.as_bytes()).is_ok(),
            Err(error) if offers > 0 && error.raw_os_error() == Some(ERROR_PIPE_BUSY) => {
                std::thread::sleep(PIPE_OFFER_WAIT);
            }
            Err(_) => return false,
        }
    }
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
///
/// **A level above nothing is a track asked back for where the launch turned the audio off.** mpv
/// is begun with `--aid=no` rather than `--volume=0` at a level of nothing (see
/// `video_launch::mpv_args`), and such a player has no track for a level to be heard at: the level
/// alone moved a property nothing played at, which is why a knob dragged up on a film the tray had
/// at 0% went on being silent. The track is asked for with the first level above nothing — `aid`
/// back to `auto`, the selection a launch with audio leaves the property at, since the app names no
/// track of its own — and with the *level already in hand*, so a track asked for first cannot open
/// the audio at mpv's own default of full for the moment the level line takes to follow.
///
/// The ask leaves a player that has its track, and that is remembered as what it is: the next step
/// of a drag, and a release's re-assertion, are the level alone (see `HeldPlayer`). A level of
/// nothing asks for nothing, since a film a hand has taken back down to silence wants the property
/// it already has and not a decoder turned off behind it.
pub(super) fn set_volume(pid: u32, level: u32) -> bool {
    let Some(held) = held_player(pid) else {
        return false;
    };

    // mpv's JSON IPC reads a message up to its newline, and a line with none is a command it
    // never sees — closing the pipe does not deliver it either — so the newline is part of
    // the line rather than of the write (see `write`).
    if !write(
        pid,
        format!(r#"{{"command":["set_property","volume",{level}],"async":true}}"#) + "\n",
    ) {
        return false;
    }

    if level > 0 && held.audio_off {
        if !write(
            pid,
            r#"{"command":["set_property","aid","auto"],"async":true}"#.to_string() + "\n",
        ) {
            return false;
        }
        // The ask landed, so the player has its track: the fact moves with it, which is what keeps
        // a drag from asking again on its next step.
        if let Ok(mut slot) = CURRENT_PIPE.lock() {
            if let Some(held) = slot.as_mut().filter(|held| held.pid == pid) {
                held.audio_off = false;
            }
        }
    }

    true
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

/// Hold mpv's own window above every other window, or let it go back into the ordinary band.
///
/// **It is the pin's `Pin Mode → Always On Top` asked of a player that was begun before the pin
/// was.** A film is launched `--ontop` for the preview's sake — a hover is a window nothing may
/// cover, and the band the pin draws over the film is this app's own pixels — and a *pin* that is
/// not Always On Top is the opposite: a window the user may put a program in front of. A film left
/// topmost there stands over that program and takes the clicks in the band they overlap, which is
/// the setting doing the opposite of what it says (see `video_launch::mpv_args`).
///
/// A film adopted by a take-up and a setting changed under a running film are both this same ask:
/// the player is told what the pin wants rather than begun again, which is the whole of what makes
/// this route not FFmpeg's (see `settle_pinned_mpv_ontop`).
pub(super) fn set_ontop(pid: u32, ontop: bool) -> bool {
    write(
        pid,
        format!(r#"{{"command":["set_property","ontop",{ontop}],"async":true}}"#) + "\n",
    )
}

/// Walk mpv's own subtitle tracks, off and round again — mpv's own word for the `T` key's step.
pub(super) fn cycle_subtitles(pid: u32) -> bool {
    write(
        pid,
        r#"{"command":["cycle","sub"],"async":true}"#.to_string() + "\n",
    )
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
    use super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME;
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
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());

        clear_noted();
        remember_pipe(41, "a-pipe", false);

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
            "the volume is set on the player's own property, and the line wants no answer — one \
             line, because a player a level above nothing was begun at has its audio already"
        );

        assert!(
            !set_volume(43, 40),
            "a pid this app holds no pipe for is nobody's player, so nothing is written"
        );
        assert_eq!(noted_lines().len(), 1, "and nothing is written for it");

        forget_pipe(41);
        forget_pipe(43);
        clear_noted();
    }

    /// **A level moved above nothing on a player the launch gave no audio asks the track back**,
    /// once per player and after the level rather than before it.
    ///
    /// A launch at a level of nothing is given `--aid=no` rather than `--volume=0` (see
    /// `video_launch::mpv_args`), so the property a level is set on is one nothing plays at: a
    /// knob dragged up on such a film moved the level and stayed silent, which is the fault this
    /// pair of lines is for. The order is the point of the second line being second — a track
    /// opened before the level lands opens the audio at mpv's own default of full — and the record
    /// moving with the ask is the point of the pair being written once and not once per pixel.
    #[test]
    fn a_level_moved_above_nothing_on_a_player_begun_without_audio_asks_for_the_track_once() {
        // The noted log is one the pin's own tests write in as well, and the three tests here that
        // read it are one at a time only because they hold the pin's own lock: without it a
        // sibling's clear lands between the two lines this test is counting.
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());

        clear_noted();
        // A player of the kind a level of nothing is begun as.
        remember_pipe(31, "a-muted-launch-pipe", true);

        // A level of nothing asks for nothing: the property already says silence, and a film the
        // knob has not been taken up on is left as the launch left it.
        assert!(set_volume(31, 0));
        assert_eq!(
            noted_lines(),
            vec![format!(
                "{}\n",
                r#"{"command":["set_property","volume",0],"async":true}"#
            )],
            "the level alone, because nothing is heard at it and no track is asked back"
        );
        clear_noted();

        // The first level above nothing moves the property and then asks for the track, in that
        // order, so that the audio the track opens at is the level just named.
        assert!(set_volume(31, 40));
        assert_eq!(
            noted_lines(),
            vec![
                format!(
                    "{}\n",
                    r#"{"command":["set_property","volume",40],"async":true}"#
                ),
                format!(
                    "{}\n",
                    r#"{"command":["set_property","aid","auto"],"async":true}"#
                ),
            ],
            "the level first, and then `aid` set back to `auto` — the selection a launch with \
             audio leaves the property at"
        );

        // And the player has a track from here: the rest of a drag is the level alone.
        assert!(set_volume(31, 55));
        assert_eq!(
            noted_lines().len(),
            3,
            "the track is asked for once per player, and not once per step of a drag"
        );

        forget_pipe(31);
        clear_noted();
    }

    /// One command of every kind is a JSON command line with the one property, the one seek or the
    /// one flag the manual names, and each says it wants no reply.
    #[test]
    fn each_command_is_the_line_mpv_is_told_with() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());

        clear_noted();
        remember_pipe(7, "a-pipe", false);

        assert!(seek_absolute(7, 90.0));
        assert!(set_paused(7, true));
        assert!(set_paused(7, false));
        assert!(set_ontop(7, false));
        assert!(cycle_subtitles(7));

        assert_eq!(
            noted_lines(),
            vec![
                format!(
                    "{}\n",
                    r#"{"command":["seek",90.000,"absolute"],"async":true}"#
                ),
                format!(
                    "{}\n",
                    r#"{"command":["set_property","pause",true],"async":true}"#
                ),
                format!(
                    "{}\n",
                    r#"{"command":["set_property","pause",false],"async":true}"#
                ),
                format!(
                    "{}\n",
                    r#"{"command":["set_property","ontop",false],"async":true}"#
                ),
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
    /// read the noted lines.
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

        remember_pipe(player.id(), &name, false);
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

    /// **Both lines of a pair written in one breath reach the player, over and over.**
    ///
    /// The tick's rewind is such a pair — a seek, and the line that settles the state the film is
    /// left in (see `video_loop_tick`) — and mpv serves one client at a time, freeing the next
    /// instance of its pipe only once the client before it has been let go of. Measured on mpv
    /// 0.41.0 with this exact write, 60 pairs back to back had the *second* line refused with
    /// `ERROR_PIPE_BUSY` six times. A refused line is a command the player never reads, and for the
    /// rewind's own pair the half that is lost is the hold: the film is taken to its top and then
    /// plays on from there where it was meant to stand, pass after pass — which is the fault this
    /// test is here to keep out (see `deliver`).
    ///
    /// Run it the way the ignore asks for — `cargo test
    /// a_real_mpv_takes_every_line_of_a_pair -- --ignored` — and one at a time beside the other test
    /// here that writes for real, for the reason its own note gives.
    #[test]
    #[ignore = "starts a real mpv, which a test run must never do; needs one installed"]
    fn a_real_mpv_takes_every_line_of_a_pair_written_back_to_back() {
        use crate::app::engine_processes;
        use crate::formats::codecs;
        use std::process::Stdio;
        use std::time::{Duration, Instant};

        let Some(program) = codecs::resolve_mpv(None, &codecs::mpv_installs()) else {
            println!("this machine has no mpv, so there is no pipe to hand a pair to");
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
                     a pair to",
                    error.kind()
                );
                return;
            }
        };

        remember_pipe(player.id(), &name, false);
        write_to_a_real_pipe(true);

        // The player puts its end of the pipe up while it starts, so the first line is tried until
        // it lands. An idle player has no film for a seek to move and answers the command with an
        // error nobody reads, which is the point: what is under test is the pipe under a pair, not
        // what mpv does with either half.
        let patience = Instant::now() + Duration::from_secs(5);
        let look = Duration::from_millis(25);
        let mut handed = false;
        while !handed && Instant::now() < patience {
            handed = seek_absolute(player.id(), 0.0);
            if !handed {
                std::thread::sleep(look);
            }
        }
        assert!(
            handed,
            "the first line has to reach the player's pipe: a pipe that cannot be opened is not a \
             player this app can tell anything"
        );

        let mut refused = 0;
        for _ in 0..60 {
            if !seek_absolute(player.id(), 0.0) || !set_paused(player.id(), true) {
                refused += 1;
            }
        }

        let _ = player.kill();
        let _ = player.wait();
        forget_pipe(player.id());
        write_to_a_real_pipe(false);

        assert_eq!(
            refused, 0,
            "every line of every pair has to be accepted: a pipe with both hands full is waited \
             out, because a line answered false is a command the film on screen never hears"
        );
    }

    /// **A real player begun the way a film at 0% is begun has its track selected back by the pair,
    /// and is playing at the level the knob is at.**
    ///
    /// This is the assertion the fault itself asks for and the one nothing else in the module can
    /// make. Every other test here asserts the *string* this app means to write — that the pair is
    /// the level and then `aid`, in that order, once per player — and the string being exactly right
    /// is not the same as a player doing what it says: a level and a track asked for on a player
    /// given no audio is the whole of what a knob dragged up from 0% does (see `set_volume`).
    ///
    /// **A film with a picture *and* a soundtrack out of two files the test writes itself.** A file
    /// with no track left at all is one mpv quits on — measured: a WAV with `--aid=no` is exit code
    /// 2 before its pipe is even there — and the app's own films always have a picture to keep the
    /// player up. So the test writes a BMP (read as a still picture) and a second of PCM WAV, and
    /// hands mpv the WAV as the picture's own audio track through its `--audio-file`: two files of a
    /// few hundred bytes rather than a muxer or a binary fixture. `--image-display-duration=inf`
    /// keeps the still from reaching its end while the pair is delivered.
    ///
    /// **The reading is what makes it fiddly, and it is a fact of the test rather than of the
    /// app.** mpv serves one client at a time, so a connection held open while `set_volume` writes
    /// has the app's own `deliver` retrying and giving up (see `deliver`): every ask here opens the
    /// pipe, asks its one property, reads the one answer and lets go again, and the pair is written
    /// with no client on the pipe at all. The player is otherwise the app's own — `video_launch::mpv_args`
    /// builds its command line, `--aid=no` among the arguments — with the two switches a machine
    /// with nobody in front of it needs: no window, no sound device.
    ///
    /// Run it the way the ignore asks for — `cargo test
    /// a_real_mpv_begun_without_audio_takes_the_track_back -- --ignored` — and one at a time beside
    /// the other tests here that write for real, for the reason their own notes give.
    #[test]
    #[ignore = "starts a real mpv, which a test run must never do; needs one installed"]
    fn a_real_mpv_begun_without_audio_takes_the_track_back() {
        use super::super::video_launch;
        use crate::app::engine_processes;
        use crate::formats::codecs;
        use std::process::Stdio;
        use std::time::{Duration, Instant};

        let Some(program) = codecs::resolve_mpv(None, &codecs::mpv_installs()) else {
            println!("this machine has no mpv, so there is no player to ask for the track back");
            return;
        };

        let folder = std::env::temp_dir().join("rust-hover-preview-mpv-aid-test");
        let picture = folder.join("a-picture.bmp");
        let soundtrack = folder.join("a-second-of-silence.wav");
        if std::fs::create_dir_all(&folder).is_err()
            || write_a_picture(&picture).is_err()
            || write_a_second_of_silence(&soundtrack).is_err()
        {
            println!(
                "the temp folder or the test's two files cannot be written, so there is no \
                      film to begin"
            );
            let _ = std::fs::remove_dir_all(&folder);
            return;
        }

        // The app's own launch for a film at 0%: the level of nothing is `--aid=no` rather than a
        // level of zero (see `video_launch::mpv_args`), and the name is the app's own pipe.
        let name = next_pipe_name();
        let args = video_launch::mpv_args(
            &picture, 0, 0, 320, 240, 0.0, 0, None, false, false, None, &name, None,
        );
        let audio_file = format!("--audio-file={}", soundtrack.to_string_lossy());
        let started = engine_processes::hidden_command(&program)
            .args(&args)
            .args([
                audio_file.as_str(),
                "--image-display-duration=inf",
                "--vo=null",
                "--ao=null",
            ])
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn();

        let mut player = match started {
            Ok(player) => player,
            Err(error) => {
                println!(
                    "mpv is installed but could not be started ({}), so there is no player to ask \
                     for the track back",
                    error.kind()
                );
                let _ = std::fs::remove_dir_all(&folder);
                return;
            }
        };

        let look = Duration::from_millis(25);
        remember_pipe(
            player.id(),
            &name,
            video_launch::mpv_launched_without_audio(0),
        );
        write_to_a_real_pipe(true);

        // The player has the film once it lists both of its tracks: a pipe that is not listening
        // yet answers nothing, and a film still being opened lists fewer tracks than the two it
        // was given. Asking about the audio before this is asking a player that may not be there.
        let patience = Instant::now() + Duration::from_secs(15);
        let mut tracks = String::new();
        while Instant::now() < patience {
            tracks = ask_a_real_mpv(&name, "track-list/count").unwrap_or_default();
            if data_of(&tracks).and_then(|count| count.parse::<u32>().ok()) == Some(2) {
                break;
            }
            std::thread::sleep(look);
        }
        let loaded = data_of(&tracks).and_then(|count| count.parse::<u32>().ok()) == Some(2);

        // The pair, by the production call and with the production record beside the pipe: a level
        // moved up on a player the launch gave no audio is the level and then the track asked back
        // (see `set_volume`). The first line is tried until it lands, as it is in the other tests
        // here that write for real: the pipe is freed a moment after the last client let go of it.
        let mut told = false;
        if loaded {
            let by = Instant::now() + Duration::from_secs(5);
            told = set_volume(player.id(), 40);
            while !told && Instant::now() < by {
                std::thread::sleep(look);
                told = set_volume(player.id(), 40);
            }
        }

        // **The player's own answer**, asked down the same pipe: the track it has selected, the
        // codec of that track, and the level it is playing at. A player that heard the level and
        // nothing else answers the second with an error — which is exactly the silent-film fault
        // this pair is for, and the reason the assertion is on mpv's answer rather than on the line
        // the app wrote.
        //
        // It is asked until the answers are those (see `the_level_was_taken`) and not once, because
        // a line written to a player and the player having read it are two moments: a read that
        // lands between the level and the ask for the track answers for the state before the pair,
        // and a player that never gets there fails the assertions below with what it did say.
        let by = Instant::now() + Duration::from_secs(5);
        let (mut aid, mut codec, mut volume) = (String::new(), String::new(), String::new());
        while loaded && Instant::now() < by {
            aid = ask_a_real_mpv(&name, "aid").unwrap_or_default();
            codec = ask_a_real_mpv(&name, "audio-codec-name").unwrap_or_default();
            volume = ask_a_real_mpv(&name, "volume").unwrap_or_default();
            if the_level_was_taken(&aid, &codec, &volume) {
                break;
            }
            std::thread::sleep(look);
        }

        let _ = player.kill();
        let _ = player.wait();
        forget_pipe(player.id());
        write_to_a_real_pipe(false);
        let _ = std::fs::remove_dir_all(&folder);

        assert!(
            loaded,
            "the player has to open the film the test wrote, which is a picture and a soundtrack \
             and nothing else: it listed \"{tracks}\" tracks, and a player that never had the film \
             cannot be asked what its soundtrack became"
        );
        assert!(
            told,
            "the level has to reach the player's pipe, or there is no pair under test at all"
        );
        assert!(
            data_of(&aid)
                .and_then(|id| id.parse::<i64>().ok())
                .is_some_and(|id| id > 0),
            "the track has to be selected back on the player, which is what the second half of the \
             pair is for: `aid` answered \"{aid}\" after a level above nothing on a player the launch \
             gave no audio at all"
        );
        assert!(
            !data_of(&codec).unwrap_or_default().is_empty(),
            "and the track has to be one the player can decode, which is what a level is heard \
             through: `audio-codec-name` answered \"{codec}\""
        );
        assert!(
            data_of(&volume)
                .and_then(|level| level.parse::<f64>().ok())
                .is_some_and(|level| (level - 40.0).abs() < 0.01),
            "and the level has to be the one the knob was moved to: `volume` answered \"{volume}\""
        );
    }

    /// A 16x16 picture in the one shape every player on the machine reads without a library: an
    /// uncompressed BMP. What the test above wants of it is a track that keeps a player up while its
    /// soundtrack is asked back — the picture itself is never drawn, since the launch is headless.
    fn write_a_picture(path: &std::path::Path) -> std::io::Result<()> {
        let side = 16u32;
        let pixels = side * side * 3;
        let mut bmp = Vec::with_capacity(54 + pixels as usize);
        bmp.extend_from_slice(b"BM");
        bmp.extend_from_slice(&(54 + pixels).to_le_bytes());
        bmp.extend_from_slice(&0u16.to_le_bytes());
        bmp.extend_from_slice(&0u16.to_le_bytes());
        bmp.extend_from_slice(&54u32.to_le_bytes());
        bmp.extend_from_slice(&40u32.to_le_bytes());
        bmp.extend_from_slice(&(side as i32).to_le_bytes());
        bmp.extend_from_slice(&(side as i32).to_le_bytes());
        bmp.extend_from_slice(&1u16.to_le_bytes());
        bmp.extend_from_slice(&24u16.to_le_bytes());
        bmp.extend_from_slice(&0u32.to_le_bytes());
        bmp.extend_from_slice(&pixels.to_le_bytes());
        bmp.extend_from_slice(&2835i32.to_le_bytes());
        bmp.extend_from_slice(&2835i32.to_le_bytes());
        bmp.extend_from_slice(&0u32.to_le_bytes());
        bmp.extend_from_slice(&0u32.to_le_bytes());
        bmp.resize(54 + pixels as usize, 0x60);
        std::fs::write(path, bmp)
    }

    /// A second of silence in the one shape every player on the machine reads: PCM WAV, one
    /// channel, 8000 samples a second, sixteen bits a sample. The smallest soundtrack mpv opens,
    /// which is all the test above wants of it — the track is the subject and the sound is not.
    fn write_a_second_of_silence(path: &std::path::Path) -> std::io::Result<()> {
        let samples = 8000u32;
        let data = samples * 2;
        let mut wav = Vec::with_capacity(44 + data as usize);
        wav.extend_from_slice(b"RIFF");
        wav.extend_from_slice(&(36 + data).to_le_bytes());
        wav.extend_from_slice(b"WAVEfmt ");
        wav.extend_from_slice(&16u32.to_le_bytes());
        wav.extend_from_slice(&1u16.to_le_bytes());
        wav.extend_from_slice(&1u16.to_le_bytes());
        wav.extend_from_slice(&8000u32.to_le_bytes());
        wav.extend_from_slice(&16000u32.to_le_bytes());
        wav.extend_from_slice(&2u16.to_le_bytes());
        wav.extend_from_slice(&16u16.to_le_bytes());
        wav.extend_from_slice(b"data");
        wav.extend_from_slice(&data.to_le_bytes());
        wav.resize(44 + data as usize, 0);
        std::fs::write(path, wav)
    }

    /// Whether a player's own answers say the level was taken: a track selected, a codec for that
    /// track, and the level the knob was moved to. It is the rule for when to stop *asking* in the
    /// test above; the assertions after it are what decide what the answers have to be.
    fn the_level_was_taken(aid: &str, codec: &str, volume: &str) -> bool {
        data_of(aid)
            .and_then(|id| id.parse::<i64>().ok())
            .is_some_and(|id| id > 0)
            && !data_of(codec).unwrap_or_default().is_empty()
            && data_of(volume)
                .and_then(|level| level.parse::<f64>().ok())
                .is_some_and(|level| (level - 40.0).abs() < 0.01)
    }

    /// One answer from a real player: open its pipe, ask the one property, read the line it answers
    /// with, and let go of the pipe again. **The letting go is the fiddly part** — mpv serves one
    /// client at a time, so a connection held open while `set_volume` writes makes its `deliver`
    /// retry and give up (see `deliver`) — so the connection lives inside the thread and is closed
    /// when the thread's own work is done, before the answer is handed back. The opening is patient
    /// for the same reason `deliver`'s is: the instance the last client held is freed a moment after
    /// the client let go of it, and an ask that lands in that moment is one nothing was asked by.
    ///
    /// The wait is bounded and a silent player is answered with nothing, because a test that hangs
    /// on a pipe no player ever writes to cannot say which of the two is at fault. `None` therefore
    /// means no answer at all — a pipe that could not be opened, a request mpv refused, or a player
    /// that wrote nothing in the time given — and every caller reads it as a fact of that kind
    /// rather than as a failure.
    fn ask_a_real_mpv(name: &str, property: &str) -> Option<String> {
        use std::io::{BufRead, Write};

        let path = format!(r"\\.\pipe\{name}");
        // The line is built here rather than inside the thread, so what the thread owns is its bytes
        // and its pipe, and nothing borrowed has to outlive the call.
        let mut line = format!(r#"{{"command":["get_property","{property}"],"request_id":1}}"#);
        line.push('\n');
        let (send, receive) = std::sync::mpsc::channel();
        std::thread::spawn(move || {
            let answer = (|| -> Option<String> {
                let mut offers = PIPE_OFFERS;
                let mut pipe = loop {
                    offers -= 1;
                    match std::fs::OpenOptions::new()
                        .read(true)
                        .write(true)
                        .open(&path)
                    {
                        Ok(pipe) => break pipe,
                        Err(error)
                            if offers > 0 && error.raw_os_error() == Some(ERROR_PIPE_BUSY) =>
                        {
                            std::thread::sleep(PIPE_OFFER_WAIT);
                        }
                        Err(_) => return None,
                    }
                };
                pipe.write_all(line.as_bytes()).ok()?;

                // A client is sent the property-change events of the properties it observed and
                // nothing else, so the first line that comes back is the reply — and the request id
                // is what makes that a fact read off the line rather than a hope about the order.
                let mut reader = std::io::BufReader::new(pipe);
                let mut answer_line = String::new();
                for _ in 0..4 {
                    answer_line.clear();
                    if reader.read_line(&mut answer_line).ok()? == 0 {
                        return None;
                    }
                    if answer_line.contains("\"request_id\":1") {
                        return Some(answer_line);
                    }
                }
                None
            })();
            let _ = send.send(answer);
        });

        receive
            .recv_timeout(std::time::Duration::from_secs(10))
            .ok()
            .flatten()
    }

    /// What the `data` field of an answer says, as the text it says it with — a number left as its
    /// digits, a string without its quotes — and nothing at all for an answer that carries no
    /// `data`, which is how mpv refuses a request.
    fn data_of(answer: &str) -> Option<String> {
        let data = answer.split("\"data\":").nth(1)?.trim_start();
        match data.strip_prefix('"') {
            Some(text) => Some(text.split('"').next()?.to_string()),
            None => Some(data.split([',', '}']).next()?.trim().to_string()),
        }
    }
}

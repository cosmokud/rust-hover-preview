//! A pinned film that was begun at a level and is playing with no sound: the ladder that tries to
//! give it one, and the one player it is ever begun again for.
//!
//! **The two roads to a silent film are one page over** (see `playback_trace`): a launch that handed
//! its player a level and a player that never opened a device from it look identical from this side
//! of the process boundary, and the difference is only ever read off the player — `aid` for a track,
//! `current-ao` and `audio-out-params` for an output, and `audio-device` for the one that was asked
//! for. So this module asks the same four questions
//! the trace asks, reads the same verdict out of them (`playback_trace::Verdict`), and then *does*
//! something about it, which is the whole of what the trace is not for.
//!
//! **The ladder, each rung taken at most once, and each rung entered on what the read-back said:**
//!
//! 1. `aid` is `no` or absent — a track asked back for with `set_property aid auto`, which is the
//!    launch's own road healed: a film the level handed `--aid=no` has no track selected, and no
//!    track is no session in Windows' Volume Mixer (see `video_launch::AudioForm`).
//! 2. no output opened from the track that is selected — `set_property audio-device auto`, which
//!    mpv answers by reloading its audio output, so a device that was refused while the film was
//!    starting gets one more attempt at it. **It is written on the second read of that shape, not on
//!    the first**: a driver named on `current-ao` with no `audio-out-params` beside it is also exactly
//!    what a player half a second into its start looks like (see
//!    `EngineAnswers::a_driver_with_no_params`), and an output reinitialised under a player that was
//!    about to open it is the guard's own fault rather than the film's. `aid=no`, and a driver of
//!    `null`, are answers and not a player mid-open, so each takes its rung on the first read.
//! 3. neither heard — **one** new player, with the arguments, the second and the pause the one it
//!    replaces had, and one read of that player. Then the ladder ends.
//!
//! **Nothing here fights the film.** No pause is toggled, no seek is asked, no level is moved, and a
//! film held by the handover's own pause is *begun again held* rather than resumed: the only things
//! this module writes are the two audio properties, and the only player it begins is a replacement
//! at the second the pin's clock already says (see
//! `pin_playback::begin_the_pinned_film_again_for_its_sound`).
//!
//! **FFmpeg's player has no pipe, so its ladder is one rung and it is the last one.** It can be told
//! nothing, and it reports the device it could not open on the standard error this app used to throw
//! away — so a launch it stands behind pipes that stream and drains it (see
//! `player_logs::drain_a_players_errors`), and a line matching [`AUDIO_FAILURES`] in a launch this
//! guard stands behind is one player begun again. Not two, and never for a line it did not see. A
//! launch it does not stand behind is given `null` and no thread, the way every ordinary launch is.
//!
//! **A film that was heard is not a film that will always be heard, so the guard no longer ends where
//! the ladder does.** A player whose sound came up is kept under a sparse schedule of read-backs —
//! +2 s, +4 s, +8 s, then one every ~10 s for as long as the player lives (see `WATCHES`) — because an
//! endpoint lost after a film started leaves no line anywhere else, and until now only the first
//! second of a film's life was ever sampled. A read that is heard again writes nothing at all: the
//! line is written where the verdict moved, so a healthy film's log is the one line it opened with.
//! A late failure climbs the two in-place rungs once more — `aid`, then `audio-device` — and may take
//! the relaunch **at most once in a player's whole life**, not once per episode (see `re_opened`): a
//! device that is genuinely gone must never be begun again for, every ten seconds, for ever. The
//! sampling ends with the player: a probe that cannot reach this guard's player reports nothing and
//! schedules nothing, and `a_player_ended` is the half that needs no probe to have been woken at all.
//!
//! **The loop never waits, and nothing here runs on it.** A player that has only just been begun has
//! no pipe yet, and waiting for one is a wait no tick may pay (see `mpv_ipc::get_property_within`,
//! which is bounded per question and only ever called from a thread). So a thread asks and leaves the
//! answer where the tick finds it — the shape `playback_trace::watch_the_engine` and the log drain
//! already have — and the tick is what writes, decides and asks for the player.

use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::Duration;

use super::playback_trace::note_playback_trace;
use super::playback_trace::{EngineAnswers, Verdict};
use super::video_launch::Engine;

/// How long a player is given before its sound is first read.
///
/// **No earlier, because the answer would be the guard's own invention**: a player that has only just
/// been begun has not opened its device yet, and a verdict read off that gap is rung two invented by
/// the guard itself. It is the same wait the trace gives a launch for the same reason, and a little
/// longer — the trace is describing a film, and this is deciding what to do to one.
const AUDIO_SETTLING: Duration = Duration::from_millis(500);

/// How long a player that was heard is left alone before its sound is read again: the first three
/// reads of the sparse schedule the module note describes.
///
/// **Sparse on purpose.** A film's first second is where its sound is decided, so the first three
/// watches are close enough to catch a device that came up and then fell away with the launch; after
/// that, one read every [`WATCH_AFTER`] is enough to notice an endpoint a later player took, and few
/// enough that a pin left playing overnight costs a handful of pipes rather than a busy loop.
const WATCHES: [Duration; 3] = [
    Duration::from_secs(2),
    Duration::from_secs(4),
    Duration::from_secs(8),
];

/// The wait every watch after the three above is given.
const WATCH_AFTER: Duration = Duration::from_secs(10);

/// The lines FFmpeg's player writes when it could not open a sound device, and nothing else.
///
/// **Each one is a whole lowercased fragment of a line and none of them is a wildcard**, so a warning
/// about anything else — a file that would not open, a decoder that fell back to software, a frame
/// that was dropped — is not in this table and never begins a player again. That strictness is the
/// point: the relaunch this table causes takes a picture off the screen for the moment it takes, so
/// only a line that *is* an audio device's failure may ask for one.
const AUDIO_FAILURES: [&str; 4] = [
    // `Could not open audio device: <the driver's own error>`, which is what ffplay prints when
    // SDL refused it — a device another player holds, or a driver that would not start.
    "could not open audio device",
    // SDL's audio subsystem failing before a device was ever asked for.
    "failed to initialize the sdl audio",
    // FFmpeg's own initialisation of what it does with a soundtrack it has decoded.
    "failed to initialize audio",
    // And the resampler's own spelling of the same refusal.
    "could not initialize audio",
];

/// Whether a line is one of the failures [`AUDIO_FAILURES`] names, matched without a case and
/// without a wildcard.
fn names_an_audio_failure(line: &str) -> bool {
    let line = line.to_ascii_lowercase();
    AUDIO_FAILURES.iter().any(|failure| line.contains(failure))
}

/// Which launches the guard stands behind: **a pin's own film, and one begun at a level**.
///
/// The origin is the launch's own — `start_video_playback`'s, rather than a hover's, a sound's or an
/// audio player's — and the form is the one rule both routes' arguments are built from (see
/// `video_launch::AudioForm`): a launch that gave its player no track at all has no ladder to climb,
/// because that silence is the level asked for rather than a fault.
pub(super) fn guarded(for_a_hover: bool, volume: u32) -> bool {
    !for_a_hover && !super::video_launch::mpv_launched_without_audio(volume)
}

/// What a guard has already spent of its ladder, so that no rung is taken twice (see `the_step`).
#[derive(Clone, Copy, Default, PartialEq, Eq, Debug)]
struct Spent {
    /// The one second read a player that answered nothing is given: a pipe still being made is not a
    /// film without sound, and it is not a film to begin again either.
    asked_again: bool,
    /// The second read a player that named a driver with nothing playing through it is given, which
    /// is the shape a player still opening its output has as readily as a silent film does: the
    /// output is not reloaded under a player that may be about to open it (see
    /// `EngineAnswers::a_driver_with_no_params`).
    opening: bool,
    /// `aid` asked back for.
    track: bool,
    /// `audio-device` handed `auto`.
    device: bool,
    /// The one player begun again.
    again: bool,
}

/// Every rung spent, the second look included: what a player begun in place of one whose own ladder had
/// reached its last rung is handed — **one read, and then nothing**, whichever way that read goes — so
/// that a film with no device at all is not begun again for ever (see `the_rebirth_of`).
const SPENT: Spent = Spent {
    asked_again: true,
    opening: true,
    track: true,
    device: true,
    again: true,
};

/// What a player has said, and no tick has read yet.
enum Report {
    /// The four answers an mpv gave (see `playback_trace::ask_the_engine`).
    Answers(EngineAnswers),
    /// A line FFmpeg's player wrote that is one of its audio failures (see `AUDIO_FAILURES`).
    Spoke,
}

/// One rung of the ladder, or an end of it (see `the_step`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Step {
    /// Heard, so there is nothing left to do — or a report this guard has nothing to say about.
    Nothing,
    /// Ask the player the same four questions once more, and write no property at all.
    AskAgain,
    /// `aid` back to `auto`: the launch's own road.
    AskForATrack,
    /// `audio-device` back to `auto`, which makes mpv reload its audio output.
    AskForADevice,
    /// One new player, at the end of the ladder.
    BeginAgain,
    /// Nothing left to try, and the film plays on as it is.
    GiveUp,
}

/// The one player under the guard: what it was begun at, what has been tried, and what it has said.
struct Guard {
    /// The player this guard is armed for, which is the only one its reports are taken from.
    pid: u32,
    /// The film it is showing, which is what a relaunch is asked about and what a replacement is
    /// recognised by (see `the_rebirth_of`).
    path: PathBuf,
    /// Which player it is, because the two have different ladders.
    engine: Engine,
    spend: Spent,
    /// What the thread last read, and no tick has acted on.
    report: Option<Report>,
    /// Whether the ladder has decided on its last rung and no tick has begun that player yet. A tick
    /// with a picture changing hands under it leaves this standing rather than spending the rung,
    /// which is why it is a flag and not the return value of one call (see `tick`).
    wants_again: bool,
    /// How many read-backs this player has had: what the trace lines are numbered by, so a run of
    /// them reads in the order they happened.
    reads: u8,
    /// How many of the sparse watches a player that was heard is kept under have been begun: what the
    /// wait before the next one is read from (see `the_wait_before_a_watch`).
    watches: u8,
    /// Whether the last read of this player was heard — what keeps a film that stays heard from
    /// writing a line every ten seconds for the whole of its length, and what makes a read that is no
    /// longer heard a late failure (see `a_late_failure`).
    heard: bool,
}

impl Guard {
    /// Whether this player is watched after it was heard: mpv's, and not the one a spent ladder began
    /// in place of another — that player is given its one read and nothing more, so a film whose
    /// device is genuinely gone is not climbed for again and again.
    fn keeps_watching(&self) -> bool {
        self.engine == Engine::Mpv && self.spend != SPENT
    }
}

/// **One guard, because there is one player on screen.** Every road that begins a film for a pin
/// arms this, so the newest launch is the one watched and a player that has been replaced is not
/// asked anything more.
static GUARD: Mutex<Option<Guard>> = Mutex::new(None);

/// The player whose relaunch the ladder has just spent, and the file it was showing.
///
/// It is one slot and it is consumed by the very next launch, because a player begun again is a
/// launch like any other and is armed like any other: without it, the replacement would start at the
/// top of the ladder and a film with no device at all would be begun again once per probe, for ever.
/// A record nothing consumes — a relaunch the pin declined, because it had already walked off the
/// film — is cleared by the next launch, which is the chain having ended.
static REBIRTH: Mutex<Option<PathBuf>> = Mutex::new(None);

/// Arm the guard for a player this app has just begun, where that player is one it stands behind
/// (see `guarded`).
///
/// **Once per launched player**, which is the arming being here and not in the tick: what is watched
/// is what was begun, and a player this app did not begin is nobody's to heal.
pub(super) fn arm(pid: u32, path: &Path, volume: u32, engine: Engine, for_a_hover: bool) {
    if !guarded(for_a_hover, volume) {
        return;
    }

    let spend = the_rebirth_of(path);
    let Ok(mut slot) = GUARD.lock() else {
        return;
    };
    *slot = Some(Guard {
        pid,
        path: path.to_path_buf(),
        engine,
        spend,
        report: None,
        wants_again: false,
        reads: 0,
        watches: 0,
        heard: false,
    });
    drop(slot);

    note_playback_trace!(
        "audio guard armed pid {pid} engine {} level {volume} file {} rungs {}",
        engine.word(),
        super::pin_trace::path_name(path),
        if spend == SPENT { "spent" } else { "whole" },
    );

    // A player with a pipe is asked what it did with the sound; FFmpeg's player has none, and is read
    // off its own stream instead (see `a_players_line`).
    if engine == Engine::Mpv {
        spawn_probe(pid);
    }
}

/// What a guard is handed of the ladder: every rung, unless the player it is watching was begun in
/// place of one whose own ladder had already reached the last rung.
fn the_rebirth_of(path: &Path) -> Spent {
    let recorded = REBIRTH.lock().ok().and_then(|mut slot| slot.take());
    the_spend_of(recorded.as_deref(), path)
}

/// The rule the record above is read by, on its own so that it can be answered without the record.
fn the_spend_of(recorded: Option<&Path>, path: &Path) -> Spent {
    if recorded == Some(path) {
        SPENT
    } else {
        Spent::default()
    }
}

/// Ask the player what it did with the sound, on a thread of its own.
///
/// **Never on the preview loop** (see the module note). The thread waits, asks, and leaves the answer
/// where the tick finds it (see `report`); the wait is the settling above plus the patience one
/// question is offered a pipe that is not there yet.
fn spawn_probe(pid: u32) {
    spawn_probe_after(pid, AUDIO_SETTLING);
}

/// The same probe after a wait of its own, which is where the sparse schedule a heard player is
/// watched on sleeps — on this thread, and on no tick (see `the_wait_before_a_watch`).
///
/// **A probe whose player is no longer this guard's asks nothing at all**: a thread that was asleep
/// when the film ended wakes to a guard that has been cleared (see `a_player_ended`) and reads no pipe
/// for a process that is not there. Nothing is retried, and nothing further is scheduled — the tick
/// that would have armed the next watch never sees a report to arm it from.
fn spawn_probe_after(pid: u32, wait: Duration) {
    let _ = std::thread::Builder::new()
        .name("audio-guard".to_string())
        .spawn(move || {
            std::thread::sleep(wait);
            if !still_watched(pid) {
                return;
            }
            report(
                pid,
                Report::Answers(super::playback_trace::ask_the_engine(pid)),
            );
        });
}

/// Whether the guard still stands for this player (see `spawn_probe_after`).
fn still_watched(pid: u32) -> bool {
    GUARD
        .lock()
        .map(|slot| slot.as_ref().is_some_and(|guard| guard.pid == pid))
        .unwrap_or(false)
}

/// The wait before a watched player's `watch`th read after it was heard (see `WATCHES`).
fn the_wait_before_a_watch(watch: u8) -> Duration {
    WATCHES.get(watch as usize).copied().unwrap_or(WATCH_AFTER)
}

/// Begin the next read of a player that was heard: one thread asleep for the schedule's next wait, and
/// whatever it leaves behind for the tick (see `spawn_probe_after`).
///
/// **The first of these is the whole of what a healthy film ever writes**: the schedule repeats until
/// the player ends, and a line per read would be a line every ten seconds for the length of a film.
fn arm_a_watch(guard: &mut Guard) {
    if guard.watches == 0 {
        note_playback_trace!("audio guard heard pid {} watching", guard.pid);
    }
    let wait = the_wait_before_a_watch(guard.watches);
    guard.watches = guard.watches.saturating_add(1);
    spawn_probe_after(guard.pid, wait);
}

/// Note what a player said, where the guard is that player's, and answer whether it was.
fn report(pid: u32, what: Report) -> bool {
    let Ok(mut slot) = GUARD.lock() else {
        return false;
    };
    match slot.as_mut().filter(|guard| guard.pid == pid) {
        Some(guard) => {
            guard.report = Some(what);
            true
        }
        None => false,
    }
}

/// One line FFmpeg's player wrote on its own error stream, read by the thread that drains it.
///
/// **The signature list is deliberately short** (see `AUDIO_FAILURES`): what this costs a run that
/// matches is a picture taken off the pin for the moment a player takes to open the file, so only a
/// line that is a device's failure may ask for one.
pub(super) fn a_players_line(pid: u32, line: &str) {
    if !names_an_audio_failure(line) {
        return;
    }
    if report(pid, Report::Spoke) {
        note_playback_trace!("audio guard signature pid {pid} line {line}");
    }
}

/// A player this app was watching has ended, which is the end of any guard on it: the stream that
/// would report for it is closed, and a guard left standing would be waiting on a process that is
/// not there (see `player_logs`, whose threads end where the player does).
pub(super) fn a_player_ended(pid: u32) {
    if let Ok(mut slot) = GUARD.lock() {
        if slot.as_ref().is_some_and(|guard| guard.pid == pid) {
            *slot = None;
        }
    }
}

/// Take one tick's worth of the guard's business, and answer the film whose player the ladder has
/// decided to begin again — once, on the tick that may begin it.
///
/// `may_begin_again` is the loop's own answer to whether a player can be begun this tick at all: a
/// picture changing hands is the pin's and is not a wait to launch under (see
/// `pin_playback::begin_the_pinned_film_again_for_its_sound`). A rung decided on a tick that may not
/// take it is held, not spent.
pub(super) fn tick(may_begin_again: bool) -> Option<PathBuf> {
    let Ok(mut slot) = GUARD.lock() else {
        return None;
    };
    let guard = slot.as_mut()?;

    // **Nothing is read while a relaunch is waiting for its tick.** A report that arrives in that
    // gap — FFmpeg's player says a thing more than once — must not be read as a rung of its own, or
    // the last rung would be given up on before it was ever taken.
    let mut finish = None;
    if !guard.wants_again {
        if let Some(report) = guard.report.take() {
            guard.reads = guard.reads.saturating_add(1);
            match the_step_of_a_read(guard, report) {
                // **Heard is no longer where a player is let go from.** A film that has its sound is
                // watched for as long as it plays, so that one which loses it later is not silent and
                // unremarked; the player a spent ladder began is the exception, and is given its one
                // read and nothing else (see `keeps_watching`).
                Step::Nothing if guard.keeps_watching() => arm_a_watch(guard),
                Step::Nothing => finish = Some("heard"),
                Step::GiveUp => finish = Some("gave up"),
                Step::AskAgain => {
                    guard.spend.asked_again = true;
                    note_playback_trace!("audio step ask-again pid {}", guard.pid);
                    spawn_probe(guard.pid);
                }
                Step::AskForATrack => {
                    guard.spend.track = true;
                    let landed = super::mpv_ipc::ask_for_a_track(guard.pid);
                    note_playback_trace!(
                        "audio step ask-for-a-track pid {} landed {}",
                        guard.pid,
                        landed as u8
                    );
                    spawn_probe(guard.pid);
                }
                Step::AskForADevice => {
                    guard.spend.device = true;
                    let landed = super::mpv_ipc::ask_for_an_audio_device(guard.pid);
                    note_playback_trace!(
                        "audio step ask-for-a-device pid {} landed {}",
                        guard.pid,
                        landed as u8
                    );
                    spawn_probe(guard.pid);
                }
                Step::BeginAgain => {
                    guard.spend.again = true;
                    guard.wants_again = true;
                }
            }
        }
    }

    let pid = guard.pid;
    let relaunch = (guard.wants_again && may_begin_again).then(|| guard.path.clone());
    if let Some(word) = finish {
        note_playback_trace!("audio guard {word} pid {pid}");
    }

    match relaunch {
        Some(path) => {
            // The player begun in its place is watched at the end of the ladder and not from its top,
            // so that the one relaunch this guard is allowed is the only one the film gets (see
            // `REBIRTH`).
            if let Ok(mut rebirth) = REBIRTH.lock() {
                *rebirth = Some(path.clone());
            }
            note_playback_trace!(
                "audio step begin-again file {}",
                super::pin_trace::path_name(&path)
            );
            *slot = None;
            Some(path)
        }
        None => {
            // The two ends of a ladder are a guard that has nothing left to watch for.
            if finish.is_some() {
                *slot = None;
            }
            None
        }
    }
}

/// What one report comes to, with the line the read-back is written as.
fn the_step_of_a_read(guard: &mut Guard, report: Report) -> Step {
    match (guard.engine, report) {
        (Engine::Mpv, Report::Answers(answers)) => {
            let verdict = answers.verdict();
            // **A player that was heard and is not any more climbs the in-place rungs once more**, and
            // only those (see `re_opened`).
            if a_late_failure(guard.heard, verdict) {
                guard.spend = re_opened(guard.spend);
            }
            // **The line is written where the verdict moved**, so a film that stays heard writes
            // nothing at all after its first read-back (see the module note). It is written before
            // the step is taken either way, so that a run reads what the player said against what was
            // then done about it even where the step itself is what fails.
            if !(guard.heard && verdict == Verdict::Heard) {
                note_playback_trace!("{}", a_read_back(guard.pid, guard.reads, &answers));
            }
            guard.heard = verdict == Verdict::Heard;

            let opening = answers.a_driver_with_no_params();
            let step = the_step(verdict, guard.spend, opening);
            // The second look a player still opening its output is given is this read's own rung, and
            // is spent here rather than in the tick, which cannot tell one second look from the other.
            // An `Unanswered` read that also names a driver spends only the other one: nothing was
            // learned about its track, so nothing is known about its output either.
            if step == Step::AskAgain && verdict == Verdict::NoDevice && opening {
                guard.spend.opening = true;
            }
            step
        }
        (Engine::Ffplay, Report::Spoke) => {
            note_playback_trace!("audio guard spoke pid {} read {}", guard.pid, guard.reads);
            the_step_of_a_signature(guard.spend)
        }
        // An answer one player gives and the other does not: nothing is known, so nothing is done.
        _ => Step::Nothing,
    }
}

/// The line one read-back is written as: the four properties the player reported, and what they come
/// to — the trace's own `engine truth` line with this ladder's numbering on it, because a film watched
/// by the guard is read per rung rather than once.
///
/// **`audio-device` is the device that was asked for and `current-ao` the driver that opened**, so the
/// line says both halves of what the cure is about: a player asked for `auto` and playing through
/// `null` is a different fault from one asked for a device it was given. Both are written as `-`
/// where a player answered nothing, which is how a build without the property has always read.
fn a_read_back(pid: u32, read: u8, answers: &EngineAnswers) -> String {
    format!(
        "audio read-back pid {pid} read {read} aid {} audio-device {} current-ao {} \
         audio-out-params {} verdict {}",
        answers.aid.as_deref().unwrap_or("-"),
        answers.audio_device.as_deref().unwrap_or("-"),
        answers.current_ao.as_deref().unwrap_or("-"),
        answers.audio_out_params.as_deref().unwrap_or("-"),
        super::playback_trace::the_verdict_word(answers.verdict()),
    )
}

/// Whether one read is a player that was heard and is not any more: the one thing that re-opens the
/// ladder's in-place rungs, and the only way a film that began with its sound can climb it at all.
fn a_late_failure(heard: bool, verdict: Verdict) -> bool {
    heard && verdict != Verdict::Heard
}

/// **What a late failure hands back, and what it never does.** The track and the device are open once
/// more on a player that has gone silent after being heard — an endpoint taken away is worth the two
/// writes it took the first time — and the last rung is *not* handed back with them: the relaunch is
/// the once-in-a-player's-life rung, so a film whose device is genuinely gone keeps its picture and
/// stops being climbed for (see `the_step`).
fn re_opened(spend: Spent) -> Spent {
    Spent {
        track: false,
        device: false,
        ..spend
    }
}

/// **The ladder's decisions, and nothing else** — what a read-back asks for next, out of the rungs
/// this guard has left.
///
/// A film that was heard ends it wherever it is found. A player that answered nothing is asked once
/// more and then left alone: nothing was learned about its sound, so nothing may be done to it. A
/// track with no output from it escalates past the track, because asking for a track it already has
/// is a property written for nothing. And both cures spent with neither heard is the last rung — one
/// new player — after which there is nothing left to say.
///
/// `opening` is the shape of a read that named a real driver on `current-ao` with no
/// `audio-out-params` beside it (see `EngineAnswers::a_driver_with_no_params`): the one read that is
/// both a silent film and a player mid-open, and the reason the device cure is written on its second
/// reading. `aid=no` and a `null` driver are answers, so each of those takes its rung on the first
/// read, exactly as it always did.
fn the_step(verdict: Verdict, spend: Spent, opening: bool) -> Step {
    match verdict {
        Verdict::Heard => Step::Nothing,
        Verdict::Unanswered if !spend.asked_again => Step::AskAgain,
        Verdict::Unanswered => Step::GiveUp,
        Verdict::NoTrack if !spend.track => Step::AskForATrack,
        Verdict::NoDevice if opening && !spend.opening => Step::AskAgain,
        Verdict::NoTrack | Verdict::NoDevice if !spend.device => Step::AskForADevice,
        _ if spend.again => Step::GiveUp,
        _ => Step::BeginAgain,
    }
}

/// The same ladder for the player with no pipe: a signature it wrote is the whole of what is known
/// about it, and one player begun again is the whole of what can be done about it.
fn the_step_of_a_signature(spend: Spent) -> Step {
    if spend.again {
        Step::GiveUp
    } else {
        Step::BeginAgain
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A file named as a film is, for the rules that turn on which one is being shown.
    fn a_film() -> PathBuf {
        PathBuf::from("a-film.mkv")
    }

    /// A guard on the player of [`a_film`], for the rules that are asked of one read rather than of
    /// the slot the guard lives in.
    fn a_guard(engine: Engine, spend: Spent) -> Guard {
        Guard {
            pid: 7,
            path: a_film(),
            engine,
            spend,
            report: None,
            wants_again: false,
            reads: 0,
            watches: 0,
            heard: false,
        }
    }

    /// **Which launch is watched at all**: a pin's own film, and one begun at a level. A hover, a
    /// sound and a film the launch itself gave no track are none of the guard's business.
    #[test]
    fn only_a_pins_film_begun_at_a_level_is_watched() {
        assert!(guarded(false, 40), "a pin's film at 40 is the fault");
        assert!(
            !guarded(true, 40),
            "a hover's film is not a pin's: nothing about it is the pin's to begin again"
        );
        assert!(
            !guarded(false, 0),
            "a level of nothing is the launch's own form, and its silence is the setting's"
        );
    }

    /// **The ladder itself**, rung by rung, taken from what the read-back said and never from a write
    /// having landed: no track is a track asked back for, a track with no output is the device, and a
    /// film still silent after both is one player begun again.
    #[test]
    fn the_ladder_escalates_on_what_the_player_said() {
        let fresh = Spent::default();

        assert_eq!(
            the_step(Verdict::Heard, fresh, false),
            Step::Nothing,
            "a film that was heard is left entirely alone"
        );
        assert_eq!(the_step(Verdict::NoTrack, fresh, false), Step::AskForATrack);
        assert_eq!(
            the_step(Verdict::NoDevice, fresh, false),
            Step::AskForADevice,
            "a track is selected and has no output: the track is already the one it wants"
        );

        let track_asked = Spent {
            track: true,
            ..fresh
        };
        assert_eq!(
            the_step(Verdict::NoTrack, track_asked, false),
            Step::AskForADevice,
            "a track asked for and still not selected leaves the device as the next thing to try"
        );
        assert_eq!(
            the_step(Verdict::NoDevice, track_asked, false),
            Step::AskForADevice
        );

        let both_cures = Spent {
            track: true,
            device: true,
            ..fresh
        };
        assert_eq!(
            the_step(Verdict::NoDevice, both_cures, false),
            Step::BeginAgain,
            "both cures spent and no sound: one new player, and it is the last rung"
        );
        assert_eq!(
            the_step(Verdict::NoTrack, both_cures, false),
            Step::BeginAgain
        );

        assert_eq!(
            the_step(Verdict::NoDevice, SPENT, false),
            Step::GiveUp,
            "a player begun in place of one that had its relaunch gets no second one"
        );
    }

    /// **A driver named with nothing playing through it is read twice before anything is written to
    /// it.** That read is `no-device` to the verdict and a player half a second into its start to the
    /// eye — the shape a captured run had 548 ms after its launch — so the first one earns the same
    /// spare read a player that answered nothing does, and only a second consecutive one of that shape
    /// earns `audio-device=auto`. A `null` driver, and no track at all, are answers rather than a
    /// player mid-open: each takes its rung on the first read, exactly as it always did.
    #[test]
    fn a_driver_named_with_nothing_playing_is_read_once_more_before_the_device_is_asked_for() {
        let fresh = Spent::default();

        assert_eq!(
            the_step(Verdict::NoDevice, fresh, true),
            Step::AskAgain,
            "one read of that shape is not yet a film that lost its device"
        );
        assert_eq!(
            the_step(
                Verdict::NoDevice,
                Spent {
                    opening: true,
                    ..fresh
                },
                true
            ),
            Step::AskForADevice,
            "a second consecutive read of it is what earns the write"
        );

        assert_eq!(
            the_step(Verdict::NoDevice, fresh, false),
            Step::AskForADevice,
            "a null driver, or a driver that answered nothing, is a device that failed on the first \
             read"
        );
        assert_eq!(
            the_step(Verdict::NoTrack, fresh, true),
            Step::AskForATrack,
            "no track outranks the driver's shape: that rung is the launch's own road"
        );
    }

    /// **A film that was heard and has gone silent climbs the ladder once more, and never its last
    /// rung twice**: the track and the device are handed back, the relaunch is not — a player whose
    /// device is genuinely gone keeps its picture, however many episodes it has.
    #[test]
    fn a_late_failure_climbs_the_in_place_rungs_again_and_never_the_relaunch_again() {
        let whole = Spent {
            asked_again: true,
            track: true,
            device: true,
            ..Spent::default()
        };

        let again = re_opened(whole);
        assert!(
            !again.track && !again.device,
            "both in-place rungs are handed back"
        );
        assert!(
            !again.again,
            "and the relaunch is the player's one, not the episode's"
        );
        assert_eq!(the_step(Verdict::NoTrack, again, false), Step::AskForATrack);
        assert_eq!(
            the_step(Verdict::NoDevice, again, false),
            Step::AskForADevice
        );

        // The rung taken and the player still silent is one new player, because this one has never had
        // its relaunch. A film that has is given nothing, however many episodes it goes through.
        assert_eq!(
            the_step(
                Verdict::NoDevice,
                Spent {
                    device: true,
                    ..again
                },
                false
            ),
            Step::BeginAgain
        );
        let spent_it = re_opened(Spent {
            again: true,
            ..whole
        });
        assert_eq!(
            the_step(
                Verdict::NoDevice,
                Spent {
                    device: true,
                    ..spent_it
                },
                false
            ),
            Step::GiveUp,
            "a device that is genuinely gone must never be begun again for"
        );

        assert!(a_late_failure(true, Verdict::NoDevice));
        assert!(a_late_failure(true, Verdict::NoTrack));
        assert!(!a_late_failure(true, Verdict::Heard));
        assert!(
            !a_late_failure(false, Verdict::NoDevice),
            "a film that was never heard is on the ladder's first pass, not a late failure"
        );
    }

    /// **The watch of a heard player is sparse, and a spent player has none**: +2 s, +4 s, +8 s, then
    /// one every ten — and the player a spent ladder began has its one read and no schedule behind it.
    #[test]
    fn a_watch_is_sparse_and_a_spent_player_has_none() {
        let waits = (0..6).map(the_wait_before_a_watch).collect::<Vec<_>>();
        assert_eq!(
            waits,
            vec![
                Duration::from_secs(2),
                Duration::from_secs(4),
                Duration::from_secs(8),
                Duration::from_secs(10),
                Duration::from_secs(10),
                Duration::from_secs(10),
            ],
            "the schedule settles into ten seconds and never asks for a tighter read than that"
        );

        assert!(
            a_guard(Engine::Mpv, Spent::default()).keeps_watching(),
            "mpv's player is watched once it has been heard"
        );
        assert!(
            !a_guard(Engine::Mpv, SPENT).keeps_watching(),
            "the player begun by a spent ladder is given its one read and nothing more"
        );
        assert!(
            !a_guard(Engine::Ffplay, Spent::default()).keeps_watching(),
            "FFmpeg's player has no pipe to be watched down"
        );
    }

    /// **A healthy film writes one line about its sound, not one every ten seconds**: the read-back is
    /// written where the verdict moved, so a player that is heard again says nothing — and a verdict
    /// that does move, later in the film's life, is written with the rung it then earns.
    #[test]
    fn a_heard_player_writes_no_line_while_it_stays_heard() {
        use super::super::playback_trace::{clear_noted, noted_lines};

        let heard = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
        };
        let mut guard = a_guard(Engine::Mpv, Spent::default());
        guard.reads = 1;

        clear_noted();
        assert_eq!(
            the_step_of_a_read(&mut guard, Report::Answers(heard.clone())),
            Step::Nothing
        );
        let first = noted_lines();
        assert_eq!(first.len(), 1, "the read that was heard is the one line");
        assert!(first[0].contains("verdict heard"));

        for _ in 0..3 {
            guard.reads = guard.reads.saturating_add(1);
            assert_eq!(
                the_step_of_a_read(&mut guard, Report::Answers(heard.clone())),
                Step::Nothing
            );
        }
        assert_eq!(
            noted_lines(),
            first,
            "a film that stays heard writes nothing after its first line"
        );

        // And the read that finds it silent is written, with the rung it earns. Its shape is a driver
        // named with nothing playing through it, so the first read of it is the spare one and the
        // second is the write (see `the_step`).
        guard.reads = guard.reads.saturating_add(1);
        let silent = EngineAnswers {
            audio_out_params: None,
            ..heard
        };
        assert_eq!(
            the_step_of_a_read(&mut guard, Report::Answers(silent.clone())),
            Step::AskAgain,
            "a driver named with nothing playing is read once more before it is written to"
        );
        let late = noted_lines();
        assert_eq!(late.len(), 2, "a verdict that moves is written down");
        assert!(late[1].contains("verdict no-device"));

        guard.reads = guard.reads.saturating_add(1);
        assert_eq!(
            the_step_of_a_read(&mut guard, Report::Answers(silent)),
            Step::AskForADevice,
            "a second read of that shape climbs the device rung"
        );
        assert_eq!(
            noted_lines().len(),
            3,
            "and a read taken inside the ladder's own pass is written too"
        );
    }

    /// **A player that answers nothing is not a film without sound**: it is asked once more, and then
    /// left alone — the one verdict the guard may not act on is the one it never got.
    #[test]
    fn a_player_that_says_nothing_is_asked_once_more_and_never_begun_again() {
        let fresh = Spent::default();
        assert_eq!(the_step(Verdict::Unanswered, fresh, false), Step::AskAgain);
        assert_eq!(
            the_step(
                Verdict::Unanswered,
                Spent {
                    asked_again: true,
                    ..fresh
                },
                false
            ),
            Step::GiveUp
        );
        assert_eq!(
            the_step(Verdict::Unanswered, SPENT, false),
            Step::GiveUp,
            "and a player begun in place of another is given no second look at all"
        );
    }

    /// FFmpeg's player has no pipe, so its ladder is a signature and then a relaunch — once, and then
    /// nothing, whatever it writes afterwards.
    #[test]
    fn a_signature_begins_the_player_once() {
        let fresh = Spent::default();
        assert_eq!(the_step_of_a_signature(fresh), Step::BeginAgain);
        assert_eq!(
            the_step_of_a_signature(Spent {
                again: true,
                ..fresh
            }),
            Step::GiveUp,
            "never twice, and the replacement is watched with the rung already spent"
        );
        assert_eq!(the_step_of_a_signature(SPENT), Step::GiveUp);
    }

    /// **The record a replacement is watched by**: the ladder is handed on only to the film the
    /// record names, and a record for another film is a chain that has ended rather than a rung.
    #[test]
    fn a_replacement_is_watched_with_the_last_rung_spent_and_only_for_its_own_film() {
        assert_eq!(the_spend_of(None, &a_film()), Spent::default());
        assert_eq!(the_spend_of(Some(&a_film()), &a_film()), SPENT);
        assert_eq!(
            the_spend_of(Some(Path::new("another-film.mkv")), &a_film()),
            Spent::default(),
            "a relaunch the pin declined is a chain that ended, not a rung of this one"
        );
    }

    /// **The strict matcher, and what strictness is for**: a device that would not open is one of the
    /// four lines below and an ordinary warning about anything else is not a reason to take a picture
    /// off the screen and begin a player again.
    #[test]
    fn only_an_audio_devices_failure_is_a_signature() {
        for line in [
            "Could not open audio device: WASAPI can't find requested audio endpoint",
            "Could not open audio device: Device unavailable",
            "Failed to initialize the SDL audio subsystem: No available audio device",
            "Failed to initialize audio filters!",
            "Could not initialize audio resampler",
            // A player that capitalises the whole of its sentence is the same sentence.
            "COULD NOT OPEN AUDIO DEVICE",
        ] {
            assert!(
                names_an_audio_failure(line),
                "{line:?} is a device that would not open"
            );
        }

        for line in [
            "Could not open file 'a-film.mkv': No such file or directory",
            "Failed to initialize the video output",
            "Failed to initialize video filters!",
            "Packet corrupt (stream = 0, dts = 1259)",
            "Audio device lost, reopening",
            "audio: no such filter",
            "Could not find codec parameters for stream 1",
            "SDL video driver lost",
            "Warning: corrupt input packet",
        ] {
            assert!(
                !names_an_audio_failure(line),
                "{line:?} is not one of the failures this guard may begin a player for"
            );
        }
    }

    /// **What a read-back is read as**: the four answers the player gave, named as the trace names
    /// them, so a field log carries the whole of the ladder's evidence — the device that was asked for
    /// beside the driver that answered it, and `-` where a player said nothing to either.
    #[test]
    fn a_read_back_is_the_players_own_answers_and_the_verdict_they_come_to() {
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
        };
        assert_eq!(
            a_read_back(7, 1, &answers),
            "audio read-back pid 7 read 1 aid 1 audio-device auto current-ao wasapi \
             audio-out-params {\"samplerate\":48000} verdict heard"
        );

        let silent = EngineAnswers::default();
        assert_eq!(
            a_read_back(9, 2, &silent),
            "audio read-back pid 9 read 2 aid - audio-device - current-ao - audio-out-params - \
             verdict unanswered",
            "a player that said nothing is written as nothing said rather than as a guess"
        );
    }
}

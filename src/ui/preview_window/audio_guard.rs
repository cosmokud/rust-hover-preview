//! A film that was begun at a level and is playing with no sound: the ladder that tries to give it
//! one, and the one player a pin's film is ever begun again for.
//!
//! **The roads to a silent film are one page over** (see `playback_trace`): a launch that handed its
//! player a level and a player that never opened a device from it look identical from this side of the
//! process boundary, and the difference is only ever read off the player — `aid` for a track,
//! `current-ao` and `audio-out-params` for an output, `audio-device` for the one that was asked for,
//! and `volume` and `mute` for a player that has all of that and was told to be quiet. So this module
//! asks the same six questions the trace asks, reads the same verdict out of them
//! (`playback_trace::Verdict`), and then *does* something about it, which is the whole of what the
//! trace is not for.
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
//! 3. neither heard, and the player says it is open — **the level the launch asked for asserted again**,
//!    with the player's own `mute` taken off beside it (`Verdict::Silent`): a track selected, an
//!    output open, and a player told to be quiet is a film that every other answer calls sound. The
//!    level itself is written by the one function that owns it (`mpv_ipc::set_volume`), so a player
//!    the launch gave no track to has it asked back for by the same line.
//! 4. neither heard — **one** new player, with the arguments, the second and the pause the one it
//!    replaces had, and one read of that player. Then the ladder ends.
//!
//! **A hover's ladder is the in-place rungs and nothing else** (see `Road`): a hover is a film that
//! comes and goes, its player is the resident's, and the wait a fresh player costs is the whole of
//! what the resident exists to save — so a hover gives up where a pin begins a player again. **And it
//! is never watched on a schedule**: a hover is short-lived and its per-serve reads are the whole of
//! what there is to know, where a pin's film is sampled at +2 s, +4 s and +8 s and then for as long
//! as it lives (see `WATCHES`).
//!
//! **Nothing here fights the film.** No pause is toggled, no seek is asked, and the only level written
//! is the one the launch already asked for, put back on a player that had been silenced under it — and
//! a film held by the handover's own pause is *begun again held* rather than resumed: the only things
//! this module writes are the audio properties, and the only player it begins is a replacement at the
//! second the pin's clock already says (see
//! `pin_playback::begin_the_pinned_film_again_for_its_sound`).
//!
//! **FFmpeg's player has no pipe, so its ladder is one rung and it is the last one.** It can be told
//! nothing, and it reports the device it could not open on the standard error this app used to throw
//! away — so a launch it stands behind pipes that stream and drains it (see
//! `player_logs::drain_a_players_errors`), and a line matching [`AUDIO_FAILURES`] in a launch this
//! guard stands behind is one player begun again. Not two, and never for a line it did not see. A
//! launch it does not stand behind is given `null` and no thread, the way every ordinary launch is.
//!
//! **A film that was heard is not a film that will always be heard, so a pin's guard no longer ends
//! where the ladder does.** A player whose sound came up is kept under a sparse schedule of read-backs
//! — +2 s, +4 s, +8 s, then one every ~10 s for as long as the player lives (see `WATCHES`) — because
//! an endpoint lost after a film started leaves no line anywhere else, and until now only the first
//! second of a film's life was ever sampled. A hover has no schedule: it comes and goes, and its
//! per-serve reads are the whole of what there is to know about it. A read that is heard again writes
//! nothing at all: the line is written where the verdict moved, so a healthy film's log is the one
//! line it opened with.
//! A late failure climbs the three in-place rungs once more — `aid`, then `audio-device`, then the
//! level — and may take the relaunch **at most once in a player's whole life**, not once per episode
//! (see `re_opened`): a device that is genuinely gone must never be begun again for, every ten
//! seconds, for ever. The sampling ends with the player: a probe that cannot reach this guard's player
//! reports nothing and schedules nothing, and `a_player_ended` is the half that needs no probe to have
//! been woken at all.
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

/// Which launches the guard stands behind: **every film begun at a level** — on a pin's road for both
/// players, and on a hover's for mpv alone — and no film whose level the launch itself resolved to
/// nothing.
///
/// The form is the one rule both routes' arguments are built from (see `video_launch::AudioForm`): a
/// launch that gave its player no track at all has no ladder to climb, because that silence is the
/// level asked for rather than a fault.
///
/// **A hover is watched too, because the resident is the player a stale audio decision lives longest
/// in**: one process handed file after file down a pipe, begun at whatever level the first hover had,
/// with every per-file option `loadfile` re-applies — `--aid=no` among them — and a pipe that only
/// asks the track back when the launch recorded the audio as off (see `mpv_resident` and
/// `mpv_ipc::set_volume`). A hover's ladder is the in-place rungs and nothing else, which is what
/// `Road` says and what `arm` records.
///
/// **FFmpeg's player is the one launch a hover is still not guarded on**, and for a reason of its own:
/// its ladder has no in-place rung at all — it can be told nothing — so a hover's guard on it could
/// only ever give up, and a hover that is piped and drained for nothing is the ordinary hover this app
/// has always made (see `video_playback`, which pipes its errors on exactly the launches that are
/// guarded).
pub(super) fn guarded(engine: Engine, for_a_hover: bool, volume: u32) -> bool {
    if super::video_launch::mpv_launched_without_audio(volume) {
        return false;
    }
    engine == Engine::Mpv || !for_a_hover
}

/// Which road armed the guard, and the whole of the difference between them: a pin's film may be
/// begun again at the end of the ladder and is watched once it was heard, and a hover's may not.
///
/// **The two are one guard and one set of questions**, so the difference is a value on the guard rather
/// than a second ladder: what a read said, what has been spent and what a probe is asked are all the
/// same facts on either road (see `the_step`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Road {
    /// A pin's own film: cured in place, begun again once, and watched for as long as it plays.
    Pin,
    /// A hover's film, on a cold launch or on the resident's serve: cured in place, and then left.
    Hover,
}

impl Road {
    /// Whether a player at the end of the ladder may be begun again (see `the_step`).
    fn relaunches(self) -> bool {
        self == Self::Pin
    }

    /// Whether a player that was heard is kept under the sparse schedule (see `WATCHES`).
    fn watches(self) -> bool {
        self == Self::Pin
    }

    /// The word a line about the guard names the road by.
    fn word(self) -> &'static str {
        match self {
            Self::Pin => "pin",
            Self::Hover => "hover",
        }
    }
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
    /// The level the launch asked for, asserted again, with the player's own mute taken off beside it.
    level: bool,
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
    level: true,
    again: true,
};

/// What a player has said, and no tick has read yet.
enum Report {
    /// The six answers an mpv gave (see `playback_trace::ask_the_engine`).
    Answers(EngineAnswers),
    /// A line FFmpeg's player wrote that is one of its audio failures (see `AUDIO_FAILURES`).
    Spoke,
}

/// One rung of the ladder, or an end of it (see `the_step`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Step {
    /// Heard, so there is nothing left to do — or a report this guard has nothing to say about.
    Nothing,
    /// Ask the player the same six questions once more, and write no property at all.
    AskAgain,
    /// `aid` back to `auto`: the launch's own road.
    AskForATrack,
    /// `audio-device` back to `auto`, which makes mpv reload its audio output.
    AskForADevice,
    /// The level the launch asked for, asserted again, and the player's own mute taken off.
    AskForALevel,
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
    /// Which road armed it, because the two roads end differently (see `Road`).
    road: Road,
    /// The level the launch asked for, which is what a `Silent` player is taken back up to — and not
    /// the level the player reports, which is the fact the verdict was read from (see
    /// `playback_trace::EngineAnswers::told_to_be_quiet`).
    volume: u32,
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
    /// Whether this player is watched after it was heard: mpv's, on a pin's road, and not the one a
    /// spent ladder began in place of another — that player is given its one read and nothing more, so
    /// a film whose device is genuinely gone is not climbed for again and again, and a hover is never
    /// put on a schedule at all (see `Road`).
    fn keeps_watching(&self) -> bool {
        self.engine == Engine::Mpv && self.road.watches() && self.spend != SPENT
    }

    /// **A pin that has taken this player over is the pin's from here**, whatever road armed the guard:
    /// the process and the reads already taken from it do not change, and the rungs a hover may never
    /// take are a pin's to take (see `a_pin_took_it_over` and `the_step`).
    fn a_pin_took_it_over(&mut self) {
        self.road = Road::Pin;
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
///
/// **A hover's serve arms it as well, and a serve is a new episode of the same pid**: the resident is
/// handed one file after another, so what the serve before it spent belongs to that file — the ladder
/// is armed whole again, with no read counted and nothing heard yet. That reset is the arming itself
/// rather than a second step, because a guard is never carried across a file: a probe of the serve
/// before it reports to nothing (see `report`), and the pin's `REBIRTH` record — the one thing that is
/// carried, and only from a relaunch to the player begun in its place — is not a hover's to spend (see
/// `the_spend_for`).
pub(super) fn arm(pid: u32, path: &Path, volume: u32, engine: Engine, for_a_hover: bool) {
    if !guarded(engine, for_a_hover, volume) {
        return;
    }

    let road = if for_a_hover { Road::Hover } else { Road::Pin };
    let guard = the_guard_a_launch_arms(pid, path, volume, engine, road);
    let spend = guard.spend;
    let Ok(mut slot) = GUARD.lock() else {
        return;
    };
    *slot = Some(guard);
    drop(slot);

    note_playback_trace!(
        "audio guard armed pid {pid} engine {} road {} level {volume} file {} rungs {}",
        engine.word(),
        road.word(),
        super::pin_trace::path_name(path),
        if spend == SPENT { "spent" } else { "whole" },
    );

    // A player with a pipe is asked what it did with the sound; FFmpeg's player has none, and is read
    // off its own stream instead (see `a_players_line`).
    if engine == Engine::Mpv {
        spawn_probe(pid);
    }
}

/// The guard one launch arms: the player, the road it was begun on, the level it was asked for, and a
/// ladder that starts at its top.
///
/// **A serve is a new episode of the same pid, so nothing of the one before it is carried**: no read is
/// counted, nothing has been heard, no watch is scheduled and no relaunch is waiting for a tick — only
/// the rungs are handed on, and only ever within one player's life (see `the_spend_for`). That is the
/// whole of the reset a serve gets, and it is the arming rather than a second step: a guard built here
/// has nothing of any other file in it.
fn the_guard_a_launch_arms(
    pid: u32,
    path: &Path,
    volume: u32,
    engine: Engine,
    road: Road,
) -> Guard {
    Guard {
        pid,
        path: path.to_path_buf(),
        engine,
        road,
        volume,
        spend: the_spend_for(road, path),
        report: None,
        wants_again: false,
        reads: 0,
        watches: 0,
        heard: false,
    }
}

/// What a guard is handed of the ladder: every rung, unless the player it is watching was begun in
/// place of one whose own ladder had already reached the last rung.
///
/// **A hover reads no record at all**, which is more than spending none of it: the `REBIRTH` slot is
/// emptied by whoever reads it, so a serve that read a pin's relaunch record would end a chain that no
/// file of its own had begun (see `the_spend_on`).
fn the_spend_for(road: Road, path: &Path) -> Spent {
    match road {
        Road::Hover => Spent::default(),
        Road::Pin => the_spend_on(road, the_rebirth_of(path)),
    }
}

/// The rule [`the_spend_for`] asks, with the record passed in rather than taken, so that it can be
/// answered without one.
///
/// **A hover spends nothing it did not spend on its own file.** The record is a pin's relaunch handed to
/// the player begun in its place, and the hover road never begins one: a serve that took it would
/// inherit rungs a film it never replaced had already used, and a hover whose audio went stale would be
/// left with no cure at all.
fn the_spend_on(road: Road, recorded: Spent) -> Spent {
    match road {
        Road::Pin => recorded,
        Road::Hover => Spent::default(),
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

/// A pin has taken a hover's player over as its own picture (see
/// `pin_film_host::promote_the_hover_film_host`): the game is the pin's from here, so the rungs a hover
/// may never take are the pin's to take again.
///
/// **Nothing is re-read and nothing is re-armed.** The player did not change and the reads already
/// taken were taken from it; what changes is only what may be done about what they said — and a hover's
/// film is exactly the one a pin is most likely to take up, so a guard left on the hover's road would
/// be a pin that can never begin its player again (see `Road`).
pub(super) fn a_pin_took_it_over(pid: u32) {
    let Ok(mut slot) = GUARD.lock() else {
        return;
    };
    let Some(guard) = slot.as_mut().filter(|guard| guard.pid == pid) else {
        return;
    };
    if guard.road == Road::Hover {
        guard.a_pin_took_it_over();
        note_playback_trace!("audio guard a pin took it over pid {pid}");
    }
}

/// A player this app was watching has ended, which is the end of any guard on it: the stream that
/// would report for it is closed, the film a hover's serve was watched for is over however long the
/// resident stands by, and a guard left standing would be waiting on a player that is not showing it
/// anything (see `player_logs`, whose threads end where the player does).
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
                // **The level is taken back up through the one function that owns it** (see
                // `mpv_ipc::set_volume`, which is also what asks a track back for a player the launch
                // gave none), and the player's own mute is taken off beside it: a level nobody can
                // hear through a mute is not the level asserted again (see `Verdict::Silent`).
                Step::AskForALevel => {
                    guard.spend.level = true;
                    let landed = super::mpv_ipc::set_volume(guard.pid, guard.volume)
                        && super::mpv_ipc::set_muted(guard.pid, false);
                    note_playback_trace!(
                        "audio step ask-for-a-level pid {} level {} landed {}",
                        guard.pid,
                        guard.volume,
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
            let verdict = answers.verdict(guard.volume);
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
                note_playback_trace!(
                    "{}",
                    a_read_back(guard.pid, guard.reads, guard.volume, &answers)
                );
            }
            guard.heard = verdict == Verdict::Heard;

            let opening = answers.a_driver_with_no_params();
            let step = the_step(verdict, guard.spend, opening, guard.road);
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

/// The line one read-back is written as: the level the launch asked for, the six properties the player
/// reported, and what they come to — the trace's own `engine truth` line with this ladder's numbering on
/// it, because a film watched by the guard is read per rung rather than once.
///
/// **`asked` is the launch's level and `volume` the player's own**, the two words the third road is read
/// by: a player reporting 0 or `mute yes` on a launch that asked for 100 is a film that is open and was
/// told to be quiet. **`audio-device` is the device that was asked for and `current-ao` the driver that
/// opened**, so the line says both halves of what the cure is about: a player asked for `auto` and
/// playing through `null` is a different fault from one asked for a device it was given. Both are
/// written as `-` where a player answered nothing, which is how a build without the property has always
/// read.
fn a_read_back(pid: u32, read: u8, asked: u32, answers: &EngineAnswers) -> String {
    format!(
        "audio read-back pid {pid} read {read} asked {asked} volume {} mute {} aid {} \
         audio-device {} current-ao {} audio-out-params {} verdict {}",
        super::playback_trace::the_level_word(answers.volume),
        super::playback_trace::the_flag_word(answers.mute),
        answers.aid.as_deref().unwrap_or("-"),
        answers.audio_device.as_deref().unwrap_or("-"),
        answers.current_ao.as_deref().unwrap_or("-"),
        answers.audio_out_params.as_deref().unwrap_or("-"),
        super::playback_trace::the_verdict_word(answers.verdict(asked)),
    )
}

/// Whether one read is a player that was heard and is not any more: the one thing that re-opens the
/// ladder's in-place rungs, and the only way a film that began with its sound can climb it at all.
fn a_late_failure(heard: bool, verdict: Verdict) -> bool {
    heard && verdict != Verdict::Heard
}

/// **What a late failure hands back, and what it never does.** The track, the device and the level are
/// open once more on a player that has gone silent after being heard — an endpoint taken away, or a
/// per-file option applied to a resident, is worth the writes it took the first time — and the last rung
/// is *not* handed back with them: the relaunch is the once-in-a-player's-life rung, so a film whose
/// device is genuinely gone keeps its picture and stops being climbed for (see `the_step`).
fn re_opened(spend: Spent) -> Spent {
    Spent {
        track: false,
        device: false,
        level: false,
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
/// new player on a pin's road, and nothing at all on a hover's — after which there is nothing left to
/// say.
///
/// `opening` is the shape of a read that named a real driver on `current-ao` with no
/// `audio-out-params` beside it (see `EngineAnswers::a_driver_with_no_params`): the one read that is
/// both a silent film and a player mid-open, and the reason the device cure is written on its second
/// reading. `aid=no` and a `null` driver are answers, so each of those takes its rung on the first
/// read.
///
/// **`Silent` is the verdict the other two are not, and it outranks them**: a track and an output are
/// the two things its player already has, so the level is what is asserted rather than a track or a
/// device asked for again (see `Verdict::Silent`). A player still quiet after that is one new player on
/// a pin's road — and nothing at all on a hover's, whose rungs stop where `Road` says.
fn the_step(verdict: Verdict, spend: Spent, opening: bool, road: Road) -> Step {
    match verdict {
        Verdict::Heard => Step::Nothing,
        Verdict::Unanswered if !spend.asked_again => Step::AskAgain,
        Verdict::Unanswered => Step::GiveUp,
        Verdict::NoTrack if !spend.track => Step::AskForATrack,
        Verdict::NoDevice if opening && !spend.opening => Step::AskAgain,
        Verdict::Silent if !spend.level => Step::AskForALevel,
        Verdict::Silent if road.relaunches() && !spend.again => Step::BeginAgain,
        Verdict::Silent => Step::GiveUp,
        Verdict::NoTrack | Verdict::NoDevice if !spend.device => Step::AskForADevice,
        _ if spend.again || !road.relaunches() => Step::GiveUp,
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
            road: Road::Pin,
            volume: 40,
            spend,
            report: None,
            wants_again: false,
            reads: 0,
            watches: 0,
            heard: false,
        }
    }

    /// **Which launch is watched at all**: any film begun at a level, on either road — the pin's own
    /// and a hover's — and no film the launch itself gave no track, and no hover on FFmpeg's player,
    /// whose ladder has no in-place rung to climb.
    #[test]
    fn every_film_begun_at_a_level_is_watched_and_a_hover_on_ffmpeg_is_not() {
        assert!(
            guarded(Engine::Mpv, false, 40),
            "a pin's film at 40 is the fault"
        );
        assert!(
            guarded(Engine::Mpv, true, 40),
            "a hover's film is watched too: the resident is where a stale audio decision lives \
             longest, and its rungs are the in-place ones"
        );
        assert!(
            !guarded(Engine::Mpv, false, 0) && !guarded(Engine::Mpv, true, 0),
            "a level of nothing is the launch's own form, and its silence is the setting's"
        );
        assert!(
            guarded(Engine::Ffplay, false, 40),
            "a pin's film FFmpeg's player draws is read off its own stream"
        );
        assert!(
            !guarded(Engine::Ffplay, true, 40),
            "a hover on FFmpeg's player is piped and drained for a rung it could never take"
        );
    }

    /// **The ladder itself**, rung by rung, taken from what the read-back said and never from a write
    /// having landed: no track is a track asked back for, a track with no output is the device, and a
    /// film still silent after both is one player begun again.
    #[test]
    fn the_ladder_escalates_on_what_the_player_said() {
        let fresh = Spent::default();
        let pin = Road::Pin;

        assert_eq!(
            the_step(Verdict::Heard, fresh, false, pin),
            Step::Nothing,
            "a film that was heard is left entirely alone"
        );
        assert_eq!(
            the_step(Verdict::NoTrack, fresh, false, pin),
            Step::AskForATrack
        );
        assert_eq!(
            the_step(Verdict::NoDevice, fresh, false, pin),
            Step::AskForADevice,
            "a track is selected and has no output: the track is already the one it wants"
        );

        let track_asked = Spent {
            track: true,
            ..fresh
        };
        assert_eq!(
            the_step(Verdict::NoTrack, track_asked, false, pin),
            Step::AskForADevice,
            "a track asked for and still not selected leaves the device as the next thing to try"
        );
        assert_eq!(
            the_step(Verdict::NoDevice, track_asked, false, pin),
            Step::AskForADevice
        );

        let both_cures = Spent {
            track: true,
            device: true,
            ..fresh
        };
        assert_eq!(
            the_step(Verdict::NoDevice, both_cures, false, pin),
            Step::BeginAgain,
            "both cures spent and no sound: one new player, and it is the last rung"
        );
        assert_eq!(
            the_step(Verdict::NoTrack, both_cures, false, pin),
            Step::BeginAgain
        );

        assert_eq!(
            the_step(Verdict::NoDevice, SPENT, false, pin),
            Step::GiveUp,
            "a player begun in place of one that had its relaunch gets no second one"
        );
    }

    /// **A player that is open and was told to be quiet is taken back up to the level the launch
    /// asked for**, and never asked for a track or a device it already has: `Silent` is the one verdict
    /// whose two other cures are beside the point.
    ///
    /// A player still quiet after the level has had everything a pin's ladder can give it, so it is one
    /// new player — once, and never on a hover, whose rungs are the in-place ones only.
    #[test]
    fn a_silent_player_has_its_level_asserted_and_only_a_pin_begins_one_again() {
        let fresh = Spent::default();

        assert_eq!(
            the_step(Verdict::Silent, fresh, false, Road::Pin),
            Step::AskForALevel,
            "a track and an output are what this player has: the level is the rung"
        );
        assert_eq!(
            the_step(
                Verdict::Silent,
                Spent {
                    level: true,
                    ..fresh
                },
                false,
                Road::Pin
            ),
            Step::BeginAgain,
            "a pin's film that will not come back up is one new player"
        );
        assert_eq!(
            the_step(Verdict::Silent, SPENT, false, Road::Pin),
            Step::GiveUp,
            "and only one: the relaunch is the player's, not the episode's"
        );
    }

    /// **A hover's ladder is the in-place rungs and nothing else.** A `no-track`, a `no-device` and a
    /// `Silent` read each earn their cure where the player stands, and every rung spent is the end of
    /// it — a hover never begins a player again, whatever its read-back said.
    #[test]
    fn a_hovers_ladder_has_no_rung_that_begins_a_player_again() {
        let fresh = Spent::default();

        assert_eq!(
            the_step(Verdict::NoTrack, fresh, false, Road::Hover),
            Step::AskForATrack
        );
        assert_eq!(
            the_step(Verdict::NoDevice, fresh, false, Road::Hover),
            Step::AskForADevice
        );
        assert_eq!(
            the_step(Verdict::Silent, fresh, false, Road::Hover),
            Step::AskForALevel
        );

        let cures_spent = Spent {
            track: true,
            device: true,
            level: true,
            ..fresh
        };
        for verdict in [Verdict::NoTrack, Verdict::NoDevice, Verdict::Silent] {
            assert_eq!(
                the_step(verdict, cures_spent, false, Road::Hover),
                Step::GiveUp,
                "{verdict:?} on a hover with every cure spent is a film left as it is: a hover is \
                 the wait the resident exists to save"
            );
            assert_eq!(the_step(verdict, SPENT, false, Road::Hover), Step::GiveUp);
        }
    }

    /// **A serve is a new episode of the same pid, so it starts at the top of the ladder** — and the
    /// pin's `REBIRTH` record is not a hover's to spend, or a hover served after a pin's declined
    /// relaunch would be handed a spent ladder and every one of its cures would be gone.
    #[test]
    fn a_serve_arms_the_ladder_whole_and_spends_no_pins_record() {
        assert_eq!(
            the_spend_on(Road::Hover, Spent::default()),
            Spent::default(),
            "a hover's serve is every rung of its own, whatever a pin left behind"
        );
        assert_eq!(
            the_spend_on(Road::Hover, SPENT),
            Spent::default(),
            "a spent record standing against this very file is still not the hover's to spend"
        );
        assert_eq!(
            the_spend_on(Road::Pin, SPENT),
            SPENT,
            "and a pin's launch still reads the record it always did"
        );
    }

    /// **The arming is the reset.** A guard built for a serve counts no read, has heard nothing, is on
    /// no schedule and has no relaunch waiting for a tick — whatever the guard it replaces had done
    /// with the file before it.
    #[test]
    fn a_serve_arms_a_ladder_with_no_read_of_the_one_before_it() {
        let guard = the_guard_a_launch_arms(7, &a_film(), 100, Engine::Mpv, Road::Hover);

        assert_eq!(guard.reads, 0, "the episode's first read is its first");
        assert_eq!(guard.watches, 0, "a hover is on no schedule at all");
        assert!(!guard.heard, "nothing has been heard yet");
        assert!(!guard.wants_again, "no relaunch is waiting for a tick");
        assert_eq!(
            guard.spend,
            Spent::default(),
            "and every rung is this file's own"
        );
        assert_eq!(guard.volume, 100, "the level it was asked for is kept");
    }

    /// **A pin that takes a hover's player over gets the pin's road**: the same process and the same
    /// reads, with the two rungs a hover may never take now the pin's — which is the difference between
    /// a film that can be begun again and one that cannot (see
    /// `pin_film_host::promote_the_hover_film_host`).
    #[test]
    fn a_player_a_pin_took_over_is_watched_as_a_pins_film() {
        let mut guard = a_guard(Engine::Mpv, Spent::default());
        guard.road = Road::Hover;

        assert!(
            !guard.keeps_watching(),
            "a hover is never put on a schedule"
        );
        let spent_on_the_hover = Spent {
            track: true,
            device: true,
            ..Spent::default()
        };
        assert_eq!(
            the_step(Verdict::NoDevice, spent_on_the_hover, false, guard.road),
            Step::GiveUp,
            "a hover with its cures spent leaves the film as it is"
        );

        guard.a_pin_took_it_over();

        assert_eq!(guard.road, Road::Pin);
        assert!(
            guard.keeps_watching(),
            "the pin's film is watched once it was heard"
        );
        assert_eq!(
            the_step(Verdict::NoDevice, spent_on_the_hover, false, guard.road),
            Step::BeginAgain,
            "and the rung a hover may not take is the pin's from the take-over on"
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
            the_step(Verdict::NoDevice, fresh, true, Road::Pin),
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
                true,
                Road::Pin
            ),
            Step::AskForADevice,
            "a second consecutive read of it is what earns the write"
        );

        assert_eq!(
            the_step(Verdict::NoDevice, fresh, false, Road::Pin),
            Step::AskForADevice,
            "a null driver, or a driver that answered nothing, is a device that failed on the first \
             read"
        );
        assert_eq!(
            the_step(Verdict::NoTrack, fresh, true, Road::Pin),
            Step::AskForATrack,
            "no track outranks the driver's shape: that rung is the launch's own road"
        );
    }

    /// **A film that was heard and has gone silent climbs the ladder once more, and never its last
    /// rung twice**: the track, the device and the level are handed back, the relaunch is not — a
    /// player whose device is genuinely gone keeps its picture, however many episodes it has.
    #[test]
    fn a_late_failure_climbs_the_in_place_rungs_again_and_never_the_relaunch_again() {
        let whole = Spent {
            asked_again: true,
            track: true,
            device: true,
            level: true,
            ..Spent::default()
        };

        let again = re_opened(whole);
        assert!(
            !again.track && !again.device && !again.level,
            "every in-place rung is handed back"
        );
        assert!(
            !again.again,
            "and the relaunch is the player's one, not the episode's"
        );
        assert_eq!(
            the_step(Verdict::NoTrack, again, false, Road::Pin),
            Step::AskForATrack
        );
        assert_eq!(
            the_step(Verdict::NoDevice, again, false, Road::Pin),
            Step::AskForADevice
        );
        assert_eq!(
            the_step(Verdict::Silent, again, false, Road::Pin),
            Step::AskForALevel,
            "a film muted after it was heard is taken back up to its level"
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
                false,
                Road::Pin
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
                false,
                Road::Pin
            ),
            Step::GiveUp,
            "a device that is genuinely gone must never be begun again for"
        );

        assert!(a_late_failure(true, Verdict::NoDevice));
        assert!(a_late_failure(true, Verdict::NoTrack));
        assert!(
            a_late_failure(true, Verdict::Silent),
            "an open output that was told to be quiet is a late failure like any other"
        );
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
            volume: Some(100.0),
            mute: Some(false),
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
        assert_eq!(
            the_step(Verdict::Unanswered, fresh, false, Road::Pin),
            Step::AskAgain
        );
        assert_eq!(
            the_step(
                Verdict::Unanswered,
                Spent {
                    asked_again: true,
                    ..fresh
                },
                false,
                Road::Pin
            ),
            Step::GiveUp
        );
        assert_eq!(
            the_step(Verdict::Unanswered, SPENT, false, Road::Pin),
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

    /// **What a read-back is read as**: the level the launch asked for, the six answers the player gave
    /// named as the trace names them, and the verdict they come to — so a field log carries the whole of
    /// the ladder's evidence: the level asked for beside the level the player is at, the device that was
    /// asked for beside the driver that answered it, and `-` where a player said nothing to either.
    #[test]
    fn a_read_back_is_the_players_own_answers_and_the_verdict_they_come_to() {
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            volume: Some(100.0),
            mute: Some(false),
        };
        assert_eq!(
            a_read_back(7, 1, 100, &answers),
            "audio read-back pid 7 read 1 asked 100 volume 100 mute no aid 1 audio-device auto \
             current-ao wasapi audio-out-params {\"samplerate\":48000} verdict heard"
        );

        // And the third road, as the read-back writes it: a player at nothing on a launch that asked for
        // 100 is the film the level rung exists for.
        let quiet = EngineAnswers {
            volume: Some(0.0),
            ..answers
        };
        assert_eq!(
            a_read_back(7, 2, 100, &quiet),
            "audio read-back pid 7 read 2 asked 100 volume 0 mute no aid 1 audio-device auto \
             current-ao wasapi audio-out-params {\"samplerate\":48000} verdict silent"
        );

        let silent = EngineAnswers::default();
        assert_eq!(
            a_read_back(9, 2, 40, &silent),
            "audio read-back pid 9 read 2 asked 40 volume - mute - aid - audio-device - current-ao - \
             audio-out-params - verdict unanswered",
            "a player that said nothing is written as nothing said rather than as a guess"
        );
    }
}

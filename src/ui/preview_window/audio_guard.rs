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
//! **The fourth road is not the player's at all, and it is asked at the same read.** A captured run
//! proved it: two players of eighteen had a Windows audio session that was **muted** while mpv reported
//! `volume 94 mute no` — two switches, and the player can see only its own — with that session's slider
//! at `mixer 1.000` and its channels at `1.000 1.000`, so nothing a hand had moved was down. So every
//! read this guard makes asks the machine as well as the player (`audio_session::look_and_cure`), on
//! the same probe thread and straight after the six engine questions, and the answer comes back folded
//! into the verdict: a player whose own answers are sound and whose session is still muted is
//! `Verdict::SessionMuted`, and the mute is taken off there and then — **whether or not
//! `RHP_PLAYBACK_TRACE` is set**, because the trace is the recorder and this is the fix. **No rung of
//! this ladder is written for it**: the ladder's writes are the player's, and a player whose own
//! answers are sound has nothing wrong with it.
//!
//! **And a session muted with its slider at nothing is left exactly as it is**: that is a hand in
//! Windows' own mixer, and putting it back would overwrite a deliberate setting — so it is written down
//! (`session turned down … left alone`) and the film plays on as the user left it. The mute switch is
//! taken as this app's business and the slider as the user's, which is the whole of the line the cure
//! draws (see `audio_session`). **FFmpeg's player is given no look**: its ladder is a line on its own
//! stream and it has no read-back to hang one on, so what a traced run writes about its session stays a
//! line and nothing more (see `audio_session::watch_the_mixer`).
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
//! 3. neither heard, and the player says it is open — **the level this app last asked this player
//!    for asserted again**, with the player's own `mute` taken off beside it (`Verdict::Silent`): a
//!    track selected, an output open, and a player told to be quiet is a film that every other answer
//!    calls sound. The level itself is written by the one function that owns it (`mpv_ipc::set_volume`),
//!    so a player the launch gave no track to has it asked back for by the same line.
//!
//! **The level is the one that was last *asked* for, not the one the film began at** (see
//! `the_level_moved`), because a hand moves it while the film plays and a level dragged to nothing is
//! not a film to cure: mpv keeps the track at `volume 0`, so a pin silenced by hand reads exactly as
//! one something else silenced, and the launch's level put back over the hand is the bug this seam
//! closes. A level of nothing is the silence the setting chose, which the verdict already exempts.
//! 4. neither heard — **one** new player, with the arguments, the second and the pause the one it
//!    replaces had, and one read of that player. Then the ladder ends.
//!
//! **A ladder that ends is written once, and the last rung is the one written whole** (see
//! `a_give_up`): a film that plays with no sound and that nothing could cure is the case there will be
//! no reproducer for, so its ending carries the launch's form and the level it asked for, the six
//! answers of the read that gave up, and the verdict they come to — one line with the whole of what is
//! known about that player in it.
//!
//! **A hover's ladder is the in-place rungs and nothing else** (see `Road`): a hover is a film that
//! comes and goes, its player is the resident's, and the wait a fresh player costs is the whole of
//! what the resident exists to save — so a hover gives up where a pin begins a player again. A hover's
//! per-serve reads are the whole of what there is to know about it, and they are the whole of what
//! there is to know about any player: no film is put on a schedule.
//!
//! **Nothing here fights the film.** No pause is toggled, no seek is asked, and the only level written
//! is the one this app last asked the player for, put back on a player silenced under it — and
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
//! **A player that answered `heard` is asked nothing further, and the sparse schedule that used to ask
//! again is gone.** It sampled a pin's film at +2 s, +4 s and +8 s and then every ten seconds for as
//! long as it played, for a failure mode that has never once been observed: a player heard at its first
//! read that went silent later. Every muted film captured so far was **born** muted and caught by that
//! first read, so the schedule bought a case that has not happened with a poll for the whole length of
//! a film. **What reads a player now is the ladder and nothing else**: the one read [`AUDIO_SETTLING`]
//! after it comes up, the ask-again rung, and every read the rungs themselves cause — after `aid`,
//! after `audio-device`, after a level, and the read-back of a player begun in another's place. Each of
//! those still asks the session, and each still cures a muted one.
//!
//! **The one wait with more than one read behind it is the ask-again rung, and it is bounded by the
//! player's own readiness rather than by a schedule.** A player is not ready at a wall clock: a 4 GB
//! file off a slow disk can spend several seconds before mpv has created its audio client at all, and a
//! player that has said nothing yet is exactly the player nobody can ask anything of. So it is asked
//! again every [`AUDIO_SETTLING`] for as long as it says nothing, and not once more the moment it says
//! **anything** — heard, no track, no device or silent (see `ASKS_OF_A_SILENT_PLAYER`) — and the cap
//! that ends the asking is a give-up, which already carries everything known. A healthy player is still
//! the one read it always was.
//!
//! **The guard ends where a read says so, and the reads end with it.** A give-up and a session the
//! mixer would not unmute are written once and the player is let go of; `a_player_ended` clears the
//! guard outright, and a probe that wakes to a guard that is not its own reads no pipe and reports
//! nothing. The relaunch stays a player's own — **at most once in its whole life**, not once per
//! episode (see `SPENT`) — so a device that is genuinely gone is never begun again for.
//!
//! **The loop never waits, and nothing here runs on it.** A player that has only just been begun has
//! no pipe yet, and waiting for one is a wait no tick may pay (see `mpv_ipc::get_property_within`,
//! which is bounded per question and only ever called from a thread). So a thread asks and leaves the
//! answer where the tick finds it — the shape `playback_trace::watch_the_engine` and the log drain
//! already have — and the tick is what writes, decides and asks for the player.

use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::Duration;

use super::audio_session::Seen;
use super::playback_trace::note_playback_trace;
use super::playback_trace::{EngineAnswers, Verdict};
use super::video_launch::{AudioForm, Engine};

/// How long a player is given before its sound is read, and before each read of it after that.
///
/// **No earlier, because the answer would be the guard's own invention**: a player that has only just
/// been begun has not opened its device yet, and a verdict read off that gap is rung two invented by
/// the guard itself. It is the same wait the trace gives a launch for the same reason, and a little
/// longer — the trace is describing a film, and this is deciding what to do to one. It is also the
/// wait before every ask-again read, because the thing an unanswered player needs is exactly this:
/// more of the time it has not had yet (see `ASKS_OF_A_SILENT_PLAYER`).
const AUDIO_SETTLING: Duration = Duration::from_millis(500);

/// **How many times a player that has answered nothing at all is asked again, and why the asking is
/// bounded by the player rather than by a clock.** A player is not ready at a wall clock: a 4 GB file
/// on a slow disk can take several seconds before mpv has created its audio client, and a guard that
/// read such a player once, saw nothing, and gave up would never look at the session that client owns
/// — the one mute nothing else in this app can see. So a player that has said nothing is asked again
/// after another [`AUDIO_SETTLING`], and again, and the asking stops the moment it says **anything**:
/// heard, no track, no device or silent. A healthy player is therefore still the one read it always
/// was, and only a player that is not ready yet keeps being asked.
///
/// Sixty asks, each of them a settling before the read and whatever the six questions themselves take,
/// puts the last of them a little past half a minute after the player came up, and that is where the
/// ladder gives up on it (see `a_give_up`): a film whose player never says anything is one nothing here
/// can be done for, and the line it leaves carries everything that is known about it.
const ASKS_OF_A_SILENT_PLAYER: u8 = 60;

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
/// The form is the one rule both routes' arguments are built from (see
/// `video_launch::launched_without_audio`): only mpv at a level of zero is given no track at all, so
/// only that launch has no ladder to climb — FFmpeg's player keeps its track at every level and is
/// turned down at the mixer (see `audio_session`), which is a silence that can still be a fault.
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
    if super::video_launch::launched_without_audio(engine, volume) {
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
    /// A pin's own film: cured in place, and begun again once.
    Pin,
    /// A hover's film, on a cold launch or on the resident's serve: cured in place, and then left.
    Hover,
}

impl Road {
    /// Whether a player at the end of the ladder may be begun again (see `the_step`).
    fn relaunches(self) -> bool {
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
    /// **How many times a player that has answered nothing at all has been asked again** (see
    /// `ASKS_OF_A_SILENT_PLAYER`) — the read the arming began is not one of them: the counter the cap is
    /// read from, and the one thing that stops a player which never speaks from being asked for ever. A
    /// player that answers anything spends none of them; the ladder takes the rung that answer earned.
    asks: u8,
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
    asks: ASKS_OF_A_SILENT_PLAYER,
    opening: true,
    track: true,
    device: true,
    level: true,
    again: true,
};

/// What a player has said, and no tick has read yet.
enum Report {
    /// The six answers an mpv gave, and what the machine's own look at its session came to (see
    /// `playback_trace::ask_the_engine` and `audio_session::look_and_cure`).
    ///
    /// **The look is boxed because a traced one carries a session's whole identity and the sessions
    /// beside it** (see `audio_session::Seen`), which is several strings and a list against the six
    /// answers' own — and this slot is held under a lock between a probe and a tick, so what it costs to
    /// hold is a line rather than an allocation per read.
    Answers {
        answers: EngineAnswers,
        seen: Box<Seen>,
    },
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

/// What a tick has decided the guard ends with, and the read that decided it.
///
/// **The last rung is the one ending written down whole**, so it carries the answers it was reached
/// on: a film nothing could cure is the one case there will be no reproducer for, and the two words
/// `gave up` on their own say nothing about which road ended there (see `a_give_up`).
enum Finish {
    /// The player was heard: the line a healthy film's ladder ends on.
    Heard,
    /// **A player whose own answers are sound and whose session is still muted**: a hand at the mixer's
    /// slider, or a mute the mixer would not take off. There is nothing left to do to the player — the
    /// session's own line was written where it was found (see `Verdict::SessionMuted`) — so this is the
    /// end of the ladder, and it says which of the two it ended on rather than reading as sound.
    SessionMuted,
    /// Nothing left to try, and the six answers the last read gave — `None` for the player whose
    /// only answer was a line on its own stream.
    GaveUp(Option<EngineAnswers>),
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
    /// **The level this app last asked this player for**: the launch's to begin with, and then
    /// whatever a hand moved it to (see `the_level_moved`). It is what a `Silent` player is taken back
    /// up to, and it is the level the verdict is read against — a level of nothing being the silence
    /// the setting chose rather than a film to cure — and it is not the level the player reports, which
    /// is the fact the verdict was read from (see `playback_trace::EngineAnswers::told_to_be_quiet`).
    volume: u32,
    /// **What the launch gave this player, in the words the launch line writes it in**: `aid=no` or
    /// `-an`, or the level with the gain measured for the file folded into it (see
    /// `video_launch::AudioForm`). It is read once, at the arming, where the level and the gain are
    /// both in hand — so that the line a give-up writes says what was *asked* of the player rather
    /// than what a later read guesses at (see `a_give_up`).
    form: String,
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
}

impl Guard {
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
pub(super) fn arm(
    pid: u32,
    path: &Path,
    volume: u32,
    gain: Option<f64>,
    engine: Engine,
    for_a_hover: bool,
) {
    if !guarded(engine, for_a_hover, volume) {
        return;
    }

    let road = if for_a_hover { Road::Hover } else { Road::Pin };
    let guard = the_guard_a_launch_arms(pid, path, volume, gain, engine, road);
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
/// counted and no relaunch is waiting for a tick — only the rungs are handed on, and only ever within
/// one player's life (see `the_spend_for`). That is the whole of the reset a serve gets, and it is the
/// arming rather than a second step: a guard built here has nothing of any other file in it.
fn the_guard_a_launch_arms(
    pid: u32,
    path: &Path,
    volume: u32,
    gain: Option<f64>,
    engine: Engine,
    road: Road,
) -> Guard {
    Guard {
        pid,
        path: path.to_path_buf(),
        engine,
        road,
        volume,
        // The same call the launch's own line is built from, so that a give-up cannot say the
        // player was handed anything other than what it was handed (see `playback_trace::a_launch`).
        form: AudioForm::of(engine, volume, gain).word(engine),
        spend: the_spend_for(road, path),
        report: None,
        wants_again: false,
        reads: 0,
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
///
/// **A probe whose player is no longer this guard's asks nothing at all**: a thread that was asleep
/// when the film ended wakes to a guard that has been cleared (see `a_player_ended`) and reads no pipe
/// for a process that is not there. Nothing is retried: a read is only ever begun by the arming or by
/// a rung, and a guard that has let its player go begins neither.
fn spawn_probe(pid: u32) {
    let _ = std::thread::Builder::new()
        .name("audio-guard".to_string())
        .spawn(move || {
            std::thread::sleep(AUDIO_SETTLING);
            let Some((asked, engine)) = still_watched(pid) else {
                return;
            };
            let answers = super::playback_trace::ask_the_engine(pid);

            // **And the machine, on this same thread and straight after the player**. The session
            // Windows has for this pid is looked at here, at no other time and for no player that is
            // not this guard's — the read-back is the one cadence either of them has, so a looked-at
            // session costs a walk of the mixer per read and nothing else (see `audio_session`). The
            // look is made whether or not the trace is asked for, because the trace is the recorder
            // and the cure is the fix; what it came to goes back with the answers, so the tick reads
            // one report rather than two.
            let seen = super::audio_session::look_and_cure(pid, asked, engine);
            if let Some(line) = super::audio_session::a_cure_line(pid, &seen) {
                note_playback_trace!("{line}");
            }

            report(
                pid,
                Report::Answers {
                    answers,
                    seen: Box::new(seen),
                },
            );
        });
}

/// The level the guard is still holding this player to, and which player it is, where this player is
/// still the one it guards (see `spawn_probe`).
///
/// **The level comes back with the answer**, because it is the half of the session rule that says a
/// launch begun without audio is never cured — and so does the engine, because which launch that is
/// depends on the player: only mpv at zero is given no track (see
/// `video_launch::launched_without_audio`). A serve re-arms this guard for the same pid at the level
/// of its new file, so a probe that carried the level it was begun at would heal a silence the
/// setting chose (see `audio_session::the_warrant`).
fn still_watched(pid: u32) -> Option<(u32, Engine)> {
    GUARD
        .lock()
        .ok()?
        .as_ref()
        .filter(|guard| guard.pid == pid)
        .map(|guard| (guard.volume, guard.engine))
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

/// **The level this app has just asked a player for, where that player is the one this guard holds.**
/// The guard stands for what was last *asked* of the player rather than for what its launch began it
/// at, because a hand moves the level while the film plays — and a level moved to nothing is the half
/// of that the guard has to know about (see `Verdict::Silent`).
///
/// **A player dragged to `volume 0` keeps its track**, which is why the read is the one this exists
/// for: `mpv_ipc::set_volume` deliberately leaves the decoder alone, so a film a hand has taken down
/// to silence answers `volume 0` with a track, an output and params — indistinguishable, to a ladder
/// that still believes the launch's level, from a film something else silenced. Ten seconds later the
/// level rung would write the launch's level back over the hand. With the record moved here, that
/// read is the silence the setting asked for and no rung fires; and a level moved to 60 makes the
/// level rung assert 60 rather than the 40 the film began at.
///
/// **The roads call this, not `mpv_ipc::set_volume`**, and the difference is the loop not closing on
/// itself. The guard's own `AskForALevel` writes through that function with this guard's lock held
/// (see `tick`), so a seam inside it would be the guard feeding itself: a level it wrote recorded as a
/// level a hand had moved, taken while the lock it already holds is taken again. So a write the guard
/// makes moves nothing, and only the app's own writes do — where each of them landed, which is the
/// rule `PinVolume::playing_at` keeps beside it: a level no player took is still owed to the next one,
/// and the player behind the bar is still at the level the guard already holds.
///
/// Nothing is written about the move: every read-back already carries the level it is read against
/// (`asked`), and a drag is a hundred moves a second, so a line per move would bury the ladder.
pub(super) fn the_level_moved(pid: u32, level: u32) {
    let Ok(mut slot) = GUARD.lock() else {
        return;
    };
    if let Some(guard) = slot.as_mut().filter(|guard| guard.pid == pid) {
        guard.volume = level;
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
            let read = the_step_of_a_read(guard, report);
            match read.step {
                // **Heard is where the guard ends, and that player is read no more**: the one read, the
                // rung it earned and the read-back of that rung are the whole of what a player is
                // asked, so a film whose sound came up is done with (see the module note). A player
                // whose own answers are sound and whose session is still muted was not heard, and its
                // end says which of the two it is (see `Finish`).
                Step::Nothing => {
                    finish = Some(if read.verdict == Some(Verdict::SessionMuted) {
                        Finish::SessionMuted
                    } else {
                        Finish::Heard
                    })
                }
                Step::GiveUp => finish = Some(Finish::GaveUp(read.answers)),
                Step::AskAgain => {
                    // **Only a player that said nothing spends one of its asks** (see
                    // `ASKS_OF_A_SILENT_PLAYER`): the spare read a player still opening its output is
                    // given is the `opening` rung, spent where the read was read.
                    if read.verdict == Some(Verdict::Unanswered) {
                        guard.spend.asks = guard.spend.asks.saturating_add(1);
                    }
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
    // **A ladder that ends is written once, and the last rung is the one written whole** (see
    // `a_give_up`): a film nothing could cure leaves the one line a run without a reproducer is read
    // from, and every other ending keeps the word it always had.
    if let Some(finish) = &finish {
        match finish {
            Finish::Heard => note_playback_trace!("audio guard heard pid {pid}"),
            Finish::SessionMuted => note_playback_trace!("audio guard session-muted pid {pid}"),
            Finish::GaveUp(answers) => {
                note_playback_trace!("{}", a_give_up(guard, answers.as_ref()));
            }
        }
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

/// What one read comes to: the rung it earns, the verdict it was read as, and the answers an ending is
/// written from.
///
/// **The verdict comes back with the step** because the two ends of a ladder are not one word: a player
/// whose own answers are sound and whose Windows session is muted is not a player that was heard, and
/// an end-of-ladder line that said so would be the one lie this module could print (see `Finish`). The
/// answers come back beside it for the same kind of reason — a tick that reaches `GiveUp` has the very
/// read that reached it to write (see `a_give_up`), and the player with no pipe hands none back.
struct AStep {
    step: Step,
    /// What the read was read as, and `None` for a player that answered nothing this app can read.
    verdict: Option<Verdict>,
    answers: Option<EngineAnswers>,
}

/// What one report comes to, with the line the read-back is written as.
fn the_step_of_a_read(guard: &mut Guard, report: Report) -> AStep {
    match (guard.engine, report) {
        (Engine::Mpv, Report::Answers { answers, seen }) => {
            // **The machine's own answer is folded into the player's**, which is the whole of the fourth
            // road: a player that is heard and a session that is muted is `session-muted`, and one
            // whose mute the look took off is heard (see `Seen::the_verdict_of`).
            let verdict = seen.the_verdict_of(answers.verdict(guard.volume));
            // **Every read is written, because a player is only ever read once past its first**: there
            // is no second read of a film that was heard for a line to be held back from (see the
            // module note). It is written before the step is taken either way, so that a run reads what
            // the player said against what was then done about it even where the step itself is what
            // fails.
            note_playback_trace!(
                "{}",
                a_read_back(guard.pid, guard.reads, guard.volume, &answers, &seen)
            );

            let opening = answers.a_driver_with_no_params();
            let step = the_step(verdict, guard.spend, opening, guard.road);
            // The second look a player still opening its output is given is this read's own rung, and
            // is spent here rather than in the tick, which cannot tell one second look from the other.
            // An `Unanswered` read that also names a driver spends only the other one: nothing was
            // learned about its track, so nothing is known about its output either.
            if step == Step::AskAgain && verdict == Verdict::NoDevice && opening {
                guard.spend.opening = true;
            }
            AStep {
                step,
                verdict: Some(verdict),
                answers: Some(answers),
            }
        }
        (Engine::Ffplay, Report::Spoke) => {
            note_playback_trace!("audio guard spoke pid {} read {}", guard.pid, guard.reads);
            AStep {
                step: the_step_of_a_signature(guard.spend),
                verdict: None,
                answers: None,
            }
        }
        // An answer one player gives and the other does not: nothing is known, so nothing is done.
        _ => AStep {
            step: Step::Nothing,
            verdict: None,
            answers: None,
        },
    }
}

/// The line one read-back is written as: the level the launch asked for, the six properties the player
/// reported, what they come to, and the machine's own answer beside them — the trace's own `engine truth`
/// line with this ladder's numbering on it, because a film watched by the guard is read per rung rather
/// than once.
///
/// **`asked` is the launch's level and `volume` the player's own**, the two words the third road is read
/// by: a player reporting 0 or `mute yes` on a launch that asked for 100 is a film that is open and was
/// told to be quiet. **`audio-device` is the device that was asked for and `current-ao` the driver that
/// opened**, so the line says both halves of what the cure is about: a player asked for `auto` and
/// playing through `null` is a different fault from one asked for a device it was given. Both are
/// written as `-` where a player answered nothing, which is how a build without the property has always
/// read.
///
/// **The last two words are the machine's, and they are what tells the four silences apart on one line**
/// (see `audio_session::the_session_word`): `verdict silent` is the player's own `volume 0` or its own
/// `mute yes`, `verdict session-muted` with `session mixer 1.000 mute yes` is Windows silencing a player
/// that reports itself unmuted, `verdict no-device` is a device that never opened, and `session mixer
/// 0.000 mute yes left-alone` is a hand on the mixer's own slider, which nothing here touches.
fn a_read_back(pid: u32, read: u8, asked: u32, answers: &EngineAnswers, seen: &Seen) -> String {
    format!(
        "audio read-back pid {pid} read {read} asked {asked} volume {} mute {} aid {} \
         audio-device {} current-ao {} audio-out-params {} verdict {} {}",
        super::playback_trace::the_level_word(answers.volume),
        super::playback_trace::the_flag_word(answers.mute),
        answers.aid.as_deref().unwrap_or("-"),
        answers.audio_device.as_deref().unwrap_or("-"),
        answers.current_ao.as_deref().unwrap_or("-"),
        answers.audio_out_params.as_deref().unwrap_or("-"),
        super::playback_trace::the_verdict_word(seen.the_verdict_of(answers.verdict(asked))),
        super::audio_session::the_session_word(seen),
    )
}

/// **The one line a ladder that gave up leaves behind: everything known about the player**, in a single
/// line, because a film that plays silently and that nothing could cure is the case there will be no
/// reproducer for.
///
/// What it carries is the whole of what was asked and the whole of what was answered: the launch's form
/// and the level it asked for (the form is the launch's own word for what the player was handed, see
/// `video_launch::AudioForm`), the six properties the player reported — `aid`, `audio-device`,
/// `current-ao`, `audio-out-params`, `volume` and `mute`, each `-` where it said nothing, which is how a
/// build without the property has always read — and the verdict they come to beside them. **The engine
/// with no pipe answers none of the six**: FFmpeg's player can only be read off its own stream, so its
/// line says the form, the level and the one thing it did say — the failure it wrote, which is `spoke`
/// (see `AUDIO_FAILURES`).
///
/// It is built from the answers rather than from a player, so a test can pin the shape of the one line a
/// field log may have nothing else to say about.
fn a_give_up(guard: &Guard, answers: Option<&EngineAnswers>) -> String {
    let verdict = match answers {
        Some(answers) => super::playback_trace::the_verdict_word(answers.verdict(guard.volume)),
        None => "spoke",
    };
    let answered = |answer: Option<&str>| answer.unwrap_or("-").to_string();
    format!(
        "audio gave up pid {} engine {} road {} reads {} form {} asked {} aid {} \
         audio-device {} current-ao {} audio-out-params {} volume {} mute {} verdict {verdict}",
        guard.pid,
        guard.engine.word(),
        guard.road.word(),
        guard.reads,
        guard.form,
        guard.volume,
        answered(answers.and_then(|answers| answers.aid.as_deref())),
        answered(answers.and_then(|answers| answers.audio_device.as_deref())),
        answered(answers.and_then(|answers| answers.current_ao.as_deref())),
        answered(answers.and_then(|answers| answers.audio_out_params.as_deref())),
        super::playback_trace::the_level_word(answers.and_then(|answers| answers.volume)),
        super::playback_trace::the_flag_word(answers.and_then(|answers| answers.mute)),
    )
}

/// **The ladder's decisions, and nothing else** — what a read-back asks for next, out of the rungs
/// this guard has left.
///
/// A film that was heard ends it wherever it is found. A player that answered **nothing at all** is
/// asked again, and again, until it says anything or the cap runs out (see
/// `ASKS_OF_A_SILENT_PLAYER`): nothing was learned about its sound, so nothing may be done to it. A
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
///
/// **`SessionMuted` asks for nothing at all**, and that is the whole of its rung: the player's own
/// answers are sound, so a track, a device or a level asserted again would be three writes to a film
/// that is not what is muted. The mute was taken off the session where the read was made, or left
/// where a hand put it (see `audio_session::look_and_cure`), and that read is the ladder's end.
fn the_step(verdict: Verdict, spend: Spent, opening: bool, road: Road) -> Step {
    match verdict {
        Verdict::Heard | Verdict::SessionMuted => Step::Nothing,
        Verdict::Unanswered if spend.asks < ASKS_OF_A_SILENT_PLAYER => Step::AskAgain,
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
    use super::super::audio_session::Cure;
    use super::super::tests::the_room_to_itself;
    use super::*;

    /// A player with a track, an output and its params, reporting `volume` from `mute` off — the shape
    /// every one of the three silences shares, which is why the level it reports is what tells them
    /// apart.
    fn a_player_at(volume: f64) -> EngineAnswers {
        EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            volume: Some(volume),
            mute: Some(false),
        }
    }

    /// Arm the one guard the suite has for a player of a test's own, at the level its launch asked for.
    ///
    /// **The slot rather than a guard in hand**, because the level a hand moves is a fact about the
    /// player this app is watching and not about a value passed down a call: `the_level_moved` is the
    /// only way in, and it answers only for the pid it holds (see `GUARD`). Each test is given a pid of
    /// its own so that a probe thread of one cannot report into the guard of another.
    fn arm_the_slot(pid: u32, level: u32) {
        let guard = the_guard_a_launch_arms(pid, &a_film(), level, None, Engine::Mpv, Road::Pin);
        if let Ok(mut slot) = GUARD.lock() {
            *slot = Some(guard);
        }
    }

    /// A file named as a film is, for the rules that turn on which one is being shown.
    fn a_film() -> PathBuf {
        PathBuf::from("a-film.mkv")
    }

    /// What a tick is handed when the player's answers are all there is to know: a look that found no
    /// session says nothing about a mute, so the read is the player's own (see
    /// `Seen::the_verdict_of`).
    fn a_report(answers: EngineAnswers) -> Report {
        Report::Answers {
            answers,
            seen: Box::new(Seen::Absent),
        }
    }

    /// The same read with a muted session, and the mute the mixer would not take off beside it — the
    /// fourth road as a tick sees it when the cure failed.
    fn a_muted_report(answers: EngineAnswers, cure: Cure) -> Report {
        Report::Answers {
            answers,
            seen: Box::new(Seen::there(Some(1.0), Some(true), cure)),
        }
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
            form: "volume=40 gain=none".to_string(),
            spend,
            report: None,
            wants_again: false,
            reads: 0,
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

    /// **The arming is the reset.** A guard built for a serve counts no read and has no relaunch
    /// waiting for a tick — whatever the guard it replaces had done with the file before it.
    #[test]
    fn a_serve_arms_a_ladder_with_no_read_of_the_one_before_it() {
        let guard = the_guard_a_launch_arms(7, &a_film(), 100, None, Engine::Mpv, Road::Hover);

        assert_eq!(guard.reads, 0, "the episode's first read is its first");
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
    fn a_player_a_pin_took_over_gets_the_pins_rungs() {
        let mut guard = a_guard(Engine::Mpv, Spent::default());
        guard.road = Road::Hover;

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

    /// **The fourth silence asks nothing of the player.** A session muted under a player whose every
    /// answer is sound is a film the ladder has no rung for: a track, a device and a level are the three
    /// things that player already has, and what is muted is the machine's session rather than the
    /// player. So the rung is none on either road and with every rung spent — the cure is the session's
    /// own, and it was taken where the read was made (see `audio_session::look_and_cure`).
    #[test]
    fn a_session_muted_under_a_sound_player_asks_nothing_of_the_player() {
        assert_eq!(
            the_step(Verdict::SessionMuted, Spent::default(), false, Road::Pin),
            Step::Nothing,
            "a track, a device and a level are what this player has: nothing here is its fault"
        );
        assert_eq!(
            the_step(Verdict::SessionMuted, Spent::default(), false, Road::Hover),
            Step::Nothing
        );
        assert_eq!(
            the_step(Verdict::SessionMuted, SPENT, false, Road::Pin),
            Step::Nothing,
            "and a player begun in place of another is not asked for a rung either"
        );

        // Which is what a tick sees: the read is written, the player is not written to, and the ladder's
        // end says the mute is the machine's rather than reading it as sound.
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
        guard.reads = 2;

        clear_noted();
        let read = the_step_of_a_read(&mut guard, a_muted_report(heard, Cure::Refused));
        assert_eq!(read.step, Step::Nothing);
        let wrote = noted_lines();
        assert_eq!(
            wrote.len(),
            1,
            "the read that found the mute is written: {}",
            wrote.join(" | ")
        );
        assert!(wrote[0].contains("verdict session-muted"));
        assert!(wrote[0].contains("session mixer 1.000 mute yes cure-failed"));
        assert_eq!(
            guard.spend,
            Spent::default(),
            "not one rung is spent on a fault that is not the player's"
        );
    }

    /// **A player that answers nothing is asked again, and the asking is bounded by the player's own
    /// readiness and not by a clock.** A 4 GB file off a slow disk can spend several seconds before
    /// mpv has created its audio client at all, and a player that has said nothing yet is the one player
    /// nobody can ask anything of — so it is asked again every [`AUDIO_SETTLING`] for as long as it says
    /// nothing, and not once more the moment it says **anything**.
    ///
    /// A player that has said nothing is never begun again, and the player a spent ladder began is given
    /// no look past its first: the one verdict the guard may not act on is the one it never got.
    #[test]
    fn a_player_that_says_nothing_is_asked_until_it_answers_or_the_cap_is_reached() {
        assert_eq!(
            AUDIO_SETTLING * ASKS_OF_A_SILENT_PLAYER as u32,
            Duration::from_secs(30),
            "sixty waits of one settling each, before the patience each read is offered: a player that \
             has said nothing is given about half a minute to say anything at all"
        );

        let mut spend = Spent::default();
        for ask in 1..=ASKS_OF_A_SILENT_PLAYER {
            assert_eq!(
                the_step(Verdict::Unanswered, spend, false, Road::Pin),
                Step::AskAgain,
                "ask {ask} of a player that has said nothing is one more it is owed"
            );
            spend.asks = spend.asks.saturating_add(1);
        }
        assert_eq!(
            the_step(Verdict::Unanswered, spend, false, Road::Pin),
            Step::GiveUp,
            "and the cap ends it: {ASKS_OF_A_SILENT_PLAYER} asks in, a player that has said nothing is a \
             film nothing here can be done for"
        );
        assert_eq!(
            the_step(Verdict::Unanswered, SPENT, false, Road::Pin),
            Step::GiveUp,
            "and a player begun in place of another is given no look at all"
        );

        // **An answer of any kind stops the asking**, whatever it is: the ladder takes the rung that
        // answer earned and the asks it never needed are left unspent.
        let asked_twice = Spent {
            asks: 2,
            ..Spent::default()
        };
        assert_eq!(
            the_step(Verdict::Heard, asked_twice, false, Road::Pin),
            Step::Nothing
        );
        assert_eq!(
            the_step(Verdict::SessionMuted, asked_twice, false, Road::Pin),
            Step::Nothing
        );
        assert_eq!(
            the_step(Verdict::NoTrack, asked_twice, false, Road::Pin),
            Step::AskForATrack
        );
        assert_eq!(
            the_step(Verdict::NoDevice, asked_twice, false, Road::Pin),
            Step::AskForADevice
        );
        assert_eq!(
            the_step(Verdict::Silent, asked_twice, false, Road::Pin),
            Step::AskForALevel
        );
        assert_eq!(
            asked_twice.asks, 2,
            "and the asks it did not need are still unspent, so a later silence is not the cap's"
        );
    }

    /// **A player that answered `heard` is asked nothing further.** The one read after it comes up,
    /// the asks a player that has said nothing is owed, and the read-backs of the rungs it earns are
    /// the whole of what the guard asks a film — so a read that comes back heard is the end of it: the
    /// guard is cleared, no thread is left asleep behind it, and a probe that reports afterwards
    /// reports to nothing.
    #[test]
    fn a_player_that_answered_heard_is_asked_nothing_further() {
        use super::super::playback_trace::{clear_noted, noted_lines};

        let _room = the_room_to_itself();
        arm_the_slot(12, 100);

        clear_noted();
        assert!(report(12, a_report(a_player_at(100.0))));
        assert_eq!(tick(false), None, "a heard film has no rung to take");
        assert_eq!(
            still_watched(12),
            None,
            "and the guard is gone: nothing stands armed to read that player again"
        );
        assert!(
            !report(12, a_report(a_player_at(100.0))),
            "a read that arrives after it has no guard to report to, and is dropped where it lands"
        );
        assert!(
            noted_lines()
                .iter()
                .any(|line| line == "audio guard heard pid 12"),
            "the end of the ladder is the one line it leaves: {:?}",
            noted_lines()
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
    /// named as the trace names them, the verdict they come to, and the machine's own answer beside them
    /// — so a field log carries the whole of the ladder's evidence: the level asked for beside the level
    /// the player is at, the device that was asked for beside the driver that answered it, and `-` where
    /// a player said nothing to either.
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
            a_read_back(7, 1, 100, &answers, &Seen::Absent),
            "audio read-back pid 7 read 1 asked 100 volume 100 mute no aid 1 audio-device auto \
             current-ao wasapi audio-out-params {\"samplerate\":48000} verdict heard session none"
        );

        // And the third road, as the read-back writes it: a player at nothing on a launch that asked for
        // 100 is the film the level rung exists for.
        let quiet = EngineAnswers {
            volume: Some(0.0),
            ..answers
        };
        assert_eq!(
            a_read_back(7, 2, 100, &quiet, &Seen::Absent),
            "audio read-back pid 7 read 2 asked 100 volume 0 mute no aid 1 audio-device auto \
             current-ao wasapi audio-out-params {\"samplerate\":48000} verdict silent session none"
        );

        let silent = EngineAnswers::default();
        assert_eq!(
            a_read_back(9, 2, 40, &silent, &Seen::Absent),
            "audio read-back pid 9 read 2 asked 40 volume - mute - aid - audio-device - current-ao - \
             audio-out-params - verdict unanswered session none",
            "a player that said nothing is written as nothing said rather than as a guess"
        );
    }

    /// **The four silences are four words on one line, and none of them can be read as another.** A
    /// player at nothing is `silent`; a player with a device it could not open is `no-device`; a player
    /// whose every answer is sound and whose Windows session is muted is `session-muted`, with the
    /// mixer's own slider and switch beside it; and a session muted by a hand, with its slider down, is
    /// written as left alone rather than as this app's failure.
    #[test]
    fn a_read_back_names_which_of_the_four_silences_it_found() {
        let heard = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            volume: Some(100.0),
            mute: Some(false),
        };

        let muted = Seen::there(Some(1.0), Some(true), Cure::Refused);
        assert_eq!(
            a_read_back(23724, 4, 100, &heard, &muted),
            "audio read-back pid 23724 read 4 asked 100 volume 100 mute no aid 1 audio-device auto \
             current-ao wasapi audio-out-params {\"samplerate\":48000} verdict session-muted \
             session mixer 1.000 mute yes cure-failed",
            "the second mute switch: every answer on the player's pipe says the film has its sound"
        );

        let turned_down = Seen::there(Some(0.0), Some(true), Cure::TurnedDown);
        assert_eq!(
            a_read_back(23724, 5, 100, &heard, &turned_down),
            "audio read-back pid 23724 read 5 asked 100 volume 100 mute no aid 1 audio-device auto \
             current-ao wasapi audio-out-params {\"samplerate\":48000} verdict session-muted \
             session mixer 0.000 mute yes left-alone",
            "a hand on the mixer's own slider is written down and not restored"
        );

        // And a session whose mute the cure took off is a film that has its sound: the write is the
        // whole of what was wrong with it.
        let cured = Seen::there(Some(1.0), Some(true), Cure::Unmuted);
        assert!(
            a_read_back(23724, 6, 100, &heard, &cured)
                .ends_with("verdict heard session mixer 1.000 mute yes cured"),
            "a mute taken off is heard, and the line still says what it was found at"
        );
    }

    /// **The one line a ladder that gave up leaves behind**: everything that was asked of the player and
    /// everything it answered, in a single line — the film nothing could cure is the case a field log is
    /// read from with no reproducer behind it, so the post-mortem has to be whole. The player with no
    /// pipe answers none of the six, and says so as `-` rather than by guessing.
    #[test]
    fn a_give_up_carries_the_whole_of_what_is_known_about_the_player() {
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("null".to_string()),
            audio_out_params: None,
            volume: Some(40.0),
            mute: Some(false),
        };
        let mut guard = a_guard(Engine::Mpv, SPENT);
        guard.reads = 3;
        assert_eq!(
            a_give_up(&guard, Some(&answers)),
            "audio gave up pid 7 engine mpv road pin reads 3 form volume=40 gain=none asked 40 \
             aid 1 audio-device auto current-ao null audio-out-params - volume 40 mute no verdict \
             no-device",
            "the form the launch gave it, the six properties the player reported and the verdict \
             they come to, on the one line a run without a reproducer is read from"
        );

        // FFmpeg's player has no pipe to answer any of the six down: what it said is a line on its own
        // stream, and the line names that rather than inventing answers.
        let mut ffplay = a_guard(Engine::Ffplay, SPENT);
        ffplay.form = "-an".to_string();
        ffplay.reads = 1;
        assert_eq!(
            a_give_up(&ffplay, None),
            "audio gave up pid 7 engine ffplay road pin reads 1 form -an asked 40 aid - \
             audio-device - current-ao - audio-out-params - volume - mute - verdict spoke"
        );
    }

    /// **The answers a give-up is written from are the ones the last rung was reached on**: the read
    /// hands them back beside its step, and the player with no pipe hands back nothing (see `a_give_up`).
    #[test]
    fn the_last_rung_hands_back_the_read_that_reached_it() {
        let mut mpv = a_guard(
            Engine::Mpv,
            Spent {
                asks: ASKS_OF_A_SILENT_PLAYER,
                ..Spent::default()
            },
        );
        let read = the_step_of_a_read(&mut mpv, a_report(EngineAnswers::default()));
        assert_eq!(read.step, Step::GiveUp);
        assert!(
            read.answers.is_some(),
            "an mpv that is given up on hands back the read it was given up on"
        );

        let mut ffplay = a_guard(Engine::Ffplay, SPENT);
        let read = the_step_of_a_read(&mut ffplay, Report::Spoke);
        assert_eq!(read.step, Step::GiveUp);
        assert!(
            read.answers.is_none(),
            "FFmpeg's player answers nothing this app can read, so there is nothing to hand back"
        );
    }

    /// **A player dragged to `volume 0` is not a film something else silenced.** mpv keeps the track
    /// when it is told a level of nothing, so a pin a hand has taken to the bottom of its knob answers
    /// with a track, an output and its params at `volume 0` — the read the `Silent` rule was written
    /// for, and the one that used to bring the launch's level back over the hand a read later.
    /// With the level moved, that read is the silence the setting chose (`mpv_launched_without_audio`),
    /// so no rung is taken, nothing is written to the player at all, and the guard lets it go.
    #[test]
    fn a_level_moved_to_nothing_is_not_a_player_that_was_silenced() {
        use super::super::mpv_ipc::{clear_noted, noted_lines};

        let _room = the_room_to_itself();
        arm_the_slot(7, 40);

        // The hand goes to the bottom of the knob and it goes there exactly: `volume 0`, and the
        // track left standing behind it (see `mpv_ipc::set_volume`).
        the_level_moved(7, 0);
        assert_eq!(
            still_watched(7),
            Some((0, Engine::Mpv)),
            "the guard stands for the level this app last asked the player for"
        );

        clear_noted();
        assert!(report(7, a_report(a_player_at(0.0))));
        assert_eq!(tick(true), None);
        assert_eq!(
            still_watched(7),
            None,
            "the read is the verdict of a launch that asked for no audio, so there was no rung to take \
             and the ladder ends there rather than reading the player again"
        );
        assert_eq!(
            noted_lines(),
            Vec::<String>::new(),
            "and not one level was written over the hand that put it at nothing"
        );
    }

    /// **A level moved after the launch is the level a `Silent` player is taken to**: the ladder's third
    /// rung asserts what this app last asked the player for, so a film moved to 60 and silenced under
    /// that is brought back to 60 rather than to the 40 its launch began at.
    #[test]
    fn a_level_moved_after_the_launch_is_what_a_silent_player_is_taken_to() {
        use super::super::mpv_ipc::{clear_noted, forget_pipe, noted_lines, remember_pipe};

        let _room = the_room_to_itself();
        arm_the_slot(8, 40);
        remember_pipe(8, "a-guarded-pipe", false);

        the_level_moved(8, 60);
        clear_noted();
        assert!(report(8, a_report(a_player_at(0.0))));
        assert_eq!(tick(false), None);
        assert_eq!(
            noted_lines(),
            vec![
                format!(
                    "{}\n",
                    r#"{"command":["set_property","volume",60],"async":true}"#
                ),
                format!(
                    "{}\n",
                    r#"{"command":["set_property","mute",false],"async":true}"#
                ),
            ],
            "the level the film was moved to, and the player's own mute taken off beside it — not the \
             40 the launch began at"
        );

        forget_pipe(8);
        a_player_ended(8);
    }

    /// **Another player's level is not this guard's to hold**: a level moved on a pid the guard is not
    /// armed for changes nothing, and the read it does hold is still the read it always was.
    #[test]
    fn a_level_moved_on_a_player_the_guard_does_not_hold_changes_nothing() {
        let _room = the_room_to_itself();
        arm_the_slot(9, 40);

        the_level_moved(10, 0);
        assert_eq!(
            still_watched(9),
            Some((40, Engine::Mpv)),
            "the guard is armed for 9, so a level moved on 10 is none of its business"
        );
        assert_eq!(
            still_watched(10),
            None,
            "and no guard was made for the player nothing armed"
        );

        let Ok(mut slot) = GUARD.lock() else {
            return;
        };
        let guard = slot
            .as_mut()
            .expect("the guard is still armed for its own pid");
        assert_eq!(
            the_step_of_a_read(guard, a_report(a_player_at(0.0))).step,
            Step::AskForALevel,
            "a silent read on the pid the guard does hold is still the film it was, at the level it \
             was begun at"
        );
        drop(slot);
        a_player_ended(9);
    }

    /// **A level the guard writes is not a hand**, which is where the seam sits: the level rung writes
    /// through `mpv_ipc::set_volume`, and a seam inside that function would be the guard feeding its
    /// own record — on the lock it already holds (see `the_level_moved`). So the record stays where the
    /// hand left it, whatever the ladder writes.
    #[test]
    fn the_guards_own_writes_do_not_move_the_level_it_holds() {
        use super::super::mpv_ipc::{forget_pipe, remember_pipe};

        let _room = the_room_to_itself();
        arm_the_slot(11, 40);
        remember_pipe(11, "a-guarded-pipe", false);

        assert!(super::super::mpv_ipc::set_volume(11, 30));
        assert_eq!(
            still_watched(11),
            Some((40, Engine::Mpv)),
            "the level rung's own write is not a level a hand moved"
        );

        assert!(report(11, a_report(a_player_at(0.0))));
        assert_eq!(tick(false), None);
        assert_eq!(
            still_watched(11),
            Some((40, Engine::Mpv)),
            "and a rung taken on a silent read leaves the record where it was"
        );

        forget_pipe(11);
        a_player_ended(11);
    }
}

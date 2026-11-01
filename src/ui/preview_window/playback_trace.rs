//! A film that plays with no sound, written down when `RHP_PLAYBACK_TRACE` asks for it: the level a
//! launch resolved to and the form it gave the player, what the player itself says it did with the
//! sound, what a player wrote on its own error stream, and whether a kill was confirmed gone.
//!
//! **Why a line is needed at all.** A pinned film that plays perfectly and has no audio — and no
//! entry in Windows' Volume Mixer at all, which is a process that never made an audio session — has
//! two roads to it that look identical from this side of the process boundary:
//!
//! - **The launch gave the player no track.** A level of nothing is `--aid=no` rather than
//!   `--volume=0` (see `video_launch::AudioForm`), and a player with no track selected never asks
//!   the device for anything, so there is no session to list.
//! - **The launch gave it a track and the device would not open.** mpv says so in its own log
//!   (`Could not open/initialize audio device -> no sound`) and plays on without it; FFmpeg's player
//!   does the same on its own stream. Neither is anything this app can see.
//!
//! The first is a fact this app decided, so the launch line records it. The second is the player's,
//! so it is read off the player: mpv's `aid` says whether a track is selected and `current-ao` and
//! `audio-out-params` say whether an output was opened from it — `current-ao` being the driver in use
//! and not the `ao` option this app never passes (see `EngineAnswers::an_output_opened`) — and the
//! pair is the discriminator: `aid=no` is the first road, `aid` a track with no output is the second
//! (see `EngineAnswers::verdict`).
//!
//! **And there is a third road, which those two answers cannot see at all**: a player that has its
//! track selected and its output open, and that was told to be quiet — its own `volume` at nothing, or
//! its own `mute` on. Every other answer a player gives about a film like that says it has sound, so
//! `volume` and `mute` are read beside them and the level the launch asked for is written with a
//! second word: `asked 100 volume 0` is a film this app gave a level to and a player that is not at
//! it (see `Verdict::Silent`).
//!
//! **And a fourth, which the player's pipe cannot reach either**: the same film with every one of those
//! answers sound and its Windows audio session muted — `volume 94 mute no` on the player and
//! `mixer 1.000 mute yes` on the session the OS keeps for it. Two mute switches, and the player can only
//! see its own; a captured run had two of eighteen players in it. The session is asked by the other
//! half of this trace (see `audio_session`), and a read where the player is heard and the session is
//! muted is `Verdict::SessionMuted`.
//!
//! **Nothing here runs unless the variable is set**, and what it writes goes to
//! `%TEMP%\rhp-playback-trace.log` — the same shape `RHP_PIN_TRACE` has (see `pin_trace`), a file of
//! its own because the question is a different one and a run that asks it should not have to read a
//! pin's road to answer it. A run without the variable builds no line at all and spawns no thread.
//! **The file is the run's own**: its first line makes it over, so what a session is read for is what
//! that session did rather than what every run before it left behind (see `write`).
//!
//! **The engine's answers are read on a thread of their own, never on the preview loop.** A player
//! that has just been begun has no pipe yet (mpv makes one some tens of milliseconds into the start,
//! measured at 40 ms warm and 211 ms cold), and waiting for one is a wait no tick may pay (see
//! `mpv_ipc::get_property_within`, which is bounded per question and only ever called from here).

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Once;
use std::time::Duration;

use super::video_launch::{AudioForm, Engine};
use super::MediaType;

static READ: Once = Once::new();
static ASKED: AtomicBool = AtomicBool::new(false);

/// Whether `RHP_PLAYBACK_TRACE` asks for a launch's sound to be written out: presence is enough, as
/// it is for the pin's own trace.
pub(super) fn asked() -> bool {
    READ.call_once(|| {
        ASKED.store(
            std::env::var_os("RHP_PLAYBACK_TRACE").is_some(),
            Ordering::Relaxed,
        );
    });
    ASKED.load(Ordering::Relaxed)
}

/// Whether a line is going anywhere at all: the log the variable names, or — in a test build — the
/// seam that reads a line back out of a run (see `NOTED`).
///
/// It is what a road asks *before* it works out what a line says, so that a run without the variable
/// pays a read of an atomic and nothing else (see `note_playback_trace`).
pub(super) fn wanted() -> bool {
    asked() || cfg!(test)
}

/// The file a run writes its lines to: `%TEMP%\rhp-playback-trace.log`.
fn the_log_file() -> PathBuf {
    std::env::temp_dir().join("rhp-playback-trace.log")
}

/// Whether this process has made its log yet: the first line of a run makes the file afresh, and
/// every later line appends to it (see `write_to`).
static STARTED: Once = Once::new();

/// Write one timestamped line to `%TEMP%\rhp-playback-trace.log`, and flush it so a run that ends
/// badly leaves what it wrote.
///
/// **The first line of a process makes the file afresh, and every later one appends to it.** A log
/// that only ever appended would be a pile of every run there has ever been: a line from yesterday's
/// launch beside today's is a line a reader cannot tell apart from this run's, and the question the
/// trace answers is about one run. A second process — a second run — starts the file over, which is
/// the rule `pin_trace` keeps.
pub(super) fn write(line: &str) {
    write_to(&the_log_file(), &STARTED, line);
}

/// The write itself, with the file and the run's own latch passed in so that a test can watch a fresh
/// start without a second process (see `STARTED`).
fn write_to(path: &Path, started: &Once, line: &str) {
    let stamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|elapsed| format!("{}.{:03}", elapsed.as_secs(), elapsed.subsec_millis()))
        .unwrap_or_else(|_| "-".to_string());
    // Truncating once, under the latch, is what keeps two threads' first lines from each making the
    // file over: every line after it appends.
    started.call_once(|| {
        let _ = std::fs::File::create(path);
    });
    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        use std::io::Write;
        let _ = writeln!(file, "{stamp} {line}");
        let _ = file.flush();
    }
}

/// One line a player wrote on its own stream — its log or its errors — under the player it came
/// from, where the trace asks for it.
///
/// The two streams are the two halves of the same evidence: mpv's log is where a device that would
/// not open says so, and FFmpeg's player says everything it has to say on the stream this app used to
/// throw away (see `player_logs`).
pub(super) fn a_players_line(pid: u32, stream: &str, line: &str) {
    if !asked() {
        return;
    }
    write(&format!("{stream} pid {pid} {line}"));
}

/// The kind of file a launch was for, in one word a line can be read by.
fn the_kinds_word(kind: Option<MediaType>) -> &'static str {
    match kind {
        Some(MediaType::Video) => "video",
        Some(MediaType::NativeVideo) => "native-video",
        Some(MediaType::Audio) => "audio",
        Some(_) => "other",
        None => "unstated",
    }
}

/// The line one launch begins with: which player, which kind of file, the level the launch resolved
/// to, and the form that level was given the player in (see `AudioForm`).
///
/// **It is the launch's own decision and not the player's answer**, which is why it is written before
/// either route begins and why the answer is a second line: the two suspected silences differ
/// *between* them, and a single line could not say which of the two a run took. The file is named
/// rather than pathed, as every line of the pin's own trace names one (see `pin_trace::path_name`).
pub(super) fn a_launch(
    engine: Engine,
    kind: Option<MediaType>,
    path: &Path,
    volume: u32,
    gain: Option<f64>,
    pinned: bool,
) -> String {
    format!(
        "launch engine {} kind {} file {} volume {} form {} pinned {}",
        engine.word(),
        the_kinds_word(kind),
        super::pin_trace::path_name(path),
        volume,
        AudioForm::of(engine, volume, gain).word(engine),
        pinned as u8,
    )
}

/// What the player's own answers come to — **the discriminator this whole module exists for**.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum Verdict {
    /// No track selected at all, so no session could ever have been asked for: the first road, and
    /// the one the launch line has already recorded.
    NoTrack,
    /// A track is selected and no output was opened from it: the device, rather than the launch —
    /// the second road, and the one that leaves nothing in the mixer.
    NoDevice,
    /// A track and an output both: the film had audio when this was read.
    Heard,
    /// **A track and an output both, and a player told to be quiet**: its own `volume` at nothing or
    /// its own `mute` on, while the launch asked for a level. The third road to a silent film, and
    /// the one every other answer calls sound — an output that is open, params that are present, and
    /// nothing coming out of it (see `EngineAnswers::told_to_be_quiet`).
    Silent,
    /// **A player whose own answers all say the film has its sound and whose Windows audio session is
    /// muted**: `mixer 1.000 mute yes` on the session's own `ISimpleAudioVolume` under an engine that
    /// reports `volume 94 mute no` — two mute switches, and this is the half the player cannot see.
    /// The fourth road to a silent film, the one a captured run proved (two players of eighteen), and
    /// the only one whose cure is not written to the player at all: the mute is taken off the session
    /// (see `audio_session::look_and_cure`, which is where a read is given this verdict).
    SessionMuted,
    /// The player answered nothing in the time it was given, which says nothing about its sound.
    Unanswered,
}

/// The six answers a started mpv is asked for, and nothing where it said nothing to one of them.
///
/// **Four are what the player did with a film's sound and two are the level it is at**, and the second
/// pair is the player's own level rather than the level the launch resolved to: the two are different
/// facts, are written with two different words on every line, and are the whole of the third road to a
/// silent film (see `Verdict::Silent`).
#[derive(Clone, Default, Debug)]
pub(super) struct EngineAnswers {
    /// `aid`: `no`, `auto`, or the track that was picked.
    pub(super) aid: Option<String>,
    /// `audio-device`: **the device the player was asked for**, which is the option rather than
    /// anything it opened — `auto` on every launch here, since none of them names a device (see
    /// `audio_guard`, whose second cure hands this property its `auto` back). It is read beside
    /// `current-ao` so that a line states both halves of the pair: what was asked for, and the driver
    /// that got it.
    pub(super) audio_device: Option<String>,
    /// `current-ao`: the audio output driver mpv is playing through, where one was opened.
    pub(super) current_ao: Option<String>,
    /// `audio-out-params`: the node of what that output is playing at, where there is one.
    pub(super) audio_out_params: Option<String>,
    /// `volume`: **the level the player is playing at**, which is not the level the launch asked for.
    /// A player handed `--volume=100` that answers 0 here was silenced by something else — a per-file
    /// option on a resident, or a mute — and no other property on this pipe would say so.
    pub(super) volume: Option<f64>,
    /// `mute`: mpv's own mute, which silences a film at any level at all and leaves every other answer
    /// saying the film has its sound.
    pub(super) mute: Option<bool>,
}

impl EngineAnswers {
    /// Whether an audio output was actually opened: a driver that is named, and not the `null` one
    /// mpv falls back to when there is no device to play through — a null output makes no session
    /// either, and reading it as sound is the one way this trace could lie about road two — and the
    /// params of what that driver is playing at.
    ///
    /// **The driver is asked for as `current-ao`, which is the driver in use; `ao` is the option, and
    /// this app never passes one.** A captured run of the old question answered `ao []` on all
    /// fifteen launches — the option's empty list, which is not the empty string — so every player
    /// read as heard, the ones mpv had put on its `null` fallback included, and the guard could not
    /// fire for the one failure it was built for. `null` reports `audio-out-params` as readily as a
    /// real driver does, which is why the driver's own name and not the params is the discriminator.
    ///
    /// **A `current-ao` that answers nothing at all keeps the answer the params alone give.** A build
    /// too old to have the property answers as a read that did not land does — nothing — and an
    /// unanswerable question is not a device that failed: the ladder's last rung takes a picture off
    /// the screen, so it must never fire on an unknown.
    pub(super) fn an_output_opened(&self) -> bool {
        match self.current_ao.as_deref() {
            Some(driver) if !driver.is_empty() => {
                driver != "null" && self.audio_out_params.is_some()
            }
            _ => self.audio_out_params.is_some(),
        }
    }

    /// **Whether the player has an output open and was told to be quiet**: its own `volume` at
    /// nothing, or its own `mute` on. It is the third road to a silent film, and the one no other
    /// answer on this pipe describes — a track, a driver and params all say the film has sound.
    ///
    /// **An answer this build cannot read is not silence.** A property an older mpv has no value for
    /// answers nothing, and no cure is written on an unknown (see `an_output_opened`, which keeps the
    /// same rule for the driver).
    fn told_to_be_quiet(&self) -> bool {
        self.volume.is_some_and(|level| level <= 0.0) || self.mute == Some(true)
    }

    /// What these answers come to, given the level the launch asked for (see `Verdict`).
    ///
    /// **`asked` is the launch's level and it belongs to this question**, because the player's own
    /// level cannot be judged without it: a player at `volume 0` on a launch that asked for no audio
    /// at all is a silence the setting chose, not a fault to heal — and
    /// `mpv_launched_without_audio` is the same rule the launch's own form is built from, so the two
    /// can never disagree about what "asked for no audio" means.
    pub(super) fn verdict(&self, asked: u32) -> Verdict {
        let Some(aid) = self.aid.as_deref() else {
            return Verdict::Unanswered;
        };

        if aid == "no" {
            return Verdict::NoTrack;
        }

        if !self.an_output_opened() {
            return Verdict::NoDevice;
        }

        if !super::video_launch::mpv_launched_without_audio(asked) && self.told_to_be_quiet() {
            return Verdict::Silent;
        }

        Verdict::Heard
    }

    /// **Whether these answers are the shape a player that is still opening its output has**: a real
    /// driver named on `current-ao` with no `audio-out-params` beside it. A driver is in use and
    /// nothing is playing through it yet, which is a player mid-open as readily as it is a film whose
    /// device was refused — a captured run read exactly this 548 ms after its launch.
    ///
    /// **It is not `an_output_opened`'s question.** A `null` driver is mpv's fallback and not an
    /// output being opened, and neither is a `current-ao` that answered nothing at all: both are
    /// answers, so the guard acts on them on the first read (see `audio_guard::the_step`).
    pub(super) fn a_driver_with_no_params(&self) -> bool {
        match self.current_ao.as_deref() {
            Some(driver) => {
                !driver.is_empty() && driver != "null" && self.audio_out_params.is_none()
            }
            None => false,
        }
    }

    /// Whether asking again could change the answer: a player that said nothing at all yet, or one
    /// whose output has not come up while a track is selected — a device opening, or a file whose
    /// header is still being read. A player with a verdict of its own is asked no further.
    fn may_answer_differently(&self, asked: u32) -> bool {
        match self.verdict(asked) {
            Verdict::Unanswered | Verdict::NoDevice => true,
            Verdict::NoTrack | Verdict::Heard | Verdict::Silent => false,
            // A player whose own answers are sound answers nothing more about them: what is muted is
            // the session, and asking the player again is not what finds it (see
            // `audio_session::look_and_cure`).
            Verdict::SessionMuted => false,
        }
    }
}

/// The word a verdict is written as, which is what a run is read for — the trace's own `engine truth`
/// line and the playback guard's read-backs are both written in it (see `audio_guard`).
///
/// **`no-device` is two faults sharing a word**: a device refused, and a player that has named its
/// driver and not yet opened it, which is why the guard reads the answers themselves before it writes
/// anything (see `EngineAnswers::a_driver_with_no_params`).
///
/// **`silent` is the verdict that reads as sound everywhere else**: a track, an output and params, with
/// the player's own level at nothing or its mute on (see `Verdict::Silent`).
///
/// **`session-muted` is that same shape one layer out** — the player unmuted and Windows' own session for
/// it muted — and the three are written apart so that a log names which of them happened: `silent` is the
/// player, `session-muted` is the machine, and `no-device` is neither having been reached at all.
pub(super) fn the_verdict_word(verdict: Verdict) -> &'static str {
    match verdict {
        Verdict::NoTrack => "no-track",
        Verdict::NoDevice => "no-device",
        Verdict::Heard => "heard",
        Verdict::Silent => "silent",
        Verdict::SessionMuted => "session-muted",
        Verdict::Unanswered => "unanswered",
    }
}

/// The level a player reported, or `-` where it reported none.
pub(super) fn the_level_word(level: Option<f64>) -> String {
    level.map_or_else(|| "-".to_string(), |level| format!("{level}"))
}

/// The flag a player reported, as `yes`/`no`, or `-` where it reported none.
pub(super) fn the_flag_word(flag: Option<bool>) -> String {
    match flag {
        Some(true) => "yes".to_string(),
        Some(false) => "no".to_string(),
        None => "-".to_string(),
    }
}

/// The line the engine's own answers are written as: the six properties it reported, and what they
/// come to.
///
/// **`asked` is the level the launch resolved to and `volume` the level the player is at**, and the two
/// are two different words on purpose: a line that reads `asked 100 volume 0 mute no` is a film this
/// app gave a level to and a player that is not at it, which is the one silence the four answers
/// before this change could not see.
///
/// **`audio-device` is the device asked for and `current-ao` the driver in use**, so one line carries
/// both what the launch (or a cure) requested and what the player got; a property an older build
/// cannot answer is written `-`, which is how such a line has always read.
///
/// **`no-track` is the launch's own fault and `no-device` is not**, which is the whole of what a run
/// of this is read for: the first is a level this app resolved to nothing, the second is a player that
/// was handed a level and could not open the device it wanted — and `silent` is a player that opened
/// one and was told not to be heard through it.
pub(super) fn the_engine_truth(pid: u32, asked: u32, answers: &EngineAnswers) -> String {
    let said = |answer: &Option<String>| answer.clone().unwrap_or_else(|| "-".to_string());
    format!(
        "engine truth mpv pid {pid} asked {asked} volume {} mute {} aid {} audio-device {} \
         current-ao {} audio-out-params {} verdict {}",
        the_level_word(answers.volume),
        the_flag_word(answers.mute),
        said(&answers.aid),
        said(&answers.audio_device),
        said(&answers.current_ao),
        said(&answers.audio_out_params),
        the_verdict_word(answers.verdict(asked)),
    )
}

/// The line a kill is written as: whether the process was gone by the time the kill returned, and which
/// of the two roads killed it.
///
/// **A kill nobody confirmed is a process that may still be holding an audio endpoint**, which is what a
/// second player opening the same device has to get past — so the answer is written beside the kill
/// rather than assumed, and nothing is waited for to have it. The two roads ask it their own way: the
/// player this app holds a handle to is asked with a wait of nothing (see
/// `video_playback::PlayerProcess::exit_confirmed`), and a process the sweep ends is asked by its exit
/// code (see `app::engine_processes::terminate_verified`). `None` is a player no question can be asked
/// about — a served hover's resident, which is the resident's process and not this handle's.
pub(super) fn a_kill(pid: u32, confirmed: Option<bool>, road: &str) -> String {
    let confirmed = match confirmed {
        Some(true) => "yes",
        Some(false) => "no",
        None => "-",
    };
    format!("kill {road} pid {pid} confirmed {confirmed}")
}

/// How long a started player is given before it is asked what it did: mpv makes its pipe some tens of
/// milliseconds into the start (measured at 40 ms warm and 211 ms cold) and selects its tracks as the
/// file is opened after that.
const ENGINE_SETTING_UP: Duration = Duration::from_millis(350);

/// How many times the four answers are asked, and how long between two askings.
///
/// **Because a verdict read off an answer that has not come yet is road two invented by the trace
/// itself**: a device being opened, or a file whose header is still being read, answers nothing at the
/// first asking. Four askings a third of a second apart is longer than either takes on this machine
/// and shorter than a film's first second.
const ENGINE_ASKS: usize = 4;
const BETWEEN_ASKS: Duration = Duration::from_millis(300);

/// How long one question is offered the pipe before it is given up on: an open that has not landed by
/// then is a player that has no pipe, and the answer is `Unanswered` rather than a guess.
const AN_ANSWER_WAIT: Duration = Duration::from_millis(400);

/// Ask a player that has just been begun what it did with the sound, and write it down.
///
/// Nothing at all happens where the trace is not asked for: no thread, no read, and no pipe opened
/// against a player nobody is watching.
pub(super) fn watch_the_engine(pid: u32, volume: u32) {
    if !asked() {
        return;
    }

    // **And the other half of the same question, which is the machine's rather than the player's**:
    // whether Windows has a session for this pid at all, what state it is in, and the mixer's own
    // slider beside the level this line reports (see `audio_session`). Both threads are begun here,
    // at the launch, because a line read only after a fault has no working film to be read against —
    // and a captured run's silent film and the twelve beside it it agreed with answered every one of
    // the six questions below identically.
    super::audio_session::watch_the_mixer(pid);

    let _ = std::thread::Builder::new()
        .name("playback-engine".to_string())
        .spawn(move || {
            std::thread::sleep(ENGINE_SETTING_UP);

            let mut answers = EngineAnswers::default();
            for asking in 0..ENGINE_ASKS {
                if asking > 0 {
                    std::thread::sleep(BETWEEN_ASKS);
                }
                answers = ask_the_engine(pid);
                if !answers.may_answer_differently(volume) {
                    break;
                }
            }

            write(&the_engine_truth(pid, volume, &answers));
        });
}

/// The six properties, each bounded (see `mpv_ipc::get_property_within`), and the one place they are
/// asked: the trace reads them once to describe a launch, and the guard reads them again at every read
/// its ladder makes — the one after a player comes up, each ask of a player that has said nothing, and
/// the read-back of every rung it took (see `audio_guard`).
///
/// The driver is `current-ao` and never `ao`: the second is the option, which no launch here sets, so
/// it answers an empty list rather than a driver (see `EngineAnswers::an_output_opened`).
/// `audio-device` is read for the other half of that pair — what was asked for — and answers nothing
/// on a build without it, which leaves the line as `-` and every verdict exactly as it was.
///
/// **`volume` and `mute` are the player's own and not the launch's**: mpv answers a level of its own
/// whatever it was handed, and the pair is the only way a film that is open and told to be quiet is
/// ever seen (see `EngineAnswers::told_to_be_quiet`).
pub(super) fn ask_the_engine(pid: u32) -> EngineAnswers {
    EngineAnswers {
        aid: super::mpv_ipc::get_property_within(pid, "aid", AN_ANSWER_WAIT),
        audio_device: super::mpv_ipc::get_property_within(pid, "audio-device", AN_ANSWER_WAIT),
        current_ao: super::mpv_ipc::get_property_within(pid, "current-ao", AN_ANSWER_WAIT),
        audio_out_params: super::mpv_ipc::get_property_within(
            pid,
            "audio-out-params",
            AN_ANSWER_WAIT,
        ),
        volume: the_number_said(super::mpv_ipc::get_property_within(
            pid,
            "volume",
            AN_ANSWER_WAIT,
        )),
        mute: the_flag_said(super::mpv_ipc::get_property_within(
            pid,
            "mute",
            AN_ANSWER_WAIT,
        )),
    }
}

/// The number mpv answered a property with, where it answered one: a scalar is written `100.000000`
/// or `0`, and a build without the property answers nothing at all.
fn the_number_said(answer: Option<String>) -> Option<f64> {
    answer?.trim().parse().ok()
}

/// The flag mpv answered a property with, where it answered one: a boolean is written `true` or
/// `false`, and `yes`/`no` are taken as the same two answers.
fn the_flag_said(answer: Option<String>) -> Option<bool> {
    match answer?.trim() {
        "true" | "yes" => Some(true),
        "false" | "no" => Some(false),
        _ => None,
    }
}

// The lines a test run wrote, so a test can hold the road to what it said rather than to what the log
// happened to be configured to keep: the same seam `pin_trace` and `mpv_ipc` keep, and for the same
// reason (see `note_playback_trace`).
//
// **It is the thread's own rather than the run's**, which is the one thing it does differently from
// its two siblings: a launch line is written by whatever thread began the player, and a test that
// cleared a table the whole suite writes into would be reading another test's launches as readily as
// its own.
#[cfg(test)]
thread_local! {
    static NOTED: std::cell::RefCell<Vec<String>> = const { std::cell::RefCell::new(Vec::new()) };
}

/// Write one line down for a test that reads it (see `NOTED`).
#[cfg(test)]
pub(super) fn note_a_line_for_a_test(line: String) {
    NOTED.with(|noted| noted.borrow_mut().push(line));
}

/// The lines this thread has written (see `NOTED`).
#[cfg(test)]
pub(super) fn noted_lines() -> Vec<String> {
    NOTED.with(|noted| noted.borrow().clone())
}

/// Forget the lines written on this thread before a test (see `NOTED`).
#[cfg(test)]
pub(super) fn clear_noted() {
    NOTED.with(|noted| noted.borrow_mut().clear());
}

/// Write one line, and build nothing at all where the trace is off. A test run keeps every line it
/// wrote, whether the log is on or not (see `NOTED`).
///
/// The paths are `$crate`'s rather than `super::`'s, which is the one thing this macro does differently
/// from `note_pin_trace`: its lines are written by this module's own tests as well as by the launches
/// beside it, and `super::playback_trace` is not a path this module has.
macro_rules! note_playback_trace {
    ($($line:tt)*) => {{
        if $crate::ui::preview_window::playback_trace::asked() {
            $crate::ui::preview_window::playback_trace::write(&format!($($line)*));
        }
        #[cfg(test)]
        $crate::ui::preview_window::playback_trace::note_a_line_for_a_test(format!($($line)*));
    }};
}

pub(super) use note_playback_trace;

#[cfg(test)]
mod tests {
    use super::*;

    /// **The one rule both launches are built from**: only mpv at a level of nothing is given no track
    /// at all. FFmpeg's player keeps its track at every level and is turned down at the mixer, so a
    /// level of nothing is a level like any other for it.
    #[test]
    fn a_level_of_nothing_is_no_track_on_mpv_and_a_level_on_ffplay() {
        assert_eq!(AudioForm::of(Engine::Mpv, 0, None), AudioForm::NoTrack);
        assert_eq!(
            AudioForm::of(Engine::Mpv, 0, Some(1.5)),
            AudioForm::NoTrack,
            "a muted film is not normalised, since there is nothing to hear"
        );
        assert_eq!(
            AudioForm::of(Engine::Ffplay, 0, None),
            AudioForm::Level {
                level: 0,
                gain: None
            },
            "FFmpeg's player plays its track at mixer 0.0 rather than being given no track"
        );
        assert_eq!(
            AudioForm::of(Engine::Mpv, 40, None),
            AudioForm::Level {
                level: 40,
                gain: None
            }
        );
        assert_eq!(
            AudioForm::of(Engine::Ffplay, 100, Some(1.25)),
            AudioForm::Level {
                level: 100,
                gain: Some(1.25)
            }
        );
    }

    /// And the words a line names the form by are each player's own answer: mpv's level in its own
    /// percent, FFmpeg's at its mixer (see `audio_session::ask_the_mixer`).
    #[test]
    fn a_form_is_named_by_the_flag_the_player_is_actually_given() {
        assert_eq!(AudioForm::NoTrack.word(Engine::Mpv), "aid=no");
        assert_eq!(
            AudioForm::of(Engine::Mpv, 40, None).word(Engine::Mpv),
            "volume=40 gain=none"
        );
        assert_eq!(
            AudioForm::of(Engine::Ffplay, 40, Some(1.25)).word(Engine::Ffplay),
            "mixer=40 gain=1.2500"
        );
    }

    /// FFmpeg's player is given the file's measured gain and **never the app's level**, which is its
    /// mixer session's; a file with no measured gain is given no filter at all.
    #[test]
    fn ffmpegs_player_is_given_its_gain_and_never_the_apps_level() {
        assert_eq!(AudioForm::NoTrack.ffplay_filter(), None);
        assert_eq!(
            AudioForm::of(Engine::Ffplay, 40, None).ffplay_filter(),
            None
        );
        assert_eq!(
            AudioForm::of(Engine::Ffplay, 0, Some(1.5)).ffplay_filter(),
            Some("volume=1.5000".to_string()),
            "the gain is the filter's whole job, at a level of nothing as much as any other"
        );
        assert_eq!(
            AudioForm::of(Engine::Ffplay, 100, Some(1.25)).ffplay_filter(),
            Some("volume=1.2500".to_string())
        );
    }

    /// **The discriminator itself.** A player that was given no track says so and nothing else can be
    /// read from what it did not do; a player with a track and no output is the device's failure; a
    /// player with both is a film that was heard; and a player that answered nothing is no answer.
    #[test]
    fn the_engines_answers_are_read_as_the_road_a_silent_film_took() {
        let answers =
            |aid: Option<&str>, current_ao: Option<&str>, params: Option<&str>| EngineAnswers {
                aid: aid.map(str::to_string),
                audio_device: None,
                current_ao: current_ao.map(str::to_string),
                audio_out_params: params.map(str::to_string),
                ..Default::default()
            };

        assert_eq!(answers(None, None, None).verdict(40), Verdict::Unanswered);
        assert_eq!(
            answers(Some("no"), None, None).verdict(40),
            Verdict::NoTrack,
            "no track selected is the launch's own doing, whatever the device would have done"
        );
        assert_eq!(
            answers(Some("1"), None, None).verdict(40),
            Verdict::NoDevice,
            "a track picked and no output opened from it is the device, not the launch"
        );
        assert_eq!(
            answers(Some("1"), Some("wasapi"), None).verdict(40),
            Verdict::NoDevice,
            "a driver named with nothing playing through it is not yet an output"
        );
        assert_eq!(
            answers(Some("1"), Some("null"), Some("{")).verdict(40),
            Verdict::NoDevice,
            "mpv's fallback to a null output makes no session either"
        );
        assert_eq!(
            answers(
                Some("auto"),
                Some("wasapi"),
                Some(r#"{"samplerate":48000}"#)
            )
            .verdict(40),
            Verdict::Heard
        );
    }

    /// **The third road, and the one no other answer could see.** A track is selected, an output is
    /// open and its params are present — everything that has always read as sound — and the player's
    /// own level is 0, or its own mute is on: a film that is playing and cannot be heard, which the
    /// guard now heals rather than calls heard.
    ///
    /// **The level a player reports is not the level the launch asked for**, and that is why the
    /// launch's level is read beside it: the same answers on a launch that asked for no audio at all
    /// are a silence the setting chose (see the next test).
    #[test]
    fn a_player_at_nothing_or_muted_is_silent_and_never_heard() {
        let open = |volume: Option<f64>, mute: Option<bool>| EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            volume,
            mute,
        };

        assert_eq!(
            open(Some(0.0), Some(false)).verdict(100),
            Verdict::Silent,
            "an open output at a level of nothing is a film nobody can hear"
        );
        assert_eq!(
            open(Some(100.0), Some(true)).verdict(100),
            Verdict::Silent,
            "a mute silences a film at the very level it was asked for"
        );
        assert_eq!(
            open(Some(0.0), None).verdict(100),
            Verdict::Silent,
            "a mute that answered nothing does not unmake a level of nothing"
        );
        assert_eq!(
            open(Some(100.0), Some(false)).verdict(100),
            Verdict::Heard,
            "the level asked for, unmuted, is the film this whole ladder is not for"
        );
        assert_eq!(
            open(None, None).verdict(100),
            Verdict::Heard,
            "a build that answers neither property is not a film told to be quiet: no cure is \
             written on an unknown"
        );
        assert!(
            !open(Some(0.0), Some(false)).may_answer_differently(100),
            "a player that said where its level stands is asked no further"
        );
    }

    /// **A launch that legitimately asked for no audio is never a silent film.** The setting said no
    /// track, the player was begun with `--aid=no` rather than a level of nothing
    /// (see `mpv_launched_without_audio`), and the silence is the level's own rather than a fault to
    /// heal — so it reads as the launch's road however the player's level happens to stand.
    #[test]
    fn a_launch_that_asked_for_no_audio_never_reads_as_silent() {
        let no_track = EngineAnswers {
            aid: Some("no".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            volume: Some(0.0),
            mute: Some(true),
            ..Default::default()
        };
        assert_eq!(
            no_track.verdict(0),
            Verdict::NoTrack,
            "a level of nothing is no track at all, which is the launch's own form"
        );

        // A track a later level asked back for on that same player: the output is open and the level
        // is still at nothing, and it is still not this app's fault to heal — the launch asked for
        // none, and the trace cannot tell this level from a hand's.
        let asked_back = EngineAnswers {
            aid: Some("auto".to_string()),
            volume: Some(0.0),
            ..no_track
        };
        assert_eq!(
            asked_back.verdict(0),
            Verdict::Heard,
            "a launch begun without audio is never read as a film to be cured"
        );
    }

    /// **The case that read as sound for as long as the driver was asked for as `ao`.** A player mpv
    /// has put on its `null` fallback reports `audio-out-params` as readily as a real driver does, so
    /// the params alone cannot tell the two apart and the driver's own name is the whole of the
    /// discriminator. `ao` — the option, which this app never passes — answered the empty list `[]`
    /// on every launch of a captured run, which is not the empty string, so this verdict was `Heard`
    /// and the guard could not fire for the one failure it exists for.
    #[test]
    fn a_null_on_current_ao_is_the_device_even_with_the_params_beside_it() {
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("null".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            ..Default::default()
        };

        assert_eq!(answers.verdict(40), Verdict::NoDevice);
        assert!(
            !answers.an_output_opened(),
            "a null output makes no session, so nothing was heard"
        );
    }

    /// And a driver that is not the fallback, with the params of what it is playing at and a level
    /// above nothing in it, is a film that was heard — the one verdict that ends the guard's ladder
    /// without touching the player.
    #[test]
    fn a_driver_on_current_ao_with_the_params_is_a_film_that_was_heard() {
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            volume: Some(100.0),
            mute: Some(false),
        };

        assert_eq!(answers.verdict(100), Verdict::Heard);
    }

    /// **A `current-ao` that answers nothing at all is not a device that failed.** A build too old to
    /// have the property leaves the params to answer on their own, so an unanswerable question reads
    /// `Heard` — because the ladder's last rung takes a picture off the screen, and a relaunch fired
    /// on an unknown is worse than a silence this app could not have seen either way.
    #[test]
    fn a_current_ao_that_answers_nothing_at_all_is_read_by_the_params_alone() {
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            current_ao: None,
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            ..Default::default()
        };

        assert_eq!(answers.verdict(40), Verdict::Heard);
        assert!(
            !answers.may_answer_differently(40),
            "a heard film is not asked about again"
        );
    }

    /// **The shape of a player still opening its output**: a real driver named on `current-ao` with no
    /// `audio-out-params` beside it yet. The guard reads that shape once more before it writes
    /// anything, because a captured run read exactly it 548 ms after its launch — and that same read is
    /// what a player merely still opening looks like, which is why the trace alone could not say
    /// whether the `audio-device=auto` write or the player was the fix.
    ///
    /// A `null` driver, a driver that answered nothing at all, and a driver with its params are none of
    /// them that shape: each is an answer, and each takes its rung on the first read.
    #[test]
    fn a_driver_named_with_nothing_playing_through_it_is_a_player_still_opening() {
        let answers = |current_ao: Option<&str>, params: Option<&str>| EngineAnswers {
            aid: Some("1".to_string()),
            current_ao: current_ao.map(str::to_string),
            audio_out_params: params.map(str::to_string),
            ..Default::default()
        };

        assert!(
            answers(Some("wasapi"), None).a_driver_with_no_params(),
            "a driver in use with nothing playing through it is a player mid-open as readily as a \
             film whose device was refused"
        );
        assert!(
            !answers(Some("wasapi"), Some(r#"{"samplerate":48000}"#)).a_driver_with_no_params(),
            "an output that named what it plays at is open"
        );
        assert!(
            !answers(Some("null"), None).a_driver_with_no_params(),
            "the fallback is not a device being opened"
        );
        assert!(
            !answers(None, None).a_driver_with_no_params(),
            "a build that answers nothing is not a player mid-open, so it takes its rung at once"
        );
        assert!(
            !answers(Some(""), None).a_driver_with_no_params(),
            "an empty answer is no driver at all"
        );
    }

    /// The track outranks the driver whichever way the driver answers: `aid=no` is the launch's own
    /// road, and an `aid` of `-` is a player that said nothing rather than one with no track.
    #[test]
    fn no_track_outranks_the_driver_and_no_answer_is_no_answer() {
        let no_track_but_heard = EngineAnswers {
            aid: Some("no".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            ..Default::default()
        };
        assert_eq!(
            no_track_but_heard.verdict(40),
            Verdict::NoTrack,
            "a track the launch never gave the player is the launch's road, sound or no sound"
        );

        let no_answer_but_heard = EngineAnswers {
            aid: None,
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            ..Default::default()
        };
        assert_eq!(
            no_answer_but_heard.verdict(40),
            Verdict::Unanswered,
            "an `aid` of `-` says nothing about the sound, so nothing may be done about it"
        );
    }

    /// A player is asked again while an answer could still change, and never once it has one: a
    /// verdict taken from a device that is still opening is a road two the trace made up.
    #[test]
    fn a_player_that_answered_is_not_asked_again() {
        let answers =
            |aid: Option<&str>, current_ao: Option<&str>, params: Option<&str>| EngineAnswers {
                aid: aid.map(str::to_string),
                current_ao: current_ao.map(str::to_string),
                audio_out_params: params.map(str::to_string),
                ..Default::default()
            };
        let heard = Some(r#"{"samplerate":48000}"#);

        assert!(
            answers(None, None, None).may_answer_differently(40),
            "nothing yet"
        );
        assert!(
            answers(Some("1"), None, None).may_answer_differently(40),
            "a track whose output has not come up yet"
        );
        assert!(
            answers(Some("1"), Some("null"), heard).may_answer_differently(40),
            "a fallback the device cure may still reload out of"
        );
        assert!(!answers(Some("no"), None, None).may_answer_differently(40));
        assert!(!answers(Some("1"), Some("wasapi"), heard).may_answer_differently(40));
        assert!(
            !answers(Some("1"), None, heard).may_answer_differently(40),
            "a driver nothing can name with the params of an output beside it is heard"
        );
    }

    /// **What a run is read for**: the launch line says the level and the form it was given, and the
    /// engine's line says what the player did with it — two lines, because they answer two questions.
    #[test]
    fn a_launch_and_the_engine_are_two_lines_about_one_film() {
        let line = a_launch(
            Engine::Mpv,
            Some(MediaType::Video),
            Path::new("a-film.mkv"),
            0,
            None,
            true,
        );
        assert_eq!(
            line,
            "launch engine mpv kind video file a-film.mkv volume 0 form aid=no pinned 1",
            "a level of nothing is written as the flag the player was handed, and the loading folder \
             is not part of a line"
        );

        let answers = EngineAnswers {
            aid: Some("no".to_string()),
            current_ao: None,
            audio_out_params: None,
            ..Default::default()
        };
        assert_eq!(
            the_engine_truth(7, 0, &answers),
            "engine truth mpv pid 7 asked 0 volume - mute - aid no audio-device - current-ao - \
             audio-out-params - verdict no-track"
        );

        // And a device that would not open names itself, with the track it never played and the
        // device it was asked for: `audio-device` is the request and `current-ao` the driver that
        // answered it, so one line says both halves of a film that opened nothing.
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: None,
            audio_out_params: None,
            ..Default::default()
        };
        assert_eq!(
            the_engine_truth(9, 40, &answers),
            "engine truth mpv pid 9 asked 40 volume - mute - aid 1 audio-device auto current-ao - \
             audio-out-params - verdict no-device",
            "a launch that gave a level and a playlist that opened no output is the device's road"
        );

        // **And the two levels are two words, which is the whole of what makes the third road
        // readable**: the level the launch resolved to is `asked`, and the one the player is at is
        // `volume`, with its own mute beside it.
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            audio_device: Some("auto".to_string()),
            current_ao: Some("wasapi".to_string()),
            audio_out_params: Some(r#"{"samplerate":48000}"#.to_string()),
            volume: Some(0.0),
            mute: Some(false),
        };
        assert_eq!(
            the_engine_truth(11, 100, &answers),
            "engine truth mpv pid 11 asked 100 volume 0 mute no aid 1 audio-device auto \
             current-ao wasapi audio-out-params {\"samplerate\":48000} verdict silent",
            "a player at nothing on a launch that asked for 100 is the film that was open and told \
             to be quiet"
        );
    }

    /// The other half of a silent film is a process still standing: a kill is written down with
    /// whether it was confirmed gone, and a kill nobody confirmed has a `-` rather than an answer.
    #[test]
    fn a_kill_is_written_down_with_the_answer_it_got_and_nothing_invented() {
        assert_eq!(
            a_kill(21, Some(true), "player"),
            "kill player pid 21 confirmed yes"
        );
        assert_eq!(
            a_kill(21, Some(false), "player"),
            "kill player pid 21 confirmed no"
        );
        assert_eq!(
            a_kill(21, None, "swept"),
            "kill swept pid 21 confirmed -",
            "a process this app holds no handle to has nothing to confirm"
        );
    }

    /// **A line the macro is asked for is the line that arrives**, and a line the launch did not ask
    /// for is not built at all: the gate is one read of an atomic, so a run that did not ask for the
    /// trace pays nothing for the lines it does not write.
    #[test]
    fn a_launchs_line_is_written_where_the_trace_asks_for_it() {
        clear_noted();

        note_playback_trace!(
            "{}",
            a_launch(
                Engine::Ffplay,
                Some(MediaType::Video),
                Path::new("a-film.mkv"),
                40,
                None,
                false,
            )
        );

        assert_eq!(
            noted_lines(),
            vec![
                "launch engine ffplay kind video file a-film.mkv volume 40 form \
                  mixer=40 gain=none pinned 0"
                    .to_string()
            ],
            "one line per launch, and the form is the one the arguments are built from"
        );
        clear_noted();
    }

    /// **A run's log is the run's own.** Two lines written by one process both land; a fresh start
    /// makes the file over rather than appending to what the last run left behind, and every line is
    /// flushed as it is written so a run that ends badly still leaves what it said.
    #[test]
    fn a_run_keeps_its_own_lines_and_a_fresh_start_makes_the_file_over() {
        let folder = std::env::temp_dir().join("rust-hover-preview-playback-trace");
        let _ = std::fs::remove_dir_all(&folder);
        let _ = std::fs::create_dir_all(&folder);
        let path = folder.join("rhp-playback-trace.log");

        // The timestamp a line opens with, so what is left is the line itself.
        let the_lines = || -> Vec<String> {
            std::fs::read_to_string(&path)
                .unwrap_or_default()
                .lines()
                .map(|line| {
                    line.split_once(' ')
                        .map_or("", |(_, line)| line)
                        .to_string()
                })
                .collect()
        };

        let a_run = Once::new();
        write_to(&path, &a_run, "first line");
        write_to(&path, &a_run, "second line");
        assert_eq!(
            the_lines(),
            vec!["first line", "second line"],
            "two lines written by one process are two lines in its log"
        );

        let a_fresh_start = Once::new();
        write_to(&path, &a_fresh_start, "third line");
        assert_eq!(
            the_lines(),
            vec!["third line"],
            "a new run makes the log over: the last run's lines are not this run's"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }
}

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
//! so it is read off the player: mpv's `aid` says whether a track is selected and `ao` and
//! `audio-out-params` say whether an output was opened from it, and the pair is the discriminator —
//! `aid=no` is the first road, `aid` a track with no output is the second (see `the_verdict`).
//!
//! **Nothing here runs unless the variable is set**, and what it writes goes to
//! `%TEMP%\rhp-playback-trace.log` — the same shape `RHP_PIN_TRACE` has (see `pin_trace`), a file of
//! its own because the question is a different one and a run that asks it should not have to read a
//! pin's road to answer it. A run without the variable builds no line at all and spawns no thread.
//!
//! **The engine's answers are read on a thread of their own, never on the preview loop.** A player
//! that has just been begun has no pipe yet (mpv makes one some tens of milliseconds into the start,
//! measured at 40 ms warm and 211 ms cold), and waiting for one is a wait no tick may pay (see
//! `mpv_ipc::get_property_within`, which is bounded per question and only ever called from here).

use std::path::Path;
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

/// Write one timestamped line to `%TEMP%\rhp-playback-trace.log`, appended and flushed so a run that
/// ends badly leaves what it wrote.
pub(super) fn write(line: &str) {
    let stamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|elapsed| format!("{}.{:03}", elapsed.as_secs(), elapsed.subsec_millis()))
        .unwrap_or_else(|_| "-".to_string());
    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(std::env::temp_dir().join("rhp-playback-trace.log"))
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
        AudioForm::of(volume, gain).word(engine),
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
    /// The player answered nothing in the time it was given, which says nothing about its sound.
    Unanswered,
}

/// The three answers a started mpv is asked for, and nothing where it said nothing to one of them.
#[derive(Clone, Default, Debug)]
pub(super) struct EngineAnswers {
    /// `aid`: `no`, `auto`, or the track that was picked.
    pub(super) aid: Option<String>,
    /// `ao`: the audio output driver mpv is playing through, where it has one.
    pub(super) ao: Option<String>,
    /// `audio-out-params`: the node of what that output is playing at, where there is one.
    pub(super) audio_out_params: Option<String>,
}

impl EngineAnswers {
    /// Whether an audio output was actually opened: a driver that is named, and not the `null` one
    /// mpv falls back to when there is no device to play through — a null output makes no session
    /// either, and reading it as sound is the one way this trace could lie about road two.
    pub(super) fn an_output_opened(&self) -> bool {
        self.ao
            .as_deref()
            .is_some_and(|ao| !ao.is_empty() && ao != "null")
            && self.audio_out_params.is_some()
    }

    /// What these answers come to (see `Verdict`).
    pub(super) fn verdict(&self) -> Verdict {
        let Some(aid) = self.aid.as_deref() else {
            return Verdict::Unanswered;
        };

        if aid == "no" {
            return Verdict::NoTrack;
        }

        if self.an_output_opened() {
            Verdict::Heard
        } else {
            Verdict::NoDevice
        }
    }

    /// Whether asking again could change the answer: a player that said nothing at all yet, or one
    /// whose output has not come up while a track is selected — a device opening, or a file whose
    /// header is still being read. A player with a verdict of its own is asked no further.
    fn may_answer_differently(&self) -> bool {
        match self.verdict() {
            Verdict::Unanswered | Verdict::NoDevice => true,
            Verdict::NoTrack | Verdict::Heard => false,
        }
    }
}

/// The word a verdict is written as, which is what a run is read for.
fn the_verdict_word(verdict: Verdict) -> &'static str {
    match verdict {
        Verdict::NoTrack => "no-track",
        Verdict::NoDevice => "no-device",
        Verdict::Heard => "heard",
        Verdict::Unanswered => "unanswered",
    }
}

/// The line the engine's own answers are written as: the three properties it reported, and what they
/// come to.
///
/// **`no-track` is the launch's own fault and `no-device` is not**, which is the whole of what a run
/// of this is read for: the first is a level this app resolved to nothing, the second is a player that
/// was handed a level and could not open the device it wanted.
pub(super) fn the_engine_truth(pid: u32, volume: u32, answers: &EngineAnswers) -> String {
    let said = |answer: &Option<String>| answer.clone().unwrap_or_else(|| "-".to_string());
    format!(
        "engine truth mpv pid {pid} volume {volume} aid {} ao {} audio-out-params {} verdict {}",
        said(&answers.aid),
        said(&answers.ao),
        said(&answers.audio_out_params),
        the_verdict_word(answers.verdict()),
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

/// How many times the three questions are asked, and how long between two askings.
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
                if !answers.may_answer_differently() {
                    break;
                }
            }

            write(&the_engine_truth(pid, volume, &answers));
        });
}

/// The three questions, each bounded (see `mpv_ipc::get_property_within`).
fn ask_the_engine(pid: u32) -> EngineAnswers {
    EngineAnswers {
        aid: super::mpv_ipc::get_property_within(pid, "aid", AN_ANSWER_WAIT),
        ao: super::mpv_ipc::get_property_within(pid, "ao", AN_ANSWER_WAIT),
        audio_out_params: super::mpv_ipc::get_property_within(
            pid,
            "audio-out-params",
            AN_ANSWER_WAIT,
        ),
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

    /// **The one rule both launches are built from**: a level of nothing is no track at all, and every
    /// level above nothing is the level, with the gain where one was measured.
    #[test]
    fn a_level_of_nothing_is_no_track_and_every_other_level_is_the_level() {
        assert_eq!(AudioForm::of(0, None), AudioForm::NoTrack);
        assert_eq!(
            AudioForm::of(0, Some(1.5)),
            AudioForm::NoTrack,
            "a muted film is not normalised, since there is nothing to hear"
        );
        assert_eq!(
            AudioForm::of(40, None),
            AudioForm::Level {
                level: 40,
                gain: None
            }
        );
        assert_eq!(
            AudioForm::of(100, Some(1.25)),
            AudioForm::Level {
                level: 100,
                gain: Some(1.25)
            }
        );
    }

    /// And the words a line names the form by are each player's own flag, so a run can be read against
    /// the command line the launch built.
    #[test]
    fn a_form_is_named_by_the_flag_the_player_is_actually_given() {
        assert_eq!(AudioForm::NoTrack.word(Engine::Mpv), "aid=no");
        assert_eq!(AudioForm::NoTrack.word(Engine::Ffplay), "-an");
        assert_eq!(
            AudioForm::of(40, None).word(Engine::Mpv),
            "volume=40 gain=none"
        );
        assert_eq!(
            AudioForm::of(40, Some(1.25)).word(Engine::Ffplay),
            "volume=40 gain=1.2500"
        );
    }

    /// FFmpeg's player is given the level as a share of full and nothing at all where the form is no
    /// track — the filter that plays the film, and not one that plays it at nothing.
    #[test]
    fn ffmpegs_player_is_given_its_filter_only_where_it_has_a_level_to_play_at() {
        assert_eq!(AudioForm::NoTrack.ffplay_filter(), None);
        assert_eq!(
            AudioForm::of(40, None).ffplay_filter(),
            Some("volume=0.40".to_string())
        );
        assert_eq!(
            AudioForm::of(100, Some(1.25)).ffplay_filter(),
            Some("volume=1.2500".to_string()),
            "the level and the gain multiply, as they do for a sound file's player"
        );
    }

    /// **The discriminator itself.** A player that was given no track says so and nothing else can be
    /// read from what it did not do; a player with a track and no output is the device's failure; a
    /// player with both is a film that was heard; and a player that answered nothing is no answer.
    #[test]
    fn the_engines_answers_are_read_as_the_road_a_silent_film_took() {
        let answers = |aid: Option<&str>, ao: Option<&str>, params: Option<&str>| EngineAnswers {
            aid: aid.map(str::to_string),
            ao: ao.map(str::to_string),
            audio_out_params: params.map(str::to_string),
        };

        assert_eq!(answers(None, None, None).verdict(), Verdict::Unanswered);
        assert_eq!(
            answers(Some("no"), None, None).verdict(),
            Verdict::NoTrack,
            "no track selected is the launch's own doing, whatever the device would have done"
        );
        assert_eq!(
            answers(Some("1"), None, None).verdict(),
            Verdict::NoDevice,
            "a track picked and no output opened from it is the device, not the launch"
        );
        assert_eq!(
            answers(Some("1"), Some("wasapi"), None).verdict(),
            Verdict::NoDevice,
            "a driver named with nothing playing through it is not yet an output"
        );
        assert_eq!(
            answers(Some("1"), Some("null"), Some("{")).verdict(),
            Verdict::NoDevice,
            "mpv's fallback to a null output makes no session either"
        );
        assert_eq!(
            answers(
                Some("auto"),
                Some("wasapi"),
                Some(r#"{"samplerate":48000}"#)
            )
            .verdict(),
            Verdict::Heard
        );
    }

    /// A player is asked again while an answer could still change, and never once it has one: a
    /// verdict taken from a device that is still opening is a road two the trace made up.
    #[test]
    fn a_player_that_answered_is_not_asked_again() {
        let answers = |aid: Option<&str>, ao: Option<&str>| EngineAnswers {
            aid: aid.map(str::to_string),
            ao: ao.map(str::to_string),
            audio_out_params: Some("{".to_string()),
        };

        assert!(answers(None, None).may_answer_differently(), "nothing yet");
        assert!(
            answers(Some("1"), None).may_answer_differently(),
            "a track whose output has not come up yet"
        );
        assert!(!answers(Some("no"), None).may_answer_differently());
        assert!(!answers(Some("1"), Some("wasapi")).may_answer_differently());
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
            ao: None,
            audio_out_params: None,
        };
        assert_eq!(
            the_engine_truth(7, 0, &answers),
            "engine truth mpv pid 7 volume 0 aid no ao - audio-out-params - verdict no-track"
        );

        // And a device that would not open names itself, with the track it never played.
        let answers = EngineAnswers {
            aid: Some("1".to_string()),
            ao: None,
            audio_out_params: None,
        };
        assert!(
            the_engine_truth(9, 40, &answers).ends_with("verdict no-device"),
            "a launch that gave a level and a playlist that opened no output is the device's road"
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
                  volume=40 gain=none pinned 0"
                    .to_string()
            ],
            "one line per launch, and the form is the one the arguments are built from"
        );
        clear_noted();
    }
}

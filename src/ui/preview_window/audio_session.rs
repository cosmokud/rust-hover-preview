//! The machine's own account of a player's sound: whether Windows has an audio session for the
//! player at all, the state that session is in, the slider the mixer is showing beside the level the
//! player is at, and the endpoint the session is on against the endpoint the room is listening through.
//!
//! **Why the OS is asked at all.** A traced run captured the fault, and every answer the player could
//! give said the film had its sound: mpv's own `ao=v` log records the device it selected
//! (`Speakers (Realtek(R) Audio)`), `Trying stereo float (32/32 bits) @ 48000hz (shared) -> ok`, a
//! buffer allocated, and no error, no retry and no `null` anywhere after it; its `aid` names a track,
//! its `current-ao` is `wasapi` with params beside it, and its own `volume` is 100 with `mute` off —
//! byte for byte the six answers of twelve launches in that same run the user did hear. And the user
//! reported that Windows' Volume Mixer listed **no session for the player** while the film was silent.
//!
//! The player's opinion is therefore exhausted; what is left is the OS's view of the stream, which
//! nothing in this app read until now. It is a different kind of evidence from anything on the
//! player's pipe, and it is the only kind that can tell the two silences apart: a stream that never
//! reached an endpoint at all, and one an endpoint took and then dropped under the player.
//!
//! **The five questions one walk answers, and what each is for:**
//!
//! 1. **Is there a session for this pid on this machine's render endpoints?** Every active render
//!    endpoint is walked and every session on each is asked its process id, because a session
//!    belongs to the process that made it — a child process's session is the child's own pid, and the
//!    players here are children. The count found is written whether or not it is one, so `session
//!    none` is a first-class answer rather than an absence (see `the_line`). **And a miss is counted
//!    against the player's own tree before it is written**: the mixer can group a child process's
//!    session under its parent, so the sessions of processes the player itself started are counted
//!    too, and `children 0` is what makes `none` mean nothing was ever started rather than something
//!    was started where this walk did not look (see [`the_children_among`]).
//! 2. **The session identifier**, which is what says whether the session is the player's own or one
//!    inherited from a process that made it first (see [`the_executable_named`], which is also where
//!    the machine's own directory layout is kept out of the file).
//! 3. **The session's state** — `Active`, `Inactive` or `Expired`. **An `Expired` session under a
//!    player that believes it opened its device is the single most valuable datum this whole
//!    investigation can produce**: it is the OS saying the stream is over while the player is still
//!    writing to it, which is a fault no answer on the player's own pipe can express.
//! 4. **The mixer's own slider and mute switch**, which are the mixer's and not the player's: mpv at
//!    `volume=100 mute no` says nothing at all about a session whose `ISimpleAudioVolume` is at
//!    nothing or muted. The two are written with two different words — `mixer` here and `volume` on
//!    the engine's line — because they are two numbers on two scales (a share of full against mpv's
//!    percent), and a film can be at 100 on one and at nothing on the other. **And the per-channel
//!    volumes beside it** (see [`Channels`]): the same session exposes `IChannelAudioVolume`, which is
//!    where a level handed to the driver shows up rather than a level applied inside the player.
//! 5. **The endpoint it is on, by name and id, beside the default endpoint read at that same
//!    instant.** Two equal pairs say the stream is on the device the user is listening through; two
//!    different ones say it landed on a device nobody is listening to, or that the default endpoint
//!    was swapped underneath it — both of which are silent films that every other line calls sound.
//!
//! **The level the driver was handed is a third number again, and it is the one nothing in this app
//! read before.** mpv's default is `--softvol=no`, so the level it was given is not applied inside the
//! player: it is handed down to the driver as a per-stream volume, and the `volume` on the engine's
//! line is mpv's own property rather than what the stream is at. The session's slider (`mixer`) is one
//! hand on that number and its per-channel volumes (`channels`) are the other, and a stream with the
//! slider at full and a channel at nothing is a silent film every other answer calls sound.
//!
//! **Nothing here is silent about a gap.** A value the OS would not give is `-`, a state this build has
//! no word for is `unknown`, an interface that would not be had is `refused`, a session that reports no
//! channel at all is `none`, and a walk that could not be made is `session refused` with the word for
//! what stopped it. The line exists to be read where nobody can reproduce the fault, so an answer that
//! is missing is written as missing rather than passed over.
//!
//! **What a line adds when the question is which session carries a mute.** The `session` line names the
//! player's session beside the other sessions the endpoints listed (see [`Neighbours`]), and both carry
//! the whole of who a session is rather than a name: `GetSessionIdentifier` and
//! `GetSessionInstanceIdentifier` entire — the two strings Windows keys its per-app volume and mute
//! memory by, so two sessions of one executable that are remembered differently are two identifiers that
//! differ somewhere in the middle — the image the pid is really running beside them, and, for each
//! neighbour, its pid, its image, its state, its mute and its slider. **A session a new one inherits a
//! mute from has to be on the same line as the victim**, and the one that matters most is a session whose
//! process has ended: it is written `dead` and never dropped (see [`Neighbour`]). A cured session is
//! written with the same identity and the same neighbours, so a fault and a clean film are compared field
//! for field inside one run (see `a_cure_line`). **Both additions are gathered only where the trace asked
//! for them**: an untraced walk reads no identity and no neighbour, so a run nobody is diagnosing pays
//! nothing for the field (see `each_session`).
//!
//! **What this was built as a reader for, the run proved, and it is a fault with a cure.** Of eighteen
//! players in one captured run, two had a session whose `ISimpleAudioVolume` was **muted** while mpv
//! reported `volume 94 mute no` — two switches, and the player can see only its own. Both of those
//! sessions were at `mixer 1.000` with `channels 1.000 1.000`: nothing a hand could have moved was
//! down, the film had no sound, and the mixer showed the user no audible session for it. So a mute
//! switch found on under a player this app launched and is guarding is **taken off** (see
//! `look_and_cure`), and it is the only write this module makes: no master volume, no per-channel
//! volume and no `IAudioEndpointVolume` is ever set, because a level the user set is not this app's
//! to put back.
//!
//! **A slider at nothing is a hand, and the mute beside it is part of that hand.** `mixer 0.000` is
//! somebody's deliberate setting in Windows' own mixer, and restoring it would overwrite a decision
//! this app has no business with — so that session is written down and left exactly as it is (see
//! `Warrant::LeaveIt`). The line this module draws is the mute switch, which a captured run found on
//! under a player that could not see it, against the slider, which is the user's.
//!
//! **COM here is this thread's own apartment, and nowhere else.** The preview loop already holds an
//! apartment for the engines that draw into it (see `pdf_preview::initialize_apartment` and the head
//! of `run_preview_window`), and the trace's other thread reads a player's pipe and holds none — so
//! the probe begins a thread of its own, and the guard's look takes an apartment on the probe thread
//! it is called from (`look_and_cure`): both take `COINIT_MULTITHREADED` and give it back in the same
//! call, and neither touches the loop's, which holds one for the engines that draw into it.
//!
//! **The line is what `RHP_PLAYBACK_TRACE` asks for; the cure is not.** A traced launch begins a
//! thread and walks the mixer once to write down what it found (see `watch_the_mixer`); a **guarded**
//! player's session is looked at for every read the playback guard already makes — after the same
//! engine ask, on the same probe thread, at no other time and for no other player — and that look runs
//! whether or not the variable is set, because the trace is the recorder and this is the fix (see
//! `audio_guard::spawn_probe`). A run with neither a trace nor a guarded player pays a read of an
//! atomic and nothing else.

use std::time::Duration;

use windows::core::{Interface, GUID, PWSTR};
use windows::Win32::Media::Audio::{
    eConsole, eRender, AudioSessionState, AudioSessionStateActive, AudioSessionStateExpired,
    AudioSessionStateInactive, IAudioSessionControl, IAudioSessionControl2,
    IAudioSessionEnumerator, IAudioSessionManager2, IChannelAudioVolume, IMMDevice,
    IMMDeviceEnumerator, ISimpleAudioVolume, MMDeviceEnumerator, DEVICE_STATE_ACTIVE,
};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, CLSCTX_ALL,
    COINIT_MULTITHREADED, STGM_READ,
};
use windows::Win32::UI::Shell::PropertiesSystem::PROPERTYKEY;

use super::playback_trace::{self, Verdict};

/// How long a launched player is given before the mixer is first walked.
///
/// **A session exists from the moment a stream is activated, and a player activates it after it has
/// opened its file**, so a walk made at the launch would answer `none` about a player that is simply
/// not there yet — a false finding is worse here than a late one. It is the same kind of wait the
/// trace gives the player's own answers and a little longer, because the session is made after the
/// output that needs it (see `playback_trace::ENGINE_SETTING_UP`).
const MIXER_SETTLING: Duration = Duration::from_millis(700);

/// How many times the mixer is walked, and how long between two walks.
///
/// **Bounded the way the pipe reads are**, and for the same reason: a walk happens once and its
/// answer is written, so the worst case is this many walks and two waits. The walks stop as soon as
/// one finds the player's session, and a machine whose mixer really does not list it is given the
/// whole of the wait before `none` is written — a miss is worth being sure about.
const MIXER_WALKS: usize = 3;
const BETWEEN_WALKS: Duration = Duration::from_millis(300);

/// How many of a session's channels are written out.
///
/// A stereo stream is two and the widest layout Windows plays is eight; a session that reports more
/// than this has the rest of them counted rather than written (`channels … of <count>`), so one line
/// stays one line on any machine while saying that there are more channels than it names.
const CHANNELS_WRITTEN: u32 = 8;

/// How many of the sessions listed beside the player's are written out.
///
/// A machine with a handful of players on it lists a handful; one with every application the user has
/// open lists more, and a line that named all of them would be a page rather than a line. **What the cap
/// leaves out is counted rather than dropped silently** (`neighbours 8 of 14`), so a carrier can never be
/// hidden by it without the line saying so; the ones written are the first the endpoints listed, in the
/// walk's own order, which is the order every other field of the line is read in.
const NEIGHBOURS_WRITTEN: usize = 8;

/// Read the mixer's side of a launched film, on a thread of its own, and write what it says.
///
/// Nothing at all happens where the trace is not asked for (see the module note): no thread, no
/// apartment and no COM, which is what makes this free on a run nobody is diagnosing.
pub(super) fn watch_the_mixer(pid: u32) {
    if !playback_trace::asked() {
        return;
    }

    let _ = std::thread::Builder::new()
        .name("playback-mixer".to_string())
        .spawn(move || {
            // **The apartment is taken here and nowhere else** (see the module note). A refusal is an
            // answer like any other and is written as one; a thread that already has an apartment
            // must still give it back below, so only a refusal skips the `CoUninitialize`.
            let taken = unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) };
            if taken.is_err() {
                let refused = Probe {
                    refused: Some(Refusal::Apartment),
                    ..Probe::default()
                };
                playback_trace::write(&the_line(pid, &refused));
                return;
            }

            std::thread::sleep(MIXER_SETTLING);

            let mut probe = walk_the_mixer(pid);
            for _ in 1..MIXER_WALKS {
                // **A walk that found the player's session has the answer**, and a thread with no
                // apartment will not have one on the next walk either — both are done. A walk that
                // could not reach the audio service is walked again, because that is exactly the case
                // a second walk can answer differently: the service is coming up with the player.
                if probe.theirs.is_some() || probe.refused == Some(Refusal::Apartment) {
                    break;
                }
                std::thread::sleep(BETWEEN_WALKS);
                probe = walk_the_mixer(pid);
            }

            playback_trace::write(&the_line(pid, &probe));

            // Safety: this thread's apartment is the one taken at the top of it, and nothing of the
            // walk outlives the function that made it — every interface is dropped inside
            // `walk_the_mixer`, before this line.
            unsafe { CoUninitialize() };
        });
}

/// What one walk of the machine's render endpoints said about a player.
///
/// Every part of it is something the line can be written from, and nothing of it holds an interface:
/// the walk is over before this is handed on, so the thread's apartment is still the only COM state
/// anything here depends on.
#[derive(Default)]
struct Probe {
    /// How many sessions the machine's render endpoints listed between them, whoever they belong to.
    sessions: usize,
    /// How many of those belonged to the player's pid.
    matching: usize,
    /// The process each listed session belonged to, where one could be read: what a miss is counted
    /// against the player's own tree by (see [`the_children_among`]).
    listed: Vec<u32>,
    /// How many of those sessions belong to a process the player itself started, counted only where
    /// none of them is the player's own (see [`the_children_among`]).
    children: usize,
    /// The first of those that is the player's, and the only one whose facts are written.
    theirs: Option<Theirs>,
    /// The sessions the endpoints listed beside the player's, which is what a mute a new session
    /// inherits would have come from (see [`Neighbours`]).
    neighbours: Neighbours,
    /// The default render endpoint at the instant of the walk, which is the half the session's own
    /// endpoint is compared against.
    default: Option<Endpoint>,
    /// What made the walk impossible, where something did — COM refused, the enumerator refused, or
    /// no endpoint could be listed at all.
    refused: Option<Refusal>,
}

/// What one session says about the player it belongs to.
struct Theirs {
    /// `Active`, `Inactive` or `Expired`, as the OS reports it.
    state: State,
    /// Who the session is, whole, rather than the name at the end of one of its identifiers (see
    /// [`Identity`]).
    identity: Identity,
    /// The mixer's own slider, as a share of full — **not** the level the player was given.
    volume: Option<f32>,
    /// The per-channel volumes of the stream, which is where a level handed to the driver shows up.
    channels: Channels,
    /// The mixer's own mute switch, which silences a film at any level at all and is the player's
    /// own mute's independent twin.
    mute: Option<bool>,
    /// The endpoint the session was listed on.
    endpoint: Endpoint,
}

/// **Who Windows believes a session is, which is the whole of what its remembered volume and mute are
/// keyed by.**
///
/// The mixer keeps a session's slider and its mute per application, and it reads an application off these
/// two identifiers rather than off the executable's name: two sessions of one `mpv.exe` that are
/// remembered differently are two identifiers that differ somewhere in the middle — a device path, or the
/// instance's own guid — which is exactly the difference a line that printed only `mpv.exe` could never
/// show. So both are written entire, and the image the pid is actually running beside them, so that an
/// identifier can be read against a real file rather than against Windows' own device spelling of one.
///
/// **Nothing here is a write and nothing here is required**: a session that will not hand over an
/// identifier has `-` where it would be, and the line keeps the rest.
#[derive(Default, Clone, PartialEq, Debug)]
pub(super) struct Identity {
    /// The executable the session identifier ends in, where it names one (see [`the_executable_named`])
    /// — the word the line carried before there was an identity to carry.
    identifier: Option<String>,
    /// `IAudioSessionControl2::GetSessionIdentifier` as Windows gave it, whole.
    session: Option<String>,
    /// `IAudioSessionControl2::GetSessionInstanceIdentifier` as Windows gave it, whole — the same
    /// endpoint and image with the instance's own guid beside them.
    instance: Option<String>,
    /// The executable the session's pid is running, by its whole path (see
    /// `app::engine_processes::image_path_of`).
    image: Option<String>,
}

/// One of the sessions the endpoints listed beside the player's own.
///
/// **Named and not described**, because the hypothesis this is written for is a negative one: a session
/// left behind by a process that has ended, still muted, is what a new session inherits from — so the one
/// neighbour that matters most is the one an ordinary walk would skip, and the word for it is `dead` (see
/// `the_neighbour_word`). Nothing is read of a neighbour that a line does not say.
#[derive(Clone, PartialEq, Debug)]
struct Neighbour {
    /// The pid the mixer listed the session under.
    pid: u32,
    /// The executable its identifier names, where it names one — the tail, as the player's own line
    /// writes it.
    identifier: Option<String>,
    /// `Active`, `Inactive` or `Expired`, as the OS reports it.
    state: State,
    /// The mixer's own mute switch for that session.
    mute: Option<bool>,
    /// The mixer's own slider for it, as a share of full.
    volume: Option<f32>,
    /// Whether the process the session was listed under still exists: `Some(false)` where the mixer
    /// lists a session for a pid that has ended, and `None` where the machine would not list its
    /// processes at all.
    gone: Option<bool>,
}

/// The sessions the endpoints listed beside the player's own, which is what a new session inherits a mute
/// from if anything does.
///
/// **A carrier has to be visible on the same line as the victim**, which is the whole of why this is
/// gathered: `found 1 of 3` says a player was listed among others and says nothing about who they were.
/// Nothing here is read where no trace asked for it (see `each_session`), so a guarded player's look on a
/// run nobody is diagnosing pays nothing for the field.
#[derive(Default, Clone, PartialEq, Debug)]
pub(super) struct Neighbours {
    /// How many sessions the endpoints listed beside the player's, whoever they belong to — the count
    /// the cap above is written against, and the whole of it where the cap left any out.
    listed: usize,
    /// The first [`NEIGHBOURS_WRITTEN`] of them, in the order the walk listed them.
    written: Vec<Neighbour>,
}

/// The per-channel volumes a session reports, which are the ones a driver-level handover moves.
///
/// **Why this is a number apart from the session's slider**: mpv runs at `--softvol=no` by default, so
/// the level this app hands it is not applied inside the player at all — it goes down to the driver as
/// a per-stream volume, and neither mpv's own `volume` nor the mixer's slider is that number. A stream
/// whose slider reads full with one of its channels at nothing is a silent film whose every other
/// answer says sound.
///
/// **An interface that would not be had is a word of its own** rather than a missing field: `refused`
/// says the per-stream level was not read, and it can never be mistaken for a stream at nothing.
enum Channels {
    /// One share of full per channel the session reported, and `None` for a channel it would not
    /// answer for; `count` is how many it says it has, which can be more than were asked for.
    Read {
        volumes: Vec<Option<f32>>,
        count: u32,
    },
    /// The session would not hand over `IChannelAudioVolume`, or would not say how many channels it
    /// has, so the per-stream level is unread.
    Refused,
}

/// One endpoint as a line names it: the name Windows shows the user, and the id the mixer keys its
/// sessions on.
#[derive(Default)]
struct Endpoint {
    name: Option<String>,
    id: Option<String>,
}

/// The state a session is in, as this app reads it.
///
/// A value of its own rather than the crate's `AudioSessionState`, because the reading of it is the
/// part worth testing and a test should not have to build a COM type to say what `Expired` is.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum State {
    Active,
    Inactive,
    Expired,
    Unknown,
}

impl State {
    /// The state the OS reported, and `Unknown` for a value this build has no word for.
    fn of(state: AudioSessionState) -> Self {
        if state == AudioSessionStateActive {
            Self::Active
        } else if state == AudioSessionStateInactive {
            Self::Inactive
        } else if state == AudioSessionStateExpired {
            Self::Expired
        } else {
            Self::Unknown
        }
    }
}

/// What stopped a walk from being made at all, which is written down rather than passed over: a probe
/// that says nothing about its own failure reads as a machine with no session, and those are two
/// different findings (see the module note).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum Refusal {
    /// The thread could not take an apartment, so no COM call was ever made.
    Apartment,
    /// The device enumerator itself could not be made.
    Enumerator,
    /// The render endpoints could not be listed.
    Endpoints,
}

impl Refusal {
    /// The word a line names a refusal by.
    fn word(self) -> &'static str {
        match self {
            Self::Apartment => "apartment",
            Self::Enumerator => "enumerator",
            Self::Endpoints => "endpoints",
        }
    }
}

/// Where a slider stops being a hand's and starts being this app's fault.
///
/// **A session muted with its slider at nothing is a setting somebody made** in the mixer's own
/// window, and the mute beside it is part of that setting rather than a fault under a player that
/// cannot see it. The line writes a slider to three places (see `the_mixer_word`), so a slider that
/// would be written `0.000` is one at nothing, and it is a hand.
const A_SLIDER_UP: f32 = 0.0005;

/// What a muted session warrants, which is the whole of the rule this module writes by.
///
/// **A pure function of three answers** — the mixer's slider, the mixer's mute switch, and the level
/// the launch asked for — so that what this module will and will not write can be read and tested with
/// no mixer on the machine (see `the_warrant`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Warrant {
    /// Nothing: the switch is off, the session would not say which way it is, or the launch asked for
    /// no audio at all — a silence the setting chose is not a fault to heal.
    Nothing,
    /// The switch is on with the slider up: the fault a captured run proved, and the one write this
    /// module makes.
    TakeItOff,
    /// The switch is on with the slider at nothing: a hand, written down and left (see
    /// [`A_SLIDER_UP`]).
    LeaveIt,
}

/// What came of one look at a session's mute switch, in the words a line is read by.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum Cure {
    /// Nothing was written, because nothing was warranted.
    Nothing,
    /// The mute was on and the mixer took `SetMute(false)`.
    Unmuted,
    /// `SetMute(false)` was refused, or the session went before it could be asked — **a failure that is
    /// written down rather than swallowed**, since the film is still silent.
    Refused,
    /// The mute is a hand's, at a slider that is down: written down and left as it was.
    TurnedDown,
}

impl Cure {
    /// The word a read-back names a cure by, and nothing at all where nothing was warranted.
    fn word(self) -> &'static str {
        match self {
            Self::Nothing => "",
            Self::Unmuted => "cured",
            Self::Refused => "cure-failed",
            Self::TurnedDown => "left-alone",
        }
    }
}

/// What one look at a guarded player's session came to (see `look_and_cure`).
///
/// **Not `Copy` any more, and it is the identity that made it so**: the two identifiers and the sessions
/// beside a player are strings and a list, and the look is moved into the report rather than handed about.
#[derive(Clone, PartialEq, Debug)]
pub(super) enum Seen {
    /// No session for this pid on any of the machine's render endpoints.
    Absent,
    /// Windows would not list its sessions at all, so the question was never put to one.
    Refused(Refusal),
    /// The player's session, the mixer's slider and mute switch as they stood, and what that switch
    /// warranted.
    There {
        mixer: Option<f32>,
        mute: Option<bool>,
        cure: Cure,
        /// The whole of who the session is, which is what a mute remembered against it is keyed by —
        /// read where a trace asked for it and left empty where none did (see [`Identity`]).
        identity: Identity,
        /// The sessions the endpoints listed beside this one, which is where a mute the player cannot see
        /// would have come from (see [`Neighbours`]).
        neighbours: Neighbours,
    },
}

impl Seen {
    /// A look that found the player's session in the state given, with the identity and the neighbours
    /// still to be filled in by the walk that found it (see `look_on_this_apartment`).
    pub(super) fn there(mixer: Option<f32>, mute: Option<bool>, cure: Cure) -> Self {
        Self::There {
            mixer,
            mute,
            cure,
            identity: Identity::default(),
            neighbours: Neighbours::default(),
        }
    }

    /// **The same look with the two answers only the end of the walk can give**: the whole identity of
    /// the session, and the sessions the endpoints listed beside it. They are attached here rather than
    /// read into the look where the session was found, because the neighbour list is still being filled
    /// for as long as the walk is running.
    fn with_identity_and_neighbours(mut self, identity: Identity, neighbours: Neighbours) -> Self {
        if let Self::There {
            identity: theirs,
            neighbours: beside,
            ..
        } = &mut self
        {
            *theirs = identity;
            *beside = neighbours;
        }

        self
    }

    /// **What a player's verdict comes to with the machine's own answer beside it**: the third
    /// silence, and the reason this module is asked at every read.
    ///
    /// A session that is muted and still muted is `Verdict::SessionMuted` — the film every answer on
    /// the player's pipe calls sound. A session whose mute the cure took off is `Heard`, because that
    /// is what the film is now: the write was the whole of what was wrong with it. And a look that
    /// said nothing about a mute — no session, a walk that refused, a switch the mixer would not read —
    /// leaves the player's own verdict exactly as it was, which is the rule
    /// `EngineAnswers::an_output_opened` keeps for an unknown, and for the same reason.
    ///
    /// **A player that is already a fault of its own keeps it**: `no-track`, `no-device` and the
    /// player's own `silent` are the ladder's road, and a session cannot soften one of them.
    pub(super) fn the_verdict_of(&self, engine: Verdict) -> Verdict {
        if engine != Verdict::Heard {
            return engine;
        }

        match self {
            Self::There {
                mute: Some(true),
                cure,
                ..
            } if *cure != Cure::Unmuted => Verdict::SessionMuted,
            _ => engine,
        }
    }
}

/// What a session's mute switch warrants, given the level the launch asked for.
///
/// **The launch's level is half the rule**, for the reason it is half the verdict: a launch that
/// resolved to nothing handed its player no track at all (see `video_launch::AudioForm`), so there is
/// no session of this app's making to have muted — and a mute on a player begun without audio is
/// nobody's but the mixer's.
///
/// **A mute the mixer would not report is not a mute**, and neither is a session whose slider could
/// not be read at all: a switch this build cannot read is not a switch to throw.
fn the_warrant(mixer: Option<f32>, mute: Option<bool>, asked: u32) -> Warrant {
    if super::video_launch::mpv_launched_without_audio(asked) || mute != Some(true) {
        return Warrant::Nothing;
    }

    if mixer.is_some_and(|mixer| mixer < A_SLIDER_UP) {
        return Warrant::LeaveIt;
    }

    Warrant::TakeItOff
}

/// The key Windows keeps an endpoint's user-visible name under.
///
/// The crate binds this key only behind a feature this app has no other use for
/// (`Devices_FunctionDiscovery`), so it is written out here by the set's canonical id and the
/// property id Windows documents — the way the listing's own four columns are (see
/// `explorer_hook::view_sort`).
const PKEY_DEVICE_FRIENDLY_NAME: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xa45c254e_df1c_4efd_8020_67d146a850e0),
    pid: 14,
};

/// Every session the machine's render endpoints list, handed to `look` with the pid it belongs to and
/// the endpoint it was listed on.
///
/// **One walk and two questions**: the probe line (see `walk_the_mixer`) and the guard's look (see
/// `look_and_cure`) are the same enumeration with a different reader behind it, and it is written once
/// so that neither can drift from the other — a session found for one is found for both, and the
/// endpoints are walked in the same order for a line as for a write.
///
/// **Every endpoint and not only the default one**, because a stream that landed on a device the user
/// is not listening to is one of the answers this module exists to give, and a walk of the default
/// endpoint alone could never see one. `default` is filled with the default render endpoint as it
/// stood at the top of the walk, before anything is enumerated, so that what a line says about it is
/// the device the room was listening through at that instant.
///
/// `look` answers whether it has found what it came for: a `true` ends the walk where it stands, which
/// is what a look for one player wants, and is why a look that is also tracing goes on — the sessions
/// beside the player's are half of what its line says (see [`Neighbours`]). The count of sessions the
/// endpoints listed is handed back whoever they belong to, and so is what stopped the walk.
///
/// **`neighbours` is `Some` only where a line will name them**, and `None` is the untraced walk this
/// module has always made: no process is listed, no identifier read and no session asked anything it was
/// not asked before. Where it is `Some`, every session the endpoints list that is not the player's own
/// is counted, and the first [`NEIGHBOURS_WRITTEN`] of them are read for the five answers a neighbour is
/// written with.
///
/// Safety: the caller's thread must hold an apartment, and nothing `look` keeps may outlive the call —
/// every interface here is dropped before this returns.
unsafe fn each_session(
    default: &mut Option<Endpoint>,
    pid: u32,
    neighbours: &mut Option<Neighbours>,
    look: &mut dyn FnMut(u32, &IAudioSessionControl, &IMMDevice) -> bool,
) -> Walked {
    let mut walked = Walked::default();

    let enumerator: IMMDeviceEnumerator =
        match CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) {
            Ok(enumerator) => enumerator,
            Err(_) => {
                walked.refused = Some(Refusal::Enumerator);
                return walked;
            }
        };

    *default = enumerator
        .GetDefaultAudioEndpoint(eRender, eConsole)
        .map(|device| the_endpoint(&device))
        .ok();

    let endpoints = match enumerator.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE) {
        Ok(endpoints) => endpoints,
        Err(_) => {
            walked.refused = Some(Refusal::Endpoints);
            return walked;
        }
    };

    let count = endpoints.GetCount().unwrap_or(0);

    // **The processes the machine is running, read once for the whole walk** — and only where a line
    // asked to name the sessions beside the player's, which is after the endpoints were listed so that a
    // walk that stops here has paid nothing for them. A session listed under a pid that has ended is the
    // neighbour a mute is inherited from, and it is asked of one snapshot rather than of a handle per
    // session: an id this app may not open is not a process that has ended (see
    // `app::engine_processes::running`).
    let live = if neighbours.is_some() {
        super::engine_processes::running()
    } else {
        None
    };

    for index in 0..count {
        let Ok(device) = endpoints.Item(index) else {
            continue;
        };
        let Ok(sessions) = the_sessions_of(&device) else {
            continue;
        };
        let listed = sessions.GetCount().unwrap_or(0);

        for at in 0..listed {
            let Ok(control) = sessions.GetSession(at) else {
                continue;
            };
            walked.sessions += 1;

            // **`GetProcessId` is the whole of the match**, and it is the second generation of the
            // interface that answers it: a session belongs to the process that made it whatever its
            // identifier says, and a session a process inherited from its parent carries the parent's
            // pid as readily as the parent's image (see `the_executable_named`, which is the
            // identifier's half of the same question).
            //
            // **It is read for every session and not only for a match**, because a miss counts the
            // sessions the player's own children made (see `the_children_among`), and because a look
            // for a player cannot be answered by a session whose pid is not said.
            let Ok(theirs) = control.cast::<IAudioSessionControl2>() else {
                continue;
            };
            let Ok(session_pid) = theirs.GetProcessId() else {
                continue;
            };

            // **The sessions beside the player's, read here so that both walks say them the same way**,
            // and never the player's own: they are put down before the look is handed the session, so a
            // look that ends the walk at the player cannot cut the list short.
            if session_pid != pid {
                if let Some(neighbours) = neighbours.as_mut() {
                    let gone = live.as_deref().map(|live| !live.contains(&session_pid));
                    if neighbours.written.len() < NEIGHBOURS_WRITTEN {
                        neighbours
                            .written
                            .push(the_neighbour(&control, session_pid, gone));
                    }
                    neighbours.listed += 1;
                }
            }

            if look(session_pid, &control, &device) {
                return walked;
            }
        }
    }

    walked
}

/// What one walk of the machine's endpoints came to, whoever its sessions belong to.
#[derive(Default)]
struct Walked {
    /// How many sessions the render endpoints listed between them.
    sessions: usize,
    /// What stopped the walk, where something did.
    refused: Option<Refusal>,
}

/// Walk every render endpoint this machine has and answer what the mixer says about one player.
///
/// **Every read it makes is a read** (see `each_session`): the probe describes and never writes, and the
/// one write in this module is the guard's, made by `look_and_cure` on a session it found for a player
/// this app is standing behind.
fn walk_the_mixer(pid: u32) -> Probe {
    let mut probe = Probe::default();
    // **The sessions beside the player's only where a line will name them**, which is what the trace is
    // (see `each_session`).
    let mut neighbours = playback_trace::asked().then(Neighbours::default);

    // Safety: every interface the walk holds is one `each_session` makes and drops before it returns,
    // on a thread whose apartment its caller took. Nothing here is written to a session.
    unsafe {
        let walked = each_session(
            &mut probe.default,
            pid,
            &mut neighbours,
            &mut |session_pid, control, device| {
                probe.listed.push(session_pid);

                if session_pid != pid {
                    return false;
                }

                probe.matching += 1;
                if probe.theirs.is_none() {
                    probe.theirs = Some(the_session(control, device, session_pid));
                }

                // **Every session of this pid is counted and the first one's facts are the ones
                // written**, so the walk goes on to the count rather than stopping at the first match.
                false
            },
        );
        probe.sessions = walked.sessions;
        probe.refused = walked.refused;
    }

    probe.neighbours = neighbours.unwrap_or_default();

    // **A miss is counted against the player's own tree before it is written**, and only a miss: a
    // session that is the player's own already answers the question the count is for.
    if probe.matching == 0 {
        probe.children = the_children_among(&probe.listed, pid);
    }

    probe
}

/// **Look at the session Windows has for one guarded player, and take its mute off where the rule says
/// a mute is this app's business.**
///
/// It is called for every read the playback guard already makes and for no other reason, so a run
/// without a trace walks the mixer exactly as often as it reads a player — which is the whole of what
/// the cure costs (see `audio_guard::spawn_probe`), and why nothing here is on a schedule of its
/// own. **A traced look reads one thing more and writes nothing more**: it goes on past the session it
/// found until the endpoints are walked out, so that the sessions beside the player's can be written on
/// the line a cure makes, and it reads a session's whole identity while that session is still in hand
/// (see [`Identity`] and [`Neighbours`]) — the writes it makes are the same one, or none.
///
/// **The apartment is the calling thread's and lasts for this call.** The preview loop holds one of its
/// own for the engines that draw into it, so the only thread this may be called on is a probe thread of
/// the guard's; nothing of the walk outlives the call, and nothing here touches an interface another
/// thread's apartment made.
///
/// A look that finds no session for the pid answers `Seen::Absent` rather than nothing, and one that
/// could not be made at all answers the part of the walk that refused — because a look that said
/// nothing would read as a film with no session, and those are two different findings.
pub(super) fn look_and_cure(pid: u32, asked: u32) -> Seen {
    // Safety: the apartment is taken here and given back at the end of this same call, whatever the
    // walk found — the pattern `watch_the_mixer` keeps on its own thread, and for the same reason. A
    // thread that already has an apartment still gives this one back, so only a refusal skips it.
    let taken = unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) };
    if taken.is_err() {
        return Seen::Refused(Refusal::Apartment);
    }

    let seen = unsafe { look_on_this_apartment(pid, asked) };

    // Safety: this thread's apartment is the one taken five lines up, and nothing of the walk outlives
    // the call that made it.
    unsafe { CoUninitialize() };
    seen
}

/// The walk and the one write, on an apartment its caller has already taken.
///
/// Safety: this thread must hold an apartment, and the session found is dropped with the walk.
unsafe fn look_on_this_apartment(pid: u32, asked: u32) -> Seen {
    let traced = playback_trace::asked();
    let mut found = None;
    let mut identity = Identity::default();
    let mut neighbours = traced.then(Neighbours::default);

    let walked = each_session(
        &mut None,
        pid,
        &mut neighbours,
        &mut |session_pid, control, _device| {
            if session_pid != pid {
                return false;
            }

            // **The first session of this pid is the one the player's sound is on**, and it is the one
            // this module finds the player by everywhere (see `walk_the_mixer`). A walk nobody traces
            // ends there; one a trace is reading goes on, because the sessions beside this one are half
            // of what its line says (see [`Neighbours`]).
            if found.is_some() {
                return !traced;
            }

            // **Who the session is, read while it is still in hand and only where a line will say it**
            // (see [`Identity`]).
            if traced {
                identity = the_identity_of(control, session_pid);
            }
            found = Some(the_cure_of(control, asked));

            !traced
        },
    );

    match (found, walked.refused) {
        (Some(seen), _) => {
            seen.with_identity_and_neighbours(identity, neighbours.unwrap_or_default())
        }
        (None, Some(why)) => Seen::Refused(why),
        (None, None) => Seen::Absent,
    }
}

/// One session as the cure reads it: the mixer's slider, its mute switch, and what that switch
/// warranted — with the one write this module makes, where the rule asks for it.
///
/// **`SetMute(false)` and nothing else**: not `SetMasterVolume`, not `IChannelAudioVolume` and not
/// `IAudioEndpointVolume`. The slider and the per-channel volumes are a hand's or a handover's, and the
/// mute switch is the one a captured run found on under a player that answered `mute no` — two
/// switches, and the film had no sound through the half the player could not see.
///
/// **An interface that would not be had is `Cure::Nothing`** rather than a guess: a session this build
/// cannot read the mute of is not a session to throw a switch on, and the line says `mute -`.
///
/// Safety: the session must be one this thread holds, on the apartment its caller took. The write is
/// made with no event context, which is what a write nobody waits on the mixer's answer for uses.
unsafe fn the_cure_of(control: &IAudioSessionControl, asked: u32) -> Seen {
    let Ok(volume) = control.cast::<ISimpleAudioVolume>() else {
        return Seen::there(None, None, Cure::Nothing);
    };

    let mixer = volume.GetMasterVolume().ok();
    let mute = volume.GetMute().ok().map(|mute| mute.as_bool());
    let cure = match the_warrant(mixer, mute, asked) {
        Warrant::Nothing => Cure::Nothing,
        Warrant::LeaveIt => Cure::TurnedDown,
        Warrant::TakeItOff => match volume.SetMute(false, std::ptr::null()) {
            Ok(()) => Cure::Unmuted,
            Err(_) => Cure::Refused,
        },
    };

    Seen::there(mixer, mute, cure)
}

/// The session manager of one endpoint, which is where its sessions are listed from.
///
/// Safety: `device` must be an endpoint this thread holds; the manager is activated for the length of
/// the walk and dropped with the enumerator it came from.
unsafe fn the_sessions_of(device: &IMMDevice) -> windows::core::Result<IAudioSessionEnumerator> {
    device
        .Activate::<IAudioSessionManager2>(CLSCTX_ALL, None)
        .and_then(|manager| manager.GetSessionEnumerator())
}

/// Read everything one session says about the player it belongs to.
///
/// **Nothing here is required**: an identifier the OS would not give, a mixer that refuses its slider, a
/// per-stream level that could not be had and a state this build has no word for are each written as
/// `-`, `refused` or `unknown` rather than ending the line, because half of this answer is worth more
/// than none of it, and a session that names its level and not its state is a line a reader can still
/// use.
///
/// Safety: the session and the endpoint must both be ones this thread holds; the two volume interfaces
/// are asked of the session itself, which is the one object that has them, and are dropped with it.
unsafe fn the_session(control: &IAudioSessionControl, device: &IMMDevice, pid: u32) -> Theirs {
    let volume = control.cast::<ISimpleAudioVolume>().ok();

    Theirs {
        state: control.GetState().map_or(State::Unknown, State::of),
        identity: the_identity_of(control, pid),
        volume: volume
            .as_ref()
            .and_then(|mixer| mixer.GetMasterVolume().ok()),
        channels: the_channels_of(control),
        mute: volume
            .as_ref()
            .and_then(|mixer| mixer.GetMute().ok())
            .map(|mute| mute.as_bool()),
        endpoint: the_endpoint(device),
    }
}

/// **The whole of who a session is**: both identifiers Windows keys its remembered volume and mute by,
/// and the image the session's pid is really running (see [`Identity`]).
///
/// Safety: the session must be one this thread holds; each identifier is read and freed here, before the
/// string it was copied into is handed on, and the process is one the caller has no right to anything but
/// its path.
unsafe fn the_identity_of(control: &IAudioSessionControl, pid: u32) -> Identity {
    let theirs = control.cast::<IAudioSessionControl2>().ok();
    let session = theirs
        .as_ref()
        .and_then(|theirs| the_string_of(|| theirs.GetSessionIdentifier()));
    let instance = theirs
        .as_ref()
        .and_then(|theirs| the_string_of(|| theirs.GetSessionInstanceIdentifier()));

    Identity {
        identifier: session.as_deref().and_then(the_executable_named),
        session,
        instance,
        image: super::engine_processes::image_path_of(pid),
    }
}

/// One of the sessions the endpoints listed beside the player's, read for the five answers a line names
/// it by and nothing else (see [`Neighbour`]).
///
/// Safety: the session must be one this thread holds; its identifier is read and freed inside
/// [`the_identifier_of`] and no read here is a write.
unsafe fn the_neighbour(control: &IAudioSessionControl, pid: u32, gone: Option<bool>) -> Neighbour {
    let volume = control.cast::<ISimpleAudioVolume>().ok();

    Neighbour {
        pid,
        identifier: the_identifier_of(control),
        state: control.GetState().map_or(State::Unknown, State::of),
        mute: volume
            .as_ref()
            .and_then(|mixer| mixer.GetMute().ok())
            .map(|mute| mute.as_bool()),
        volume: volume
            .as_ref()
            .and_then(|mixer| mixer.GetMasterVolume().ok()),
        gone,
    }
}

/// The per-channel volumes of one session, which is the level the driver was handed.
///
/// `IChannelAudioVolume` is asked of the session object itself, as `ISimpleAudioVolume` is: the stream's
/// own interface is reachable only through the `IAudioClient` of the process that opened it, which this
/// app has no handle on, and the session is the object the OS exposes to anybody who can enumerate one.
///
/// Safety: the session must be one this thread holds; the interface is dropped with it and nothing of it
/// is written — `GetChannelCount` and `GetChannelVolume` are reads, and no `Set` on it is ever called.
unsafe fn the_channels_of(control: &IAudioSessionControl) -> Channels {
    let Ok(channels) = control.cast::<IChannelAudioVolume>() else {
        return Channels::Refused;
    };
    let Ok(count) = channels.GetChannelCount() else {
        return Channels::Refused;
    };

    Channels::Read {
        volumes: (0..count.min(CHANNELS_WRITTEN))
            .map(|index| channels.GetChannelVolume(index).ok())
            .collect(),
        count,
    }
}

/// How many of the sessions the endpoints listed belong to a process the player itself started.
///
/// **The mixer can group a child process's session under its parent**, so a player whose stream was made
/// by something it spawned leaves a session that belongs to its tree rather than to the pid that was
/// asked about. Counting those is what makes a miss unambiguous: `children 1` says something this player
/// is part of did reach an endpoint, and `children 0` says nothing of it ever did — two different faults
/// that were one word until now (see `the_line`).
///
/// It is walked from the machine rather than read from a record, because what a pid started is not
/// anything this app keeps (see `app::engine_processes::children_of`).
fn the_children_among(listed: &[u32], player: u32) -> usize {
    let children = super::engine_processes::children_of(player);

    listed
        .iter()
        .filter(|session_pid| children.contains(session_pid))
        .count()
}

/// The executable a session identifier names, where it names one.
///
/// Safety: the identifier is read off a session this thread holds and freed here, before the string it
/// was copied into is handed on.
unsafe fn the_identifier_of(control: &IAudioSessionControl) -> Option<String> {
    let theirs = control.cast::<IAudioSessionControl2>().ok()?;
    let identifier = the_string_of(|| theirs.GetSessionIdentifier())?;

    the_executable_named(&identifier)
}

/// One string Windows hands over in a block the caller owns, copied out and freed.
///
/// Safety: `read` must be a call that hands back a `PWSTR` in the caller's task allocator — every
/// identifier and endpoint id in this module does — and that block is freed here, before this returns.
unsafe fn the_string_of(read: impl FnOnce() -> windows::core::Result<PWSTR>) -> Option<String> {
    let raw = read().ok()?;
    let text = raw.to_string().ok();
    CoTaskMemFree(Some(raw.0 as *const _));

    text
}

/// The executable a session identifier names, and nothing where it names none.
///
/// **A session identifier is the endpoint's id, a `|#`, and the image the session is keyed to**:
/// `…|#\Device\HarddiskVolume4\…\mpv.exe%b{…}`. The file name at the end of that image is the whole of
/// what a line needs from it — a session keyed to `mpv.exe` is the player's own, and one keyed to the
/// player's parent was inherited from it — and **the device path before it is not written at all**: a
/// machine's own directory layout is not a fact a diagnostic log has any use for, and the log is a file
/// that gets sent to somebody. It is the same rule a film's own name keeps (see
/// `pin_trace::path_name`).
///
/// An identifier that names no program at all — the machine's own sounds, or a session grouped under an
/// application's own id — answers nothing, and the line writes `-` for it.
fn the_executable_named(identifier: &str) -> Option<String> {
    let tail = identifier.rsplit(['\\', '/']).next()?;
    let name = tail.split('%').next().unwrap_or_default().trim();

    if name.to_ascii_lowercase().ends_with(".exe") {
        Some(name.to_string())
    } else {
        None
    }
}

/// Read one endpoint's name and id.
///
/// Safety: `device` must be an endpoint this thread holds; the id comes back in a block this function
/// frees, and the name out of a property store that is dropped with it.
unsafe fn the_endpoint(device: &IMMDevice) -> Endpoint {
    Endpoint {
        name: the_name_of(device),
        id: the_id_of(device),
    }
}

/// The name Windows shows the user for an endpoint, which is what the user read in the mixer.
///
/// Safety: the store and the property are both the endpoint's own and are released here.
unsafe fn the_name_of(device: &IMMDevice) -> Option<String> {
    let store = device.OpenPropertyStore(STGM_READ).ok()?;
    let value = store.GetValue(&PKEY_DEVICE_FRIENDLY_NAME).ok()?;
    let name = value.to_string();
    let name = name.trim();

    if name.is_empty() {
        None
    } else {
        Some(name.to_string())
    }
}

/// The id the mixer keys an endpoint's sessions on, which is the half of the pair a locale cannot
/// change.
///
/// Safety: the string comes back in a block this function allocated through the caller's task
/// allocator, and is freed here after it has been copied out.
unsafe fn the_id_of(device: &IMMDevice) -> Option<String> {
    let raw = device.GetId().ok()?;
    let id = raw.to_string().ok();
    CoTaskMemFree(Some(raw.0 as *const _));
    id
}

/// The word a session's state is written as — **`expired` being the one this probe was built for**.
fn the_state_word(state: State) -> &'static str {
    match state {
        State::Active => "active",
        State::Inactive => "inactive",
        State::Expired => "expired",
        State::Unknown => "unknown",
    }
}

/// The mixer's slider as a share of full, to three places, or `-` where the mixer would not say.
///
/// A share and not a percent, deliberately: the engine's line writes mpv's own level as the percent it
/// is, and a session at `0.000` must never be read as the `0` of a player at `volume 0` — nor a player
/// at `100` as a mixer fully up.
fn the_mixer_word(volume: Option<f32>) -> String {
    volume.map_or_else(|| "-".to_string(), |volume| format!("{volume:.3}"))
}

/// A value Windows handed over, written in quotes so that the spaces inside a path stay inside the field
/// — **and not through `Debug`**, which would double every backslash of the one string an identity is read
/// for. A quote of Windows' own would end the field, so it is written as an apostrophe; neither an
/// identifier nor an image path has ever held one.
fn the_quoted(value: Option<&str>) -> String {
    match value {
        Some(value) => format!("\"{}\"", value.replace('"', "'")),
        None => "-".to_string(),
    }
}

/// **The whole of who a session is**, written after the `identifier` the line has always carried: both
/// identifiers Windows keys its remembered volume and mute by, entire, and the image the session's pid is
/// really running (see [`Identity`]).
///
/// The short name stays where it was, so a run is still grepped for `identifier mpv.exe`; the two
/// identifiers beside it are what says whether two sessions of that one executable are the same session to
/// the mixer, which is the difference a line could not show before.
fn the_identity_word(identity: &Identity) -> String {
    format!(
        "session-id {} instance-id {} image {}",
        the_quoted(identity.session.as_deref()),
        the_quoted(identity.instance.as_deref()),
        the_quoted(identity.image.as_deref()),
    )
}

/// One neighbour as a field of its own: pid, the executable its identifier names, its state, its mute and
/// its slider — and **`dead` where the process the mixer listed it under has ended**, which is the
/// neighbouring session this whole addition exists to catch.
///
/// A neighbour whose process is still there says nothing about it, and one the machine would not list its
/// processes for is `-`: the same word every unread value on this line gets, and the only one that can
/// never be mistaken for a session left behind.
fn the_neighbour_word(neighbour: &Neighbour) -> String {
    let mut fields = vec![
        neighbour.pid.to_string(),
        neighbour.identifier.as_deref().unwrap_or("-").to_string(),
        the_state_word(neighbour.state).to_string(),
        format!("mute {}", playback_trace::the_flag_word(neighbour.mute)),
        the_mixer_word(neighbour.volume),
    ];
    match neighbour.gone {
        Some(true) => fields.push("dead".to_string()),
        Some(false) => {}
        None => fields.push("-".to_string()),
    }

    format!("[{}]", fields.join(" "))
}

/// The sessions listed beside the player's: how many were written out of how many there were, and each of
/// the ones written as a field of its own.
///
/// **The elision is written, never silent**: `neighbours 8 of 14` is six that were not named, so the cap
/// can never read as a machine that had only eight sessions on it (see [`NEIGHBOURS_WRITTEN`]).
/// `neighbours 0 of 0` is a player whose session was alone on every endpoint — an answer, and not a gap.
fn the_neighbours_word(neighbours: &Neighbours) -> String {
    let written = neighbours
        .written
        .iter()
        .map(the_neighbour_word)
        .collect::<Vec<_>>()
        .join(" ");
    let counted = format!(
        "neighbours {} of {}",
        neighbours.written.len(),
        neighbours.listed
    );

    if written.is_empty() {
        counted
    } else {
        format!("{counted} {written}")
    }
}

/// **The two additions a line that names a session carries**: its whole identity, and the sessions the
/// endpoints listed beside it.
///
/// Both are on every such line — the probe's and a cure's — so that a victim, a cured session and a film
/// that never had the fault are compared field for field inside one run.
fn the_identity_and_neighbours(identity: &Identity, neighbours: &Neighbours) -> String {
    format!(
        "{} {}",
        the_identity_word(identity),
        the_neighbours_word(neighbours)
    )
}

/// The per-channel volumes of a stream, written one after another in the mixer's own channel order.
///
/// **`refused` is the word for an interface that would not be had**, and it is not `-`: the per-stream
/// level was not read at all, which is a fact about this probe, where a `-` in the row is a channel the
/// mixer would not answer for. `none` is a session that says it has no channel at all, and `of <count>`
/// is written where the stream has more channels than the line names (see `CHANNELS_WRITTEN`).
fn the_channels_word(channels: &Channels) -> String {
    let Channels::Read { volumes, count } = channels else {
        return "refused".to_string();
    };

    if volumes.is_empty() {
        return "none".to_string();
    }

    let mut word = volumes
        .iter()
        .map(|volume| the_mixer_word(*volume))
        .collect::<Vec<_>>()
        .join(" ");
    if volumes.len() < *count as usize {
        word.push_str(&format!(" of {count}"));
    }

    word
}

/// One endpoint written as `"<name>" <id>`, or `-` where Windows would not name it at all.
///
/// The name is quoted because it is the machine's own sentence and holds the spaces and parentheses it
/// likes (`Speakers (Realtek(R) Audio)`); the id beside it is what a run is compared by, since the name
/// is a locale's and the id is not.
///
/// **An endpoint neither half of which could be read is one `-`**, not two: it is the same word a
/// machine with no endpoint at all gets, because both are the same answer to a reader — nothing was
/// said about a device — and only an endpoint that answered one half is written as a pair.
fn the_endpoint_word(endpoint: Option<&Endpoint>) -> String {
    let Some(endpoint) = endpoint else {
        return "-".to_string();
    };

    if endpoint.name.is_none() && endpoint.id.is_none() {
        return "-".to_string();
    }

    let name = match endpoint.name.as_deref() {
        Some(name) => format!("{name:?}"),
        None => "-".to_string(),
    };
    let id = endpoint.id.as_deref().unwrap_or("-");

    format!("{name} {id}")
}

/// The session's half of a read-back: the mixer's slider and mute switch for one player, and what the
/// switch warranted where it warranted anything (see `Seen`).
///
/// **It is where the three silences are told apart on one line.** `verdict session-muted` with
/// `session mixer 1.000 mute yes` is Windows silencing a player that reports itself unmuted — the fault
/// a captured run proved, and the one the look takes off. `session mixer 0.000 mute yes left-alone` is
/// a hand on the mixer's own slider, written down and not restored. And `verdict silent` beside
/// `session mixer 1.000 mute no` is the player's own `volume 0` or its own `mute yes`, which is the
/// ladder's level rung; `no-device` beside a clear session is a device that never opened. The verdict
/// says what was found and this clause says what it was found in, so neither has to be guessed at.
///
/// The three words a look can fail in are words of their own: `session none` is the mixer answering and
/// the player not in it, and `session unreadable <why>` is the machine not answering at all — a silent
/// probe would read as the first of those, which is a finding about the player rather than about this
/// question (see `Refusal`).
pub(super) fn the_session_word(seen: &Seen) -> String {
    match seen {
        Seen::Absent => "session none".to_string(),
        Seen::Refused(why) => format!("session unreadable {}", why.word()),
        Seen::There {
            mixer, mute, cure, ..
        } => {
            let written = format!(
                "session mixer {} mute {}",
                the_mixer_word(*mixer),
                playback_trace::the_flag_word(*mute),
            );
            let cure = cure.word();
            if cure.is_empty() {
                written
            } else {
                format!("{written} {cure}")
            }
        }
    }
}

/// The line a look at a session writes of its own, where it wrote to the mixer or found a mute it
/// deliberately did not write to — **so that a cure is never silent and a failure is never swallowed**.
///
/// **It is written where the read-back is not.** A read that found a mute and took it off goes back as
/// `Heard` (see `Seen::the_verdict_of`), so the read-back says the film has its sound and nothing of
/// what it cost — and a cure written only on the read-back would leave no trace at all in the log it is
/// read from, which is exactly the line a run of this needs: the second mute switch, the one the player
/// cannot see, being taken off.
///
/// **A look that found nothing writes nothing here**: the read-back's clause carries it, and a line per
/// read about a machine whose mixer lists no session for a film is noise rather than evidence — the
/// more so for a player that has said nothing, which is asked again until it speaks (see
/// `audio_guard::ASKS_OF_A_SILENT_PLAYER`).
///
/// **A look that could not be made at all is a failure and is written as one**, because the cure it might
/// have been needed for was never attempted: a thread with no apartment, an enumerator that would not be
/// had and endpoints that would not be listed each get a line of their own, once per read, which is the
/// rate the read-back itself keeps (see `Refusal`).
pub(super) fn a_cure_line(pid: u32, seen: &Seen) -> Option<String> {
    let Seen::There {
        mixer,
        mute,
        cure,
        identity,
        neighbours,
    } = seen
    else {
        return match seen {
            Seen::Absent => None,
            Seen::Refused(why) => {
                Some(format!("session look refused pid {pid} why {}", why.word()))
            }
            Seen::There { .. } => None,
        };
    };

    let mixer = the_mixer_word(*mixer);
    // **A cure carries the identity and the neighbours of the session it was written to**, so the cured
    // session and a clean one are compared field for field in the same run (see
    // `the_identity_and_neighbours`).
    let beside = the_identity_and_neighbours(identity, neighbours);
    match cure {
        Cure::Nothing => None,
        Cure::Unmuted => Some(format!(
            "session unmuted pid {pid} mixer {mixer} was mute {beside}"
        )),
        Cure::Refused => Some(format!(
            "session unmute failed pid {pid} mixer {mixer} mute {} {beside}",
            playback_trace::the_flag_word(*mute),
        )),
        Cure::TurnedDown => Some(format!(
            "session turned down pid {pid} mixer {mixer} mute {} left alone {beside}",
            playback_trace::the_flag_word(*mute),
        )),
    }
}

/// The line one walk of the mixer is written as — **the pure half of this module**, and the whole of
/// what a run is read from.
///
/// **A session that is there is written as the state the mixer reports it in, and a session that is not
/// is a sentence of its own.** `session none` is a finding rather than an absence: it is what the user's
/// Volume Mixer showed while the film was silent, and `found 0 of 3` beside it says the mixer was
/// answering — three other sessions were listed — which is a different fault from `found 0 of 0`.
///
/// **`expired` is the datum this line exists for.** It is the OS saying the stream ended while the
/// player still believes it is playing, and no answer on the player's own pipe can say so.
///
/// **`mixer` is the mixer's slider and `volume` on the engine line is the player's own level**: they are
/// different numbers on different scales, and a film at `volume 100 mute no` whose session is at
/// `mixer 0.000` or `mute yes` is a silence this app never caused and no read of mpv can see.
///
/// **`channels` is the third hand on that number, and the one a driver-level handover moves.** mpv at its
/// default `--softvol=no` does not apply the level this app gave it: the driver does, through the
/// per-stream volume, and a stream whose slider is at full with a channel at nothing is a silent film
/// that every answer on the player's own pipe calls sound (see `Channels`).
///
/// **`children` is written on the miss line and only there**, because a miss is the one answer that
/// needs qualifying: the mixer can group a child process's session under its parent, so `children 1`
/// says something the player started did reach an endpoint and `children 0` says nothing it is part of
/// ever did (see `the_children_among`).
///
/// **The endpoint pair is the whole of the fifth question**, and it is written on every line that could
/// read one: `on` is where the session was listed and `default` the device the room is listening
/// through at that same instant. Two equal pairs say the stream is on the right device; two different
/// ones say it landed on another, or that the default was swapped under it.
///
/// **And two additions the run that proved the fault asked for** (see the module note): `session-id`,
/// `instance-id` and `image` are the whole of who the session is rather than the name it ends in, and
/// `neighbours` names the sessions the endpoints listed beside it — each with its pid, image, state,
/// mute and slider, and `dead` on one whose process has ended (see `the_identity_word` and
/// `the_neighbour_word`). They are written on the session line and on a cure's, and nowhere else: the
/// miss line is about a session that is not there, and the refusal line about a walk that was never
/// made. **Neither is read on a run nobody is tracing**, and a session that will not name itself writes
/// `-` where an identifier would be: `neighbours 0 of 0` is a player alone on every endpoint, which is an
/// answer rather than a gap.
///
/// A walk that could not be made says so: `session refused` with the word for what stopped it, because a
/// silent probe would read as `session none` and those are two different findings.
fn the_line(pid: u32, probe: &Probe) -> String {
    if let Some(refused) = probe.refused {
        return format!(
            "session refused pid {pid} why {} default {}",
            refused.word(),
            the_endpoint_word(probe.default.as_ref()),
        );
    }

    let found = format!("found {} of {}", probe.matching, probe.sessions);
    let default = the_endpoint_word(probe.default.as_ref());

    let Some(theirs) = probe.theirs.as_ref() else {
        return format!(
            "session none pid {pid} {found} children {} default {default}",
            probe.children,
        );
    };

    format!(
        "session {} pid {pid} {found} identifier {} mixer {} channels {} mute {} on {} default {default} {}",
        the_state_word(theirs.state),
        theirs.identity.identifier.as_deref().unwrap_or("-"),
        the_mixer_word(theirs.volume),
        the_channels_word(&theirs.channels),
        playback_trace::the_flag_word(theirs.mute),
        the_endpoint_word(Some(&theirs.endpoint)),
        the_identity_and_neighbours(&theirs.identity, &probe.neighbours),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    /// An endpoint as a machine names one, for the lines below.
    fn an_endpoint(name: &str, id: &str) -> Endpoint {
        Endpoint {
            name: Some(name.to_string()),
            id: Some(id.to_string()),
        }
    }

    /// The speakers a captured run played through, by the name the user saw and the id the mixer keys
    /// its sessions on.
    fn the_speakers() -> Endpoint {
        an_endpoint(
            "Speakers (Realtek(R) Audio)",
            "{0.0.0.00000000}.{1e4b6c1d-0f2a-4b8b-9c3d-5a6f7e8d9c0b}",
        )
    }

    /// The same machine's other device, which is the one a film can land on with nobody listening.
    fn the_headset() -> Endpoint {
        an_endpoint(
            "Headphones (2- USB Audio)",
            "{0.0.0.00000000}.{7c2d9f10-3a4b-4c5d-8e6f-102132435465}",
        )
    }

    /// The two channels of a stereo stream at full, which is what a player that opened its device
    /// cleanly reports through the per-stream volume.
    fn stereo() -> Channels {
        Channels::Read {
            volumes: vec![Some(1.0), Some(1.0)],
            count: 2,
        }
    }

    /// **Who Windows keys a player's session by**: the endpoint and the image, both identifiers entire —
    /// the two of them differing only in the guid an instance of the same application carries — and the
    /// image the pid is really running.
    fn an_identity(image: &str) -> Identity {
        let keyed = format!(r"\Device\HarddiskVolume4\Program Files\Players\{image}");

        Identity {
            identifier: Some(image.to_string()),
            session: Some(format!(
                "{{0.0.0.00000000}}.{{1e4b6c1d}}|#{keyed}%b{{00000000-0000-0000-0000-000000000000}}"
            )),
            instance: Some(format!(
                "{{0.0.0.00000000}}.{{1e4b6c1d}}|#{keyed}%b{{4c1b8a3e-5d6f-4a7b-8c9d-0e1f2a3b4c5d}}"
            )),
            image: Some(format!(r"C:\Program Files\Players\{image}")),
        }
    }

    /// One session as the endpoints listed it beside a player's: a pid, the executable its identifier
    /// names, its state, its mute and its slider — and what the machine said about the process it was
    /// listed under.
    fn a_neighbour(
        pid: u32,
        image: Option<&str>,
        state: State,
        mute: Option<bool>,
        volume: Option<f32>,
        gone: Option<bool>,
    ) -> Neighbour {
        Neighbour {
            pid,
            identifier: image.map(str::to_string),
            state,
            mute,
            volume,
            gone,
        }
    }

    /// The sessions a captured run had beside a player's: the user's browser, and one left behind by a
    /// process that has ended with its mute still on — the carrier the line is written to name.
    fn some_neighbours() -> Neighbours {
        Neighbours {
            listed: 4,
            written: vec![
                a_neighbour(
                    880,
                    Some("msedge.exe"),
                    State::Inactive,
                    Some(false),
                    Some(1.0),
                    Some(false),
                ),
                a_neighbour(
                    1204,
                    Some("mpv.exe"),
                    State::Expired,
                    Some(true),
                    Some(0.0),
                    Some(true),
                ),
            ],
        }
    }

    /// The identity and the neighbours word a player's session line carries, as the fixtures above make
    /// them — written once, so that the probe's line and a cure's line are held to the same field.
    fn the_mpv_beside() -> String {
        the_identity_and_neighbours(&an_identity("mpv.exe"), &some_neighbours())
    }

    /// What a player's session says on a machine whose mixer is answering: a state, who the session is,
    /// the mixer's own slider, the stream's own channels, its mute, the endpoint it was listed on, and the
    /// sessions that stood beside it.
    fn a_session(state: State, volume: Option<f32>, mute: Option<bool>) -> Probe {
        Probe {
            sessions: 3,
            matching: 1,
            listed: vec![9504, 1, 2],
            theirs: Some(Theirs {
                state,
                identity: an_identity("mpv.exe"),
                volume,
                channels: stereo(),
                mute,
                endpoint: the_speakers(),
            }),
            neighbours: some_neighbours(),
            default: Some(the_speakers()),
            ..Probe::default()
        }
    }

    /// **The line a healthy film writes, which is the record a silent one is compared against.** The
    /// state the mixer reports, the count it was found among, the image its identifier names, the
    /// mixer's own slider, the per-stream channel volumes, its mute, and the endpoint pair — the session
    /// on the same device as the default one, written twice so a swap is the difference between two
    /// words on one line. And the two additions every session line carries: the whole of who the session
    /// is, and the sessions that stood beside it (see `the_identity_and_neighbours`).
    #[test]
    fn a_session_that_is_there_is_written_with_its_state_level_and_both_endpoints() {
        let line = the_line(9504, &a_session(State::Active, Some(0.75), Some(false)));

        assert_eq!(
            line,
            "session active pid 9504 found 1 of 3 identifier mpv.exe mixer 0.750 \
             channels 1.000 1.000 mute no \
             on \"Speakers (Realtek(R) Audio)\" \
             {0.0.0.00000000}.{1e4b6c1d-0f2a-4b8b-9c3d-5a6f7e8d9c0b} \
             default \"Speakers (Realtek(R) Audio)\" \
             {0.0.0.00000000}.{1e4b6c1d-0f2a-4b8b-9c3d-5a6f7e8d9c0b} \
             session-id \"{0.0.0.00000000}.{1e4b6c1d}|#\\Device\\HarddiskVolume4\\Program \
             Files\\Players\\mpv.exe%b{00000000-0000-0000-0000-000000000000}\" \
             instance-id \"{0.0.0.00000000}.{1e4b6c1d}|#\\Device\\HarddiskVolume4\\Program \
             Files\\Players\\mpv.exe%b{4c1b8a3e-5d6f-4a7b-8c9d-0e1f2a3b4c5d}\" \
             image \"C:\\Program Files\\Players\\mpv.exe\" neighbours 2 of 4 \
             [880 msedge.exe inactive mute no 1.000] \
             [1204 mpv.exe expired mute yes 0.000 dead]"
        );
    }

    /// **Who Windows believes a session is, written whole and never shortened to the name at the end of
    /// it.** Both identifiers are what the mixer's per-app volume and mute are remembered against, so two
    /// sessions of one `mpv.exe` that are remembered differently differ somewhere inside them — and the
    /// image the pid is really running is written beside them, so that an identifier can be read against
    /// a file rather than against Windows' own device spelling of one.
    #[test]
    fn a_session_is_written_with_the_whole_identity_windows_keys_it_by() {
        let word = the_identity_word(&an_identity("mpv.exe"));

        assert!(
            word.starts_with(
                r#"session-id "{0.0.0.00000000}.{1e4b6c1d}|#\Device\HarddiskVolume4\Program Files\Players\mpv.exe%b{00000000-0000-0000-0000-000000000000}" "#
            ),
            "the whole identifier, and not the executable at the end of it: {word}"
        );
        assert!(
            word.contains(
                r#"instance-id "{0.0.0.00000000}.{1e4b6c1d}|#\Device\HarddiskVolume4\Program Files\Players\mpv.exe%b{4c1b8a3e-5d6f-4a7b-8c9d-0e1f2a3b4c5d}""#
            ),
            "the instance's own identifier beside it: {word}"
        );
        assert!(
            word.ends_with(r#"image "C:\Program Files\Players\mpv.exe""#),
            "and the image the pid is running, by a path a reader can look at: {word}"
        );

        // A session that will name neither identifier is three `-`s rather than a gap: a field this
        // build could not read is written as missing, in the same way as every other unread value here.
        assert_eq!(
            the_identity_word(&Identity::default()),
            "session-id - instance-id - image -"
        );
    }

    /// **The sessions beside a player's, each with the five answers that would name a carrier**: its pid,
    /// the executable its identifier names, its state, its mute and its slider — and **`dead` where the
    /// process the mixer listed it under has ended**, which is the neighbouring session this whole line
    /// exists to catch: a session left behind with its mute still on is what a new one inherits from.
    #[test]
    fn a_neighbour_is_written_with_its_pid_image_state_mute_and_slider() {
        assert_eq!(
            the_neighbours_word(&some_neighbours()),
            "neighbours 2 of 4 [880 msedge.exe inactive mute no 1.000] \
             [1204 mpv.exe expired mute yes 0.000 dead]"
        );

        // A neighbour that answered nothing is still a field of its own: a session that will not name an
        // image, will not say its mute or its slider, and whose process the machine would not list.
        assert_eq!(
            the_neighbour_word(&a_neighbour(7, None, State::Unknown, None, None, None)),
            "[7 - unknown mute - - -]",
            "every gap is a `-`, which can never be read as a session left behind"
        );
    }

    /// **The cap is counted against what there was, and never silently.** A machine with fourteen
    /// sessions on it and eight written says six were left out, so the list can never be read as the
    /// whole machine; and a player whose session was alone on every endpoint is a count of nothing
    /// rather than a gap.
    #[test]
    fn a_neighbour_list_that_was_cut_says_how_many_were_left_out() {
        let many = Neighbours {
            listed: 14,
            written: (0..NEIGHBOURS_WRITTEN as u32)
                .map(|at| {
                    a_neighbour(
                        900 + at,
                        Some("mpv.exe"),
                        State::Active,
                        Some(false),
                        Some(1.0),
                        Some(false),
                    )
                })
                .collect(),
        };
        let word = the_neighbours_word(&many);

        assert!(word.starts_with("neighbours 8 of 14 "), "{word}");
        assert_eq!(
            word.matches('[').count(),
            NEIGHBOURS_WRITTEN,
            "and exactly the cap is named, whatever the count beside it: {word}"
        );
        assert_eq!(
            the_neighbours_word(&Neighbours::default()),
            "neighbours 0 of 0"
        );
    }

    /// **The datum the probe was built for.** A player that believes it opened its device, with an
    /// output and params of its own, and a session the OS has already expired: the state word is the
    /// whole of the evidence, it is at the front of the line, and no read of mpv could produce it.
    #[test]
    fn an_expired_session_says_the_stream_ended_under_a_player_that_still_thinks_it_is_playing() {
        let line = the_line(9504, &a_session(State::Expired, Some(1.0), Some(false)));

        assert!(
            line.starts_with("session expired pid 9504 "),
            "the state is the first thing a line says: {line}"
        );
        assert!(
            line.contains("mixer 1.000 channels 1.000 1.000 mute no"),
            "a session can be expired with the mixer and the stream both still up: {line}"
        );
    }

    /// **The mixer is a second hand on the same volume**: a player at `volume 100 mute no` whose session
    /// is at nothing, or whose session is muted, is a silence nothing on the player's pipe describes —
    /// and the two are written with words of their own so a reader cannot take one for the other.
    #[test]
    fn the_mixers_own_slider_and_mute_are_written_apart_from_the_players_level() {
        assert_eq!(the_mixer_word(Some(0.0)), "0.000");
        assert_eq!(
            the_mixer_word(Some(0.5)),
            "0.500",
            "a share of full, not the percent mpv reports"
        );
        assert_eq!(the_mixer_word(Some(1.0)), "1.000");
        assert_eq!(
            the_mixer_word(None),
            "-",
            "a mixer that would not say is no answer, not a level of nothing"
        );

        let muted = the_line(9504, &a_session(State::Active, Some(0.0), Some(true)));
        assert!(
            muted.contains("mixer 0.000") && muted.contains("mute yes"),
            "the slider and the switch are the mixer's own pair: {muted}"
        );
    }

    /// **The per-stream level is a third number again, and the one a driver-level handover moves.**
    /// mpv at its default `--softvol=no` does not apply the level inside the player, so a stream whose
    /// mixer slider is at full and whose channel is at nothing is the silent film this probe exists to
    /// catch — and an interface the session would not hand over is `refused` rather than a gap.
    #[test]
    fn the_streams_own_channels_are_written_beside_the_mixers_slider() {
        let handed_down = Probe {
            theirs: Some(Theirs {
                state: State::Active,
                identity: an_identity("mpv.exe"),
                volume: Some(1.0),
                channels: Channels::Read {
                    volumes: vec![Some(0.0), None],
                    count: 2,
                },
                mute: Some(false),
                endpoint: the_speakers(),
            }),
            default: Some(the_speakers()),
            ..a_session(State::Active, Some(1.0), Some(false))
        };

        let line = the_line(9504, &handed_down);
        assert!(
            line.contains("mixer 1.000 channels 0.000 -"),
            "a slider at full over a channel at nothing, and a channel that would not answer: {line}"
        );

        // A stream with more channels than the line names says how many it has rather than pretending
        // to have named them all.
        let wide = Channels::Read {
            volumes: vec![Some(1.0), Some(1.0)],
            count: 12,
        };
        assert_eq!(the_channels_word(&wide), "1.000 1.000 of 12");

        assert_eq!(
            the_channels_word(&Channels::Refused),
            "refused",
            "an interface that would not be had is said so rather than left out"
        );
        assert_eq!(
            the_channels_word(&Channels::Read {
                volumes: Vec::new(),
                count: 0,
            }),
            "none",
            "a session with no channel at all is a word of its own"
        );
    }

    /// **A stream on a device nobody is listening to is visible as two different pairs.** The default
    /// endpoint was read at the same instant as the walk, so a session listed on the headset while the
    /// room listens through the speakers is the finding itself rather than a silence nobody can place.
    #[test]
    fn a_session_on_another_endpoint_is_written_beside_the_default_one() {
        let probe = Probe {
            theirs: Some(Theirs {
                state: State::Active,
                identity: an_identity("ffplay.exe"),
                volume: Some(1.0),
                channels: stereo(),
                mute: Some(false),
                endpoint: the_headset(),
            }),
            default: Some(the_speakers()),
            ..a_session(State::Active, Some(1.0), Some(false))
        };

        let line = the_line(9504, &probe);

        assert!(
            line.contains(
                "on \"Headphones (2- USB Audio)\" \
                 {0.0.0.00000000}.{7c2d9f10-3a4b-4c5d-8e6f-102132435465}"
            ),
            "the endpoint it landed on: {line}"
        );
        assert!(
            line.contains(
                "default \"Speakers (Realtek(R) Audio)\" \
                 {0.0.0.00000000}.{1e4b6c1d-0f2a-4b8b-9c3d-5a6f7e8d9c0b}"
            ),
            "beside the one the room is listening through: {line}"
        );
        assert_ne!(
            the_endpoint_word(probe.theirs.as_ref().map(|theirs| &theirs.endpoint)),
            the_endpoint_word(probe.default.as_ref()),
            "the two pairs being different is the whole of what the line is read for"
        );
    }

    /// **A miss is a finding, not an absence.** No session for the player while the mixer lists three
    /// others is exactly what the user reported, and it is a sentence of its own — `session none` — with
    /// the count that says the mixer was answering.
    #[test]
    fn a_player_with_no_session_is_written_as_none_rather_than_left_out() {
        let nobody = Probe {
            sessions: 3,
            matching: 0,
            theirs: None,
            default: Some(the_speakers()),
            ..Probe::default()
        };

        assert_eq!(
            the_line(9504, &nobody),
            "session none pid 9504 found 0 of 3 children 0 default \"Speakers (Realtek(R) Audio)\" \
             {0.0.0.00000000}.{1e4b6c1d-0f2a-4b8b-9c3d-5a6f7e8d9c0b}"
        );

        // And a machine with nothing to enumerate at all is the same sentence with nothing found,
        // which is a different fault from a mixer that listed sessions and not this one.
        let no_endpoint = Probe {
            sessions: 0,
            matching: 0,
            theirs: None,
            default: None,
            ..Probe::default()
        };
        assert_eq!(
            the_line(9504, &no_endpoint),
            "session none pid 9504 found 0 of 0 children 0 default -"
        );
    }

    /// **A miss is qualified before it is written.** The mixer can group a child process's session under
    /// its parent, so a session something the player started made is a different finding from a machine
    /// on which nothing this player is part of ever reached an endpoint — and `children` is what says
    /// which of the two a `none` is.
    #[test]
    fn a_miss_says_whether_a_process_the_player_started_made_a_session() {
        let grouped = Probe {
            sessions: 3,
            matching: 0,
            listed: vec![41, 42, 43],
            children: 1,
            theirs: None,
            default: Some(the_speakers()),
            ..Probe::default()
        };

        assert!(
            the_line(9504, &grouped).contains("found 0 of 3 children 1"),
            "a session a child of the player made is started somewhere, not nowhere: {}",
            the_line(9504, &grouped)
        );
    }

    /// **A probe that could not be made says so.** A silent failure would read as `session none`, which
    /// is a finding about the player rather than about this question — so each way the walk can fail is
    /// its own word, and the default endpoint is written where one could be read before it failed.
    #[test]
    fn a_walk_that_could_not_be_made_says_which_part_of_it_refused() {
        for (refused, word) in [
            (Refusal::Apartment, "apartment"),
            (Refusal::Enumerator, "enumerator"),
            (Refusal::Endpoints, "endpoints"),
        ] {
            let probe = Probe {
                refused: Some(refused),
                ..Probe::default()
            };
            assert_eq!(
                the_line(9504, &probe),
                format!("session refused pid 9504 why {word} default -"),
                "every refusal has a word of its own"
            );
        }

        // A walk that got as far as the default endpoint and then lost the endpoints still carries the
        // pair, so the line says which device the machine had at that instant.
        let half_way = Probe {
            refused: Some(Refusal::Endpoints),
            default: Some(the_speakers()),
            ..Probe::default()
        };
        assert!(the_line(9504, &half_way).contains("default \"Speakers (Realtek(R) Audio)\""));
    }

    /// **The identifier is read for the image it names and for nothing else.** It is the one answer that
    /// says whether a session is the player's own or was inherited, and the device path Windows writes
    /// beside the image is the machine's own business rather than a log's.
    #[test]
    fn a_session_identifier_is_read_for_the_executable_it_names() {
        assert_eq!(
            the_executable_named(
                r"{0.0.0.00000000}.{1e4b6c1d}|#\Device\HarddiskVolume4\Program Files\mpv\mpv.exe%b{00000000-0000-0000-0000-000000000000}"
            )
            .as_deref(),
            Some("mpv.exe"),
            "the image at the end of the identifier, and not the path in front of it"
        );
        assert_eq!(
            the_executable_named(
                r"{0.0.0.00000000}.{1e4b6c1d}|#\Device\HarddiskVolume3\Tools\FFPLAY.EXE%b{}"
            )
            .as_deref(),
            Some("FFPLAY.EXE"),
            "an image is a program whatever case it is written in"
        );
        assert_eq!(
            the_executable_named("mpv.exe").as_deref(),
            Some("mpv.exe"),
            "an identifier that is nothing but an image names it"
        );
        assert_eq!(
            the_executable_named("{0.0.0.00000000}.{1e4b6c1d}|#%b{}"),
            None,
            "a grouped or inherited session names no image at all"
        );
        assert_eq!(the_executable_named("-"), None);
    }

    /// The four states a session can be read in, each with the word the line carries — a state this
    /// build has no value for being `unknown` rather than nothing.
    #[test]
    fn every_state_the_mixer_can_report_has_a_word() {
        assert_eq!(the_state_word(State::of(AudioSessionStateActive)), "active");
        assert_eq!(
            the_state_word(State::of(AudioSessionStateInactive)),
            "inactive"
        );
        assert_eq!(
            the_state_word(State::of(AudioSessionStateExpired)),
            "expired"
        );
        assert_eq!(
            the_state_word(State::of(AudioSessionState(7))),
            "unknown",
            "a value this build has no word for is not a state it invents one for"
        );
    }

    /// A session that answered nothing at all is still a line: the state it could not be asked, the
    /// mixer's `-` pair, the stream's level refused rather than guessed at, and the endpoint it was
    /// listed on.
    #[test]
    fn a_session_that_answered_nothing_is_written_as_nothing_rather_than_dropped() {
        let silent = Probe {
            theirs: Some(Theirs {
                state: State::Unknown,
                identity: Identity::default(),
                volume: None,
                channels: Channels::Refused,
                mute: None,
                endpoint: Endpoint::default(),
            }),
            default: None,
            ..a_session(State::Unknown, None, None)
        };

        assert_eq!(
            the_line(9504, &silent),
            "session unknown pid 9504 found 1 of 3 identifier - mixer - channels refused mute - \
             on - default - session-id - instance-id - image - neighbours 2 of 4 \
             [880 msedge.exe inactive mute no 1.000] \
             [1204 mpv.exe expired mute yes 0.000 dead]"
        );
    }

    /// **The rule the cure is written by, and the three answers it turns on.** A muted session under a
    /// launch that asked for a level is the fault a captured run proved, and it is taken off. A switch
    /// that is already off is nothing to do — and so is one the mixer would not report, because a
    /// switch this build cannot read is not a switch to throw.
    #[test]
    fn a_muted_session_under_a_launch_that_asked_for_a_level_is_taken_off() {
        assert_eq!(
            the_warrant(Some(1.0), Some(true), 100),
            Warrant::TakeItOff,
            "muted at a slider nobody touched is the second mute switch, and this app's own doing"
        );
        assert_eq!(
            the_warrant(Some(0.05), Some(true), 40),
            Warrant::TakeItOff,
            "any slider that is up at all is the same finding"
        );
        assert_eq!(
            the_warrant(None, Some(true), 100),
            Warrant::TakeItOff,
            "a slider that would not answer is not a slider at nothing: the mute is still the fault"
        );

        assert_eq!(
            the_warrant(Some(1.0), Some(false), 100),
            Warrant::Nothing,
            "a session that is not muted has nothing to take off"
        );
        assert_eq!(
            the_warrant(None, None, 100),
            Warrant::Nothing,
            "a mute the mixer would not report is never written to"
        );
    }

    /// **A launch that asked for no audio is never cured**, and the level is half the rule for the
    /// reason it is half the verdict: the player was begun with no track at all (see
    /// `video_launch::AudioForm`), so a mute on a session it does not have is nobody's but the mixer's
    /// — and a film the setting silenced is not a fault to heal.
    #[test]
    fn a_launch_that_asked_for_no_audio_never_unmutes_a_session() {
        assert_eq!(the_warrant(Some(1.0), Some(true), 0), Warrant::Nothing);
        assert_eq!(
            the_warrant(Some(1.0), Some(true), 100),
            Warrant::TakeItOff,
            "while the same session under a launch that asked for a level is the fault itself"
        );
    }

    /// **A hand on the mixer's own slider is written down and not restored.** The same muted session at
    /// `mixer 0.000` is somebody's deliberate setting, and putting the slider back would overwrite a
    /// decision this app has no business with — so the look records it and the write it makes is none.
    #[test]
    fn a_session_muted_with_its_slider_at_nothing_is_left_as_a_hand_left_it() {
        assert_eq!(
            the_warrant(Some(0.0), Some(true), 100),
            Warrant::LeaveIt,
            "the slider is down, so the mute is part of the same hand"
        );
        assert_eq!(
            the_warrant(Some(0.0004), Some(true), 100),
            Warrant::LeaveIt,
            "a slider the line would write as 0.000 is a slider at nothing"
        );

        let turned_down = Seen::there(Some(0.0), Some(true), Cure::TurnedDown);
        assert_eq!(
            the_session_word(&turned_down),
            "session mixer 0.000 mute yes left-alone"
        );
        let turned_down =
            turned_down.with_identity_and_neighbours(an_identity("mpv.exe"), some_neighbours());
        assert_eq!(
            a_cure_line(9504, &turned_down),
            Some(format!(
                "session turned down pid 9504 mixer 0.000 mute yes left alone {}",
                the_mpv_beside()
            )),
            "a film this app left alone still says so, so that nobody reads the silence as this \
             module's failure"
        );
    }

    /// **The cure is written down, and so is the cure that failed.** A mute taken off is the line a run
    /// of this is read from — the second switch, the one the player cannot see — and a `SetMute` the
    /// mixer refused is written as a failure rather than swallowed, because the film is still silent
    /// and the next read will try again.
    #[test]
    fn a_cure_and_a_cure_that_failed_are_each_written_down() {
        let muted = Seen::there(Some(1.0), Some(true), Cure::Unmuted)
            .with_identity_and_neighbours(an_identity("mpv.exe"), some_neighbours());
        assert_eq!(
            a_cure_line(23724, &muted),
            Some(format!(
                "session unmuted pid 23724 mixer 1.000 was mute {}",
                the_mpv_beside()
            )),
            "a cure carries the same identity and neighbours as the probe's line, so a cured session \
             and a clean one are compared field for field in one run"
        );
        assert_eq!(
            the_session_word(&muted),
            "session mixer 1.000 mute yes cured",
            "the read-back says both what was found and what was done about it"
        );

        let refused = Seen::there(Some(1.0), Some(true), Cure::Refused)
            .with_identity_and_neighbours(an_identity("mpv.exe"), some_neighbours());
        assert_eq!(
            a_cure_line(23724, &refused),
            Some(format!(
                "session unmute failed pid 23724 mixer 1.000 mute yes {}",
                the_mpv_beside()
            ))
        );
        assert_eq!(
            the_session_word(&refused),
            "session mixer 1.000 mute yes cure-failed"
        );

        // And a session that needed nothing says so on the read-back and writes no line of its own,
        // which is what keeps a healthy film's log to the one line it opened with.
        let clear = Seen::there(Some(1.0), Some(false), Cure::Nothing);
        assert_eq!(a_cure_line(23724, &clear), None);
        assert_eq!(the_session_word(&clear), "session mixer 1.000 mute no");
    }

    /// **An unreadable session is nothing done to it, and the line says which part of the walk
    /// refused**: a look that found no session, and one whose walk could not be made, are two
    /// findings that must never read as each other.
    #[test]
    fn a_session_that_could_not_be_looked_at_is_never_written_to() {
        assert_eq!(the_session_word(&Seen::Absent), "session none");
        assert_eq!(a_cure_line(23724, &Seen::Absent), None);

        for (refused, word) in [
            (Refusal::Apartment, "apartment"),
            (Refusal::Enumerator, "enumerator"),
            (Refusal::Endpoints, "endpoints"),
        ] {
            let seen = Seen::Refused(refused);
            assert_eq!(
                the_session_word(&seen),
                format!("session unreadable {word}")
            );
            assert_eq!(
                a_cure_line(23724, &seen),
                Some(format!("session look refused pid 23724 why {word}")),
                "a look that could not be made threw no switch, and says so rather than passing over"
            );
        }
    }

    /// **The third silence, and what it folds into a verdict.** A session still muted under a player
    /// that says it is heard is `session-muted` — the word a log is read for — and a session whose mute
    /// the cure took off is `heard`, because that is what the film is now. A look that said nothing
    /// about a mute leaves the player's own verdict exactly as it was, and a fault of the player's own
    /// is never softened by a session.
    #[test]
    fn a_muted_session_reads_as_session_muted_and_never_as_heard() {
        let muted = |cure| Seen::there(Some(1.0), Some(true), cure);

        assert_eq!(
            muted(Cure::Unmuted).the_verdict_of(Verdict::Heard),
            Verdict::Heard,
            "a mute taken off is a film with its sound: the write was the whole of the fault"
        );
        for cure in [Cure::Refused, Cure::TurnedDown] {
            assert_eq!(
                muted(cure).the_verdict_of(Verdict::Heard),
                Verdict::SessionMuted,
                "{cure:?} is a film that is still silent, and no read of mpv could say so"
            );
        }

        let clear = Seen::there(Some(1.0), Some(false), Cure::Nothing);
        assert_eq!(clear.the_verdict_of(Verdict::Heard), Verdict::Heard);
        assert_eq!(
            Seen::Absent.the_verdict_of(Verdict::Heard),
            Verdict::Heard,
            "a look that found no session cannot unmake a heard film"
        );
        assert_eq!(
            Seen::Refused(Refusal::Apartment).the_verdict_of(Verdict::Heard),
            Verdict::Heard
        );
        for engine in [
            Verdict::NoTrack,
            Verdict::NoDevice,
            Verdict::Silent,
            Verdict::Unanswered,
        ] {
            assert_eq!(
                muted(Cure::Refused).the_verdict_of(engine),
                engine,
                "{engine:?} is the ladder's road and a session cannot soften it"
            );
        }
    }
}

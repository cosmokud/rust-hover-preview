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
//! **This is a reader and never a cure.** Nothing here sets a session's volume, unmutes anything, or
//! touches `IAudioEndpointVolume`: the mixer's own answer is the evidence, and changing what the
//! user's mixer says would destroy the thing being collected. The one module that does write to a
//! player is `audio_guard`, and it writes to the player rather than to the OS.
//!
//! **COM here is this thread's own apartment, and nowhere else.** The preview loop already holds an
//! apartment for the engines that draw into it (see `pdf_preview::initialize_apartment` and the head
//! of `run_preview_window`), and the trace's other thread reads a player's pipe and holds none — so
//! this probe begins a thread, takes `COINIT_MULTITHREADED` at the top of it, gives it back at the
//! end, and touches nothing of the loop's.
//!
//! **Nothing runs unless `RHP_PLAYBACK_TRACE` asks for a line** (see `playback_trace::asked`): no
//! thread, no apartment, no COM call and no line, so a run without the variable pays a read of an
//! atomic. And what it says is written on that thread: a walk of the machine's endpoints is a
//! microsecond-scale call but the wait for a player's session to exist is not, and no tick may pay it
//! (see `playback_trace`, whose own thread is the same shape).

use std::time::Duration;

use windows::core::{Interface, GUID};
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

use super::playback_trace;

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
    /// The executable the session identifier names, where it names one (see
    /// [`the_executable_named`]).
    identifier: Option<String>,
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
enum Refusal {
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

/// Walk every render endpoint this machine has and answer what the mixer says about one player.
///
/// **Every endpoint and not only the default one**, because a stream that landed on a device the user
/// is not listening to is one of the answers this line exists to give, and a walk of the default
/// endpoint alone could never see one. The default endpoint is read first and at that instant, so that
/// the pair the line carries compares the session's endpoint with the device the room is listening
/// through now rather than with a default read a moment later.
fn walk_the_mixer(pid: u32) -> Probe {
    let mut probe = Probe::default();

    // Safety: every interface below is one this function holds and drops before it returns, on a
    // thread whose apartment its caller took. Every call is a read: no session is set, muted, or
    // otherwise written to.
    unsafe {
        let enumerator: IMMDeviceEnumerator =
            match CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) {
                Ok(enumerator) => enumerator,
                Err(_) => {
                    probe.refused = Some(Refusal::Enumerator);
                    return probe;
                }
            };

        probe.default = enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole)
            .map(|device| the_endpoint(&device))
            .ok();

        let endpoints = match enumerator.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE) {
            Ok(endpoints) => endpoints,
            Err(_) => {
                probe.refused = Some(Refusal::Endpoints);
                return probe;
            }
        };

        let count = endpoints.GetCount().unwrap_or(0);
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
                probe.sessions += 1;

                // **`GetProcessId` is the whole of the match**, and it is the second generation of
                // the interface that answers it: a session belongs to the process that made it
                // whatever its identifier says, and a session a process inherited from its parent
                // carries the parent's pid as readily as the parent's image (see
                // `the_executable_named`, which is the identifier's half of the same question).
                //
                // **It is read for every session and not only for the match**, because a miss counts
                // the sessions the player's own children made (see `the_children_among`).
                let Ok(theirs) = control.cast::<IAudioSessionControl2>() else {
                    continue;
                };
                let Ok(session_pid) = theirs.GetProcessId() else {
                    continue;
                };
                probe.listed.push(session_pid);

                if session_pid != pid {
                    continue;
                }

                probe.matching += 1;
                if probe.theirs.is_none() {
                    probe.theirs = Some(the_session(&control, &device));
                }
            }
        }

        // **A miss is counted against the player's own tree before it is written**, and only a miss: a
        // session that is the player's own already answers the question the count is for.
        if probe.matching == 0 {
            probe.children = the_children_among(&probe.listed, pid);
        }
    }

    probe
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
unsafe fn the_session(control: &IAudioSessionControl, device: &IMMDevice) -> Theirs {
    let volume = control.cast::<ISimpleAudioVolume>().ok();

    Theirs {
        state: control.GetState().map_or(State::Unknown, State::of),
        identifier: the_identifier_of(control),
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
    let raw = theirs.GetSessionIdentifier().ok()?;
    let identifier = raw.to_string().ok();
    CoTaskMemFree(Some(raw.0 as *const _));
    identifier.as_deref().and_then(the_executable_named)
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
        "session {} pid {pid} {found} identifier {} mixer {} channels {} mute {} on {} default \
         {default}",
        the_state_word(theirs.state),
        theirs.identifier.as_deref().unwrap_or("-"),
        the_mixer_word(theirs.volume),
        the_channels_word(&theirs.channels),
        playback_trace::the_flag_word(theirs.mute),
        the_endpoint_word(Some(&theirs.endpoint)),
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

    /// What a player's session says on a machine whose mixer is answering: a state, an identifier, the
    /// mixer's own slider, the stream's own channels, its mute, and the endpoint it was listed on.
    fn a_session(state: State, volume: Option<f32>, mute: Option<bool>) -> Probe {
        Probe {
            sessions: 3,
            matching: 1,
            listed: vec![9504, 1, 2],
            theirs: Some(Theirs {
                state,
                identifier: Some("mpv.exe".to_string()),
                volume,
                channels: stereo(),
                mute,
                endpoint: the_speakers(),
            }),
            default: Some(the_speakers()),
            ..Probe::default()
        }
    }

    /// **The line a healthy film writes, which is the record a silent one is compared against.** The
    /// state the mixer reports, the count it was found among, the image its identifier names, the
    /// mixer's own slider, the per-stream channel volumes, its mute, and the endpoint pair — the session
    /// on the same device as the default one, written twice so a swap is the difference between two
    /// words on one line.
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
             {0.0.0.00000000}.{1e4b6c1d-0f2a-4b8b-9c3d-5a6f7e8d9c0b}"
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
                identifier: Some("mpv.exe".to_string()),
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
                identifier: Some("ffplay.exe".to_string()),
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
                identifier: None,
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
             on - default -"
        );
    }
}

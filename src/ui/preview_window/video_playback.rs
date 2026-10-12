//! Starting and stopping a film: the launch itself, the wait for its window and for its first
//! frame, and putting the player's window back in front.

use super::*;

/// The films whose most recently begun player was drawing subtitles: a sidecar, the film's own
/// track, or the small files this app copied out of it (see `subtitle_filter`).
///
/// What asks is the pin's take-up, and it asks because a pin does not begin a player of its
/// own over a hover's: the `ffplay` a hover began is adopted as it stands, subtitles and all —
/// so a player that was begun *without* them is one a pinned window would keep drawing without
/// them, and the take-up is the moment to begin it again (see `reload_adopted_subtitles`). The
/// list holds only the films whose player is drawing; an absence is the answer that the player
/// last begun for that film is not — which is also the answer for a film no player was begun
/// for at all, and the one the take-up wants for it, because a film with no player up has
/// nothing on screen that is wrong.
static DRAWING_SUBTITLES: Lazy<Mutex<Vec<PathBuf>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// How many films the list holds before it is emptied, the same bound and the same reasoning
/// as the measure lists' own: a film forgotten by the clear is one whose next take-up begins
/// its player again although it was already drawing the subtitles, which is a blink rather
/// than a fault (see `note_player_subtitles`).
const DRAWING_SUBTITLES_MAX_ENTRIES: usize = 64;

/// Write down what the player just begun for this film was handed: subtitles drawn, or none.
///
/// It is asked of at the take-up rather than of the player, which cannot be asked anything
/// once it is running — and it is written on every launch, so what stands for a film is always
/// that film's latest player (see `reload_adopted_subtitles`).
pub(super) fn note_player_subtitles(path: &Path, drawing: bool) {
    let Ok(mut films) = DRAWING_SUBTITLES.lock() else {
        return;
    };

    films.retain(|held| held != path);

    if !drawing {
        return;
    }

    if films.len() >= DRAWING_SUBTITLES_MAX_ENTRIES {
        films.clear();
    }

    films.push(path.to_path_buf());
}

/// Whether the player most recently begun for this film was drawing subtitles.
pub(super) fn player_draws_subtitles(path: &Path) -> bool {
    DRAWING_SUBTITLES
        .lock()
        .map(|films| films.iter().any(|held| held == path))
        .unwrap_or(false)
}

/// The player a preview holds: the process this app began, or — under test — the id one would
/// have been given.
///
/// The two are one type so that every caller treats them alike, which is what a seam at the
/// launch has to leave alone: the id is what the bookkeeping names the player by (see
/// `note_player_started`), and the handle is what ends it. A test run must begin nothing, so what
/// the seam hands back in place of a process is the noting of one (see `spawn_player`).
///
/// **A noted id can never end a real process, and that is what makes one safe to keep.** Nothing
/// is ever killed by an id alone: `terminate_player_pid` verifies the id still runs the image name
/// it was recorded under before it touches anything (`process_image_matches`), and a Windows
/// process id is always a multiple of four — so an odd stand-in is refused twice over.
pub(super) enum PlayerProcess {
    /// The player this app began.
    #[cfg_attr(test, allow(dead_code))] // The test seam notes a player instead of beginning one.
    Started(Child),
    /// The player a test noted in place of beginning one.
    #[cfg(test)]
    Noted(u32),
}

impl PlayerProcess {
    /// The id the player is booked in and ended under.
    pub(super) fn id(&self) -> u32 {
        match self {
            Self::Started(child) => child.id(),
            #[cfg(test)]
            Self::Noted(pid) => *pid,
        }
    }

    /// Whether the player has ended, asked of a real handle and of a note alike: a noted player
    /// is not a process and never ends, which is what a caller waiting on one expects.
    pub(super) fn try_wait(&mut self) -> std::io::Result<Option<std::process::ExitStatus>> {
        match self {
            Self::Started(child) => child.try_wait(),
            #[cfg(test)]
            Self::Noted(_) => Ok(None),
        }
    }

    /// End the player, where there is one to end.
    pub(super) fn kill(&mut self) -> std::io::Result<()> {
        match self {
            Self::Started(child) => child.kill(),
            #[cfg(test)]
            Self::Noted(_) => Ok(()),
        }
    }
}

/// Begin the player a launch built its command for, or — under test — note that one would have
/// been begun and begin nothing.
///
/// The one place the two builds differ, and it is a compile-time branch rather than a runtime
/// check: the release build spawns, and the test build hands back an id no process holds, so the
/// release binary carries no part of the test arm. What the noting leaves out is the process and
/// nothing else — the launch books the player in through `note_player_started` all the same.
#[cfg(not(test))]
pub(super) fn spawn_player(command: &mut std::process::Command) -> Option<PlayerProcess> {
    command.spawn().ok().map(PlayerProcess::Started)
}

/// The test build's launch: note a player instead of beginning one (see `spawn_player`).
#[cfg(test)]
pub(super) fn spawn_player(_command: &mut std::process::Command) -> Option<PlayerProcess> {
    Some(PlayerProcess::Noted(next_noted_pid()))
}

/// The id a noted player carries, one apart from the last.
///
/// An odd number, which no Windows process id ever is — they are multiples of four — so a noted
/// id names no real process even before the image-name check refuses it (see `PlayerProcess`).
#[cfg(test)]
fn next_noted_pid() -> u32 {
    static NEXT: AtomicU32 = AtomicU32::new(u32::MAX);
    NEXT.fetch_sub(2, Ordering::Relaxed)
}

/// Start ffplay for video preview, at the level `volume` names and with the film's own measured
/// loudness folded into it where `Normalize` is on for videos (see `normalizing_video`).
///
/// The level is the caller's answer rather than a read of the configuration, because the two
/// callers keep different ones: a preview that is beginning is played at `Volume → Video`, and a
/// pinned one is played at the level its own window is holding — the one the tray named when that
/// pin was taken up, moved by whatever the hand on its volume control has done since, and never
/// written back to the setting (see `PinVolume`). `subtitle` is the same kind of answer for the
/// same reason: it is what a pinned window is remembering, and `None` is what a hover asks for
/// (see `next_subtitle`).
#[allow(clippy::too_many_arguments)] // Each one is a fact about the launch, and a struct of them would be a type for one call.
pub(super) fn start_video_playback(
    path: &PathBuf,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    start: f64,
    volume: u32,
    subtitle: Option<usize>,
) -> Option<PlayerProcess> {
    let volume = volume.min(100);

    // Which player this file is begun with, asked here rather than left to the caller because
    // what a player is told is this function's whole business: the two routes' command lines have
    // nothing in common but the file (see `video_route`).
    let route = video_route(path);

    // The soundtrack's own measured gain, where `Normalize` is on for videos and one has been
    // measured. It is read here rather than below because both routes want it, and on the same
    // terms: a film nothing has measured is played as the file holds it, and what measures it is a
    // read on a thread of its own — the hover after this one is the one that hears it (see
    // `spawn_gain_scan`). A muted film asks for nothing, so nothing is measured for one either.
    let gain = if volume == 0 || !normalizing_video() {
        None
    } else {
        match audio_track::gain(path) {
            // A gain of one is a gain, not the absence of one: a soundtrack already standing at
            // the target is played through the filter like every other measured file.
            Some(gain) => Some(gain),
            None => {
                spawn_gain_scan(path);
                None
            }
        }
    };

    if route == VideoRoute::Mpv {
        return start_mpv_video_playback(path, x, y, width, height, start, volume, gain);
    }

    // Use ffplay for video playback - borderless, positioned at preview location
    let mut cmd = engine_processes::hidden_command("ffplay");

    // Whether the film is decoded on the graphics card rather than on a core, which is the
    // `Video` toggle under `Performance → Hardware Acceleration` in the tray.
    //
    // It is read here rather than at the top of the function because this is where a player is
    // actually begun, and a setting that is read at a hover's beginning is read for the *next*
    // hover: a film already playing was launched with the answer that was on then, and it is kept
    // playing with it rather than being relaunched under the user's feet (see
    // `video_hw_accel_device`).
    if let Some(device) = video_hw_accel_device() {
        cmd.args(["-hwaccel", device]);
    }

    // If volume is 0, disable audio completely for better performance
    if volume == 0 {
        cmd.arg("-an");
    } else {
        // The level, with the soundtrack's own measured gain folded into it where `Normalize` is on
        // for videos and one has been measured: the two multiply, exactly as they do for a sound
        // file (see `start_audio_player`).
        let volume_filter = match gain {
            Some(gain) => format!("volume={:.4}", volume as f64 / 100.0 * gain),
            None => format!("volume={:.2}", volume as f64 / 100.0),
        };
        cmd.args(["-af", &volume_filter]);
    }

    // The geometry is read from the cache and never probed for here: this runs on the
    // preview thread, which is the one thread that must not wait for two external
    // processes — and by the time a player is started for a hover, the probe that sized
    // that hover has already answered (see `probe_video_geometry`). A file whose answer is
    // that there is nothing to measure gets no filter at all, which is the frame as the
    // file holds it.
    //
    // **The sidecar is read out of this same answer rather than out of a lookup of its own, and
    // that is what keeps this thread off the film's folder.** Finding it is a `read_dir` of that
    // folder (`video_launch::sidecar_for`), and this is inside the launch — so a walk here would be
    // paid again by every seek, resize, volume change and track change. The probe resolves it
    // once per file and version, beside the two processes it already runs (see
    // `probe_video_geometry`), and one lookup here answers both facts the chain needs. The copied
    // subtitle files and the answer that their one pass has failed come out of the same read, for
    // the same reason once more: what the launch draws is the probe's answer, not a second look.
    let measured = match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => Some(geometry),
        _ => None,
    };

    let vf = measured.as_ref().map(|geometry| match geometry.crop {
        Some(crop) => format!(
            "crop={}:{}:{}:{},setsar=1",
            crop.width, crop.height, crop.x, crop.y
        ),
        None => "setsar=1".to_string(),
    });

    // The subtitles are drawn into the same filter chain rather than named beside it, because
    // `-sst` is inert on this build: the subtitle stream is demuxed, no subtitle filter is put
    // in the graph, so nothing is drawn whatever track is chosen. What draws is the `subtitles`
    // filter, which initialises libass — and which is appended *after* the crop rather than put in
    // front of it, so the lettering is laid over the cropped picture rather than cropped along
    // with it.
    //
    // Only a small file is ever named — the sidecar beside the film or the copy made out of it —
    // and a track the caller was given chooses which copy, where there is a choice. A file with
    // neither source is a film played without subtitles, which is deliberate and measured: the
    // film's own stream would stream the whole container before a first frame and cost the
    // preview itself (see `video_launch::subtitle_filter`, where that refusal is written out).
    let streams = video_subtitles(path);
    let sidecar = measured
        .as_ref()
        .and_then(|geometry| geometry.sidecar.as_deref());
    let derived = measured
        .as_ref()
        .and_then(|geometry| geometry.derived.as_ref());
    let subtitles = video_launch::subtitle_filter(sidecar, derived, subtitle.or(streams.chosen()));

    // **The copy of this film's own tracks is asked for here, where the film is actually put on
    // screen**, and not by the probe: the pass that makes one is a whole read of the film, so it
    // belongs to a window that is showing it rather than to a hover that may only be passing
    // over. Being asked at the launch is also what gives a dropped copy its second chance:
    // every launch asks again, so a film the user came back to is copied again rather than
    // never (see `subtitle_files::request_extraction`).
    subtitle_files::request_extraction(path, measured.as_ref());

    // What this player is drawn with is written down for the take-up that may adopt it: a pin
    // does not begin a player of its own over a hover's, so a player begun without subtitles
    // is one a pinned window keeps drawing without them until something begins it again — and
    // only the take-up knows it is a pin's player now (see `reload_adopted_subtitles`).
    note_player_subtitles(path, subtitles.is_some());
    let vf = match (vf, subtitles) {
        (None, None) => None,
        (Some(chain), None) => Some(chain),
        (None, Some(filter)) => Some(filter),
        (Some(chain), Some(filter)) => Some(format!("{chain},{filter}")),
    };

    if let Some(vf) = vf.as_deref() {
        cmd.args(["-vf", vf]);
    }

    // Where the player is asked to start. A pinned preview's transport bar is the one caller that
    // asks for anything but the beginning: FFmpeg's player can be told nothing once it is running,
    // so a seek is this player ended and another one begun at the second the bar was dragged to.
    //
    // It is written *after* the file rather than before it, and that placement is not a style
    // choice. Measured on FFmpeg 9.0.2: a `-ss` written before the input makes the `subtitles`
    // filter draw nothing at all — the same frame comes back byte-identical with and without it —
    // while the same `-ss` after the input changes the picture. A time-shifted anime release
    // carries its subtitle track offset from the picture, so a seek that moved the video without
    // drawing the track would leave the two disagreeing about where they are, which is the fault
    // the relaunch names the track for in the first place.
    //
    // Moving the seek past the input does not fix the loop, which is why it is not asked to: with
    // the seek on either side, `-loop 0` wraps the player back to *the seek* rather than to the
    // beginning, so an eight-second film begun at six plays its last two and a half seconds for
    // ever. The loop is therefore taken away from a player that was seeked and given by this app
    // instead (see `video_launch::loop_is_ours` and `video_loop_tick`).
    let seeked = start > 0.0;

    // `-loop 0` is asked for only where it is right. A player begun at the beginning wraps to the
    // beginning by itself and so never reaches its own end, which is what keeps a hover — which
    // begins at zero — from closing its own preview by reaching the last frame of the file.
    if !video_launch::loop_is_ours(seeked) {
        cmd.args(["-loop", "0"]);
    }

    // Which subtitle track, where there is a choice to have made. It is threaded into the filter
    // above rather than named beside the command, because on this build nothing beside the command
    // draws anything (see `video_launch::subtitle_filter`); it is written down by the pin that
    // remembers it rather than by the player, which reports nothing about what it did with the
    // number (see `PinTransport::subtitle`).
    //
    // The index is the specifier's own: FFmpeg counts subtitle streams from zero among themselves,
    // so `s:0` is the first subtitle stream whatever the video and audio streams around it happen
    // to be numbered — which is also the numbering a next-track press walks in, and the one
    // `si=` is given in (see `next_subtitle`).

    let child = cmd
        .args([
            "-err_detect",
            "ignore_err", // Ignore header/stream errors
            "-fflags",
            "+genpts+discardcorrupt+igndts", // Handle missing timestamps & corrupt data
            "-framedrop",                    // Drop undecodable frames instead of stalling
            "-noborder",                     // No window border
            "-left",
            &x.to_string(),
            "-top",
            &y.to_string(),
            "-x",
            &width.to_string(),
            "-y",
            &height.to_string(),
            "-autoexit",
            "-loglevel",
            "quiet",
        ])
        .arg(path)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null());

    // The seek is written after the input file rather than before it, which is the whole of why it
    // is written here and not above: a seek written before the input makes the `subtitles` filter
    // draw nothing, and one written after it does not (measured — see the note over the filter
    // chain). A player begun at the beginning is given no `-ss` rather than `-ss 0`, so that
    // nothing is asked for that would be answered with the position it is already at.
    let child = match (start > 0.0).then(|| format!("{start:.3}")) {
        Some(second) => spawn_player(child.args(["-ss", &second])),
        None => spawn_player(child),
    };

    // After spawning, try to set WS_EX_NOACTIVATE on the player's window
    // to prevent it from stealing focus, and book the player in: its window is the monitor's, its
    // name is what its id is verified as, and its loop — where the loop is this app's to give — is
    // clocked from here (see `note_player_started`).
    if let Some(ref player) = child {
        note_player_started(
            player_image(VideoRoute::Ffplay),
            path,
            player,
            (x, y, x + width, y + height),
            start,
            video_launch::loop_is_ours(seeked),
        );
    }

    child
}

/// Begin an mpv video, in a window of its own over the preview box.
///
/// It is the ffplay launch's own shape — a player spawned and booked in, its window found by pid
/// and styled, its process recorded — with the two things that are mpv's: what it is told (see
/// `video_launch::mpv_args`, where every flag is argued for) and the fact that its loop is always
/// this app's to give. mpv is launched with `--keep-open=yes`, so a film that reaches its end
/// pauses there rather than exiting — and a player that exits is the only thing the tick can see
/// a pass ending by, so nothing would ever begin the film again where the app left the loop to
/// the player (see `video_loop_tick`). It is true of a hover as much as of a pin: the ffplay
/// route's hover loops by FFmpeg's own `-loop 0`, and mpv's is begun again by the tick.
#[allow(clippy::too_many_arguments)] // Each one is a fact about the launch, and a struct of them would be a type for one call.
fn start_mpv_video_playback(
    path: &Path,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    start: f64,
    volume: u32,
    gain: Option<f64>,
) -> Option<PlayerProcess> {
    // The switches are read here, at the launch, for the reason the ffplay route reads its own
    // here: a film already playing was started with the answer that stood then and is kept
    // playing with it (see `video_hw_accel_wanted`, `subtitles_wanted`). The picked mpv is read
    // here for the same reason.
    let chosen = CONFIG
        .lock()
        .map(|config| config.mpv_path.clone())
        .unwrap_or_default();
    let program = video_hw::mpv_program(&chosen, &codecs::mpv_installs())?;

    // The pipe this player is told things on while it runs, named before the process exists and
    // remembered against it once it does, so that a kill drops the name with the player (see
    // `mpv_ipc`).
    let ipc = mpv_ipc::next_pipe_name();

    // **The box is named on the scale of the display mpv's window is created on, which is the
    // primary's** — that is the size Windows scales to the display the film is moved to, and
    // naming it there is the whole of what makes the window the box rather than twice it (see
    // `video_launch::mpv_box_size`, where the measurement behind this is written out). `(0, 0)` is
    // the primary display's own corner, and the primary is the display a player's window is created
    // on whatever box the film is watched in. The box itself is what is written down below and what
    // the tick asserts: only the player's geometry is in the other scale.
    let (asked_width, asked_height) =
        video_launch::mpv_box_size(width, height, dpi_at(x, y), dpi_at(0, 0));

    let mut command = engine_processes::hidden_command(program);
    command
        .args(video_launch::mpv_args(
            path,
            x,
            y,
            asked_width,
            asked_height,
            start,
            volume,
            gain,
            video_hw_accel_wanted(),
            subtitles_wanted(),
            video_launch::mpv_input_conf(),
            &ipc,
        ))
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null());
    let child = spawn_player(&mut command)?;
    mpv_ipc::remember_pipe(child.id(), &ipc);

    // **mpv picks its own tracks, so "was it asked for subtitles" is the honest answer here** —
    // and it is the answer the take-up is asked for: a pin adopts a hover's player rather than
    // beginning one of its own, and this list is what says whether the player it adopted is
    // drawing what the user asked for (see `reload_adopted_subtitles`). On the ffplay route the
    // same fact is whether a filter was named, because nothing else there can draw a subtitle.
    note_player_subtitles(path, subtitles_wanted());

    note_player_started(
        MPV_IMAGE_NAME,
        path,
        &child,
        (x, y, x + width, y + height),
        start,
        true,
    );

    Some(child)
}

/// Book in a player this app has just begun, whichever engine it is.
///
/// The four things are the same four for both players, which is why they are one function: the
/// window is the monitor's to keep in step, the image name is written down beside the id so that
/// what a sweep or a retire verifies before a kill is the player actually begun (see
/// `note_player_process`), the process is recorded for the run that never gets to end it —
/// recorded as the player rather than as an engine, because what ends a player is its hover
/// ending, and a tier being let go of is not its to receive (see `engine_processes`) — and the
/// loop is clocked from here, where the loop is this app's to give rather than the player's own.
///
/// The clock starts at the second the player was actually begun at and not at the moment the
/// process appeared: the header is read before the window is up, and a film of a few hundred
/// megabytes spends a good part of a second in there. The box goes in beside it because a rewind
/// is a relaunch, and a relaunch has to be told where to put its window — the film is not drawn
/// into a surface of this app's, so the replacement's window has to be placed rather than laid
/// out (see `video_loop_action`).
fn note_player_started(
    image: &'static str,
    path: &Path,
    player: &PlayerProcess,
    content: ScreenRegion,
    start: f64,
    own_loop: bool,
) {
    let pid = player.id();
    set_noactivate_for_process(pid, image);
    engine_processes::record_player(image, pid);

    if own_loop {
        note_video_loop(pid, path, content, start, video_duration(path));
    } else {
        forget_video_loop(pid);
    }
}

/// A player that has been started and has not put its window up yet.
///
/// A video's preview is the player's own window, and a player is a process: what stands
/// in for the video until that window is there is the waiting spinner, at the pointer it
/// is shown at for every other kind of wait, and the media the player was started for is
/// held here until it is. Holding it here rather than in `CURRENT_MEDIA` is what keeps
/// the spinner on screen: the frame the player will play into would otherwise be what
/// this app's window is showing, and a video's window is the one the player draws.
pub(super) struct VideoStart {
    /// The frame the player plays into, with the process it was started for.
    pub(super) media: MediaData,
    pub(super) path: PathBuf,
    pub(super) pid: u32,
    pub(super) started: Instant,
}

/// A video the media engine plays whose preview is up but has no frame of the file in it
/// yet, and what is needed to put that preview up once there is one.
///
/// A preview of a video is loaded with a placeholder frame — the frame the engine's own
/// frames land in rather than a picture of anything (see `take_native_video_frame`) — and
/// the engine is asked to play rather than made to: `Load` and `Play` answer before the
/// engine has read the file's header, so a load lands with a placeholder that nothing has
/// been drawn into yet, and the first frame arrives on a tick of its own.
///
/// Opening the window on that placeholder is what is seen as a flash of the backdrop at
/// the start of a hover, and it is not a transparency that could be drawn round: the frame
/// the engine's frames land in is zeroed, and a frame read out of it and forced opaque is
/// an opaque black rectangle — the backdrop flash in a harder form. So nothing is read from
/// it until the engine reports a frame of its own (see `video_player::Session::copy_into`),
/// and the install holds the window back for that report instead. What it is held with is
/// the arrival of the first frame the engine hands over.
#[derive(Clone, Copy)]
pub(super) struct FirstFrameWait {
    /// The hover that held the preview back. A newer hover installs its own media and puts
    /// its own window up, and this is not what reveals that one.
    pub(super) generation: u64,
    /// The hide count the install was under (see `HIDDEN_EPOCH`): a hide since is the
    /// pointer having left the file, and no window is put up for a hover that has gone.
    pub(super) epoch: u64,
    /// Where this hover laid the preview out, which is where the frame is painted: while
    /// the wait is outstanding the window is the wait's own — the spinner's box at the
    /// pointer, or whatever the hover before this one left on screen.
    pub(super) pos: (i32, i32),
}

/// Where this tick puts a held-back preview's first frame up, when this tick is the one it
/// comes up on — and `None` for every tick that is not, including a tick with no wait at all.
///
/// The tick a video's first frame lands on is the only tick with two painters pointed at it.
/// The tick's own answers `needs_repaint`, because a frame was taken; the reveal's answers
/// because the frame is the thing the window was being held back for. They compose the same
/// frame into the same surface, so one of them is a whole frame of copying and a whole hand-off
/// to the compositor for a window nobody can tell the difference of — and at the size of a 4K
/// display that is thirty megabytes copied twice, once a hover.
///
/// The reveal's is the one kept, and not for the saving. It paints at the box this hover was
/// laid out at, and for as long as the wait was outstanding the window was the wait's own: the
/// spinner's box at the pointer, or whatever the hover before this one left on screen. A paint
/// at the window's own rectangle would be at neither place, and the reveal would correct it —
/// which is what a duplicate paint costs besides the copying.
///
/// Everything that can mean *not this tick* is asked here rather than left to the arm that
/// acts on the answer, and the four are the whole of it. A newer generation is a hover that
/// installed its own media since, which this one does not reveal. A hide count that has moved
/// is the pointer having left the file, and no window is put up for a hover that has gone (see
/// `HIDDEN_EPOCH`). A pin up is a window that is the pin's window now, painted by the pin's
/// own painter, which is what the tick's paint will have done. And no frame in hand is a wait
/// that simply goes on. Anything else that can change what a paint would put on screen — the
/// window being moved or resized, another preview installed over it, the surface being rebuilt
/// at a new size — reaches a paint through a tick that sets `needs_repaint` or through the arm
/// that reveals, and neither is answered by this question.
pub(super) fn first_frame_lands(
    wait: &FirstFrameWait,
    generation: u64,
    epoch: u64,
    pinned: bool,
    frame_in_hand: bool,
) -> Option<(i32, i32)> {
    (wait.generation == generation && wait.epoch == epoch && !pinned && frame_in_hand)
        .then_some(wait.pos)
}

/// How long a player is given to put its window up before the wait for it is given up
/// on: a player that has not by then is one that will not, and a spinner that never ends
/// is worse than the desktop it leaves behind.
pub(super) const VIDEO_START_WAIT_SECS: u64 = 10;

/// How the wait for a player stands: whether the video is on screen now, whether the
/// wait is over some other way, or whether the player is still starting.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PlayerWait {
    /// The player's window is up: what is on screen is the video from here, and the media
    /// it was started for is the preview.
    Arrived,
    /// The player is not coming: it died, its hover moved on, or it has taken longer than
    /// a start ever does.
    Abandoned,
}

/// How the wait for a player stands. `None` is a player still starting, which is a wait
/// that goes on.
///
/// A window that is up arrives, the cap included: the video is on screen whether or not
/// the start took long, and ending a player that has one would take a picture away. What
/// is read before that is a player that is gone — what a process leaves behind is a
/// handle and not a window, so a window with no player behind it is not a preview — and
/// a start that has run past `VIDEO_START_WAIT_SECS` with no window to show for it is one
/// this app stops watching, because a spinner that never ends is worse than the desktop
/// it leaves behind.
pub(super) fn player_wait(
    window_up: bool,
    player_alive: bool,
    waited: Duration,
) -> Option<PlayerWait> {
    if window_up && player_alive {
        return Some(PlayerWait::Arrived);
    }

    if !player_alive || waited >= Duration::from_secs(VIDEO_START_WAIT_SECS) {
        return Some(PlayerWait::Abandoned);
    }

    None
}

/// Stop video playback, on the thread the engine belongs to.
pub(super) fn stop_video_playback(media: &mut MediaData) {
    // A video the media engine is playing has no process and no window of its own: letting
    // the engine go is the whole of stopping it, and it is done here because this is where
    // every path that ends a video already comes through — the pointer leaving the file,
    // another preview taking its place, the `Videos` gate closing, a display change, a
    // resume from sleep, and the app itself.
    //
    // A sound is stopped here for the same reason and by the same call: the engine that plays
    // one is this app's own, and a sound FFmpeg plays instead is a process in `video_process`
    // below — the same field, killed the same way, because what ends either is the hover
    // ending.
    //
    // The engine's ending is this thread's to perform and no other's — a session belongs to
    // the thread that started it (see `video_player`'s `SESSION`) — so a take-down that runs
    // on some other thread kills the player process and leaves the media engine for the
    // thread that owns it (see `kill_player_process` and `hide_preview`).
    //
    // What a sound had played of its file is written down here rather than where it is heard:
    // a hover that is ending is a sound whose position is settled, and the memory the mode that
    // resumes one reads back is asked to keep what it has before the clock that measured it is
    // taken down (see `audio_seek::flush`).
    audio_seek::flush();

    if media.media_type.is_native_video() || media.media_type.is_audio() {
        video_player::stop();
    }

    kill_player_process(media);
}

/// End the player process a preview holds, if it has one, and forget its window.
///
/// It is the half of a take-down that any thread may perform — a process is ended with a
/// handle, whatever thread holds it — which is what lets the hide on the Explorer hook's
/// thread end the `ffplay` a hover started without reaching for the media engine, whose
/// session is the preview thread's alone (see `stop_video_playback` and `hide_preview`).
pub(super) fn kill_player_process(media: &mut MediaData) {
    if let Some(ref mut process) = media.video_process {
        // The pipe a player was told things on goes with the player: nothing is left behind for a
        // command to reach once the process is ended (see `mpv_ipc`).
        mpv_ipc::forget_pipe(process.id());
        // Kill only, never wait: a process stuck in kernel I/O would block the
        // caller (possibly the Explorer hook thread) indefinitely. The leftover
        // process checks confirm death and clear VIDEO_PID.
        let _ = process.kill();
    }
    media.video_process = None;
    // Clear the video window HWND. VIDEO_PID stays recorded until the process
    // is confirmed gone, so a surviving ffplay can still be found and killed.
    VIDEO_HWND.store(0, Ordering::SeqCst);
}

/// What is to be done with a player's window, in the one word both of the places that put one
/// somewhere read.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PlayerOrder {
    /// Put the window at its box and raise it above everything: the film is the picture, and
    /// nothing of this app's is standing over it.
    Raise,
    /// Put the window at its box, at the top of the ordinary band: above every ordinary window —
    /// Explorer, and whatever else the user brings up — and still below the topmost pin, whose own
    /// pixel the band the film stands in is (see `film_band_is_hit_testable`, `player_insert_after`).
    Carry,
    /// Leave the window exactly as it stands: a strip of the pin's chrome or its volume popup is
    /// drawn where the film is, or the film is parked for the length of a gesture.
    Leave,
}

/// The placement decision the last call to `player_order` on this thread came to.
///
/// A test run has no screen to read a z-order off, and the answer a raiser acts on is the whole of
/// what it can be asked here — so this is the record a test reads to know that every raiser asked
/// the question at all (see `last_player_order`). It is the thread's own because the suite runs in
/// parallel and a raiser on another test's thread must not answer for this one's. The release
/// binary carries no part of it.
#[cfg(test)]
thread_local! {
    static LAST_PLAYER_ORDER: std::cell::Cell<Option<PlayerOrder>> =
        const { std::cell::Cell::new(None) };
}

/// The placement decision the last call to `player_order` on this thread came to (see
/// `LAST_PLAYER_ORDER`).
#[cfg(test)]
pub(super) fn last_player_order() -> Option<PlayerOrder> {
    LAST_PLAYER_ORDER.with(|order| order.get())
}

/// Write down the decision `player_order` came to, for this thread's test eye alone (see
/// `LAST_PLAYER_ORDER`).
#[cfg(test)]
fn note_player_order(order: PlayerOrder) {
    LAST_PLAYER_ORDER.with(|held| held.set(Some(order)));
}

/// What is to be done with the player's window, from the facts both of the places that put one
/// somewhere can read.
///
/// **A pinned film mpv draws is the one answer the player is never raised over the pin.** Its band
/// is the pin's own pixel rather than a hole onto the player's window, so a press on the film has to
/// reach the pin — which it cannot while the player is the window the point belongs to. The film is
/// therefore carried with the band, at the top of the ordinary band: over every ordinary window, and
/// never in front of the pin, where every other answer is the one it has always been: a hover must
/// go on falling through, and FFmpeg's player keeps the band it has always kept. Which player is up
/// is read off `video_process_image` — the same record the carry itself is read off — rather than
/// off the media, which is a fact about the file on screen and not about whose window stands in its
/// band.
///
/// The raise is held off for the same reasons it always was: a strip of the pin's chrome or the
/// volume popup is drawn where the film is, and a film parked for a gesture is deliberately not on
/// screen. `parked` is asked here for the monitor, which nobody can hold off from outside; the
/// placement road is the one road that must put a parked window back, so it asks this with no park
/// (see `unpark_player_window`).
///
/// The media engine's route never reaches the carry at all: its band is painted into this app's own
/// window and it begins no player, so there is no window of somebody else's in the band for either
/// raiser to place and both find none (see
/// `only_an_mpv_players_window_is_carried_under_the_pins_chrome`).
pub(super) fn player_order(
    image: &str,
    pinned: bool,
    chrome_up: bool,
    volume_open: bool,
    parked: bool,
) -> PlayerOrder {
    // The pin's own film first, because it is the pin's own pixel the press has to land in and
    // that is not the chrome's to refuse; a park is still refused, since a parked film is not on
    // screen to be carried.
    let order = if parked {
        PlayerOrder::Leave
    } else if pinned && image == MPV_IMAGE_NAME {
        PlayerOrder::Carry
    } else if chrome_up || volume_open {
        PlayerOrder::Leave
    } else {
        PlayerOrder::Raise
    };

    #[cfg(test)]
    note_player_order(order);

    order
}

/// The z-order an order is placed with — and the whole of the difference between a film moved under
/// the pin and a film raised over it.
///
/// A `SetWindowPos` places a window *after* the handle it is handed, and there are two bands worth
/// naming: the topmost band, where the pin stands and where every raised film goes, and the top of
/// the ordinary band, which is where a carried film goes. `HWND_TOPMOST` would put the film over
/// the pin and take the press on the picture with it; **`HWND_TOP` is the top of the ordinary
/// band**, so the film is above every ordinary window — Explorer, and whatever else the user brings
/// up — and still under the topmost pin, which is exactly the carry: the band is the pin's own
/// pixel, so the film has to be seen there without being the window the point belongs to (see
/// `player_order`, `film_band_is_hit_testable`).
///
/// **The z-order bit must not be left alone.** This is what the carry used to do, with
/// `SWP_NOZORDER`: `SetWindowPos` ignores its insert-after argument under that flag, so a carried
/// film kept the order it happened to have, and Explorer covered it the moment it was activated —
/// while the pin's band went on taking the press over a film nobody could see. `HWND_NOTOPMOST` is
/// no answer either: it is a no-op for a window that is not currently topmost, and the carry clears
/// the topmost bit in the style immediately before the call (see `player_ex_style_for`).
///
/// It is a function so that a test can read it: a screen is the only other place the order of two
/// windows can be seen, and the suite has none.
pub(super) fn player_insert_after(order: PlayerOrder) -> HWND {
    // A `Leave` never reaches a placement — both raisers refuse it before asking (see
    // `ensure_video_window_topmost`, `apply_noactivate_to_hwnd`) — and the raise is the ordinary
    // answer for every other order.
    match order {
        PlayerOrder::Carry => HWND_TOP,
        PlayerOrder::Raise | PlayerOrder::Leave => HWND_TOPMOST,
    }
}

/// Put the ffplay window where it belongs and on top of everything, re-discovering and re-asserting
/// its style by pid so that a window ffplay recreated is styled again rather than inherited.
///
/// **A pinned film mpv draws is carried here rather than raised over the pin**, which is the one
/// answer `player_order` gives that is not the raise: the pin's window is the one the point in the
/// band belongs to, so the player is put where the band is and at the top of the ordinary band —
/// over every ordinary window, and still behind the pin (see `player_insert_after`).
///
/// **This is the one place a player's window is raised, and the volume popup and the pin's
/// chrome are guarded here rather than by their callers.** That is the whole of the
/// arrangement: what either one floats over is the media, and the media of a video FFmpeg
/// plays *is* this window — a popup is drawn over the band, and a strip of chrome is
/// painted into the pin's own rows, which for this kind are rows of this same window, so
/// a raise while either is up puts the film over the thing the user is looking at. The
/// tick was guarding the popup here,
/// which covered the periodic re-assertion and nothing else — a seek, a resize settling and the
/// first appearance of a pinned player all raise the window from their own call sites, and those are
/// exactly the moments a popup is most likely to be open, because all three of them are things a
/// pinned window does while it is being looked at. Guarding here rather than in four places is what
/// makes it true rather than nearly true.
///
/// **The drag's park is guarded by its callers rather than here, and the reason is that this
/// function is one of the ways a park ends.** A relaunch is a new player with a window of its own,
/// and it is the relaunch that puts that window up and clears the flag (see `restart_pinned_player`);
/// a guard here would answer the unpark itself with "parked" and leave the replacement hidden behind
/// a band painted flat for it. The callers are the two that spend a drag's life raising a window that
/// should not be on screen — a drag's own pointer moves and the tick's re-assertion — and they read
/// the same flag (see `pin_player_is_parked`); the monitor thread's own raise, which nobody can hold
/// off from outside, is guarded where it is raised (see `apply_noactivate_to_hwnd`).
pub(super) fn ensure_video_window_topmost(x: i32, y: i32, width: i32, height: i32) -> bool {
    // What is to be done with the window, from the one answer both of the places that put it
    // somewhere read: a film mpv is drawing under a pin is carried, not raised over the pin, and a
    // raise the pin's own chrome or popup has held off is nothing this road places (see
    // `player_order`).
    //
    // The park is deliberately not asked of it here: this is the road that must put a parked
    // window back (see `unpark_player_window`), so the park is the callers' refusal and the
    // monitor's rather than this one's (see the note over the function).
    let order = player_order(
        video_process_image(),
        pinned(),
        pin_chrome_up(),
        pin_volume_open(),
        false,
    );
    if order == PlayerOrder::Leave {
        return false;
    }

    let carrying = order == PlayerOrder::Carry;

    // Re-discover/re-apply style by PID each time to survive ffplay window recreation
    // and keep topmost state resilient over time.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid != 0 {
        unsafe {
            let _ = try_apply_noactivate_style(pid);
        }
    }

    let hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
    if hwnd_val == 0 {
        return false;
    }

    unsafe {
        let hwnd = HWND(hwnd_val as *mut std::ffi::c_void);
        if hwnd.is_invalid() {
            VIDEO_HWND.store(0, Ordering::SeqCst);
            return false;
        }

        // Re-assert desired style bits in case ffplay modified them — through the one list, so a
        // window carried under the pin's own band is not asserted a topmost bit here while the
        // monitor leaves it off (see `player_ex_style_for`, `player_order`).
        let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
        let new_style = player_ex_style_for(current_style, carrying);
        if new_style != current_style {
            SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_style);
        }

        // The z-order the order is placed with — the carry's own word is `HWND_TOP`, the top of the
        // ordinary band, and not `SWP_NOZORDER`, which would leave a carried film under Explorer
        // (see `player_insert_after`).
        let insert_after = player_insert_after(order);

        let _ = SetWindowPos(
            hwnd,
            insert_after,
            x,
            y,
            width,
            height,
            SWP_NOACTIVATE | SWP_SHOWWINDOW,
        );
        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    }

    true
}

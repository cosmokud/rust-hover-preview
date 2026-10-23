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

/// Whether a hovered film loops, read when a player is started rather than when the
/// setting changes: a preview is a few seconds long, and the next one is played at
/// whatever the setting is by then.
pub(super) fn current_video_loop() -> bool {
    CONFIG
        .lock()
        .map(|cfg| cfg.video_loop)
        .unwrap_or(DEFAULT_VIDEO_LOOP)
}

/// What a *pinned* film does at the end of it — the film's own end-of-file setting, read the
/// way the hover's loop is and at the same moment.
///
/// `Pause` is the film left where it ended, `Loop` the film beginning again, and `Continue`
/// the file the walk takes next — which is the file beside it, or a random film of the
/// folder where the pin's own shuffle says so (see `pin_video_loops` and `pin_walk`).
pub(super) fn current_pin_mode_video_end_of_file() -> PinEndOfFile {
    CONFIG
        .lock()
        .map(|cfg| cfg.pin_mode_video_end_of_file)
        .unwrap_or(DEFAULT_PIN_MODE_VIDEO_END_OF_FILE)
}

/// Whether the film `path` is showing goes round on its own: a hovered film loops where its
/// own loop switch says so, and a pinned one where the pin's end of file is a loop — every
/// other film plays its pass once, and the end of it is what either moves the pin on, steps
/// through the folder, or stops it (see `loop_ended_pinned_player`).
pub(super) fn pin_video_loops(path: &Path) -> bool {
    if pinned() && pinned_path().as_deref() == Some(path) {
        current_pin_mode_video_end_of_file() == PinEndOfFile::Loop
    } else {
        current_video_loop()
    }
}

/// Whether the pinned film `path` is one that stops at the beginning of itself at the end of a pass,
/// on the one road this app can take a *running* player back to its first frame: mpv's pipe.
///
/// `End of File → Pause` is the one end of file that neither begins the pass again nor steps on, and
/// what it does at the end of a pass is what a loop does — take the film back to its first frame
/// down the pipe the player listens on (see `video_loop_tick`) — and then leave it standing there.
///
/// **And the route is part of the question rather than an accident of it.** The routes under mpv are
/// players nothing of this app's can be told once they are going, so the tick's rewind answers one of
/// them by *ending* it and beginning another from zero, playing (see `video_loop_tick`) — a `Pause`
/// that plays the film again, pass after pass, which is the one thing this end of file exists not to
/// do. What those routes do instead is stop the film where it stands when its end arrives, which is a
/// hold of the player rather than a rewind of the film (see `loop_ended_pinned_player`).
///
/// **At the top and not at the end**, which is the whole of what keeps the press that plays it again
/// a clean one. A film held at the top is resumed by that press with a single command, the same
/// resume every other held film already has (see `toggle_pinned_playback`). A film held at its own
/// end has to be taken to the top *by* the press, and a seek written against a player that has
/// stopped at the end of its file is answered by mpv as a restart of the film's whole chain — its
/// own log has the decoder reinitialised and the first frame shown again (`hr-seek, skipping to
/// 0.000000`, `Reinit context to 320x240`, `first video frame after restart shown`, measured in
/// mpv 0.41.0's own log) — and the picture the chain owns goes with it, which is the band showing
/// the desktop for the moment in between. The picture the pin is left showing is the same either
/// way: the frame the press would have put there.
pub(super) fn pin_video_pauses_at_the_top(path: &Path) -> bool {
    pinned()
        && pinned_path().as_deref() == Some(path)
        && current_pin_mode_video_end_of_file() == PinEndOfFile::Pause
        && video_route(path) == VideoRoute::Mpv
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
    #[cfg_attr(test, allow(dead_code))]
    // The test seam notes a player instead of beginning one.
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

/// Whether the test build spawns a real player where a test asks it to, the way the mpv pipe
/// writes to a real one where one does (see `WRITE_TO_A_REAL_PIPE`): the noting above is what
/// keeps every other test off the desktop, and a loop that watches real windows come and go is
/// the one test that cannot run on noted ones.
#[cfg(test)]
static SPAWN_REAL_PLAYER: AtomicBool = AtomicBool::new(false);

/// Begin real players until told otherwise, or note them again. Only a loop watching real
/// windows sets this; every other test inherits noted players, which begin nothing, show
/// nothing and can never end a real process (see `spawn_player`).
#[cfg(test)]
pub(super) fn spawn_real_player(real: bool) {
    SPAWN_REAL_PLAYER.store(real, Ordering::SeqCst);
}

/// The test build's launch: note a player instead of beginning one (see `spawn_player`).
#[cfg(test)]
pub(super) fn spawn_player(command: &mut std::process::Command) -> Option<PlayerProcess> {
    if SPAWN_REAL_PLAYER.load(Ordering::Acquire) {
        return command.spawn().ok().map(PlayerProcess::Started);
    }
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

/// The place a player's window is begun at: the box the caller asked for, or — while the pin is a
/// bubble — the same box moved off the desktop altogether.
///
/// **A player puts its window up the moment it has one, and that is not a window anybody wants
/// behind a bubble.** The picture a bubble shows is painted by this app out of frames read beside
/// the player (see `bubble_stream`), so the player's own window is wanted nowhere on the screen
/// until the pin is put back up — and a window that appears where the pin's band was is a black
/// rectangle over the desktop, because a window that has just been shown has drawn nothing yet.
/// Hiding it once it exists is a frame too late, which is what the monitor thread's hide is and
/// what the eye catches (see `apply_noactivate_to_hwnd`).
///
/// So a player begun while the pin is collapsed is begun *off* the desktop instead of being hidden
/// afterwards: `SM_XVIRTUALSCREEN` is the leftmost edge the desktop has, so a window of this width
/// entirely beyond it is over no display at all. The *size* is deliberately kept and only the
/// place moved — a player renders at the size it was begun at, so a window begun at nothing would
/// come back up as a film scaled to nothing — and the restore places it at the band in its own
/// tick rather than a tick later (see `restore_pin`, which places its siblings as it puts the
/// window up).
///
/// What this is not is a rule about hovers: a hover has no pin to be a bubble of, so its player is
/// begun at the box its preview was laid out at, where it is the picture and is wanted.
pub(super) fn bubble_players_box(x: i32, y: i32, width: i32) -> (i32, i32) {
    if !(pinned() && pin_is_collapsed()) {
        return (x, y);
    }

    let left = unsafe { GetSystemMetrics(SM_XVIRTUALSCREEN) };
    let right_of_the_desktop = left - width.max(1) - 1;
    (right_of_the_desktop, y)
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

    // **And where its window goes is asked before either route is, because it is the same question
    // for both** (see `bubble_players_box`).
    let (x, y) = bubble_players_box(x, y, width);

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

    // **And no `-hwaccel`, which is not an oversight.** It names a *device* rather than asking
    // for a decoder, and ffplay brings its Vulkan renderer up the moment the flag is present:
    // measured on this machine, every candidate device dies of an access violation in the graphics
    // driver with no picture at all, where the same launch without the flag plays through in
    // software. So hardware decode is mpv's switch alone and FFmpeg's player decodes on a core
    // (see `video_hw_accel_wanted`).

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

    // Whether the film goes round on its own, read here, at the launch, for the reason
    // the volume beside it is: a film already playing was launched with the answer that
    // stood then and is kept playing with it (see `pin_video_loops`).
    let loop_ = pin_video_loops(path);

    // `-loop 0` is asked for only where it is right. A player begun at the beginning wraps to the
    // beginning by itself and so never reaches its own end, which is what keeps a hover — which
    // begins at zero — from closing its own preview by reaching the last frame of the file.
    // And only where the end of file says the film goes round: a film that is not to be looped plays
    // its pass once, and its exit is what the pin's end answers — with a pass begun again, the
    // next file, or nothing, as the end of file says (see `loop_ended_pinned_player`).
    if !video_launch::loop_is_ours(seeked) && loop_ {
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

/// The window an mpv player begun for this film is told to draw into, as `--wid` is given it: the
/// pin's own child for a pinned mpv film, and nothing for every other launch — a hover's film, a
/// pin standing in the preview's window, and FFmpeg's player and the media engine's own frame,
/// which are not this route at all (see `pin_film_host::host_for_the_pins_film`).
///
/// **The one answer both of the launch's questions are read from.** A player given a window of
/// this app's own is told it and no place of its own, and it is owed no topmost ask either: the
/// pin is the host's order, so the ask that is the desktop's to answer is the top-level road's
/// alone (see `mpv_ipc::set_ontop`).
pub(super) fn mpv_launch_host(path: &Path) -> Option<u32> {
    pin_film_host::host_for_the_pins_film(path).map(|host| host.0 as usize as u32)
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

    // The window this player draws into, where it is the pin's to give: a pinned mpv film is a
    // child of the pin's own window, and a player told about it has no top-level window for the
    // tick to place (see `mpv_launch_host`).
    let host = mpv_launch_host(path);

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
            host,
        ))
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null());
    let child = spawn_player(&mut command)?;
    // **The host holds this player and no older one**, recorded before the player has a window of
    // its own in there: the wait for the film's first frame is a wait for this player's child and
    // not the last film's, which is what the pid beside the host is for (see `pin_film_host`).
    if host.is_some() {
        pin_film_host::note_hosted_player(child.id());
    }
    // The pipe the player is told on, and the one thing a later command has to know about the
    // launch: whether it gave the player any audio at all (see `mpv_launched_without_audio`,
    // `mpv_ipc::remember_pipe` and `set_volume`).
    mpv_ipc::remember_pipe(
        child.id(),
        &ipc,
        video_launch::mpv_launched_without_audio(volume),
    );

    // **A player begun for a pin is told the pin's own answer about the topmost band before it has
    // drawn anything.** `--ontop` is the launch's flag for a hover's window, and a pin that does
    // not hold the front wants the opposite of it: asking here rather than at the launch's
    // arguments is what keeps one command line for both, and the pipe is already remembered so the
    // ask is on the player this app holds (see `mpv_ipc::set_ontop`).
    //
    // **And a player in the pin's own host is not asked at all**: it has no place in the desktop's
    // order for the answer to be about, and the pin is the host's order (see `mpv_launch_host`).
    if host.is_none() && pinned() && !pin_holds_the_front() {
        let _ = mpv_ipc::set_ontop(child.id(), false);
    }

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
        // mpv's loop is always this app's to give — it pauses on its last frame rather than
        // exiting — so the clock is kept for every film mpv is begun with, whatever the end of file says
        // at the launch: the end of file is asked when the end comes, which is the only way a film the
        // *hover* played once can be looped by the pin it becomes (see `video_loop_action`).
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
/// **`our_loop` is whether the loop of this player is this app's to give, and not whether it is
/// going round.** The end of file is read at the tick instead, which is what lets a film begun for a hover
/// be looped — or stopped — by the pin it becomes: mpv is begun once and kept, while FFmpeg's
/// player wraps to the beginning by itself only where it was begun at one, so that half stays a
/// fact of the launch (see `loop_is_ours` and `pin_video_loops`).
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
    our_loop: bool,
) {
    let pid = player.id();
    set_noactivate_for_process(pid, image);
    engine_processes::record_player(image, pid);

    if our_loop {
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

/// How the wait for a *swap's* player stands: the same three answers, at a swap's own bound
/// instead of a hover's.
///
/// It is `player_wait`'s question asked of the one road that already has something on screen. A
/// hover holds back a window with a spinner behind it and nothing under the spinner that could
/// arrive sooner, so ten seconds is a start nobody has to watch; a swap is waiting with the
/// file the pin is leaving still up, at the box the pin is still at, and every tick of the wait
/// is a tick of that old picture. So the bound is the one the cover a step raises is settled on
/// — a player begun for a replacement has to open the file, seek, decode hardware and present,
/// and a wait past that is a player this app stops watching (see `PIN_PARK_SWAP_TIMEOUT` and
/// `settle_pinned_park`).
///
/// What the wait is *for* is the same fact the cover above it is settled on: the player's own
/// window is the only sign a player of this kind gives that its film is on screen, so a window
/// that is up has arrived, cap or no cap (see `player_wait`).
pub(super) fn player_swap_wait(
    window_up: bool,
    player_alive: bool,
    waited: Duration,
) -> Option<PlayerWait> {
    if window_up && player_alive {
        return Some(PlayerWait::Arrived);
    }

    if !player_alive || waited >= PIN_PARK_SWAP_TIMEOUT {
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
    /// Put the window at its box in the topmost band, and put the pin back over it where one is
    /// up: the film is the picture, and the window that keeps the press on it is the pin's (see
    /// `place_player_window`).
    Raise,
    /// Leave the window exactly as it stands: a strip of the pin's chrome or its volume popup is
    /// drawn where the film is, or the film is parked for the length of a gesture.
    Leave,
}

// The placement decision the last call to `player_order` on this thread came to.
//
// A test run has no screen to read a z-order off, and the answer a raiser acts on is the whole of
// what it can be asked here — so this is the record a test reads to know that every raiser asked
// the question at all (see `last_player_order`). It is the thread's own because the suite runs in
// parallel and a raiser on another test's thread must not answer for this one's. The release
// binary carries no part of it.
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

// How many times a placement on this thread has put the pin back over the film it raised.
//
// The other half of the same invariant, and the half `LAST_PLAYER_ORDER` cannot carry: the answer
// is one fact and what a placement does with it is two — a window raised, and the pin put back
// over it — so a test that read only the answer could not tell a raise that leaves the pin under
// the film from one that does not. It is read and cleared together, so what a test asks is what
// the placement it caused did rather than what an earlier one did (see `take_pin_re_raises`).
#[cfg(test)]
thread_local! {
    static PIN_RE_RAISES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// How many times the pin has been put back over a raised film on this thread since this was last
/// asked (see `PIN_RE_RAISES`).
#[cfg(test)]
pub(super) fn take_pin_re_raises() -> usize {
    PIN_RE_RAISES.with(|held| held.replace(0))
}

/// Count one placement that put the pin back, for this thread's test eye alone (see
/// `PIN_RE_RAISES`).
#[cfg(test)]
fn note_pin_re_raise() {
    PIN_RE_RAISES.with(|held| held.set(held.get() + 1));
}

/// What is to be done with the player's window, from the facts both of the places that put one
/// somewhere can read.
///
/// **The answer is the raise for every film, mpv's included.** mpv's window is topmost like
/// FFmpeg's, and a film standing anywhere else is the fault this answer exists to not have:
/// Explorer is raised over whatever stands in the ordinary band the moment the user clicks into
/// it, while the band painted for the film goes on taking the press (see `film_band_is_hit_testable`).
/// Which film a pin draws is read off `video_process_image` — the same record the placement is
/// read off — rather than off the media, which is a fact about the file on screen and not about
/// whose window stands in its band.
///
/// The raise is held off for the same reasons it always was: a strip of the pin's chrome, the
/// volume popup or the card's menu is drawn where the film is, and a film parked for a gesture is
/// deliberately not on screen. `parked` is asked here for the monitor, which nobody can hold off
/// from outside; the placement road is the one road that must put a parked window back, so it asks
/// this with no park (see `unpark_player_window`).
///
/// **The hold-offs stand for mpv's film as they do for FFmpeg's.** A raise put over the pin
/// covers the caption and the transport bar the pin draws into the band over the media (see
/// `pin_overlay_chrome`), then the pin is put back over the film — a two-step reorder per tick,
/// per monitor pass and per drag move, which is the blink the chrome was seen doing. What keeps
/// the film visible under the chrome instead is mpv's own window holding its band from the first
/// frame it draws (see `--ontop` in `mpv_args`), and what keeps the press the pin's is the pin's
/// own raise, which the chrome block and the volume arm already do (see `event_loop`). What the
/// film still owes a band that has moved is the *place*, and that travels without a raise (see
/// `ensure_pinned_sibling_box`).
///
/// **And the card's menu is a hold-off of the same kind, for a window of its own** (see
/// `pin_menu_window`): the panel is drawn over the film, and every raise it gets is
/// `render_pin_menu`'s own (see `event_loop`, `open_pin_menu`). A raise of the film is therefore
/// the film over the panel for as long as the next tick takes — a panel that blinks wherever it
/// overlaps the picture, which is where a hand reads it (see `pin_menu_open`).
///
/// The media engine's route never reaches a placement at all: its band is painted into this app's
/// own window and it begins no player, so there is no window of somebody else's in the band for
/// either raiser to place and both find none (see
/// `only_an_mpv_players_window_is_placed_with_the_pins_chrome_up`).
pub(super) fn player_order(
    _image: &str,
    _pinned: bool,
    chrome_up: bool,
    volume_open: bool,
    menu_open: bool,
    parked: bool,
) -> PlayerOrder {
    // A park, a strip of the pin's chrome, the volume popup and the card's menu are all refused for
    // every window there is: a parked film is not on screen to be raised, and on any of the other
    // routes a raise puts the film over what this app is drawing — the caption and the bar in the
    // pin's strips over it, and the panel of a card's menu in a window of its own over them. mpv's film
    // holds its own band (`--ontop`) while the pin's own raise carries the press, and a band that
    // has moved is placed without a raise (see the note above and `ensure_pinned_sibling_box`).
    let order = if parked || chrome_up || volume_open || menu_open {
        PlayerOrder::Leave
    } else {
        PlayerOrder::Raise
    };

    #[cfg(test)]
    note_player_order(order);

    order
}

/// Put a player's window at `box_` in the topmost band, and put the pin back over it where one is
/// up.
///
/// **This is the one place either raiser puts a player's window somewhere**, so the two cannot
/// disagree about the band it goes in or about what has to be standing over it (see
/// `ensure_video_window_topmost`, `apply_noactivate_to_hwnd`).
///
/// **The raise is `HWND_TOPMOST`, and never a placement left to `SWP_NOZORDER`.** `SetWindowPos`
/// ignores its insert-after argument under that flag, so a window placed that way keeps whatever
/// order it happened to have — and any ordinary window the user activates, Explorer first among
/// them, is then raised over the film, while the band painted for it goes on taking the press (see
/// `film_band_is_hit_testable`).
///
/// **And the pin is put back over the film it just raised.** Both windows are topmost, so which of
/// the two stands in front is a question of which was raised last, and this app raises the player
/// on every tick and on every monitor pass — so the order is put in rather than left to whoever
/// raised last. What the pin's window owes above the film is the caption and the transport bar of
/// a video pin, which are drawn *into* the band over the media (see `pin_overlay_chrome`), and the
/// press on the picture, which has to be the pin's. This is this app's half of the belt-and-braces
/// with mpv's own `--ontop`: mpv holds its window up from the first frame it draws, and this is
/// what answers for the pin standing over it.
///
/// **A `Leave` never reaches here** — both raisers refuse it before asking (see `player_order`) —
/// so the pin is not re-raised for a park, and the three hold-offs a `Leave` stands for are
/// answered where they are caused: a strip of the pin's chrome and an open volume popup each come
/// with a raise of the pin's own window (see `event_loop`'s chrome block, `toggle_pin_volume`), and
/// an open card's menu is a window of its own that the tick raises over the film every tick (see
/// `render_pin_menu`). Nor is a
/// hover's placement followed by one, because a hover has no pin to put over anything. What is
/// left is the pin being put back exactly where a pinned film was raised.
///
/// Where the window goes is the caller's own box and flags, so the monitor's
/// `SWP_NOMOVE`/`SWP_NOSIZE` and a tick's box are the same placement either way.
pub(super) unsafe fn place_player_window(
    hwnd: HWND,
    order: PlayerOrder,
    box_: ScreenRegion,
    flags: SET_WINDOW_POS_FLAGS,
) {
    // **A film the pin draws into its own host is not placed from the outside.** Its window is a
    // child of the pin's and the host's box is the band (see `ensure_pinned_sibling_box`), so there
    // is no place of the desktop's for this road to put it in and nothing for the pin to be raised
    // back over (see `pin_film_host`).
    if pin_film_host::the_pinned_player_was_begun_in_the_host() {
        return;
    }

    let (x, y, width, height) = box_;
    let _ = SetWindowPos(hwnd, player_insert_after(), x, y, width, height, flags);

    // A raise with no pin standing is a hover's, and there is nothing to put over it — and a pin
    // that does not hold the front is raised over nothing at all: the film was just put *below* it,
    // which is the whole of the order that pair owes (see `player_insert_after`).
    let re_raise = order == PlayerOrder::Raise && pin_holds_the_front();

    #[cfg(test)]
    if re_raise {
        note_pin_re_raise();
    }

    if re_raise {
        raise_pinned_window(HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _));
    }
}

/// What a player's window is put under: the pin's own window while a pin is up and keeps the
/// ordinary band, and the topmost band where there is no pin to go under or the pin is Always On
/// Top.
///
/// **It is the pin's presence that asks this, never whether the pin holds the front.** A hover's
/// film holds the band because nothing of this app's stands over it, and so does the film of a pin
/// that is Always On Top, whose band is the setting's. What a pin in the ordinary band owes its
/// film is the place directly below itself — the front it holds or gave back is the pin's own and
/// moves no window of somebody else's.
///
/// **The pin's handle rather than a z-order command, because the pin is not to move.** With the
/// setting off the pin is an ordinary window, and one raised above all the others every time its
/// chrome came up — or every time a film was placed — is a window that is in front of whatever the
/// user is working in without ever having been asked to be, which is the setting saying the
/// opposite of what it did. What the chrome is owed is the *film's* place, and the film is the half
/// of the pair that may be given one: placed directly below the pin it is seen through the band and
/// drawn over by the caption and the bar, and the pair keeps exactly the place in the z-order the
/// user's own clicks gave it. That an insert-after handle that is not topmost also takes the window
/// *out* of the topmost band is the other half of it, and is what turns a film adopted from a
/// hover — whose window wore that band for the preview's sake — into an ordinary window again.
///
/// **A page the engine draws is not asked this**: it is put in front of the pin rather than behind
/// it, because a browser's surface blended through the band's hole is a resize that goes black (see
/// `Host::put_window`). What the two windows share is the band, not the side of the pin they stand
/// on.
pub(super) fn player_insert_after() -> HWND {
    // The band is the *hover's* to hold, and the setting's while the pin is Always On Top. What a
    // pin in the ordinary band keeps is the film under it, holding the front or having given it
    // back. Asked whether the pin holds the front instead, a demoted pin's film was put back in the
    // topmost band by the next placement — every tick — which is a picture floating over every
    // program with `Always On Top` off.
    if !pinned() || pin_always_on_top() {
        return HWND_TOPMOST;
    }

    let pin = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
    if pin.is_invalid() {
        // A pin that has no window of its own yet has nothing for the film to go under: it is
        // put in the ordinary band, which is where it was going to be anyway.
        HWND_NOTOPMOST
    } else {
        pin
    }
}

/// Put a pinned film directly below the pin's own window, so the caption and the transport bar
/// are drawn over the picture and neither window is raised above the program the user is working
/// in.
///
/// **This is the whole of what raising a pin that does not hold the front means** (see
/// `raise_pinned_window`): the pin has no fight to win and no business being raised, and the film
/// is the window that may be moved to give the chrome its place.
pub(super) unsafe fn put_player_under_the_pin() {
    let player = HWND(VIDEO_HWND.load(Ordering::SeqCst) as *mut _);
    if player.is_invalid() || !IsWindow(player).as_bool() {
        return;
    }

    let pin = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
    if pin.is_invalid() || pin == player {
        return;
    }

    // Safety: the handle is a player this app started — or the window the engine draws into — and
    // the place it is put is a z-order position under this app's own window, with neither window
    // moved or sized. A window that has since gone is refused by the call rather than acted on.
    let _ = SetWindowPos(
        player,
        pin,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
    );
}

/// Whether the player's window is already directly below the pin's, which is the answer that keeps
/// the tick from asking the Shell for an order it already has.
///
/// A pin with no player's window of its own — a picture, a page, the engine's own rows — has
/// nothing out of place to report: the answer is yes.
pub(super) unsafe fn player_is_under_the_pin(pin: HWND) -> bool {
    let player = HWND(VIDEO_HWND.load(Ordering::SeqCst) as *mut _);
    if player.is_invalid() {
        return true;
    }

    GetWindow(player, GW_HWNDPREV).is_ok_and(|above| above == pin)
}

/// Tell a pinned mpv film whether its own window is to hold the topmost band, which is the band
/// the pin holds rather than the player's own.
///
/// A film is launched `--ontop` because a hover is a window nothing may cover, and a film *adopted*
/// by a pin keeps what it was begun with — so a pin that does not hold the front would go on
/// standing over the program the user is working in, taking the clicks in the band they overlap,
/// while the pin belongs behind it. The same ask answers the band given back or owed again under a
/// film that is already running.
///
/// Nothing is asked where no mpv is the film on screen: FFmpeg's player is placed by this app and
/// has no opinion of its own to correct (see `mpv_ipc::set_ontop`).
pub(super) fn settle_pinned_mpv_ontop() {
    if !pinned() {
        return;
    }

    // **A film in the pin's own host has no band of the desktop's to be told about.** Its window is a
    // child of the pin's, so the pin is the order it stands in and the flag `--ontop` would hold is
    // not a window the user can see over anything (see `mpv_launch_host`).
    if pin_film_host::the_pinned_player_was_begun_in_the_host() {
        return;
    }

    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid == 0 || video_process_image() != MPV_IMAGE_NAME {
        return;
    }

    let _ = mpv_ipc::set_ontop(pid, pin_holds_the_front());
}

/// Put the ffplay window where it belongs and on top of everything, re-discovering and re-asserting
/// its style by pid so that a window ffplay recreated is styled again rather than inherited.
///
/// **A player's window is raised here — and the pin is put back over it with it**, for every film
/// this app shows, mpv's included (see `place_player_window`, which is where both of those happen).
///
/// **This is the one place a player's window is raised, and the volume popup and the pin's
/// chrome are guarded here rather than by their callers.** That is the whole of the
/// arrangement: what either one floats over is the media, and the media of a video FFmpeg
/// plays *is* this window — a popup is drawn over the band, and a strip of the pin's chrome
/// stands over it in a window of its own, so a raise while either is up puts the film over the
/// thing the user is looking at. The
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
///
/// **What a hold-off refuses is the raise, and never the band.** A player's window that is not on
/// screen is owed the band whatever is up over it — the band's rows are this app's own transparent
/// pixels, painted for that window to show through (see `film_band_is_hit_testable`), so one left
/// hidden is a hole with the desktop standing in it, which is what a park leaves behind and what the
/// swap out of a park fills *before* the cover comes down (see `unpark_pinned_player`). A window
/// that is not being drawn is not carried over a strip the pin is showing by being put back
/// (see `put_a_player_that_is_not_on_screen_back`).
pub(super) fn ensure_video_window_topmost(x: i32, y: i32, width: i32, height: i32) -> bool {
    // A collapsed pin has no band on screen, so nothing of the *pin's* is placed — not even the
    // hidden window a hold-off would otherwise put back: that place is a chrome-less film over
    // Explorer beside the bubble (see `hide_pinned_windows`).
    //
    // **And a collapsed pin places no window of anybody else's at all.** What a bubble shows of a
    // film is a picture it paints itself out of the frames read beside the player (see
    // `bubble_stream`), so there is nothing of the player's to carry into the circle and nothing to
    // put back where something has moved it.
    if pin_is_collapsed() {
        return false;
    }
    // **A film drawn into a child of the pin's has no window of the desktop's to order, style or
    // raise.** The host is the picture and the pin is the order it stands in, so showing it is the
    // whole of what the band is owed — and the box this road was handed is a top-level player's,
    // not the host's. It is asked of the pairing rather than of the host alone: a pin that took up
    // a hover's top-level player has a host standing empty until a relaunch fills it, and that
    // player still has a window this road places (see `pin_film_host`).
    if pin_film_host::the_pinned_player_was_begun_in_the_host() {
        // The band is owed its picture back, and the picture is the host — but only once the player
        // has put a child of its own in it: a host shown empty is a black rectangle over the band,
        // which is the film the hover was playing, blanked (see `pin_film_host::create`).
        //
        // **Where the host stands is not this road's to say**, and the box it was handed is a
        // top-level player's rather than the band's: the band's box is the pin's, and it is given
        // by the roads that read it (see `ensure_pinned_sibling_box`).
        if pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)) {
            // SAFETY: the host is this app's own window, made a child of the pin's, and `ShowWindow`
            // takes any window and reports rather than faults.
            unsafe { pin_film_host::set_the_pin_film_host_visible(true) };
        }
        return true;
    }
    // What is to be done with the window, from the one answer both of the places that put one
    // somewhere read: everything held off — the pin's own chrome, its popup or its menu — is
    // nothing this road places, and a raise puts the pin back over the film it raised (see
    // `place_player_window`).
    //
    // The park is deliberately not asked of it here: this is the road that must put a parked
    // window back (see `unpark_player_window`), so the park is the callers' refusal and the
    // monitor's rather than this one's (see the note over the function).
    let order = player_order(
        video_process_image(),
        pinned(),
        pin_chrome_up(),
        pin_volume_open(),
        pin_menu_open(),
        false,
    );
    if order == PlayerOrder::Leave {
        // The raise is what is held off and the band is still owed (see the note over this
        // function): a window that is not on screen goes back where it stands in the band, with no
        // raise for anything to be put back over.
        put_a_player_that_is_not_on_screen_back(x, y, width, height);
        return false;
    }

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

        // Re-assert desired style bits in case ffplay modified them, through the one list the
        // monitor asserts too (see `player_ex_style_for`).
        let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
        let new_style = player_ex_style_for(current_style);
        if new_style != current_style {
            SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_style);
        }

        place_player_window(
            hwnd,
            order,
            (x, y, width, height),
            SWP_NOACTIVATE | SWP_SHOWWINDOW,
        );
        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    }

    true
}

/// Put a player's window back in the band it is owed where it is not on screen, without the raise
/// the hold-off refuses and without touching the order either window stands in.
///
/// **The half of a hold-off that cannot mean a hole.** What a strip of the pin's chrome and an open
/// volume popup hold this road off from is the *raise*: a raise is what carries the film over the
/// strip the pin paints into the band (see `player_order`, `pin_overlay_chrome`). A window that is
/// not on screen is the other thing entirely, and the band is painted for it rather than over it —
/// this app's own transparent rows, which are the desktop when nothing of the player's is drawing in
/// them (see `film_band_is_hit_testable`).
///
/// **Which is the state a park leaves, and the reason the park is not asked about here.** The window
/// goes back up before the flag that hides it comes down, so that no frame falls between a band that
/// is a hole and the window that fills it (see `unpark_pinned_player`, `settle_pinned_park_onto`); a
/// road that asked the flag would refuse the very place the swap is made of. Every other caller asks
/// the park itself and never reaches this with a parked window (see `event_loop`'s tick,
/// `show_current_replacement`, `ensure_pinned_sibling_box`).
///
/// The place is `SWP_NOZORDER`, which is what keeps the film under the pin that is standing over it:
/// the order is left exactly where it is, there is no raise here, and so there is nothing to put back
/// over anything (see `place_player_window`, and `ensure_pinned_sibling_box`, which carries a film
/// that travels with the band the same way for the same reason).
fn put_a_player_that_is_not_on_screen_back(x: i32, y: i32, width: i32, height: i32) {
    let hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
    if hwnd_val == 0 {
        return;
    }

    unsafe {
        let hwnd = HWND(hwnd_val as *mut std::ffi::c_void);
        if !IsWindow(hwnd).as_bool() || IsWindowVisible(hwnd).as_bool() {
            return;
        }

        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            x,
            y,
            width.max(1),
            height.max(1),
            SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW,
        );
    }
}

/// Whether the player's window is standing over the pin right now.
///
/// **The one thing about those two windows' order that this app cannot keep for itself.** Both are
/// topmost, so which of the two stands in front is which of them was raised last (see
/// `place_player_window`) — and the film is not raised by this app alone: mpv holds its own band
/// and stands itself back up on it (see `mpv_args`), which leaves the film over the pin, and with
/// it over the caption and the transport bar the pin draws into that band (see
/// `pin_overlay_chrome`). Everything this app holds off while a strip of that chrome is up (see
/// `player_order`) is what keeps *this app* from putting the film there; nothing keeps the player
/// from putting itself there.
///
/// The walk goes up from the pin and stops at the first window that is not topmost, which is where
/// the band both of them are in ends, so what it costs is the length of that band and not of the
/// desktop. A player's window that is not on screen covers nothing, which is the answer a parked
/// film owes (see `unpark_player_window`), and a band with no player's window in it at all — a
/// hover, or the media engine painting this app's own rows — is nobody standing over anything.
pub(super) unsafe fn player_stands_over_the_pin(pin: HWND) -> bool {
    let player = HWND(VIDEO_HWND.load(Ordering::SeqCst) as *mut std::ffi::c_void);
    if player.is_invalid() || !IsWindowVisible(player).as_bool() {
        return false;
    }

    let mut above = GetWindow(pin, GW_HWNDPREV).unwrap_or_default();

    while !above.is_invalid() {
        if above == player {
            return true;
        }

        // What stands over the pin stands over it from inside the topmost band, and the band
        // ends at the first window that is not of it: a window above that one is not over the
        // pin, whatever it is.
        if GetWindowLongPtrW(above, GWL_EXSTYLE) & WS_EX_TOPMOST.0 as isize == 0 {
            return false;
        }

        above = GetWindow(above, GW_HWNDPREV).unwrap_or_default();
    }

    false
}

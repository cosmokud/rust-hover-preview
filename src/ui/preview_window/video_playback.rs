//! Starting and stopping a film: the launch itself, the wait for its window and for its first
//! frame, and putting the player's window back in front.

use super::*;

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
) -> Option<Child> {
    let volume = volume.min(100);

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
        // file (see `start_audio_player`). A film nothing has measured is played as the file holds
        // it, and what measures it is a read on a thread of its own — the hover after this one is
        // the one that hears it (see `spawn_gain_scan`).
        let gain = if normalizing_video() {
            match audio_track::gain(path) {
                // A gain of one is a gain, not the absence of one: a soundtrack already standing
                // at the target is played through the filter like every other measured file.
                Some(gain) => Some(gain),
                None => {
                    spawn_gain_scan(path);
                    None
                }
            }
        } else {
            None
        };

        // Convert percentage to ffplay volume filter (0-100 maps to 0.0-1.0)
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
    // `probe_video_geometry`), and one lookup here answers both facts the chain needs.
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
    // `-sst` is inert on this build: the subtitle stream is demuxed and no subtitle filter is put
    // in the graph, so nothing is drawn whatever track is chosen. What draws is the `subtitles`
    // filter, which initialises libass — and which is appended *after* the crop rather than put in
    // front of it, so the lettering is laid over the cropped picture rather than cropped along
    // with it.
    //
    // The track named is the one the caller was given, where there is a choice. A caller that has
    // chosen nothing is answered with the file's own default rather than with nothing at all,
    // because the player is about to be launched with a specifier either way and the one it would
    // have picked for itself is the one the picture was drawn with last time (see
    // `SubtitleStreams::chosen`).
    let streams = video_subtitles(path);
    let sidecar = measured.and_then(|geometry| geometry.sidecar);
    let subtitles = video_launch::subtitle_filter(
        path,
        sidecar.as_deref(),
        streams.count,
        subtitle.or(streams.chosen()),
    );
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
        Some(second) => child.args(["-ss", &second]).spawn().ok(),
        None => child.spawn().ok(),
    };

    // After spawning, try to set WS_EX_NOACTIVATE on the ffplay window
    // to prevent it from stealing focus
    if let Some(ref child_process) = child {
        set_noactivate_for_process(child_process.id());

        // The player is this app's own child, and it is taken charge of the way the
        // engines are: one that is up when the app is killed, or crashes, is not left
        // playing with nothing to close it — the job ends it there and then, and the
        // record is what answers for the run that never got to end it. It is recorded
        // as the player rather than as an engine, because what ends a player is its
        // hover ending: a tier being let go of — a preview type switched off, a worker
        // given up on — is not its to receive; see `engine_processes`.
        engine_processes::record_player(VIDEO_PROCESS_IMAGE_NAME, child_process.id());

        // A player that was seeked has to be given its loop from out here, because it was not
        // given one at all — see `video_launch::loop_is_ours`. The clock it is measured against
        // starts here, at the second the player was actually begun at, and not at the moment the
        // process appeared: the header is read before the window is up, and a film of a few
        // hundred megabytes spends a good part of a second in there.
        //
        // The box goes in beside the clock because the rewind is a relaunch and a relaunch has to
        // be told where to put its window — the film is not drawn into a surface of this app's, so
        // the replacement's window has to be placed rather than laid out (see `video_loop_action`).
        if video_launch::loop_is_ours(seeked) {
            note_video_loop(
                child_process.id(),
                path,
                (x, y, x + width, y + height),
                start,
                video_duration(path),
            );
        } else {
            forget_video_loop(child_process.id());
        }
    }

    child
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

/// Put the ffplay window where it belongs and on top of everything, re-discovering and re-asserting
/// its style by pid so that a window ffplay recreated is styled again rather than inherited.
///
/// **This is the one place a player's window is raised, and the volume popup is guarded here rather
/// than by its callers.** That is the whole of the arrangement: what the popup floats over is the
/// media, and the media of a video FFmpeg plays *is* this window, so a raise while the popup is open
/// puts the film over the thing the user is adjusting its level with. The tick was guarding this,
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
    // For as long as a volume popup is open the player is left where it is, and the pin's window is
    // the one on top. Nothing is asked of the order on the way out — the next tick of the caller asks
    // for the player again once the popup is closed (see `pin_volume_open`, `toggle_pin_volume`).
    if pin_volume_open() {
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

        // Re-assert desired style bits in case ffplay modified them
        let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
        let new_style = current_style
            | WS_EX_NOACTIVATE.0 as isize
            | WS_EX_TOOLWINDOW.0 as isize
            | WS_EX_TOPMOST.0 as isize;
        if new_style != current_style {
            SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_style);
        }

        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
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

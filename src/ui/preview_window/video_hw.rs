//! What a film is decoded on and how it loops: the hardware device probe, the loop a pinned
//! film is given, and the route a file takes to a player rather than to a picture.

use super::*;

/// Whether a video is decoded on the graphics card at all: the `Video` toggle under
/// `Performance → Hardware Acceleration` in the tray, read as the switch it is.
///
/// It is the whole of what mpv is told about decoding — its `--hwdec=auto` is mpv's own safe list
/// with mpv's own software fallback, so the device the ffplay route has to measure is not a
/// question for it (see `video_launch::mpv_args`) — and it is read at the launch for the reason
/// below: a setting read at a hover's beginning is read for the *next* hover.
pub(super) fn video_hw_accel_wanted() -> bool {
    CONFIG
        .lock()
        .map(|config| config.video_hw_accel)
        .unwrap_or(DEFAULT_VIDEO_HW_ACCEL)
}

/// The `mpv.exe` a launch runs: the path `mpv_path` names where it is still a file, and the first
/// install this machine has otherwise — `None` only where there is neither.
///
/// `chosen` is the setting's own text, empty for "nobody has picked one", which means the same as
/// the first install. The setting is read at the launch for the reason above: a film already
/// playing was started with the path that stood then and is kept playing with it.
pub(super) fn mpv_program(chosen: &str, installs: &[PathBuf]) -> Option<PathBuf> {
    let chosen = (!chosen.trim().is_empty()).then_some(Path::new(chosen));

    codecs::resolve_mpv(chosen, installs)
}

/// The device a video is decoded on, or nothing for a video decoded on a core.
///
/// This is the `Video` toggle under `Performance → Hardware Acceleration` in the tray, read here
/// rather than in the tray, because a setting read at a hover's beginning is read for the *next*
/// hover: a film already playing was launched with the answer that stood then and is kept playing
/// with it rather than being relaunched under the user's feet.
///
/// It is **read, never waited for**: the probe runs on a thread of its own and this is on the
/// preview thread, which is the one thread that must not be waiting on an external process (see
/// `HwAccelProbe::read`). A preview begun before the answer has landed is software decoded, which is
/// the answer that works on every machine and the one a slow first hover should get anyway.
///
/// The answer is kept for as long as it answers the question that is being asked, which is the whole
/// of what `HwAccelProbe` is: the question is about this machine's drivers, about one build of
/// FFmpeg, and about *the setting as it stands*, and a tray toggle makes the third of those a
/// different question (see `forget_video_hw_accel_answer`).
pub(super) fn video_hw_accel_device() -> Option<&'static str> {
    VIDEO_HW_ACCEL.read(|| {
        let on = CONFIG
            .lock()
            .map(|config| config.video_hw_accel)
            .unwrap_or(DEFAULT_VIDEO_HW_ACCEL);

        probe_hwaccel_device(on)
    })
}

/// The one answer to "which device should a video be decoded on" for the run of the app.
///
/// It is a struct rather than a lone `OnceLock` because there are two parties with different rights
/// over it, and folding them into one is what made the first version of this answer `None` for ever.
/// A `OnceLock` completed by the thread that *reads* it is completed before the thread that
/// *writes* it has run, so the write is refused and every launch in the run is software decoded —
/// and the mistake is invisible to any test that reads the answer twice, because `None` is a
/// perfectly good answer for a machine where nothing survives the probe.
///
/// So the two rights are two fields. `asked` is the reader's to set and the writer's never to touch:
/// it says the question has been put, and a reader that finds it behind `current` is the one that
/// puts it — a compare-and-swap, so any number of readers racing this way still ask once. The answer
/// itself belongs to the probe alone, which is why it carries the generation it was found for: a
/// probe still running when the user flips the toggle would otherwise land an answer to a question
/// that is no longer being asked, and a film decoded on a card the user has just switched off is a
/// worse fault than one decoded in software.
pub(super) struct HwAccelProbe {
    /// The generation of the question a probe has been put for, and `0` for none — a reader that
    /// finds this behind `current` is the one that asks.
    pub(super) asked: AtomicU32,
    /// The generation of the question as it stands, which the tray moves on when the setting is
    /// switched. It starts at one so that zero can mean "no question has ever been asked".
    pub(super) current: AtomicU32,
    /// What a probe found, and which question it was found for. The generation is half the field for
    /// the same reason it is in `asked`: an answer is only an answer to the question it was asked,
    /// and a reader that cannot tell which question that was would have to take the newest on trust.
    pub(super) found: Mutex<Option<(u32, Option<&'static str>)>>,
}

impl HwAccelProbe {
    /// The answer as it stands, putting the question if it has never been put.
    ///
    /// **This never waits for the probe**, which is the whole bargain and the reason the answer is
    /// not behind a `OnceLock` the reader completes: the probe takes up to three seconds a device and
    /// walks four of them, and this is on the preview thread — the one thread in this app that must
    /// not be blocked on an external process, because it is the thread that has to keep answering
    /// Explorer. So a reader that arrives before the answer gets nothing, and nothing is a working
    /// answer: FFmpeg decodes in software and the film plays.
    ///
    /// The lock below is not a wait for anything slow. It is held by the probe for the length of one
    /// pointer-sized write, which is the same arrangement every other piece of this app's own state
    /// uses on this thread (`VIDEO_LOOP`, `PIN_STATE`).
    ///
    /// The record is taken by `&'static` rather than borrowed because the probe outlives this call by
    /// a thread, and a borrow of a caller's frame is not what a thread can write through. That is a
    /// constraint on where the record lives rather than a trick: it is a `static`, so every caller has
    /// one.
    pub(super) fn read(
        &'static self,
        probe: impl FnOnce() -> Option<&'static str> + Send + 'static,
    ) -> Option<&'static str> {
        let generation = self.current.load(Ordering::Acquire);

        // Asked once per generation however many readers arrive at once: the swap is what makes one
        // of them the asker, and the losers read rather than ask, which is the answer they would have
        // got had they waited a moment anyway.
        if self.asked.load(Ordering::Acquire) != generation
            && self.asked.swap(generation, Ordering::AcqRel) != generation
        {
            std::thread::spawn(move || {
                let found = probe();

                let Ok(mut answer) = self.found.lock() else {
                    return;
                };

                if self.current.load(Ordering::Acquire) == generation {
                    *answer = Some((generation, found));
                }
            });
        }

        self.found
            .lock()
            .ok()
            .filter(|answer| {
                answer
                    .as_ref()
                    .is_some_and(|(asked, _)| *asked == generation)
            })
            .and_then(|answer| answer.and_then(|(_, found)| found))
    }

    /// Make every answer found so far an answer to a question no longer being asked.
    ///
    /// This is what switching the setting does, and it is why the answer is not simply kept for the
    /// run of the app. The question is about the machine *and* about the setting: a probe run with
    /// the setting off names no device at all, so a run that kept its first answer would answer the
    /// switched question with the answer to the question before it.
    ///
    /// It is a bump and not a clear, because the answer may be written by a probe that is still
    /// running and which must be told which question it was asked rather than simply being ignored:
    /// `read` is what discards it, by comparing the generation it was found for against the current
    /// one. That is also why a launch made between the switch and the new answer gets nothing rather
    /// than the old device — software decoding is the answer that is right when the setting is off.
    pub(super) fn forget(&self) {
        self.current.fetch_add(1, Ordering::AcqRel);
    }
}

/// Where the answer to "which device should a video be decoded on" is kept while it is on its way.
pub(super) static VIDEO_HW_ACCEL: HwAccelProbe = HwAccelProbe {
    asked: AtomicU32::new(0),
    current: AtomicU32::new(1),
    found: Mutex::new(None),
};

/// Stop answering the hardware-acceleration question with what was found for the setting as it was
/// before the user switched it.
///
/// The tray calls this when the `Video` row is toggled, and it is the whole of what makes that row
/// work within a session rather than only across restarts. Without it the answer found for the
/// setting as it stood at the first hover would stand for the rest of the run, so the row would be a
/// question about the next run of the app rather than about the next preview — and the promise its
/// own note makes, that the next hover is the first one decoded differently, would be true only
/// after a restart the note never mentioned.
pub fn forget_video_hw_accel_answer() {
    VIDEO_HW_ACCEL.forget();
}

/// The first device of the ones worth trying that puts a frame on a screen without dying, and
/// nothing at all where none of them does.
///
/// `None` is the answer on a machine where every name in the list is fatal, which on the machine
/// this was written on is all of them — `ffplay` takes `-hwaccel` only alongside its Vulkan
/// renderer, and a renderer that cannot be brought up here dies of an access violation rather than
/// falling back (see `video_launch::HWACCEL_DEVICES` for the whole of the measurement). Falling
/// back is therefore this app's job and not FFmpeg's, and a software-decoded preview is a preview.
/// FFmpeg's own fallback is still behind the answer for a stream the device will not take, and the
/// cost of that is a slower film rather than a dead one.
pub(super) fn probe_hwaccel_device(enabled: bool) -> Option<&'static str> {
    video_launch::hw_accel_candidates(enabled)
        .iter()
        .copied()
        .find(|device| probe_one_hwaccel(device))
}

/// Whether one device decoded something and drew it without FFmpeg's renderer giving up on it.
///
/// The source is FFmpeg's own `testsrc` pattern rather than a file: it asks nothing of the disk and
/// answers the same on a machine with no video on it. It is the right limit because both of the
/// ways this option fails happen before a film is looked at — the renderer is brought up when the
/// option is parsed — so the answer is about the machine and the build rather than about the file.
///
/// A clean exit is the whole of the answer. Nothing about the picture is checked, because on a
/// machine where the device *is* usable the failures being ruled out here are not ones that would
/// show up in a tenth of a second of synthetic video: what they are is a process that is gone.
///
/// **A player that survives the budget is ended rather than left to finish.** It has drawn its tenth
/// of a second and not died, which is the answer being asked for, so there is nothing more to wait
/// for — but a `Child` dropped here is not a player ended: a `Drop` that closes a handle does not
/// kill a process, so a probe that timed out would leave a player on screen for ever, on a machine
/// that is precisely the slow one where nobody would have noticed it was meant to be a probe. It is
/// killed rather than waited on, because this runs on the probe thread and the app has no reason to
/// hold a thread open for a process it has already learned all it needed; the kill is confirmed by
/// image name and is not ours to confirm by waiting (see `engine_processes::terminate_verified`).
#[cfg(not(test))]
pub(super) fn probe_one_hwaccel(device: &str) -> bool {
    let Ok(mut child) = engine_processes::hidden_command("ffplay")
        .args([
            "-hwaccel",
            device,
            "-f",
            "lavfi",
            "-i",
            "testsrc=size=64x64:rate=1:duration=0.1",
            "-x",
            "16",
            "-y",
            "16",
            "-noborder",
            "-autoexit",
            "-loglevel",
            "warning",
        ])
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
    else {
        return false;
    };

    // Bounded, because a probe that can wait for ever is a probe that can hang a thread of the app
    // on something it cannot answer. The player is given longer to be alive than the tenth of a
    // second of video it is asked to draw, because the failure being caught is an exit rather than
    // a frame — and `try_wait` rather than `wait`, so the bound is this function's own and not a
    // thread's.
    let pid = child.id();
    let deadline = Instant::now() + Duration::from_millis(3000);
    let outcome = loop {
        match child.try_wait() {
            Ok(Some(status)) => break status.success(),
            Ok(None) if Instant::now() < deadline => {
                std::thread::sleep(Duration::from_millis(20));
            }
            // Still running with nothing left of the budget: it has drawn its tenth of a second and
            // not died, which is the answer the probe is asking for — the failures being ruled out
            // here are exits, not hangs.
            Ok(None) => break true,
            // A poll that cannot be read is not a poll that saw the player alive, and a device
            // named on an answer nobody read is a device that can be a fault on its own: this probe
            // exists to *rule devices in*, so an indeterminate state has to rule nothing out and
            // name nothing. It is not the exit either — nothing is known about the player — which is
            // why it is its own arm rather than folded into the one above.
            Err(_) => break false,
        }
    };

    // Only the player that outlived its budget has anything to end. One that exited was reaped by
    // `try_wait` above, and asking to be killed as well would be a kill aimed at an id the system is
    // free to hand to something else in the meantime.
    if child.try_wait().ok().flatten().is_none() {
        engine_processes::terminate_verified(pid, "ffplay.exe", None);
    }

    outcome
}

/// The test build's probe: begin nothing, and answer as a machine whose probe found no device does.
///
/// A test run must start no process (see `video_playback::spawn_player`), and naming no device is a
/// working answer on a machine where nothing survives (`probe_hwaccel_device` says so): the walk
/// stops at the same place, and the preview is decoded on a core.
#[cfg(test)]
pub(super) fn probe_one_hwaccel(_device: &str) -> bool {
    false
}

/// A player this app has to loop itself: which one it is, what it is playing, the second it was
/// begun at, how long the file is, and when it was begun.
///
/// It is here rather than in the pin's transport because it has to answer for a hover as much as
/// for a pin. A hover begins at zero and is given FFmpeg's own loop, so it is never in here; but
/// the transport is the pin's own record of a *pin*, and a pinned file that is seeked has to keep
/// looping after the walk that began it has long since been answered.
///
/// **The file and the box are in the record because the rewind is a relaunch.** A loop this app
/// gives cannot be given by posting a key — FFmpeg's player has no key that goes to the beginning,
/// every seek key it binds being a fixed increment (see `video_launch::rewind_launch_seconds`) — so
/// beginning the film again is the whole of the rewind, and a relaunch cannot be made out of a pid
/// and a clock. The length is here for the same reason it always was: the question "how near the end
/// is it" cannot be asked of a file nothing has measured, so a file with no length is a player with
/// nothing to be saved from.
pub(super) struct VideoLoop {
    /// The player the clock belongs to, so a relaunch leaves the old one's clock behind rather
    /// than rewinding a player that is no longer the one on screen.
    pub(super) pid: u32,
    /// The file, which is what the rewind begins again.
    pub(super) path: PathBuf,
    /// The box the player's window fills, which is where the player that replaces it is begun: the
    /// window is another program's, so it has to be told where to be rather than laid out by this
    /// app's own surface.
    pub(super) content: ScreenRegion,
    /// The second the player was begun at, which is where the first pass starts.
    pub(super) from: f64,
    /// How long the file is, or nothing for one nothing has measured.
    pub(super) duration: Option<f64>,
    /// When the pass began, from which the position is read.
    pub(super) at: Instant,
}

impl VideoLoop {
    /// Where in the file this pass has got to, off this app's own clock — which is the only position
    /// a player that reports nothing at all can be measured by (see `transport_clock`).
    pub(super) fn position(&self) -> f64 {
        self.from + self.at.elapsed().as_secs_f64()
    }
}

/// The one loop this app is giving, if one is being given.
///
/// A single slot rather than one per player because only one player is ever up: a relaunch ends the
/// player it replaces before the next one is begun (see `retire_replaced_player`), and a preview
/// that has been superseded is not a loop worth keeping.
pub(super) static VIDEO_LOOP: Mutex<Option<VideoLoop>> = Mutex::new(None);

/// Start the clock on the loop of a player that has to be given one.
pub(super) fn note_video_loop(
    pid: u32,
    path: &Path,
    content: ScreenRegion,
    from: f64,
    duration: Option<f64>,
) {
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = Some(VideoLoop {
            pid,
            path: path.to_path_buf(),
            content,
            from,
            duration,
            at: Instant::now(),
        });
    }
}

/// Forget the loop of a player that has to be given one only when it is not `pid`.
///
/// The guard is what keeps a relaunch's own bookkeeping from clearing the clock of the player that
/// replaced it: the outgoing player is retired after the incoming one is up, and an unconditional
/// clear on every end would leave the new player with no loop at all.
pub(super) fn forget_video_loop(pid: u32) {
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if slot.as_ref().is_some_and(|loop_| loop_.pid == pid) {
            *slot = None;
        }
    }
}

/// Move the loop's clock onto the moment the player's window came up, which is when the film began
/// playing rather than when its process was spawned.
///
/// The clock is started at the spawn because that is the one fact a launch has, and the film is a
/// good part of a second behind it: the file is opened, the first frame is decoded and the window is
/// made (measured: mpv 0.41.0 on this machine is 0.2–0.4s from the process to its first frame with
/// `--hwdec=auto`, and that is before there is any window to put up at all). A clock running ahead
/// of the picture reaches the end of the pass early, and the rewind is posted a tenth of a second
/// before the end *it* measures — so the film pays that startup time off its tail on every pass (see
/// `rewind_due`).
///
/// The pid guard is this file's usual one: a window coming up for a player that has already been
/// replaced is not a clock for the player that took its place.
pub(super) fn rebase_video_loop_clock(pid: u32) {
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if let Some(loop_) = slot.as_mut().filter(|loop_| loop_.pid == pid) {
            loop_.at = Instant::now();
        }
    }
}

/// Move the loop's clock onto a second the player was taken to while it ran, which is the same
/// correction as the window's rebase above with the position named as well as the moment.
///
/// A seek is what asks this, and it is asked on the player's side of the same question: mpv's
/// seek is a command rather than a relaunch (see `seek_pinned_playback`), so the clock a pass is
/// measured by is not begun at the second the player was taken to the way a relaunch's is (see
/// `note_video_loop`) — it goes on counting from where the pass began, and a seek taken while
/// that count is near the film's end is then undone by the very next tick, which reads the film
/// as having reached its end and wraps it back to the beginning. The second is named as well as
/// the moment because it is where the pass is now played from, and a clock left at the old
/// second would wrap the film early by the distance the seek moved it.
///
/// The pid guard is the same one every rebase keeps: a clock belonging to a player this app no
/// longer holds is left where it stands.
pub(super) fn rebase_video_loop_clock_to(pid: u32, seconds: f64) {
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if let Some(loop_) = slot.as_mut().filter(|loop_| loop_.pid == pid) {
            loop_.from = seconds;
            loop_.at = Instant::now();
        }
    }
}

/// Where the player behind `pid` has got to in the file it is playing, off the loop's own clock.
///
/// It is the one position a player that reports nothing can be read at (see `mpv_ipc`), and the one
/// a take-up needs: a pin taken up on a film the hover has been playing for a while is not a film at
/// its beginning, and a bar — and every end of file read off its clock — started at zero would
/// count the whole hover's stretch of the film again before the pass it is drawing was over.
///
/// Nothing for a player the record is not about, which is the answer a route that gives the loop to
/// its player rather than to this app keeps (see `video_launch::loop_is_ours`), and what makes this
/// safe to ask for any kind of player.
pub(super) fn video_loop_playhead(pid: u32, path: &Path) -> Option<f64> {
    VIDEO_LOOP.lock().ok().and_then(|slot| {
        slot.as_ref()
            .filter(|loop_| loop_.pid == pid && loop_.path == path)
            .map(|loop_| loop_.position())
    })
}

/// What a tick does about the loop of a player this app is looping itself.
pub(super) enum VideoLoopAction {
    /// The film is playing on and is nowhere near its end.
    CarryOn,
    /// The film is held, so the clock underneath the record is rebased onto the second it stands at
    /// and nothing is begun.
    Held,
    /// Begin the file again, from `seconds` of it, in `content` — the whole of a rewind.
    Rewind {
        path: PathBuf,
        content: ScreenRegion,
        seconds: f64,
    },
}

/// What one tick does about a loop, from the clock alone.
///
/// It is asked of the record rather than worked into the tick because the tick has to *do* the
/// relaunch and this has to be right about what to relaunch, and a decision with three answers
/// written inline beside a process being spawned is a decision nothing can be asked about.
///
/// The position is read off the clock rather than off the pin's transport, so that a file with no
/// pin in front of it is measured the same way — and it is read against the *pid on screen*, so a
/// clock left behind by a player that has been replaced cannot rewind the player that replaced it.
///
/// The hold is the other refusal, and it is what a gesture over the picture does: a drag holds the
/// player with a pause key (see `video_drag_hold_apply`), so a film a second from its end would
/// otherwise be rewound *underneath the hand holding it*, and the second the film was held at would
/// keep counting up in the clock underneath this record — so that a film held for a minute near its
/// end was past its end the moment it was let go of, and was rewound on the very next tick. Rebasing
/// is what a release does to the transport's own clock, and it is the same arithmetic for the same
/// reason (see `PinTransport::released`).
///
/// **And the film's own end of file is asked here rather than remembered with the record, which is the
/// whole of what makes a pinned loop reach a player the pin did not begin.** A record is made where
/// a player is begun, and mpv is the one player a take-up does not begin again — so a film the
/// hover played once on carried the *hover's* answer into the pin: with the hover's loop off and
/// `End of File → Loop` on there was nothing to rewind at all, and with the hover's loop on and
/// the pin's end of file a step on the rewind undid a step the pin had asked for. `loops` is that question,
/// asked of the setting the film on screen is played under (see `pin_video_loops`).
pub(super) fn video_loop_action(loop_: &VideoLoop, playing: bool, loops: bool) -> VideoLoopAction {
    if !playing {
        return VideoLoopAction::Held;
    }

    if !loops {
        return VideoLoopAction::CarryOn;
    }

    if !video_launch::rewind_due(true, loop_.duration, loop_.position()) {
        return VideoLoopAction::CarryOn;
    }

    VideoLoopAction::Rewind {
        path: loop_.path.clone(),
        content: loop_.content,
        seconds: video_launch::rewind_launch_seconds(),
    }
}

/// What the tick does about the loop of a player this app is looping itself, answering whether it
/// began the file again.
///
/// The clock is moved on before the key is posted rather than after, so that a tick which runs
/// again inside the same margin — a slow tick, a loop of this app's own — does not post the same
/// rewind twice and cut a second off the film it just began.
///
/// The record is *taken out* rather than updated, and that is what lets the rewind be a relaunch:
/// `restart_pinned_player` begins a player, and beginning one takes this very lock (through
/// `forget_video_loop`), so a record updated in place under a held lock would deadlock on the first
/// pass. Taking it out also drops the clock of the player being replaced, which is right in its own
/// right — a clock whose player is gone is a clock with nothing to measure.
///
/// `loops` is the film's own end of file, read by the caller, which is the one that knows what is on
/// screen: this tick is asked for a pinned film as much as for a hovered one, and a pinned film's
/// loop is the one road mpv has (see `video_loop_action`).
pub(super) fn video_loop_tick(playing: bool, loops: bool) -> bool {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let Ok(mut slot) = VIDEO_LOOP.lock() else {
        return false;
    };
    let Some(loop_) = slot.as_mut().filter(|loop_| loop_.pid == pid && pid != 0) else {
        return false;
    };

    match video_loop_action(loop_, playing, loops) {
        VideoLoopAction::CarryOn => false,
        VideoLoopAction::Held => {
            // Rebased onto the second the film is standing at, which is the same correction the
            // transport's own clock gets when the hand lets go of a held film.
            loop_.from = loop_.position();
            loop_.at = Instant::now();
            false
        }
        VideoLoopAction::Rewind {
            path,
            content,
            seconds,
        } => {
            // **mpv takes the wrap as a command rather than as a relaunch.** The film goes back to
            // the second the rewind names down the pipe the player already listens on, so the
            // process is the one that has been playing all along and no window comes down for it
            // (see `mpv_ipc`). The clock is rebased onto that second rather than taken out, so the
            // pass that follows is measured from where it actually began.
            if video_route(&path) == VideoRoute::Mpv
                && mpv_ipc::seek_absolute(loop_.pid, seconds)
            {
                // **And it is set going in the same breath, which is the half a seek cannot do.**
                // A film that has reached its end is *paused* there (`--keep-open=yes`, see
                // `mpv_args`), and a seek is a position for that player rather than a new pass: mpv
                // keeps the pause, and says so in its own log — `playback restart complete @
                // 0.000000, audio=playing, video=playing (paused)` (measured on mpv 0.41.0 down
                // this app's own pipe). So a wrap that lands a moment after the pass has ended is a
                // film standing at its first frame, silent and going nowhere, until the user presses
                // pause and play at it. The tick only rewinds a film that is playing (see
                // `video_loop_action`), so playing is what it is owed back.
                mpv_ipc::set_paused(loop_.pid, false);
                loop_.from = seconds;
                loop_.at = Instant::now();
                // **And the pin's own bar is moved onto that second with the clock.** A wrap is
                // not a relaunch on this road — no player is begun again, so `transport.begun`
                // never runs and nothing else writes the second the pass now runs from — and a bar
                // left counting from the launch it was taken up at stands a whole pass away from
                // the picture after the first wrap (see `rebase_video_loop_clock_to`).
                if pinned_path().as_deref() == Some(path.as_path()) {
                    update_pin_transport(|transport| transport.released(seconds));
                }
                return true;
            }

            *slot = None;
            drop(slot);

            // The film is played on from the beginning rather than sought back to it, because a
            // seek cannot say "zero" on this player — every key it binds is a fixed step, and a
            // step on a film longer than the step walks the film backwards instead of returning it
            // to the start (see `video_launch::rewind_launch_seconds`). `holding` is false because
            // the tick only rewinds a film that is playing.
            restart_pinned_player(&path, content, seconds, false);
            true
        }
    }
}

/// Who plays a video on this machine, which is one of four answers rather than a set of
/// preferences: the media engine Windows has, FFmpeg's `ffplay`, mpv, or nothing at all.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum VideoRoute {
    /// The media engine Windows has, drawing into this app's own window.
    MediaEngine,
    /// FFmpeg's `ffplay`, in a window of its own.
    Ffplay,
    /// mpv, in a window of its own: mpv's own window filling the preview box, the shape
    /// `Ffplay` gives ffplay's.
    Mpv,
    /// No player: nothing installed here will take the file, so there is no preview of it. This
    /// is the answer `[ffmpeg]` has always meant on a machine without FFmpeg, and it is an answer
    /// rather than a failure — a file nothing plays is a file with no preview, not a broken one.
    NoPreview,
}

/// The engines a video may be played by, most preferred first: the order `Best` walks and the
/// order an explicit choice falls through. mpv is at the top — `Best` prefers it wherever the
/// machine has it — and the two under it are the pair the size rule chooses between, which is
/// why `Hybrid` names neither of them through this list (see `VideoEngine`).
pub(super) const VIDEO_ENGINES: [VideoEngine; 3] =
    [VideoEngine::Mpv, VideoEngine::Ffmpeg, VideoEngine::Native];

/// The name-only half of whether the media engine can play `path`: the file's name is one the
/// `[video]` list carries (see `video_formats::claims_video_name`).
fn named_in_the_video_list(path: &Path) -> bool {
    let Ok(config) = CONFIG.lock() else {
        return false;
    };
    let extensions = config.video_extensions.clone();
    drop(config);
    video_formats::claims_video_name(path, &extensions)
}

/// Whether the media engine Windows has will play `path`: its name is the engine's to ask about
/// and the engine opens it (see `video_player::plays`). It is `pub(super)` because the probe asks
/// it too — a native choice runs FFmpeg's passes only for a film the engine will not take, which is
/// the one it hands on (see `video_probe::probe_video_geometry`).
pub(super) fn named_for_the_media_engine(path: &Path) -> bool {
    named_in_the_video_list(path) && video_player::plays(path)
}

/// Which engine plays a video, from the configured choice, the fallback switch, what this machine
/// has and how big the file is. The closures are the machine's answers (`installed`) and the
/// file's (`can_play`), and `pixels` is the file's resolution where the probe read one — so the
/// whole rule is one function the tests can hand stub answers to.
///
/// `Best` prefers mpv wherever the machine has it, whatever the file weighs and whether it could be
/// measured at all; a machine without it answers the file's own rule — the media engine for a film
/// at or below `VIDEO_FFMPEG_ABOVE_PIXELS`, and FFmpeg's player for a bigger one or for a film
/// nobody measured, where the player more forgiving than the probe is a better answer than a box
/// nothing would be drawn into (see `probe_video_geometry`). `Hybrid` is that size rule stated
/// plainly and nothing else, so mpv is not part of it. A chosen engine is itself alone with the
/// fallback off, and with it on is followed by the others it may fall through to: the pair it always
/// walked, so an explicit `Native` or `Ffmpeg` lands on the other of the two and never on mpv,
/// while a chosen `Mpv` falls through to both. A choice the machine cannot supply at all is ignored
/// as `Best`.
pub(super) fn resolve_video_engine(
    choice: VideoEngine,
    fallback: bool,
    pixels: Option<u64>,
    installed: impl Fn(VideoEngine) -> bool,
    can_play: impl Fn(VideoEngine) -> bool,
) -> Option<VideoEngine> {
    // A choice the machine cannot supply is ignored as if it were `Best`: a row that is greyed
    // names a player that is not here, and there is nothing to prefer about it.
    let choice = if !installed(choice) {
        VideoEngine::Best
    } else {
        choice
    };

    let mut candidates: Vec<VideoEngine> = match choice {
        // `Best` prefers mpv outright where the machine has it: every film, whatever it weighs and
        // whether it could be measured at all. A machine without it has nothing to prefer about
        // it, and the rule below answers as it always has.
        VideoEngine::Best if installed(VideoEngine::Mpv) => VIDEO_ENGINES.to_vec(),
        VideoEngine::Best | VideoEngine::Hybrid => {
            // A film nobody measured is handed over where FFmpeg's player is here: `ffprobe` could
            // not read it, and a player more forgiving than the probe is a better answer than a box
            // nothing would be drawn into (see `probe_video_geometry`). The order's own filter
            // leaves the media engine where there is no player to hand it to. It is a rule between
            // the media engine and FFmpeg's player, so mpv is not in it even where the machine has
            // one.
            let handed_over = pixels.map_or(true, |pixels| pixels > VIDEO_FFMPEG_ABOVE_PIXELS);
            let prefer = if handed_over {
                VideoEngine::Ffmpeg
            } else {
                VideoEngine::Native
            };
            let mut list = vec![prefer];
            list.extend(
                VIDEO_ENGINES
                    .iter()
                    .copied()
                    .filter(|engine| *engine != prefer && *engine != VideoEngine::Mpv),
            );
            list
        }
        chosen if fallback => {
            let mut list = vec![chosen];
            // An explicit choice keeps the meaning it had before mpv existed: its fall-through is
            // the pair it always walked, so an explicit `Native` or `Ffmpeg` lands on the other of
            // the two and never on mpv. A chosen `Mpv` is the exception — it falls through to both
            // the others, the way `Best` reaches past it.
            list.extend(VIDEO_ENGINES.iter().copied().filter(|engine| {
                *engine != chosen && (chosen == VideoEngine::Mpv || *engine != VideoEngine::Mpv)
            }));
            list
        }
        chosen => vec![chosen],
    };

    candidates.retain(|engine| installed(*engine));
    candidates.into_iter().find(|engine| can_play(*engine))
}

/// Which engine plays this file, on this machine, with these lists.
///
/// The routing rule is `resolve_video_engine`, and what stands around it here is only what the
/// machine and the file are: the choice and the fallback switch off the configuration, the
/// installs the machine has, and the two questions about the file — how big it is, for the hybrid,
/// and whether the media engine will take it. The name is asked before the engine is (see
/// `named_in_the_video_list`), and the engine is asked only where the name is one of its own,
/// which is what keeps a film only FFmpeg's list carries from being opened by an engine that has
/// nothing to do with it.
pub(super) fn video_route(path: &Path) -> VideoRoute {
    let (choice, fallback) = engine_choice();

    // The two players are asked through the pair a test can stand in or out (`ffplay_is_here`,
    // `mpv_is_here`) rather than through `VideoEngine::installed`, so that the routes that differ
    // only by which player is on the machine can be asked about at all: mpv is not installed on
    // the machine this is written on, and neither arrangement it makes is reachable without one
    // (see `stand_mpv_in`). The answers are the machine's own in a run.
    let here = |engine: VideoEngine| match engine {
        VideoEngine::Ffmpeg => ffplay_is_here(),
        VideoEngine::Mpv => mpv_is_here(),
        VideoEngine::Best | VideoEngine::Native | VideoEngine::Hybrid => engine.installed(),
    };

    match resolve_video_engine(
        choice,
        fallback,
        cached_video_source_pixels(path),
        here,
        |engine| match engine {
            VideoEngine::Ffmpeg | VideoEngine::Mpv => here(engine),
            VideoEngine::Native => named_for_the_media_engine(path),
            VideoEngine::Best | VideoEngine::Hybrid => false,
        },
    ) {
        Some(VideoEngine::Ffmpeg) => VideoRoute::Ffplay,
        Some(VideoEngine::Mpv) => VideoRoute::Mpv,
        Some(VideoEngine::Native) => VideoRoute::MediaEngine,
        _ => VideoRoute::NoPreview,
    }
}

/// The configured engine choice a route is read off: the setting's own in a run, and whatever a
/// test stands in its place under one.
///
/// It is asked through here rather than inline for the same reason the machine's answers are
/// (`ffplay_is_here`, `mpv_is_here`): `video_route` falls an engine the machine has not got back
/// to the machine's choice, and a test that stands the installs in still inherits whichever
/// engine that run's `config.ini` names — an explicit `Native`, `Ffmpeg` or `Hybrid` walks a
/// fall-through that never reaches mpv, so an mpv arm standing its installs in would answer a
/// route with the config file's rather than the road's (see `stand_video_engine_in`).
fn engine_choice() -> (VideoEngine, bool) {
    #[cfg(test)]
    if let Some(stood) = video_engine_choice_stand_in() {
        return stood;
    }

    CONFIG
        .lock()
        .map(|c| (c.video_engine, c.video_engine_fallback))
        .unwrap_or((DEFAULT_VIDEO_ENGINE, DEFAULT_VIDEO_ENGINE_FALLBACK))
}

/// Stand a configured engine choice in place of the file's, or give the file's answer back with
/// `None` — one slot for the whole process, so it is put back by the tests that set it rather
/// than left to whichever test the runner reaches next, the same rule the install slots keep
/// (see `stand_mpv_in`, `stand_ffplay_in`).
#[cfg(test)]
pub(super) fn stand_video_engine_in(choice: Option<(VideoEngine, bool)>) {
    if let Ok(mut held) = VIDEO_ENGINE_STAND_IN.lock() {
        *held = choice;
    }
}

/// The choice [`video_route`] reads while a test stands one in (see `stand_video_engine_in`).
#[cfg(test)]
fn video_engine_choice_stand_in() -> Option<(VideoEngine, bool)> {
    VIDEO_ENGINE_STAND_IN.lock().ok().and_then(|held| *held)
}

/// The choice [`video_route`] reads while a test stands one in (see `stand_video_engine_in`).
#[cfg(test)]
static VIDEO_ENGINE_STAND_IN: Mutex<Option<(VideoEngine, bool)>> = Mutex::new(None);

/// Whether the media engine Windows has plays this file, which is one of the players a layout, a
/// load and a pin each ask about before they do anything else for a video.
///
/// It is one answer and not a chain, and everything about a video follows from it — whether the
/// frames are drawn by this app or by a player, whether a pin of one is resized and maximized or
/// only moved, and whether its transport bar is a control or a read-out (see `pin_frame` and
/// `pin_transport_kind`). It is also where the engine is asked about a file at all: where the
/// route settles on FFmpeg's player — a film above the hybrid's threshold, or every film when the
/// choice names FFmpeg's player — or on mpv, which the install alone settles, nothing opens the
/// file, so `video_player::plays` and its per-file memo are reached only for a file the route gives
/// the media engine.
///
/// The write-back is what makes the question affordable where it is asked: the probe opens the
/// file and builds a decoder chain for it, so the answer is held, and a hover that asks twice —
/// the layout, the load, a pin — is a lookup after the first (see `video_player::plays`). The
/// geometry probe opens the file as well, so this is asked beside it on the thread that exists
/// for keeping both off the one that draws (see `spawn_video_probe`).
pub(super) fn media_engine_plays(path: &Path) -> bool {
    video_route(path) == VideoRoute::MediaEngine
}

/// The nominal width of the box a page of HTML is drawn in.
pub(super) const HTML_PAGE_WIDTH: u32 = 1280;

/// The nominal height of the box a page of HTML is drawn in.
pub(super) const HTML_PAGE_HEIGHT: u32 = 800;

/// The box a page of HTML is drawn in, which is this app's own: a page asks for no size of
/// its own, and what the layout gives it is this box scaled by `document_scale` — so a
/// `Fit to Screen` is the room the display has and a share of it is that share of the room
/// (see `page_is_painted`, which keeps a page that is text off this box).
///
/// Nothing is read to answer it, which is what a specimen's box is not: a page is laid out by
/// the browser at the size it is given, so there is no size in the file to measure.
pub(super) fn html_page_box() -> (u32, u32) {
    (HTML_PAGE_WIDTH, HTML_PAGE_HEIGHT)
}

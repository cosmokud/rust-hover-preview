//! Handing a film over: retiring the player that is up while its replacement is launched, and
//! what the end of a pinned film means.

use super::*;

/// A player that has been begun to take the place of another, and the player being replaced.
///
/// It exists because a relaunch is two things at once and the order they happen in is what the
/// user sees. Ending the old player first and beginning a new one leaves a window with nothing
/// in it for as long as the new player takes to open the file, seek, and put a window up —
/// which on a 1440p HEVC file is long enough to read as a flash of the desktop through the pin.
/// Beginning the new one first means its window arrives over the old one and the old one is ended
/// only once there is something to replace it with, so the picture on screen is continuous and
/// the swap is invisible.
///
/// So the old player is not ended where it is replaced; it is parked here, with the player that
/// has taken its place and the moment it was begun, and ended by the loop once the arrival is
/// answered (see `settle_video_retirement`). What the loop asks is a question about two
/// processes, and it is asked of this rather than of the state so that it can be answered with
/// nothing up.
///
/// Only one player is ever parked this way, which is why the record holds one and not a list: a
/// relaunch that arrives while another is parked ends the player in between rather than stacking
/// a third wait on top of the first (see `retire_replaced_player`).
///
/// What the overlap costs is *sound*, and it is a price paid on purpose rather than an oversight:
/// the player on screen is a whole player, not a window, and nothing can take the sound out of it
/// while it is there. Its own window cannot even be addressed — `video_window_for` answers out of
/// the one published handle, which through the overlap names whichever window the monitor found
/// last, so a key posted to the parked player would reach it only on the ticks where its own window
/// happens to be that one, and this app could not tell whether it arrived (see
/// `transport_playing`). The other trade is worse: a replacement begun silent cannot be un-silenced
/// afterwards, because the level of a running player is a key this app has no way to confirm and a
/// mute is a toggle rather than a setting, so a film begun at zero stays at zero for good.
///
/// So two audio streams run at once for as long as the swap takes — bounded by the same
/// `VIDEO_START_WAIT_SECS` the picture's hold is bounded by, and in practice a few hundred
/// milliseconds, a player having to open the file, seek and put its window up — and the overlap ends
/// the moment the wait does, whether that is the replacement arriving or the pin being gone.
#[derive(Clone, Copy)]
pub(super) struct VideoRetirement {
    /// The player being replaced, verified still to be `ffplay.exe` before it is ended however
    /// long ago it was begun — a handle this old is a process id that has since been reused
    /// (see `terminate_ffplay_pid`).
    pub(super) retiring: u32,
    /// The player begun in its place, whose window arriving is what ends it.
    pub(super) replacement: u32,
    /// When that player was begun, which bounds the wait exactly as `VIDEO_START_WAIT_SECS`
    /// bounds a first start: a player that is never going to put a window up must not leave the
    /// one it was replacing playing over the desktop for ever.
    pub(super) started: Instant,
}

/// The player being replaced by a relaunch, if one is waiting to be ended.
///
/// It is held behind the loop's own state rather than in the media, for the reason the media is
/// never read for a question like this one: every thread here may ask, and the loop is the one
/// that has to be the only one ending a player (see `settle_video_retirement`).
pub(super) static VIDEO_RETIREMENT: Lazy<Mutex<Option<VideoRetirement>>> =
    Lazy::new(|| Mutex::new(None));

/// Park a player that a relaunch has taken the place of, so that it is ended once its
/// replacement has arrived rather than before it was begun.
///
/// It is a no-op for a run with nothing playing and for a relaunch of a player that was not
/// running — a hold that is being let go of has no window of its own to keep on screen, so
/// there is nothing to overlap and the arriving player is all there is.
///
/// A relaunch that arrives while a player is already parked is answered by
/// `retirement_after_relaunch`, and the lock is held across both: the record is one slot, so a
/// relaunch and the loop's settling of that slot cannot be allowed to interleave, or the relaunch
/// parks a player the settling then never looks at (see `settle_video_retirement`).
pub(super) fn retire_replaced_player(retiring: u32, replacement: u32) {
    if retiring == 0 || replacement == 0 || retiring == replacement {
        return;
    }

    let Ok(mut held) = VIDEO_RETIREMENT.lock() else {
        return;
    };

    let (parked, ending) = retirement_after_relaunch(
        *held,
        retiring,
        replacement,
        video_window_for(retiring).is_some(),
        Instant::now(),
    );
    if let Some(ending) = ending {
        // Ended the same way a settled retirement is ended, and for the same reason: this player
        // has left the media and `VIDEO_PID` already names its replacement, so this call is the
        // only one that will ever end it or stop holding it (see `end_retired_player`).
        end_retired_player(ending);
    }
    *held = parked;
}

/// The wait that is left when one player is retired while a replacement for another may already be
/// running — and the player, if any, that has to be ended at once for that wait to be the right one.
///
/// The record holds one player, so a relaunch that arrives while a wait is running has to say what
/// became of the player it displaced, and there is only ever one question to answer it with: whose
/// window is on screen. Nothing else distinguishes the two cases, because in both of them a
/// replacement is beginning over a player that is still on screen.
///
/// * The player being retired has a window, so it covers the one already parked, and that one is
///   ended at once — it has nothing left to be kept on screen for.
/// * It has no window, so the player already parked is still the only picture there is. Ending
///   *that* one would leave the band empty for the whole of the new start, which is the flash of
///   the desktop this record exists to prevent; the player in between has put no window up to
///   lose, so it is the one that goes, and what stays parked is the oldest player with the newest
///   player waited for — one wait, still bounded by the oldest player's own start, rather than a
///   second wait stacked on a player nothing will ever look at again.
///
/// `now` is the moment this relaunch began its replacement, which is only used when there is no
/// wait already running: a wait that is already running is bounded by the wait that began with the
/// first of these players, and re-basing it on this one would let it outlive the start it stands
/// for by however many relaunches arrived inside it.
pub(super) fn retirement_after_relaunch(
    previous: Option<VideoRetirement>,
    retiring: u32,
    replacement: u32,
    retiring_up: bool,
    now: Instant,
) -> (Option<VideoRetirement>, Option<u32>) {
    let Some(previous) = previous else {
        return (
            Some(VideoRetirement {
                retiring,
                replacement,
                started: now,
            }),
            None,
        );
    };

    // The same player parked twice over is the case the record was built for: the wait is the
    // older player's and the newer player is what is being waited for, and nothing is ended —
    // the player is the one on screen, which is the whole of what it was parked for.
    if previous.retiring == retiring {
        return (
            Some(VideoRetirement {
                retiring,
                replacement,
                started: previous.started,
            }),
            None,
        );
    }

    if retiring_up {
        return (
            Some(VideoRetirement {
                retiring,
                replacement,
                started: previous.started,
            }),
            Some(previous.retiring),
        );
    }

    (
        Some(VideoRetirement {
            retiring: previous.retiring,
            replacement,
            started: previous.started,
        }),
        Some(retiring),
    )
}

/// Whether the player standing in the band is a different process from the one a park captured its
/// frame for — which is the whole of what "the replacement has not put a frame up yet" can be asked
/// of anything this app knows.
///
/// **It is asked of the pids and not of the window, because a window says nothing about whether the
/// process behind it has decoded anything.** A player that has just been begun has a window of its
/// own within a few milliseconds — SDL makes it before it opens the file — and that window is
/// visible, correctly sized and empty: which is exactly the state a park that hands the band back
/// on visibility alone leaves the desktop behind for as long as the decode takes (see
/// `park_swap_arm`). A pid is the one fact that tells a player carrying the picture the park
/// captured from a player that has not opened the file yet, and it is a fact this app wrote down
/// itself rather than read out of a process of somebody else's.
///
/// A pid of zero on either side is not a replacement but a player that is not there: `VIDEO_PID` is
/// cleared by every path that ends a player, so a cleared one says the pin has nothing playing
/// rather than that something new has taken its place — and a park whose player has gone has no
/// replacement to wait for (see `PIN_PARK_SWAP_TIMEOUT`).
pub(super) fn player_replaced_since(parked: u32) -> bool {
    let now = VIDEO_PID.load(Ordering::Acquire);
    parked != 0 && now != 0 && parked != now
}

/// Whether a player being replaced can be ended now, or is still the one on screen.
///
/// Everything in this is about not ending the only picture there is: a replacement with a
/// window of its own is a replacement that has arrived; a replacement that is *gone* never will
/// be, and waiting longer leaves the old player playing over a file nothing is going to show;
/// and a replacement that has been starting for longer than a start ever takes is a player this
/// app has already given up on once today (`player_wait`), so the wait is bounded by the same
/// number and ends the old one rather than outliving it.
pub(super) fn retire_ready(
    replacement_up: bool,
    replacement_alive: bool,
    waited: Duration,
) -> bool {
    replacement_up || !replacement_alive || waited >= Duration::from_secs(VIDEO_START_WAIT_SECS)
}

/// What is left of a retired player once it has been asked to end.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum RetireEnd {
    /// Confirmed gone. The record of it goes with it, and there is nothing left to wait over.
    Gone,
    /// Still there. The request has not taken, so the player stays parked and is asked again.
    Asked,
}

/// Whether a retired player is gone or still there, from whether it was alive to be asked.
///
/// It is the same two answers `kill_stray_video_process` gives, and for the same reason: ending a
/// player is a *request* (see `terminate_ffplay_pid`), and this is the one place that asks for one
/// this app has already stopped holding — the `Child` went when the player was parked (see
/// `restart_pinned_player`) and `VIDEO_PID` names the replacement by now, so nothing else in the
/// preview would ever ask it again. Which is what makes the second answer mean anything: an
/// unconfirmed request keeps the player parked rather than dropping it, and a confirmed one stops
/// holding an id for a process that is gone.
pub(super) fn retire_end(alive: bool) -> RetireEnd {
    if alive {
        RetireEnd::Asked
    } else {
        RetireEnd::Gone
    }
}

/// End a retired player and say what is left of it, giving up its process id only once it is
/// confirmed gone.
///
/// It is the one place a player this app has stopped holding is ended, so it is also the one place
/// that has to stop holding it: `record_player` writes every player down for the next run to answer
/// for (see `engine_processes`), and a retired player is the only one that leaves without a
/// confirmation of its own — `is_video_process_running` reads the handle in the media and
/// `kill_stray_video_process` reads `VIDEO_PID`, and both name the replacement by the time a
/// retirement is settled. Without this, every seek, resize and level change that ends in a relaunch
/// leaves one dead pid behind in the record and in the state file it is written into, for the rest
/// of the run.
pub(super) fn end_retired_player(pid: u32) -> RetireEnd {
    terminate_ffplay_pid(pid);
    let end = retire_end(is_ffplay_pid_alive(pid));
    if end == RetireEnd::Gone {
        // The player is confirmed gone, so the record of it goes with it rather than being left
        // for the next run to look for.
        engine_processes::forget(pid);
    }
    end
}

/// A player that a relaunch has taken the place of, ended now that there is a window to replace
/// it with, and a player that died being settled in the bar that was drawn against it.
///
/// It answers out of the loop rather than out of whichever thread happened to begin the
/// replacement, because ending a player is process discipline and this is the one place in the
/// preview that owns it: a seek is asked for from the pinned window's own procedure and a
/// resize from the loop's tick, and both arrive here rather than each taking a player's life
/// into its own hands.
///
/// The wait it waits on is the same wait the first launch waits on (`VideoStart`, bounded by
/// `VIDEO_START_WAIT_SECS`), which is what makes a relaunch no more visible than the start it
/// replaces: on both sides of this is one hold — the spinner on a hover, the player's own still-
/// on-screen frame on a pin — for as long as a player takes to put its window up.
///
/// The lock is held across the whole of it, and the record is read under it rather than taken out
/// of it. That is what closes the race between here and a relaunch: a retirement taken out and put
/// back afterwards is a gap in which a relaunch on the pin's own window thread parks a player of
/// its own, and the put-back then overwrites it — a player left alive and unowned, its window on
/// screen for as long as the app runs. Reading it under the lock leaves nothing to interleave
/// with, and every path here is a handful of non-blocking Windows calls, so the wait a relaunch
/// asks on this thread is a wait measured in microseconds.
pub(super) fn settle_video_retirement() {
    let Ok(mut held) = VIDEO_RETIREMENT.lock() else {
        return;
    };
    let Some(pending) = *held else {
        return;
    };

    // The handle the arriving player published is what says it has arrived, and it has to be the
    // handle that belongs to *that* player rather than to any window on screen: the two overlap,
    // the old window is on screen throughout, and `VIDEO_HWND` names whichever of them the
    // monitor found last (see `video_window_for`).
    let arrival = video_window_for(pending.replacement).is_some();

    // A pin that is no longer up is the other end of the wait, and it ends it at once. What this
    // player is being kept on screen *for* is a window that has gone — a pin closed, a walk that
    // stepped off, a hover that took the file over — and a replacement the pin is not waiting for
    // any more is one whose arrival will never be answered by a window of this pin's, so waiting
    // the full `VIDEO_START_WAIT_SECS` for it would leave a stray window on the desktop for ten
    // seconds after the pin it belonged to had closed.
    if !pinned()
        || retire_ready(
            arrival,
            is_ffplay_pid_alive(pending.replacement),
            pending.started.elapsed(),
        )
    {
        // A player whose end was not confirmed keeps its place in the record and is asked again on
        // the next tick, because nothing else here would ask it: the wait is over either way, and
        // what is left of it is only the pid (see `end_retired_player`).
        if end_retired_player(pending.retiring) == RetireEnd::Gone {
            *held = None;
        }
    }
}

/// Settle a transport bar against the player actually behind it.
///
/// Everything a bar says about a video FFmpeg plays is something this app asserted when it began
/// a player or sent it a key, and an assertion is not a reading: a player that has ended — a film
/// watched to its end, a window the user closed from the taskbar, a process killed from outside —
/// leaves every claim behind standing. So the loop, which already asks whether the player is alive
/// to decide whether the *pin* has come apart (see `pin_media_is_alive`), asks it once more here
/// to decide what the *bar* is allowed to say, and overwrites the claim where the two disagree.
///
/// This is why the pause glyph cannot get stuck. A file held and then lost is not a file that is
/// playing, and `PinTransport::player_gone` writes it as a held file rather than as a running one:
/// the second the film had reached is kept, so the bar goes on showing where it stopped and a
/// press starts it from there — but the button stops claiming there is something playing to pause.
pub(super) fn settle_pinned_transport() {
    if current_media_type() != Some(MediaType::Video) || !pinned() {
        return;
    }

    if is_video_process_running() {
        return;
    }

    let pinned = pin_state();
    let pin = pinned.as_ref().and_then(|pinned| pinned.pin());
    let played = pin.and_then(|pin| pin_playhead(&pin.transport).or(pin.transport.paused_at));
    let Some(played) = played else {
        // Nothing was ever claimed and nothing is running: a bar with nothing behind it, which is
        // the state a pin that is only just starting is in and is not a reconciliation to make.
        return;
    };

    update_pin_transport(|transport| transport.player_gone(played));
}

/// How near its end a film has to be before this app reads a player that has gone as having been
/// watched to the end rather than as having been killed.
///
/// The clock underneath `PinTransport::started` is this app's own, counting from the moment it
/// began a player, and it is a *wall* clock: a 2560x1440, 144 fps HEVC file on a machine that
/// cannot decode it in real time plays back behind the clock, because `-framedrop` drops the
/// frames it has no time for rather than waiting. So the two clocks drift, by seconds over a long
/// film, and a decision made on an exact comparison would either begin the film again seconds
/// before it finished — cutting the last scene off — or, on a machine that is behind rather than
/// ahead, never recognise the end at all and leave a pin showing a dead player.
///
/// Five seconds is the width of that drift as measured rather than a round number picked for
/// tidiness, and it is deliberately lopsided. A film restarted up to five seconds early costs the
/// viewer the tail of one loop out of however many the preview runs for; a film restarted never
/// costs a frozen pin and a play button over a picture that is not moving, which is the failure
/// this number exists to prevent. The other side of that trade is a viewer who closes the window
/// during the last five seconds of a film and finds it beginning again, which is indistinguishable
/// from wanting it to and is what a looping preview does anyway.
pub(super) const FILM_END_GRACE_SECONDS: f64 = 5.0;

/// What a pinned video's player being gone means, from what this app claimed about it.
///
/// Three answers because a player can leave for three reasons and this app has to tell them
/// apart with nothing to ask it: FFmpeg's player is not asked whether it is still there, and all
/// this app has is a clock it started and a length the probe read.
///
/// The distinction that matters is `Finished` against `Gone`, and it is the whole of why the loop
/// is this app's. `-loop 0` was dropped from the launch because an input `-ss` sends the player's
/// own loop back to the keyframe it landed on rather than to the start of the film — asked for
/// 720.000 s on a file 723.754 s long, it began at 718.757 s and looped there forever, which is a
/// preview that has stopped previewing anything but its own last five seconds. So the player plays
/// once and this app begins it again, and the only question each time is whether the film finished
/// or the player was killed.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PlayerEnd {
    /// The film was watched to its end and is to be begun again from the top.
    Finished,
    /// The player is gone for some other reason, and the film holds where it stopped.
    Gone,
}

/// Whether a player that has ended ended because its film did.
///
/// `reached` is how much of the film this app's clock believes has gone by, counted from the
/// moment this app began the player and deliberately *not* wrapped at the length of the file — the
/// wrap in `transport_clock` is for drawing a playhead that goes round, and a decision about
/// whether the film is over must be able to see past it.
///
/// Every condition here is a thing that could be otherwise, and each is one that has been: a held
/// film whose player was killed is not a film that finished, a film whose length the probe never
/// read cannot be judged against its end, and a player this app asked to end is a player whose
/// replacement is already on its way.
pub(super) fn player_end(
    playing: bool,
    reached: f64,
    duration: Option<f64>,
    retiring: bool,
) -> PlayerEnd {
    let long_enough = duration
        .filter(|duration| *duration > 0.0)
        .is_some_and(|duration| reached >= duration - FILM_END_GRACE_SECONDS);

    if playing && !retiring && long_enough {
        PlayerEnd::Finished
    } else {
        PlayerEnd::Gone
    }
}

/// Begin a pinned film again from the beginning, answering whether that was what to do.
///
/// This is the loop FFmpeg's player is no longer asked for, and it is one tick of the preview loop
/// rather than a thread of its own: a player that has ended is a process that is gone, the loop
/// already asks every tick whether the pinned media is alive, and the relaunch it performs is the
/// same relaunch a seek and a resize perform (see `restart_pinned_player`), so the film comes back
/// over the hold that keeps its window up while the new player puts one of its own — the gap at the
/// loop boundary is the same gap the first start of the preview has, and it is covered by the same
/// thing.
///
/// It is asked *first* in a pinned tick, before anything else that reads the player as gone, and
/// that order is the whole of how this loop works. Three other things in the tick notice a player
/// that has exited — a bar that must stop claiming to be playing (`settle_pinned_transport`), a
/// file written down as failed before it ever drew a frame (`pin_media_failed_before_a_frame`),
/// and a pin whose media has come apart (`pin_media_is_alive`) — and each of them is right about
/// what it sees and wrong about what it means: a film that has finished looks exactly like one
/// whose player has been killed. Answered after any of them, the loop would be restarting a pin
/// that has already closed itself.
pub(super) fn loop_ended_pinned_player() -> bool {
    if current_media_type() != Some(MediaType::Video) {
        return false;
    }

    // A player that is still running has not ended, whatever the clock says. The grace below is for
    // a player that has *gone* and whose going is read as the film finishing; it is not a licence to
    // cut the tail off a film that is still playing, and without this check a film shorter than the
    // grace would be begun again on the tick after it began.
    if is_video_process_running() {
        return false;
    }

    let Some((path, content, transport, _)) = pinned_playback_state() else {
        return false;
    };
    if transport.pending_hold {
        return false;
    }

    let Some(started) = transport.started else {
        // Nothing was ever running, so nothing ended: this is a bar with nothing behind it and it
        // is `settle_pinned_transport`'s to answer, not this one's.
        return false;
    };

    let retiring = VIDEO_RETIREMENT
        .lock()
        .is_ok_and(|pending| pending.is_some());
    let reached = started.1 + started.0.elapsed().as_secs_f64();
    if player_end(
        transport.paused_at.is_none(),
        reached,
        transport.duration,
        retiring,
    ) != PlayerEnd::Finished
    {
        return false;
    }

    restart_pinned_player(&path, content, 0.0, false);
    true
}

/// Give a player that has been told nothing yet the hold that was written down for it.
///
/// It is the other half of carrying a hold through a relaunch, and it exists because a relaunch
/// cannot finish the job itself: the player it begins has no window for a moment, and a hold this
/// app could already deliver is a key posted to that player's own window (see `ffplay_key_pause`).
/// So the relaunch writes the hold down as owed — the file *is* meant to be held, and a bar that
/// forgot that would draw a play button over a paused film — and this is what delivers it on the
/// tick that finds the window up, which is the first tick after the one that began it.
///
/// What the hold is re-based onto is the player's own clock rather than the second the relaunch
/// began at, which is the whole of the correction: a player told to hold a moment after it started
/// has got to that moment's worth of film since, and a bar left at the relaunch's second would be
/// drawing the film a fraction behind where it stopped (see `transport_clock`).
///
/// A press of the play button takes an owed hold back rather than racing it, which is what
/// `PinTransport::released` is for, and a player that died takes it with the claim it belonged to
/// (see `PinTransport::player_gone`). What is left owing after that is a player that never put a
/// window up, and the bar's own answer to that is a play button — which begins the file again
/// rather than leaving it stuck (see `settle_pinned_transport`).
pub(super) fn settle_pending_hold() {
    if current_media_type() != Some(MediaType::Video) {
        return;
    }

    let Some((_, _, transport, _)) = pinned_playback_state() else {
        return;
    };
    if !pending_hold_delivers(transport.pending_hold, transport.drag_held) {
        return;
    }

    // Nothing is done until there is a window to post a key to, which is the whole of the wait:
    // a key posted into nothing is dropped, and the press that put the hold here would be
    // answered by a fallback that begins a *third* player (see `toggle_pinned_playback`).
    if ffplay_key_pause() {
        let at = transport_clock(&transport).unwrap_or(0.0);
        update_pin_transport(|state| state.held(at));
    }
}

/// Whether a hold a relaunch wrote down as owed is this tick's to deliver.
///
/// **A gesture that is still holding the film delivers it instead, and the two are the same key.**
/// A hold is the pause key, and the pause key is a toggle: a relaunch begun under a gesture's hold
/// — which is what a resize's settle is, and a seek taken from a hand that is still down — writes
/// the hold down as owed because the player it has just begun has no window to post it through yet.
/// The tick that delivers it and the tick that lets the gesture go of it are the same tick, so both
/// post, and two toggles on one player is a film playing over a bar with a pause glyph on it — which
/// is what a paused video did on a resize and did not do on a move, the whole difference being that
/// a move has no relaunch behind it to owe anything.
///
/// The claim is the transport's own rather than a flag beside it, because it is the one fact that
/// reconciles the two: every path that begins, holds, lets go of or loses a player takes it with
/// them (see `video_drag_hold_claim`).
pub(super) fn pending_hold_delivers(owed: bool, gesture_held: bool) -> bool {
    owed && !gesture_held
}

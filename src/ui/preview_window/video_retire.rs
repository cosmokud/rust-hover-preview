//! Handing a film over: retiring the player that is up while its replacement is launched, and
//! what the end of a pinned film means.

use super::*;
use crate::formats::routing::NavCategory;

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
    /// (see `terminate_player_pid`).
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
/// player is a *request* (see `terminate_player_pid`), and this is the one place that asks for one
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
    terminate_player_pid(pid);
    let end = retire_end(is_player_pid_alive(pid));
    if end == RetireEnd::Gone {
        // The player is confirmed gone, so the record of it goes with it rather than being left
        // for the next run to look for.
        engine_processes::forget(pid);
    }
    end
}

/// End a player a newer gesture has superseded, before it can publish:
/// kill-on-supersede rather than land-and-retire.
///
/// A relaunch that a second gesture strands behind a standing cover must
/// never be shown nor placed — at a stale box it lingers outside the preview,
/// playing, until the next release. So the superseded player is terminated
/// here, while the cover is still up, and reaped where that can be confirmed:
/// its record forgotten and `VIDEO_PID` cleared where it still names it, so
/// no show path can find it and no audio survives it.
///
/// Returns whether nothing is left: pid zero is already nothing, and a kill
/// that is confirmed gone leaves nothing either. A kill not yet confirmed
/// answers false, and the settle asks again on the next tick — termination is
/// a request, not a wait.
pub(super) fn kill_superseded_player(pid: u32) -> bool {
    if pid == 0 {
        return true;
    }

    terminate_player_pid(pid);

    if is_player_pid_alive(pid) {
        return false;
    }

    // Confirmed gone, so the record of it goes with it rather than being left
    // for the next run to look for — the same bookkeeping a settled
    // retirement does (see `end_retired_player`).
    engine_processes::forget(pid);
    clear_video_process_state(pid);
    true
}

/// Players a pin's teardown asked to die whose kill is not yet confirmed.
///
/// A teardown cannot wait — termination is a request, not a wait — and the pin
/// is going away, so no settle will ever retry the pending record it leaves
/// behind. Clearing that record from under an unconfirmed kill loses the only
/// retry the kill had: an orphan ffplay with zombie audio. So an unconfirmed
/// kill is parked here instead, owned by this list rather than by any pin, and
/// retried to confirmation by the loop's own settle, which runs with no pin up
/// (see `settle_video_retirement`).
pub(super) static ORPHANED_VIDEO_KILLS: Lazy<Mutex<Vec<u32>>> =
    Lazy::new(|| Mutex::new(Vec::new()));

/// Park an unconfirmed kill where the loop retries it to confirmation.
///
/// A kill confirmed on the spot is reaped here instead and never parked: its
/// record forgotten and `VIDEO_PID` cleared where it still names it, the same
/// bookkeeping `kill_superseded_player` does. Only a player still dying after
/// being asked is owned by the list — never dropped, never waited on.
pub(super) fn retire_orphaned_player(pid: u32) {
    if pid == 0 {
        return;
    }
    if end_retired_player(pid) == RetireEnd::Gone {
        clear_video_process_state(pid);
        return;
    }
    if let Ok(mut held) = ORPHANED_VIDEO_KILLS.lock() {
        if !held.contains(&pid) {
            held.push(pid);
        }
    }
}

/// Retry every orphaned kill, keeping only the ones still dying.
///
/// Only confirmed-gone kills leave the list, and only by being reaped — the
/// same terminate-confirm-forget as every other kill path (see
/// `end_retired_player`).
fn settle_orphaned_video_kills() {
    let Ok(mut held) = ORPHANED_VIDEO_KILLS.lock() else {
        return;
    };
    held.retain(|&pid| {
        if end_retired_player(pid) != RetireEnd::Gone {
            return true;
        }
        // Confirmed gone: its pid goes with it where it still names it, so no
        // show path can find a dead player after its pin has gone.
        clear_video_process_state(pid);
        false
    });
}

#[cfg(test)]
pub(super) fn orphaned_video_kills() -> Vec<u32> {
    ORPHANED_VIDEO_KILLS
        .lock()
        .map(|held| held.clone())
        .unwrap_or_default()
}

#[cfg(test)]
pub(super) fn orphaned_video_kills_push_for_test(pid: u32) {
    if let Ok(mut held) = ORPHANED_VIDEO_KILLS.lock() {
        if !held.contains(&pid) {
            held.push(pid);
        }
    }
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
    // The orphan reaper runs with no pin up: a teardown hands unconfirmed
    // kills here precisely because no settle will retry them afterwards, so
    // this must not wait on the retirement below having anything to do.
    settle_orphaned_video_kills();

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
            is_player_pid_alive(pending.replacement),
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

    // A gesture that killed owns the dead interval: the frozen hold, the
    // claim and the stopped clock are the press's to write and the end's to
    // take back, so a player nothing is behind reconciles nothing here.
    if gesture_snapshot_active() {
        return;
    }

    if is_video_process_running() {
        return;
    }

    // Read in one look and let the lock go before the reconcile below
    // re-locks: the reconcile writes through `update_pin_transport`, which
    // takes the pin's lock itself, and a guard held across it is a thread
    // waiting on a lock it owns (see `pin_media_is_alive` for the rule).
    let played = {
        let Some(pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin() else {
            return;
        };
        pin_playhead(&pin.transport).or(pin.transport.paused_at)
    };
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

/// Whether a pass that is still being played has reached the end of its file, by this app's own
/// clock.
///
/// It is the end for a player that *stays up* when the film ends rather than exiting — mpv, which
/// is launched with `--keep-open=yes` and holds the last frame (see `start_mpv_video_playback`) —
/// and it is asked with no grace at all, because the five seconds `player_end` allows are the width
/// needed to read a *gone* player's going as the film finishing rather than as a kill (see
/// `FILM_END_GRACE_SECONDS`). A player that is still up has done no going, the length of the file
/// is the file's own, and the pass is over when the clock reaches it — no earlier, which is the
/// whole of what an end of file that steps on needs: five seconds read off the tail of every film
/// is a next file begun while the one before it is still talking.
pub(super) fn pass_ended(reached: f64, duration: Option<f64>) -> bool {
    duration
        .filter(|duration| *duration > 0.0)
        .is_some_and(|duration| reached >= duration)
}

/// Answer the end of a pinned film's pass: stop it, step on, or begin it again from the top.
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
///
/// **And it is where the end of a film that leaves its player up is answered at all.** mpv stands on
/// its last frame when the film ends (`--keep-open=yes`) instead of exiting, so the reconciliation
/// that catches a player which has *gone* never runs for it, and nothing else in the tick would tell
/// the bar the film was over — which is why a walk that this end of file asks for is asked here, off
/// this app's own clock and with no grace cut off the tail of the film (see `pass_ended`).
///
/// A pass that ends in a *hold* is not answered here: `End of File → Pause` takes the film back to the
/// first frame and holds it there from the tick's own rewind, before its end is ever read off this
/// clock, so that the press which plays it again is the plain resume of a held film rather than a seek
/// against a player that has stopped at the end of its file (see `pin_video_pauses_at_the_top` and
/// `video_loop_tick`).
pub(super) fn loop_ended_pinned_player(
    anchor: Option<NavCategory>,
    wait: &mut Option<PinWait>,
    busy: bool,
) -> bool {
    // An advance already on its way is not asked for again: the clock underneath keeps
    // counting past the end while the walk is out, so without this every tick would ask
    // another walk and the pin would step once per answer rather than once per end.
    if busy {
        return false;
    }

    if current_media_type() != Some(MediaType::Video) {
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

    // The pin's own end of file, read where the end is answered: `Pause` is a film that
    // stops at its end — neither begun again nor stepped on — `Loop` is the same film again,
    // and `Continue` is the walk's next file, taken in the folder's own order or at random
    // where the pin's shuffle says so (see `ask_pin_walk`).
    let end_of_file = current_pin_mode_video_end_of_file();

    // A film being held is not a film whose pass has ended, whatever the clock underneath says:
    // the start a hold keeps counting from is the one the clock runs on, so a film held for longer
    // than it lasts would otherwise be read as finished under the hand (see `PinTransport::held`).
    if transport.paused_at.is_some() {
        return false;
    }

    let retiring = VIDEO_RETIREMENT
        .lock()
        .is_ok_and(|pending| pending.is_some());
    let reached = started.1 + started.0.elapsed().as_secs_f64();

    // A player that is still up when the clock says the pass is over is the one player that stays
    // up at the end of its film — mpv — and its end is the film's own end rather than the five
    // seconds the gone-player reading allows for. A player that has gone keeps that reading, and
    // with it the retirement check: the `retiring` player is the one being replaced, and a going
    // that has to be told from the film finishing is exactly what that flag is for (see
    // `player_end`). The player still up is the replacement, whose own clock was begun at the
    // second it was given (see `restart_pinned_player`), so no retirement of another process is
    // anything to this reading.
    let live = is_video_process_running();
    let ended = if live {
        pass_ended(reached, transport.duration)
    } else {
        player_end(true, reached, transport.duration, retiring) == PlayerEnd::Finished
    };
    if !ended {
        return false;
    }

    // **Where the rewind can take the film home, `Pause` is that rewind's to answer and not this
    // reading's.** A film the rewind holds at its top is a film whose pass never reaches its end at
    // all, and one whose end this reading *did* catch — a tick that landed after the film had already
    // stopped — is taken back to the top by that same rewind later in the same tick (see
    // `pin_video_pauses_at_the_top` and `video_loop_tick`). Writing the film down as *held* here, which
    // is what this reading used to do, is what would stop it: a film this app believes is held is a
    // film the rewind leaves alone (see `video_loop_action`). And a seek to the top written from here
    // would be a seek against a player that has stopped at the end of its file — the restart of the
    // film's whole chain, picture and all, that the rewind exists to avoid.
    //
    // **And a film nothing can take home is still this reading's**, asked of the same predicate the
    // rewind will ask itself (`rewind_due`) rather than a paraphrase of it: a film shorter than the
    // margin a rewind is posted within is never rewound at all (see `video_launch`), so its end is the
    // arm below, which is the bar's only answer to a player that stands on its last frame rather than
    // going.
    if pin_video_pauses_at_the_top(&path)
        && video_launch::rewind_due(true, transport.duration, reached)
    {
        return false;
    }

    if end_of_file == PinEndOfFile::Pause {
        // **The player that stays up is the end of file that never used to arrive.** A player
        // that exits is reconciled by `settle_pinned_transport`, which finds nothing behind the
        // bar and writes the film down as held — but mpv holds its last frame instead of going, so
        // nothing was ever reconciled: the clock counted past the end of the film for as long as
        // the pin was up, the bar swept round, and the play/pause button claimed a film that was
        // playing. So the pass's end is written down here, where the end of file is answered: the
        // clock stops at the film's own end, the bar stands there, and the button becomes the play
        // button a finished film is owed. It is the arm a film the rewind cannot take home is left
        // with — one whose length nothing has read (see the note above).
        if !live {
            // A player that has gone is the reconcile's to answer, and it writes the second the
            // *wrapped* clock stands at — which is where the bar and the next press expect it.
            return false;
        }
        let held = transport.duration.unwrap_or(reached).min(reached);
        // **A player this app cannot command is held by key, and that is the only hold there is for
        // one.** The routes under mpv run the player they were launched with, and that player is the
        // *hover's* — `-loop 0`, given because the hover's own loop is on, and kept by the pin, which
        // begins no player of its own over one that is already going — so the film wraps itself at
        // the end of every pass unless the player is stopped, and the one thing that stops it is the
        // key the pause button already posts (see `video_key_pause`). A player with no window to post
        // that key through is ended instead, which is the same pair the press is (see
        // `hold_pinned_player_without_a_key`).
        if video_route(&path) == VideoRoute::Ffplay {
            if video_key_pause() {
                update_pin_transport(|transport| transport.held(held));
            } else {
                hold_pinned_player_without_a_key(held);
            }
            return true;
        }
        update_pin_transport(|transport| transport.player_gone(held));
        return true;
    }

    if end_of_file != PinEndOfFile::Loop {
        // The film played through and the end of file is a step on, not a loop: the walk's
        // next file is asked for — the same ask the caption's own **Next** button makes, with
        // the same wait for its answer. A lingering player (mpv paused on its last
        // frame, where FFmpeg's has exited) is left for the swap the walk lands to take
        // down, the way a walk over a playing film is.
        if ask_pin_walk(path.clone(), 1, anchor) {
            *wait = Some(PinWait::new());
        }
        // And the film is held where it stands until the walk lands, which is the same write the
        // end of file that pauses makes and for a second reason besides: the walk is a folder read
        // and the clock under the bar would otherwise keep its own count past the end while the pin
        // waits — a bar that wraps round while the next file is being looked for reads as a film
        // that is still playing, and the second the walk was asked from is the end of the film.
        if let Some(end) = transport.duration {
            update_pin_transport(|transport| transport.player_gone(end));
        }
        return true;
    }

    // A loop of a film the player is still holding is a wrap the tick gives it, down the pipe the
    // player listens on (`video_loop_tick`): the wrap is asked *after* this function in the tick,
    // and it is the one road mpv's loop has. A player that has gone is begun again here, because
    // there is nothing left to tell: the grace this reading allows is for a player that has gone
    // and whose going is read as the film finishing, not a licence to cut the tail off a film that
    // is still playing.
    if live {
        return false;
    }

    restart_pinned_player(&path, content, 0.0, false);
    true
}

/// Answer what the film the media engine is playing has asked at the end of its pass: step the walk
/// on, hold the bar on the film the engine has stopped on, or do nothing.
///
/// It is the pin's own ask rather than the film's, so it is the work of the answer that is here — the
/// session let go, the walk asked for, the bar written down — and the reading of what the film's end
/// of file wants is `pinned_native_video_end`, which a run needs an engine for and this does not.
///
/// It answers whether the pass's end was acted on at all, which is what a tick can say about a film
/// whose end it has read and a film it has not.
pub(super) fn answer_native_pass_end(
    end: NativePassEnd,
    path: &Path,
    anchor: Option<NavCategory>,
    wait: &mut Option<PinWait>,
) -> bool {
    match end {
        NativePassEnd::Nothing => false,
        NativePassEnd::Walk => {
            // The session is let go — which is also what stops the end being read again on the sixty
            // ticks a second this loop turns — and the next file of the folder is asked for, the same
            // ask the caption's own **Next** button makes and the same wait for its answer.
            video_player::stop();
            ask_pin_walk(path.to_path_buf(), 1, anchor);
            *wait = Some(PinWait::new());
            true
        }
        NativePassEnd::Hold => {
            // **The engine stands on the last frame of a film it was told to play once and says
            // nothing more**, so this is the bar's only way of hearing that the film is over: the
            // second it is held at is the film's own length, and a hold is what stops the clock and
            // turns the button into the play button a finished film is owed (see `pin_playhead` and
            // `pin_is_playing`). The press takes a film held here back to its top (see
            // `toggle_pinned_playback`).
            //
            // Nothing is written where the bar already says the film is held, which is every tick
            // after the first: the hold is a transition, and a film whose length nothing has read has
            // no second for one — the same corner the reading of a film mpv stands on is left with
            // (see `loop_ended_pinned_player`).
            //
            // shortcut: the clock half of the reading this answer comes from can go on saying the film
            // is at its end for a moment after this app has taken it to another second, so a press over
            // a finished film can leave the bar held over a film that is playing again until the press
            // after it; upgrade when the engine's own seeked event is watched (see `video_ended`).
            let Some((_, _, transport, _)) = pinned_playback_state() else {
                return false;
            };
            if transport.paused_at.is_some() {
                return false;
            }
            let Some(at) = pin_duration(&transport) else {
                return false;
            };
            update_pin_transport(|transport| transport.held(at));
            true
        }
    }
}

/// What the end of a pass of the film the media engine plays asks of the pin.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum NativePassEnd {
    /// Nothing: a pass that is still being played, and a film the engine is looping itself — which
    /// reports no end of one at all (see `apply_native_video_loop`).
    Nothing,
    /// The walk's next file, which is the ask the caption's own **Next** button makes.
    Walk,
    /// The bar held on the film the engine has stopped on, which is `Pause` left where it stands.
    Hold,
}

/// What the film the media engine is playing for the pin asks at the end of its pass, from the two
/// facts that decide it.
///
/// The engine was handed the end of file the film on screen is played under, so its end is its own
/// to report: the engine's word for it, or its clock at the length of the file, is the end of the
/// film (see `video_player::video_ended` and `apply_native_video_loop`).
///
/// It is a function of its two facts rather than of a session because a session is a thread-local
/// nothing but the running app holds (see `pinned_native_video_end`, which is the reading itself).
pub(super) fn native_pass_end(end_of_file: PinEndOfFile, ended: bool) -> NativePassEnd {
    if !ended {
        return NativePassEnd::Nothing;
    }

    match end_of_file {
        PinEndOfFile::Loop => NativePassEnd::Nothing,
        PinEndOfFile::Continue => NativePassEnd::Walk,
        PinEndOfFile::Pause => NativePassEnd::Hold,
    }
}

/// What the pass of the film the media engine is playing for the pin asks of the pin now, read where
/// the end is answered.
///
/// A path is part of the question because the engine holds one film at a time — a file the engine is
/// not playing has not ended, whatever the clock of the file on screen says (see
/// `video_player::playing_path`).
pub(super) fn pinned_native_video_end(path: &Path) -> NativePassEnd {
    if !(pinned() && pinned_path().as_deref() == Some(path))
        || video_player::playing_path().as_deref() != Some(path)
    {
        return NativePassEnd::Nothing;
    }

    native_pass_end(
        current_pin_mode_video_end_of_file(),
        video_player::video_ended(),
    )
}

/// Give a player that has been told nothing yet the hold that was written down for it.
///
/// It is the other half of carrying a hold through a relaunch, and it exists because a relaunch
/// cannot finish the job itself: the player it begins has no window for a moment, and a hold this
/// app could already deliver is a key posted to that player's own window (see `video_key_pause`).
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
    if video_key_pause() {
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

//! What the pin plays: where the playhead is, seeking it, toggling it, stepping a subtitle, and
//! handing the picture from a player to its replacement without taking it off the screen. A
//! gesture asks nothing of a player (see `RestartReason`).

use super::pin_trace::{note_pin_trace, path_name};
use super::*;

/// Where a pinned preview's playhead is, in seconds: the engine's own answer where the media
/// engine is playing it, and this app's clock over the player's start where FFmpeg's is — the
/// same clock a sound's card is drawn from, and for the same reason (see `audio_clock`).
pub(super) fn pin_playhead(transport: &PinTransport) -> Option<f64> {
    if let Some(paused) = transport.paused_at {
        return Some(paused);
    }

    match current_media_type() {
        Some(MediaType::NativeVideo) => video_player::position().or(Some(0.0)),
        Some(MediaType::Video) => transport_clock(transport),
        _ => None,
    }
}

/// The second a player of this app's has reached, off its own clock and nothing else.
///
/// It is `pin_playhead` with the hold ignored, and it is what a hold is *re-based onto* once the
/// player has actually been told to hold: the second written when the pause began is the second
/// the film was at when the press happened, and the film has been playing since — so it is where
/// the file was, not where it stopped (see `settle_pending_hold`).
pub(super) fn transport_clock(transport: &PinTransport) -> Option<f64> {
    transport.started.map(|(at, from)| {
        let elapsed = from + at.elapsed().as_secs_f64();
        match transport.duration {
            Some(duration) if duration > 0.0 => elapsed % duration,
            _ => elapsed,
        }
    })
}

/// Where a pinned preview's playhead is, read out of the pin that is up: the same answer as
/// above, for a caller that has no transport in hand — a resize deciding where the player it is
/// about to begin again should start (see `relayout_pinned_media`).
pub(super) fn pinned_playhead() -> Option<f64> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    pin_playhead(&pin.transport)
}

/// How long the pinned file plays: the length the probe read, or the engine's own answer where
/// the engine is the one that knows.
pub(super) fn pin_duration(transport: &PinTransport) -> Option<f64> {
    transport.duration.or_else(|| {
        (current_media_type() == Some(MediaType::NativeVideo))
            .then(video_player::duration)
            .flatten()
    })
}

/// Whether a pinned preview is playing: the engine's own answer, or whether a player of this
/// app's is running.
///
/// A file FFmpeg plays is answered from what this app has written down *and* from whether the
/// process is still there, because the first of those is an assertion and the second is the only
/// thing about a player that can be observed — this player reports nothing, which is why what it
/// is doing is kept rather than asked for (see `transport_playing`).
///
/// The process is read off the published pid rather than through `is_video_process_running`
/// because the bar is drawn from a paint, and asking that would take the media's lock: a paint
/// is reached with the media held in one place and a lock a thread already owns is not a wait
/// but a stop (see `pin_media_is_alive` for the same rule at length). What the pid is asked is
/// weaker in one way and stronger in another — weaker, because it is only cleared once a death
/// has been *confirmed*, so it can outlive a process by a check; stronger, because it is an
/// atomic and so can be read anywhere, from any thread, at any point in a tick.
pub(super) fn pin_is_playing(transport: &PinTransport) -> bool {
    if transport.paused_at.is_some() {
        return false;
    }

    match current_media_type() {
        Some(MediaType::NativeVideo) => video_player::is_playing(),
        Some(MediaType::Video) => {
            transport_playing(transport, VIDEO_PID.load(Ordering::Acquire) != 0)
        }
        _ => false,
    }
}

/// Whether the sound behind a card is playing, which is the one question the card's own play
/// button answers and the one its menu's hold row is named from.
///
/// A player of this app's is the process behind the sound, asked of the media rather than off
/// the published pid: a held player was killed without a wait, so its pid outlives it until a
/// death is confirmed and would answer "playing" for a sound nobody can hear. The lock is safe
/// here — nothing that builds the card's menu runs with the media held, which is the one thing
/// the bar a pid is read from cannot promise (see `pin_is_playing`).
///
/// The engine's own is asked whether it is *held* rather than whether it is playing, because a
/// session that was asked to hold is a session that is still there (see `a_sound_that_is_held`).
pub(super) fn a_sound_is_playing() -> bool {
    let running = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.video_process.is_some()))
        .unwrap_or(false);

    running || (video_player::is_playing() && !video_player::is_paused())
}

/// What the transport's own actions need of the pin: the file being played, the box the player's
/// window fills, where its playback is, and the level it is playing at.
pub(super) fn pinned_playback_state() -> Option<(PathBuf, ScreenRegion, PinTransport, PinVolume)> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    Some((pin.path.clone(), pin.content, pin.transport, pin.volume))
}

/// How many players `restart_pinned_player` has begun.
///
/// A test seam only: the re-entrant double this counts would otherwise spawn
/// (or fail to spawn) identically twice, leaving no state behind to tell one
/// relaunch from two.
#[cfg(test)]
pub(super) static RESTART_COUNT: AtomicU64 = AtomicU64::new(0);

/// How many players have been begun (see `RESTART_COUNT`).
#[cfg(test)]
pub(super) fn restart_count() -> u64 {
    RESTART_COUNT.load(Ordering::Acquire)
}

/// Forget how many players have been begun, so each test counts its own ends.
#[cfg(test)]
pub(super) fn clear_restart_count() {
    RESTART_COUNT.store(0, Ordering::SeqCst);
    LAST_RESTART_REASON.store(0, Ordering::SeqCst);
}

/// The reason the last player was begun again, as the code of the line the trace wrote.
///
/// The same one question the trace answers — a pid that changed with no reason said out loud is the
/// fault this field is read for (see `relaunch_pinned_player`) — and the question a handover's own
/// line is read against, so a seek's picture is written where a seek is what began the player (see
/// `pin_film_thumbnail::note_the_seek_picture`).
pub(super) static LAST_RESTART_REASON: AtomicU64 = AtomicU64::new(0);

/// The reason the last player was begun again, if one was (see `LAST_RESTART_REASON`).
#[cfg(test)]
pub(super) fn last_restart_reason() -> Option<RestartReason> {
    RestartReason::of_code(LAST_RESTART_REASON.load(Ordering::Acquire))
}

/// The reason the last player was begun again, **taken**: what a handover reads once to say whether a
/// seek began the source it is about to draw, so one reason is never read as a second seek's (see
/// `pin_film_thumbnail::hand_the_picture_over`).
pub(super) fn take_the_last_restart_reason() -> Option<RestartReason> {
    RestartReason::of_code(LAST_RESTART_REASON.swap(0, Ordering::AcqRel))
}

/// Why a player is begun again, which is the one thing every relaunch says out loud.
///
/// The set is exactly the user actions that need a new player on a route with no IPC: a seek, a
/// level, a subtitle track, the end of a file, and the engine that cannot take an instruction any
/// other way. **There is no gesture reason and no resize reason, and that is the rule rather than
/// an omission**: a press, a move, a drag and a resize leave the player running, because the film
/// is a thumbnail of a source window the gesture does not move (see `pin_film_thumbnail`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum RestartReason {
    Seek,
    Volume,
    Subtitle,
    End,
    Engine,
}

impl RestartReason {
    /// The word the trace line and the test seam are written with.
    pub(super) fn as_str(self) -> &'static str {
        match self {
            RestartReason::Seek => "seek",
            RestartReason::Volume => "volume",
            RestartReason::Subtitle => "subtitle",
            RestartReason::End => "end",
            RestartReason::Engine => "engine",
        }
    }

    /// The code this reason is written down as (see `LAST_RESTART_REASON`).
    fn code(self) -> u64 {
        match self {
            RestartReason::Seek => 1,
            RestartReason::Volume => 2,
            RestartReason::Subtitle => 3,
            RestartReason::End => 4,
            RestartReason::Engine => 5,
        }
    }

    fn of_code(code: u64) -> Option<RestartReason> {
        match code {
            1 => Some(RestartReason::Seek),
            2 => Some(RestartReason::Volume),
            3 => Some(RestartReason::Subtitle),
            4 => Some(RestartReason::End),
            5 => Some(RestartReason::Engine),
            _ => None,
        }
    }
}

/// End the player a pinned video is playing in and begin another one at a second of the file.
///
/// It is what a seek and a resume from a pause both are with FFmpeg, whose player can be told
/// nothing once it is running: the same bargain the sound path makes, where a file dropped in
/// half way is a player started at that second (see `start_audio_player`). It is also what the one
/// thing about a running player that *can* be changed is asked by — the level, and the subtitle
/// track — so the player that begins is given the pin's own level rather than the tray's setting
/// (see `PinVolume`), and the pin's own track rather than whatever the player would pick for
/// itself (see `PinTransport::subtitle`).
///
/// The order of the two is the whole of what the user sees, and it is the reason the player being
/// replaced is *taken out of the media* rather than ended here. It is left running and drawing —
/// its window is the source the pin's host draws — until the source of the player begun in its
/// place has taken the picture over, which the loop ends it for (see `retire_replaced_player`).
/// Ending it first would leave the pin's band showing the desktop for as long as the new player
/// takes to open the file, seek, and put a window up, which on a 1440p file is long enough to read
/// as a flash, and which is why nothing here goes through `stop_video_playback`: that call is for
/// a player that has ended, and this one has not.
///
/// What *is* stopped is the media's own background work, which belongs to the player being
/// replaced and has nothing to do with whether that player is still there.
///
/// `holding` is a decision this function does not make for itself, because it is not the function's
/// business. A seek, a resize settling and a change of track are all relaunches of a file that is
/// meant to go on being held, and the press that cannot be answered by a key is a relaunch of a file
/// the hand has just let go of — so which of the two this is belongs to the caller, and the only
/// thing that has to be true of either is that it is said out loud (see `PinTransport::begun`).
pub(super) fn restart_pinned_player(
    path: &PathBuf,
    content: ScreenRegion,
    seconds: f64,
    holding: bool,
    reason: RestartReason,
) {
    relaunch_pinned_player(path, content, seconds, holding, false, reason);
}

/// Begin the pin's player on the very frame the hover is showing, and hold that hover there until
/// the pin can show the same one.
///
/// It is the relaunch above with the one difference a take-up from a hover has: the player being
/// replaced is the hover's, and it is the picture on screen until the pin's own player has drawn
/// that frame. So it is **frozen** where it stands rather than retired for the loop to end (see
/// `pin_film_host::FrozenHover`), the second it froze at is **read back off the player** rather than
/// estimated from this app's clock, and the player begun in its place is begun standing on that
/// second rather than playing from it (see `video_launch::mpv_args` and
/// `pin_film_host::settle_pin_handover`).
pub(super) fn begin_the_pinned_film_at_the_hovers_frame(
    path: &PathBuf,
    content: ScreenRegion,
    seconds: f64,
    holding: bool,
) {
    relaunch_pinned_player(path, content, seconds, holding, true, RestartReason::Engine);
}

/// The relaunch both roads above are, where `handover` says the player being replaced is a hover's.
fn relaunch_pinned_player(
    path: &PathBuf,
    content: ScreenRegion,
    seconds: f64,
    holding: bool,
    handover: bool,
    reason: RestartReason,
) {
    #[cfg(test)]
    {
        RESTART_COUNT.fetch_add(1, Ordering::AcqRel);
    }
    LAST_RESTART_REASON.store(reason.code(), Ordering::SeqCst);
    // **Every relaunch says why it happened, in one line**, so a pid that changed can be read
    // against the user action that asked for it. A player that changes with no line here is a
    // bug, and the words are exactly the actions that need one (see `RestartReason`).

    let width = (content.2 - content.0).max(1);
    let height = (content.3 - content.1).max(1);
    let volume = pinned_volume_level();
    let subtitle = pinned_subtitle();

    // The player this relaunch is replacing, read from the record rather than from the handle in
    // the media. The record is this app's own account of the player it last started, and it is the
    // right account to read here for a second reason as well as the obvious one: it is also the
    // account of a player whose handle was dropped without a kill ever being confirmed, which is
    // precisely the process a second one must not be stacked up behind (see
    // `kill_stray_video_process`). A handle would have said nothing about that one, because there
    // is no handle left to ask.
    //
    // It is read before the new player is started, because starting one overwrites the record with
    // its own pid — and a player being retired deliberately does *not* clear it, only a death
    // being confirmed does, so it still names the process that was there a moment ago.
    let replaced = VIDEO_PID.load(Ordering::SeqCst);

    // **The picture on screen is frozen before anything is begun over it**, and it is the whole of
    // the placeholder a relaunch is owed: the player being replaced is left standing until its
    // replacement has drawn (see `retire_replaced_player`), and a player left playing walks the film
    // on through the whole wait and then jumps back to the second that was asked for, which reads as
    // a stall and a jump rather than as a wait behind a still frame (see
    // `freeze_the_outgoing_picture`).
    freeze_the_outgoing_picture(replaced, freeze_road());

    // **A relaunch is a new generation**, so a replacement an earlier one began can never publish
    // under this one, and the unpublished one is killed BEFORE this replacement is spawned rather
    // than after it. A spawn first leaves a start-to-kill window in which the stale player can
    // still publish: its window arrives, the monitor finds it, and something shows it at a stale
    // box beside its replacement. Killed first, it is dying before its replacement exists. Only a
    // confirmed kill clears the record; a kill not yet taken stays pending and the next launch asks
    // again (see `reap_superseded_relaunch`).
    PIN_PLAYER_GENERATION.fetch_add(1, Ordering::AcqRel);
    let killed = reap_superseded_relaunch();

    // The handle is taken and the process left alone: nothing else in here can reach the player, so
    // the only thing that can end it is the loop that was told about it.
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.cancel_background_work();
            media.video_process = None;
        }
    }

    // **A hover's second is read off the player rather than counted by this app**, and the read is
    // asked after the freeze above and before the pipe is dropped: the frame the user is looking at
    // is mpv's own last one, and the clock this app keeps is an estimate of where that frame is —
    // good enough to begin a film on and not good enough to *stand* one on, which is what the
    // handover asks (see `pin_film_host::the_hosted_frame_is_presented`). A read that does not
    // answer leaves the estimate as the second to begin on, which is the launch this road had before
    // the hover was frozen at all.
    let frozen = if handover {
        freeze_the_hover(replaced)
    } else {
        None
    };
    let seconds = frozen.unwrap_or(seconds);

    // The player being replaced keeps nothing back: the pipe a command would go down is the
    // replacement's once one is named, and the replaced player's name is not reusable (see
    // `mpv_ipc`). The frozen hover's last line is the one above — what is left for it is a kill,
    // not a command (see `pin_film_host::retire_the_frozen_hover`).
    mpv_ipc::forget_pipe(replaced);

    let process = start_video_playback(
        path, content.0, content.1, width, height, seconds, handover, volume, subtitle, None,
    );
    let pid = process.as_ref().map(|player| player.id()).unwrap_or(0);
    note_pin_trace!(
        "player restarted for {} from {replaced} to {pid}",
        reason.as_str()
    );

    // **A copy the film was waiting on rides this relaunch.** The launch names the film's copy in its
    // filter chain out of the film's own geometry, so this one line is the whole of what the preload
    // costs the relaunch it waited for (see `preload_the_subtitles_for_the_next_relaunch`).
    if take_the_pending_subtitles(path) {
        note_pin_trace!(
            "subtitle carried into a relaunch pid {pid} file {}",
            path_name(path)
        );
    }

    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.video_process = process;
        }
    }

    // The player being replaced is retired once the replacement exists, and it is retired *and still
    // standing* — a relaunch that ended it instead would leave the band empty for the whole of the
    // wait, and the wait is as long as a player takes to open the file. It is a frozen frame it holds
    // there, not a film still playing (see `freeze_the_outgoing_picture`).
    if pid != 0 {
        // ...unless it is the superseded one just killed: nothing waits on its arrival and
        // nothing will ever show it.
        if killed != Some(replaced) {
            if handover {
                // **The hover stays the picture until the pin has drawn it.** What retires it is the
                // handover, once the pin's window is up over the same frame; ending it here would
                // leave the band the desktop for as long as the pin's player takes to open the file
                // and decode, which is the flash the whole wait exists to keep out (see
                // `pin_film_host::settle_pin_handover`).
                pin_film_host::hold_the_hover_for_the_handover(replaced, seconds);
            } else {
                retire_replaced_player(replaced, pid);
            }
        }
    } else if handover {
        // A pin's player that did not come up leaves no swap to make, so the hover it froze is ended
        // here rather than left frozen for a handover that is never going to happen — the same road
        // an unconfirmed kill goes down, and the only one that can end a player this app has stopped
        // holding (see `retire_orphaned_player`).
        retire_orphaned_player(replaced);
    } else {
        // A replacement that did not come up leaves nothing to arrive, so the player on screen is
        // ended here instead of waiting for an arrival that is never going to happen — which is the
        // road the sweep is still for, and the only road that reaches it.
        kill_stray_video_process();
    }

    // The window the new player puts up is placed by the tick, which re-asserts it every two
    // hundred milliseconds; a seek is one window gone and another arriving, so it is asked for
    // now rather than at the next of those. **It is a place and not a show for a film drawn in the
    // pin's host**: the host is up already and the registration standing in it is the picture, so
    // this cannot put a window that has decoded nothing in front of one that is still drawing (see
    // `show_current_replacement`).
    //
    // **And not a superseded generation either** (see `show_current_replacement`): a newer step's
    // bump kills the in-flight relaunch, and a bump landing between the tag below and this show is
    // re-checked after it — the condemned window hidden again rather than left published — with the
    // residual stated honestly where the helper is.
    note_pinned_relaunch(pid);
    show_current_replacement(pid, content);

    update_pin_transport(|transport| transport.begun(seconds, pid != 0, holding));
    with_pin(|pin| pin.volume.playing_at = volume);
}

/// Hold the hover where it stands and read back the second its picture is at.
///
/// **The freeze is the hold the relaunch already made** (see `freeze_the_outgoing_picture`), and what
/// is left here is the read: mpv leaves the frame it was showing on screen when it is told to hold (and
/// a suspended player is not told anything at all), so the hover window the user is looking at keeps
/// that frame for as long as the take-up takes — nothing is captured, nothing is painted over it, and
/// nothing hides it (see `begin_the_pinned_film_at_the_hovers_frame`).
///
/// `None` where the player did not answer with its position, which leaves the caller's own clock as
/// the second to begin the pin on: a frozen picture this app cannot read is still a frozen picture,
/// and the launch is the one it would have been before the read existed (see
/// `mpv_ipc::get_property`).
fn freeze_the_hover(pid: u32) -> Option<f64> {
    mpv_ipc::get_property(pid, "time-pos").and_then(|answer| answer.parse().ok())
}

/// **Freeze the picture on screen before its replacement is launched**, so that the whole wait for a
/// new source to draw is a still frame rather than the film walking on and then jumping back to the
/// second the relaunch asked for.
///
/// It is a hold of the player being replaced and nothing else: no capture, no scaling, no painting
/// over the band and no hide. The registration the host wears goes on drawing that window (see
/// `pin_film_thumbnail::the_source_is_standing`), and a player that is held keeps the frame it was
/// showing, which is exactly the placeholder the wait is owed. The player is never resumed: it is
/// ended once its replacement has taken the picture over (see `settle_video_retirement`).
///
/// **The two engines are held differently, and the difference is what each can be told.** mpv takes
/// `set_property pause true` down the pipe it listens on and holds its last frame. FFmpeg's player has
/// no command channel, and the one thing it can be told — the pause key — is a *toggle* that clears
/// the very window the Shell draws the pin through, which is a registration of a black window (see
/// `FreezeRoad`). So a player with no pipe is **suspended** where it stands instead: it stops drawing
/// without being given anything to redraw (see `ffplay_window::suspend_player_threads`).
///
/// **And whether the frame survived is sampled on the spot**, of the player's own client, so a run
/// says whether the picture DWM goes on drawing is really there. `ok` is a window that kept its
/// frame; `black` is a window with nothing in it, which is the answer that leaves the capture road —
/// grab the last presented frame into a still painted in the film host — as the only thing left (see
/// `pin_film_thumbnail::the_sources_picture_is_lit`).
///
/// Answers whether a player on screen was frozen, which is nothing on screen where the player had no
/// window of its own to be told through: that wait is the walk-and-jump it always was rather than a
/// failure.
pub(super) fn freeze_the_outgoing_picture(pid: u32, road: FreezeRoad) -> bool {
    if pid == 0 || road == FreezeRoad::Off {
        return false;
    }

    if mpv_ipc::set_paused(pid, true) {
        note_pin_trace!("outgoing picture frozen pid {pid} by the pipe");
        return true;
    }

    if road == FreezeRoad::Key {
        // **A film already held is already frozen, and the key toggles**: a key posted to a player
        // that is holding would set it going again, which is the one thing this must not do. What
        // says so is the hold this app has written down for the transport — including a hold still
        // owed to a player that has not been given it yet, which is a player that is still playing
        // (see `PinTransport`). A suspension has no such toggle and no such question (see
        // `FreezeRoad`).
        if the_player_on_screen_is_held() {
            return true;
        }

        if video_key_pause_for(pid) {
            note_pin_trace!("outgoing picture frozen pid {pid} by the key");
            return true;
        }

        return false;
    }

    // **An engine with no command channel is suspended rather than told** (see `FreezeRoad`): the
    // handful of calls that stop it drawing is the whole of the freeze, and the frame it last put up
    // is left exactly where it is.
    let threads = suspend_player_threads(pid);
    if threads == 0 {
        return false;
    }

    let sample = match pin_film_thumbnail::the_sources_picture_is_lit(pid) {
        Some(true) => "ok",
        Some(false) => "black",
        None => "none",
    };
    note_pin_trace!("outgoing picture frozen pid {pid} suspended threads {threads} sample {sample}");
    true
}

/// The road a relaunch freezes its outgoing player by, from `RHP_FREEZE`, read once per run.
///
/// **`suspend` is the road that ships** (see `freeze_the_outgoing_picture`): a player with no command
/// channel is suspended where it stands, so the frame it last presented is untouched.
///
/// **`key` is the dead road, and it is kept as the A/B arm rather than deleted** because the
/// measurement that killed it is one line to repeat: FFmpeg's player clears its own window when the
/// pause key pauses it, so the registration the host went on wearing drew a black window for the
/// whole wait — three relaunches traced with the freeze line 11 ms before each restart, the outgoing
/// window standing as the source throughout, and the placeholder black. A key is a toggle and not a
/// setting besides: a held film is already a still frame, and a second key would set it going.
///
/// `off` freezes nothing, which is the road before there was a freeze at all: the film walks on under
/// the registration and jumps back to the second the relaunch asked for when the new source draws.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum FreezeRoad {
    Suspend,
    Key,
    Off,
}

/// `RHP_FREEZE` as a freeze road: `key` and `off` are the two named arms, and anything else —
/// including nothing at all — is the suspension (see `FreezeRoad`).
pub(super) fn freeze_road_from(setting: Option<&str>) -> FreezeRoad {
    match setting.map(str::trim) {
        Some("key") => FreezeRoad::Key,
        Some("off") => FreezeRoad::Off,
        _ => FreezeRoad::Suspend,
    }
}

/// The freeze road this run was begun on: a question about the run rather than about a relaunch, so
/// it is read once and every relaunch of the run agrees on it (see `FreezeRoad`).
static FREEZE_ROAD: Lazy<FreezeRoad> =
    Lazy::new(|| freeze_road_from(std::env::var("RHP_FREEZE").ok().as_deref()));

/// The freeze road of this relaunch: the run's own, or the answer a test stood in (see
/// `stand_the_freeze_road_in`).
pub(super) fn freeze_road() -> FreezeRoad {
    #[cfg(test)]
    if let Some(road) = FREEZE_ROAD_STAND_IN.lock().ok().and_then(|held| *held) {
        return road;
    }
    *FREEZE_ROAD
}

/// The road a test is answered with in place of the run's own, which `RHP_FREEZE` names and a suite
/// does not set (see `freeze_road`).
#[cfg(test)]
static FREEZE_ROAD_STAND_IN: Mutex<Option<FreezeRoad>> = Mutex::new(None);

/// Stand a freeze road in place of the run's own, or give the run's own answer back with `None`, so
/// each test puts the slot back rather than leaving it to the test the runner reaches next.
#[cfg(test)]
pub(super) fn stand_the_freeze_road_in(road: Option<FreezeRoad>) {
    if let Ok(mut held) = FREEZE_ROAD_STAND_IN.lock() {
        *held = road;
    }
}

/// Whether the player the pin is drawing is already holding its film, which is what says the pause
/// key must not be posted to it (see `freeze_the_outgoing_picture`).
fn the_player_on_screen_is_held() -> bool {
    pinned_playback_state()
        .is_some_and(|(_, _, transport, _)| transport.paused_at.is_some() && !transport.pending_hold)
}

/// Whether the pinned window is showing a held file, for a relaunch that has no transport of its
/// own in hand: the two callers that have none are a seek taken from the bar and the settle of a
/// resize, and a resize is exactly as much a relaunch of a held file as a seek is (see
/// `restart_pinned_player`).
pub(super) fn pinned_is_held() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport.paused_at.is_some()))
        .unwrap_or(false)
}

/// The subtitle track a pinned window is showing, out of what the probe read of the file on
/// screen: nothing for a pin with no file up, and for a file the probe never answered for
/// nothing either (see `video_subtitles`).
pub(super) fn pinned_subtitle() -> Option<usize> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    pin.transport.subtitle
}

/// Take a pinned video to a second of its file.
///
/// The two engines are asked differently: the media engine is told where to go and gets there on
/// its own, while FFmpeg's player is ended and begun again at that second — which is why a drag
/// on the bar shows where it is being taken rather than seeking as it moves, and the seek is made
/// where the pointer lets go.
///
/// The restart is *kept* rather than replaced by a key sent to the player, and the reason is
/// what a drag on the bar means. The bar knows a second — four minutes thirty-one and a half, say
/// — and every key FFmpeg's player has for moving is a step of a fixed size: ten seconds left,
/// ten right, a minute up or down. None of them is "go to this second", and rounding a drag to
/// the nearest multiple of ten would be a seek that lands somewhere the hand did not ask for and
/// cannot be undone by dragging the bar further. `-ss` says the second itself, lands there, and
/// lands there on the *next* frame rather than on the one a step of playback happened to arrive
/// at, which for a film at 144 frames a second is the difference between the frame the hand let
/// go on and one up to seventy milliseconds either side of it.
///
/// It is also the only way to land a seek with the film's own subtitles in the right place. A
/// time-shifted anime release carries its subtitle track offset from the picture, and a seek that
/// moved the video without naming a track would leave the two disagreeing about where they are —
/// so the track travels with the seek because the relaunch names it, which is the same fact B5
/// needs and the reason one relaunch path serves both (see `restart_pinned_player`).
///
/// What it must not fight is the loop's own re-hover checks, and it does not: nothing here
/// touches the media the hover installed, only the player behind it, and the player that a hover
/// would have started for a *different* file is swept first (see `kill_stray_video_process`).
pub(super) fn seek_pinned_playback(path: &PathBuf, content: ScreenRegion, seconds: f64) {
    match current_media_type() {
        Some(MediaType::NativeVideo) => {
            video_player::seek(seconds);

            // A file this side has paused is drawn at the second it was stopped at rather than at
            // the engine's own position (see `pin_playhead`), so a seek made while it is paused has
            // to move that second with it. The engine goes where it was taken either way; what
            // this is for is the bar, which would otherwise spring back to the second the pause
            // began at the moment the hand let go — the file seeked, and the bar saying otherwise.
            update_pin_transport(|transport| transport.sought(seconds));
        }
        // A held file is still held afterwards: the bar is drawn from a second this side wrote
        // down, so a seek taken while it is held has to move that second with it — which for
        // FFmpeg's player is the same statement as telling the player it is to be begun at that
        // second held (see `PinTransport::begun` and `settle_pending_hold`).
        Some(MediaType::Video) => {
            if video_route(path) == VideoRoute::Mpv
                && mpv_ipc::seek_absolute(VIDEO_PID.load(Ordering::SeqCst), seconds)
            {
                // mpv is taken to the second while it runs: the player stays and the bar's own
                // clock is rebased onto where it was taken rather than a player begun at that
                // second. A held file stays held, because mpv's own seek keeps its pause, and the
                // rebase happens only where the command reached the player: a seek that did not
                // is the relaunch below, or the bar would claim a second mpv never went to (see
                // `mpv_ipc`).
                update_pin_transport(|transport| {
                    transport.sought(seconds);
                    transport.started = Some((Instant::now(), seconds));
                });
                // And the clock a *pass* is measured by is moved with the bar, because a seek
                // down this road is a command to a player that goes on running rather than a
                // relaunch with a fresh clock: left where it was, the loop's own reading of the
                // film would still stand at the second the seek left behind, and a seek taken
                // into the last moments of a film would be undone by the next tick, which wraps
                // the film back to zero (see `rebase_video_loop_clock_to`).
                video_hw::rebase_video_loop_clock_to(VIDEO_PID.load(Ordering::SeqCst), seconds);
                // **And the one bounded read that says what the seek's frame is drawn from**: a seek
                // in place moves the picture inside a source whose client is the one placement's, and
                // a frame drawn smaller than the pin's box there would be a number in this line (see
                // `pin_film_thumbnail::note_the_seek_picture`).
                pin_film_thumbnail::note_the_seek_picture(
                    VIDEO_PID.load(Ordering::SeqCst),
                    seconds,
                );
            } else {
                // **A pipe this app could not write to is the relaunch road after all**, and the
                // player on screen is left running behind it: the registration standing in the host
                // is the picture until the replacement's own source takes it over, and only then is
                // the old player ended (see `relaunch_pinned_player`).
                restart_pinned_player(
                    path,
                    content,
                    seconds,
                    pinned_is_held(),
                    RestartReason::Seek,
                );
            }
        }
        _ => {}
    }
}

/// Pause a pinned video, or set it going again.
///
/// The media engine is asked to hold where it is. FFmpeg's player is asked as well, which used
/// to be impossible: it can be told nothing (see §"Where the commands went" in `ARCHITECTURE.md`),
/// so a hold was that player *ended* with the second it had reached kept for the player that
/// took its place. It can be asked one thing — a key, posted to its own window the way Windows
/// would have delivered it — and that is enough, because a player that is asked to hold does
/// hold, and a player that is holding is still there to be let go of.
///
/// Which changes what a resume costs: it used to be a second start, with the second it stopped
/// at typed into `-ss` and the window rebuilt; it is now a key to the player that never stopped.
/// That is the whole of the difference in what the user sees, and it is why the hold keeps the
/// start rather than giving it up (see `PinTransport::held`).
///
/// The fallback is the arrangement this function had before, and it is not a formality: a player
/// that has no window — one still starting, or one whose window has gone the way its process
/// went — cannot be given a key, and a player that cannot be given a key has to be *ended* to
/// hold a file, exactly as before. So the shape of the state is the same either way and only the
/// cost differs, which is what keeps one function answer for both.
///
/// One case is answered by neither, and it is the same gap a relaunch leaves. A hold the relaunch
/// wrote down and has not delivered yet is not a player holding — the player is playing, and the
/// key that would let it go of the hold has not been sent — so a press in that moment is not a
/// request to un-pause anything: it is the hand taking the hold back, and the film carries on from
/// where the relaunch put it (see `PinTransport::pending_hold`).
pub(super) fn toggle_pinned_playback(
    path: &PathBuf,
    content: ScreenRegion,
    transport: PinTransport,
) {
    let playing = pin_is_playing(&transport);
    let playhead = pin_playhead(&transport).unwrap_or(0.0);

    match current_media_type() {
        Some(MediaType::NativeVideo) => {
            video_player::set_paused(playing);
            update_pin_transport(|state| {
                if playing {
                    state.held(playhead);
                } else {
                    // The engine reports its own position, so the second the hold began at is
                    // only what the *bar* was drawn at and not what the file is at; re-basing
                    // onto it is a no-op for the engine's own reading and is the whole of the
                    // correction for the clock a player of this app's is measured by.
                    state.released(playhead);
                }
            });
        }
        Some(MediaType::Video) => {
            // mpv is told to hold and to go over its own pipe; FFmpeg's player is posted a key to
            // its own window. Both leave a player that keeps running and reports nothing, so both
            // are a hold rather than a relaunch (see `mpv_ipc` and `video_key_pause`).
            let mpv = video_route(path) == VideoRoute::Mpv;
            let hold_the_player = |paused: bool| {
                if mpv {
                    mpv_ipc::set_paused(VIDEO_PID.load(Ordering::SeqCst), paused)
                } else {
                    video_key_pause()
                }
            };

            if transport.pending_hold {
                // The player is playing and no key has reached it, so there is nothing to let go
                // of: what this press does is stop the hold from ever being sent.
                update_pin_transport(|state| state.released(transport.paused_at.unwrap_or(0.0)));
            } else if playing {
                if hold_the_player(true) {
                    update_pin_transport(|state| state.held(playhead));
                } else {
                    hold_pinned_player_without_a_key(playhead);
                }
            } else {
                // A film held at its own end is a film that has *finished* rather than one paused
                // in the middle of itself (see `loop_ended_pinned_player`): the player stands on
                // the last frame, and letting it go there plays that frame and ends the pass again
                // at once. So the press begins the film from the top, which is what a play button
                // over a finished film is for.
                let held = transport.paused_at.unwrap_or(0.0);
                let finished = transport
                    .duration
                    .is_some_and(|duration| duration > 0.0 && held >= duration);
                let from = if finished { 0.0 } else { held };
                if finished && mpv {
                    // mpv stops on its last frame rather than pausing there (see `mpv_args`), so
                    // the seek to the top is what begins the film again; the play below goes with
                    // it for a player a press had paused, which a seek leaves paused.
                    mpv_ipc::seek_absolute(VIDEO_PID.load(Ordering::SeqCst), from);
                }
                // **A finished film on the route mpv is not is a player begun again rather than one
                // told to go on**, and that is the whole of what a finished film leaves this press to
                // decide: a key cannot take FFmpeg's player anywhere, so resuming the film that is
                // standing at its own end plays the last frames of the file out of it and lets it
                // close its own window — measured, and the band is a hole the player's window is the
                // picture through. The pass the press asks for is therefore a new player at the top,
                // which is the same relaunch a seek on that route is (see `restart_pinned_player`).
                if finished && !mpv {
                    restart_pinned_player(path, content, from, false, RestartReason::End);
                } else if hold_the_player(false) {
                    update_pin_transport(|state| state.released(from));
                    // And the clock a pass is measured by moves onto the second the film is played
                    // from: a resumed film is a pass with its whole length still to run, and a
                    // clock left where the hold began would wrap it early (see
                    // `rebase_video_loop_clock_to`).
                    video_hw::rebase_video_loop_clock_to(VIDEO_PID.load(Ordering::SeqCst), from);
                } else {
                    // A player begun to resume a file is not one to be told to hold, or the press
                    // would arrive twice over: once here, and once as a hold the relaunch owed. So
                    // the relaunch is told, at its own call site, that this is a file being let go.
                    restart_pinned_player(path, content, from, false, RestartReason::Engine);
                }
            }
        }
        _ => {}
    }
}

/// A file whose player could not be given a key, held the only other way there is: the player
/// is ended and the second it had reached kept for the player that takes its place.
///
/// This is what a hold was before a video FFmpeg plays could be paused at all, and it is kept
/// for the case that still needs it — a player with no window to post a key to. Nothing is lost
/// by falling back to it: `PinTransport::held` is the same field either way, so the bar looks
/// the same, and what differs is only which player will be there when the file is let go.
pub(super) fn hold_pinned_player_without_a_key(playhead: f64) {
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.cancel_background_work();
            stop_video_playback(media);
        }
    }
    update_pin_transport(|state| {
        state.player_gone(playhead);
    });
}

/// Take a pinned film to its next subtitle track, and begin the player again naming it.
///
/// The number is written down *before* the player is begun, so that a relaunch that failed to
/// start still leaves the choice standing: a player that did not come up is not a reason to
/// forget which track the user asked for, and a key pressed again after it gets what it asked
/// for rather than the next track on from one it never reached.
///
/// Everything else is refused rather than attempted. A file with no subtitle streams has no next
/// track, and a file the probe has not answered for is treated as having none — guessing at a
/// track in a file whose header has not been read is a `-sst` the player may refuse outright,
/// which is a player that exits (see `video_subtitles`). The engine's own videos are refused too:
/// it shows no subtitles at all, in frame-server mode or any other (see `video_player`), so there
/// is nothing on screen for a track to change. And a film mpv is drawing is stepped over the pipe
/// it listens on rather than by a relaunch: mpv found and chose its own tracks, so there is no
/// number of this app's to write down, and `cycle sub` is mpv's own word for walking them (see
/// `mpv_ipc`).
pub(super) fn step_pinned_subtitle() {
    if current_media_type() != Some(MediaType::Video) {
        return;
    }

    let Some((path, content, transport, _)) = pinned_playback_state() else {
        return;
    };

    // **mpv picks and draws its own tracks, and `cycle sub` is its own word for walking them.**
    // There is no number here for this app to write down — nothing reports back which track mpv
    // chose — but the pipe it listens on is the whole of what the step needs, so the step is a
    // command rather than a player begun again. A command that did not reach the player leaves the
    // step to the road below, which is how FFmpeg's player's track has always been walked (see
    // `mpv_ipc`).
    if video_route(&path) == VideoRoute::Mpv
        && mpv_ipc::cycle_subtitles(VIDEO_PID.load(Ordering::SeqCst))
    {
        return;
    }

    let Some(next) = next_subtitle(transport.subtitle, video_subtitles(&path).count) else {
        return;
    };

    if Some(next) == transport.subtitle {
        return;
    }

    update_pin_transport(|state| state.subtitle = Some(next));
    restart_pinned_player(
        &path,
        content,
        pinned_playhead().unwrap_or(0.0),
        transport.paused_at.is_some(),
        RestartReason::Subtitle,
    );
}

/// Hand a film's copied subtitles to the player that is drawing it, where that player can be told.
///
/// **This is the whole of what a late copy costs a film mpv plays: nothing.** The copy is a whole
/// subtitle with no streams of its own, mpv takes it down its pipe and draws it, and the player, its
/// window and its picture never move — so there is no black frame, and no second decode of the film
/// (see `mpv_ipc::add_subtitles`). FFmpeg's player has no pipe to be told on and answers `false` here,
/// which is what leaves the relaunch as its answer.
fn add_the_subtitles_to_the_player(path: &Path, pid: u32, track: Option<usize>) -> bool {
    if video_route(path) != VideoRoute::Mpv {
        return false;
    }
    let Some(copy) = video_subtitle_copy(path, track) else {
        return false;
    };
    if !mpv_ipc::add_subtitles(pid, &copy) {
        return false;
    }
    note_pin_trace!(
        "subtitle added without a restart pid {pid} file {}",
        path_name(&copy),
    );
    true
}

/// The film whose ready copy of its subtitles has not been drawn yet: **a preload for the next
/// relaunch**, because the engine that cannot be told can only be given a copy by being begun again,
/// and a player begun again for a copy alone is a second decode and a picture that changes hands for a
/// subtitle (see `preload_the_subtitles_for_the_next_relaunch`).
static PRELOADED_SUBTITLES: Lazy<Mutex<Option<PathBuf>>> = Lazy::new(|| Mutex::new(None));

/// Write a film's ready copy down as **pending for its next relaunch**, and say so.
///
/// It is the whole of what a late copy costs a film FFmpeg's player draws. That player can be told
/// nothing, so a copy can only reach the picture by the player being begun again — and a player begun
/// again for the copy alone changes the picture's hands and decodes the file a second time for a
/// subtitle nobody asked for at that moment. What the copy does instead is wait: a launch builds its
/// filter chain out of the film's own geometry, which already names the copy once the extraction has
/// written it (see `video_launch::subtitle_filter`), so the *next* relaunch for a reason of the user's
/// — a seek, a level, a track, an engine — draws it on the way past (see `take_the_pending_subtitles`).
///
/// This is the "preload in the background, activate at the next seek" a copy was asked for, and it is
/// why a film with a pending copy is never begun again on the copy's own account (see
/// `reload_pinned_subtitles`).
fn preload_the_subtitles_for_the_next_relaunch(path: &Path) {
    let already = PRELOADED_SUBTITLES
        .lock()
        .ok()
        .and_then(|pending| pending.clone());
    if already.as_deref() == Some(path) {
        return;
    }
    if let Ok(mut pending) = PRELOADED_SUBTITLES.lock() {
        *pending = Some(path.to_path_buf());
    }
    note_pin_trace!(
        "subtitle preloaded for the next relaunch file {}",
        path_name(path),
    );
}

/// Take a film's pending copy where it is this film's: the relaunch that carries it into the picture.
fn take_the_pending_subtitles(path: &Path) -> bool {
    let Ok(mut pending) = PRELOADED_SUBTITLES.lock() else {
        return false;
    };
    if pending.as_deref() != Some(path) {
        return false;
    }
    *pending = None;
    true
}

/// Forget the film whose copy is waiting for a relaunch, so each test carries its own.
#[cfg(test)]
pub(super) fn clear_pending_subtitles() {
    if let Ok(mut pending) = PRELOADED_SUBTITLES.lock() {
        *pending = None;
    }
}

/// Begin a pinned window's player again so a film's own subtitles, copied out of it a moment
/// ago, are drawn — where the pin is showing that film right now.
///
/// **This is the one reload a hover deliberately does not have.** A hover whose copy was still
/// coming was answered without subtitles and is not begun again under the pointer — that was the
/// user's own choice, and the small files are what the hover *after* it draws (see
/// `video_launch::subtitle_filter`). A pinned window is a different answer: the user has said
/// with the pin that this film is the one being watched, so what its player was handed is worth
/// correcting the moment it can be — and the way it is corrected is that the copy **rides the next
/// relaunch**: the picture is not changed hands and the film is not decoded a second time for a
/// subtitle (see `preload_the_subtitles_for_the_next_relaunch`).
///
/// Everything else is refused rather than attempted: no pin, a pin on another file, a pin whose
/// media is not a film this app plays with FFmpeg (the engine's own videos are not this app's to
/// reload, and no extraction runs for one), and a pin whose file has no player up are all
/// nothing to begin again.
///
/// **mpv is never begun again for this**, which is the difference the pipe makes: the copy is handed to
/// the player that is drawing the film (see `add_the_subtitles_to_the_player`), so the line above is
/// the relaunch road and not this one. FFmpeg's player is begun again for nothing at all: the copy is
/// preloaded and rides the next relaunch the user asks for (see
/// `preload_the_subtitles_for_the_next_relaunch`).
///
/// Answers whether the copy was answered at all, which is what the tests read.
pub(super) fn reload_pinned_subtitles(path: &Path) -> bool {
    if current_media_type() != Some(MediaType::Video) {
        return false;
    }

    let Some((pinned, _, transport, _)) = pinned_playback_state() else {
        return false;
    };

    if pinned.as_path() != path {
        return false;
    }

    if add_the_subtitles_to_the_player(
        &pinned,
        VIDEO_PID.load(Ordering::SeqCst),
        transport.subtitle,
    ) {
        return true;
    }
    if video_route(&pinned) == VideoRoute::Mpv {
        // A film mpv plays is never begun again for its subtitles: the pipe is the whole of what the
        // copy needs, and a relaunch is the black frame this road removed (see
        // `add_the_subtitles_to_the_player`).
        return false;
    }

    preload_the_subtitles_for_the_next_relaunch(&pinned);

    true
}

/// Begin the adopted player of a freshly made pin again, where the film's copied subtitles are
/// ready and the player predates them.
///
/// **This is the corner the ready-message cannot reach.** A copy that landed while a pin was up
/// is answered by `reload_pinned_subtitles` there and then; a copy that landed with no pin up
/// has already had its message and gone — and the hover player that predates it is adopted by
/// the pin as it stands (a take-up re-parents the `ffplay` the hover began rather than starting
/// one, see `take_up_pinned_window`), so the pinned window would draw on without the subtitles
/// for as long as nothing else begins the player again. The take-up is the moment to correct
/// that, and what it asks is the two facts the correction is made of: the copy is ready (see
/// `video_copy_ready`) and the player was begun without subtitles (see
/// `player_draws_subtitles`).
///
/// A film whose copy is still coming is deliberately not this: there is nothing to draw yet,
/// and the message will find the pin when there is. And the correction is a preload rather than a
/// relaunch (see `preload_the_subtitles_for_the_next_relaunch`).
pub(super) fn reload_adopted_subtitles(path: &Path) -> bool {
    if current_media_type() != Some(MediaType::Video) {
        return false;
    }

    let Some((pinned, _, transport, _)) = pinned_playback_state() else {
        return false;
    };

    if pinned.as_path() != path {
        return false;
    }

    if !video_copy_ready(path) {
        return false;
    }

    // **A player that can be told is told**, and this is the road where a copy lands on a player that
    // predates it: mpv takes the file down its pipe with no relaunch at all (see
    // `add_the_subtitles_to_the_player`).
    if add_the_subtitles_to_the_player(
        &pinned,
        VIDEO_PID.load(Ordering::SeqCst),
        transport.subtitle,
    ) {
        return true;
    }
    if video_route(&pinned) == VideoRoute::Mpv {
        return false;
    }

    if player_draws_subtitles(path) {
        return false;
    }

    // **A player that cannot be told preloads the copy rather than being begun again**, whether it is
    // the player its own launch began or one a hover's host was taken up with: the film keeps the
    // picture it has, and the copy rides the next relaunch of the user's (see
    // `preload_the_subtitles_for_the_next_relaunch`).
    preload_the_subtitles_for_the_next_relaunch(&pinned);

    true
}

/// Move the hover's already-running host into the pin, and give the film the pin's level.
///
/// **Nothing is begun and nothing is frozen.** The player in the host is the one that was drawing
/// the hover: the same child — or, on the thumbnail road, the same source window, which the Shell
/// draws into the very host that becomes the pin's — the same pid, the same playhead. What changes is
/// the box the picture is cropped to and the level the pin plays it at; the transport was already
/// written at the take-up from `video_loop_playhead`, so the bar continues from where the hover had
/// reached rather than from a frozen second.
pub(super) fn adopt_the_hovered_film_in_the_host(host: HWND) -> bool {
    unsafe { pin_film_host::take_up_the_pin_film_host(Some(host)) };
    if pin_film_host::film_host() != Some(host) {
        // The window went between the take-up's own reading and this call, so there is no film in
        // the pin for the level below to be about (see `take_up_the_pin_film_host`).
        return false;
    }

    // This also asks `aid` back where the hover was begun muted and the pin has a level to hear.
    let volume = pinned_volume_level();
    mpv_ipc::set_volume(VIDEO_PID.load(Ordering::SeqCst), volume);
    with_pin(|pin| pin.volume.playing_at = volume);

    true
}

/// Begin the film a pin took up from a hover in the host it is drawn in, where the player the
/// hover began is still a window of the desktop's.
///
/// **A pin does not keep a hover's player for a film it draws in a window of its own.** The host is
/// where the picture belongs from the take-up on, and a player begun before the pin existed was
/// given a top-level window, `--ontop` and a place of its own — none of which is what a film the
/// pin draws is. So the take-up is a relaunch, held where the film was held and at the level the
/// pin's own bar says, with the one thing about it that a seek and a resize do not have: the hover
/// is **frozen** where it stands, the second its picture is at is read back off the player, and the
/// pin's player is begun **standing on that frame** rather than playing from it (see
/// `begin_the_pinned_film_at_the_hovers_frame`).
///
/// So the hover's own last frame is the placeholder, and it stays the picture until the pin's player
/// has that same frame drawn — not a window made — and the tick's handover swaps the two in one pass
/// (see `pin_film_host::settle_pin_handover`). Nothing covers it and nothing hides it: this app's own
/// frame there would be a guess at one mpv is already showing.
///
/// The player being replaced is retired by that road and not ended here: it is the picture on screen
/// until the hosted one has the same frame in it (see `pin_film_host::FrozenHover`).
///
/// Nothing is begun for a film the host already holds — every show of a hosted pin would otherwise
/// be a player begun again — and nothing for a pin with no player up at all, which is a film that
/// has not started rather than one to move (see `pin_film_host::host_for_the_pins_film`).
pub(super) fn begin_the_adopted_film_in_the_host() -> bool {
    let Some((path, content, transport, _)) = pinned_playback_state() else {
        return false;
    };

    if !pin_film_host::the_film_is_shown_in_a_host(&path, None)
        || pin_film_host::film_host().is_none()
        || pin_film_host::the_pinned_player_was_begun_in_the_host()
        || VIDEO_PID.load(Ordering::SeqCst) == 0
    {
        return false;
    }

    let seconds = pin_playhead(&transport).unwrap_or(0.0);
    let holding = transport.paused_at.is_some();

    // **FFmpeg's player cannot prove a frozen frame, so its take-up adopts the player the step
    // already began.** A player this pin launched unhosted is still the right player at the right
    // second: taking its window into the host is the same pid and no restart at zero. A player
    // whose window has not been published yet is noted into the host here, and the monitor adopts
    // it as it appears (see `adopt_a_stranded_ffplay_into_the_pin_host`).
    if video_route(&path) != VideoRoute::Mpv {
        if the_pinned_media_holds_the_player(VIDEO_PID.load(Ordering::SeqCst)) {
            pin_film_host::adopt_a_stranded_ffplay_into_the_pin_host(
                VIDEO_PID.load(Ordering::SeqCst),
                &path,
            );
            return false;
        }
        // A hover's player is not the pin's: the pin begins its own, once, in the host.
        restart_pinned_player(&path, content, seconds, holding, RestartReason::Engine);
        return false;
    }

    begin_the_pinned_film_at_the_hovers_frame(&path, content, seconds, holding);
    true
}

/// Whether the pin's media holds the player `pid`: a player a step of this pin began and installed
/// rather than a hover's, which the take-up still replaces with one of the pin's own.
fn the_pinned_media_holds_the_player(pid: u32) -> bool {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.video_process.as_ref().map(|process| process.id()))
    }) == Some(pid)
}

/// Write an answer about the transport back into the pin, if there is still one.
pub(super) fn update_pin_transport(change: impl FnOnce(&mut PinTransport)) {
    with_pin(|pin| change(&mut pin.transport));
}

/// Hold a pinned sound where it stands, or set it going again: what a Space in a window the
/// keyboard is in is, and the only pause a sound's card has (see `pinned_key_command`).
///
/// The two players are answered the way each of them takes a pause, which is the same answer the
/// bubble parks one with (see `bubble_playback_to_park`) because they are the same two players:
/// the engine Windows has is told to hold where it is and keeps reporting that second for itself,
/// so nothing is written down for it — and whether it is holding is asked of the engine itself,
/// because a session that is paused is a session that is still there and the hold is the engine's
/// own state (see `video_player::is_paused`); a player this app started cannot be told anything
/// and is ended instead, with the second it had got to kept — on the card, so that the sound is
/// still drawn where the hand left it, and for the player that takes its place when the key is
/// pressed again.
///
/// What this has in common with the bubble's is the mechanics and not the setting:
/// `Pin Mode → Pause on Minimize` is about what a collapse does, and this is about a key pressed in
/// a window the user is in. A sound with no player behind it is left alone on both counts — there
/// is no playback to hold — which is the answer a file nothing will play gives.
pub(super) fn toggle_pinned_audio(
    started: &mut Option<Instant>,
    offset: &mut f64,
    paused: &mut Option<f64>,
) {
    let Some((path, _)) = pinned_media_owner() else {
        return;
    };

    // Which player is playing the file is the answer of the player itself, and it is the answer the
    // card's clock is measured against as well (see `playing_player` and `audio_clock`).
    let Some(track) = audio_track::playable(&path) else {
        return;
    };

    match playing_player(&path, &track) {
        Player::Native => {
            // A sound the card holds at the end of its file is a sound with no session behind
            // it: the engine's own was let go the moment it ended (see
            // `pinned_native_audio_end`), and the second the card stands at is the beginning of
            // the file. The press is therefore the sound started again from that second rather
            // than a hold lifted, which is the same bargain the player this app started is
            // given below (see the `Ffmpeg` arm).
            if paused.is_some() {
                video_player::play_audio(&path, pinned_audio_level(), 0.0, pin_audio_loops());
                *paused = None;
            } else {
                // A press is a toggle, and the engine's own answer is the only one there is about
                // whether it is holding the sound: it is asked to hold where it stands, or to go
                // on from where it is holding. A session that is paused is a session that is
                // still there, so asking whether one is *playing* answered yes for a sound that
                // was already held — a pause of a pause, which is a sound held for ever with no
                // way to let it go (see `video_player::is_paused`). A session that is not there
                // is asked to hold, which pauses nothing rather than starting a sound nobody
                // asked for.
                let held = a_sound_that_is_held(*paused, Player::Native, video_player::is_paused());

                video_player::set_paused(!held);
            }
        }
        Player::Ffmpeg => {
            if let Some(from) = *paused {
                // A key while a sound is held is a sound let go: another player is begun at the
                // second it stopped at. A player that does not come up is a card left held, which
                // is the answer a machine with no sound to play with gives (see
                // `restart_pinned_audio`).
                if restart_pinned_audio(&path, from).is_some() {
                    *started = Some(Instant::now());
                    *offset = from;
                    *paused = None;
                }
            } else if started.is_none() {
                // A sound with no player behind it has no playback of ours to hold: a decoder that
                // would not have the file never got one. The key is then a key that did nothing,
                // which is what a sound that is not playing is.
                return;
            } else {
                // And a key while it is playing is a sound held: the player is ended, and the
                // second it had got to is where the card is left standing.
                let played = audio_clock(&path, *started, *offset, None).0.unwrap_or(0.0);
                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                    if let Some(media) = current.as_mut() {
                        kill_player_process(media);
                    }
                }

                *started = None;
                *offset = played;
                *paused = Some(played);
            }
        }
    }

    // The card is at the second the key moved it to, and is asked for again at once rather than
    // at the cadence a clock is watched at (see `AUDIO_CARD_DIRTY`).
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

/// Which of a pin's two players a play/pause pressed in the window is a press on, which is a
/// question about the file on screen and not about the key: the same Space is a hold on a video
/// and a hold on a sound, and which of the two it is comes from what the window is showing.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinToggle {
    /// The pin's own transport: a video held where it stands, or set going again.
    Video,
    /// A sound held where it stands, or set going again: the clock behind its card, and the only
    /// pause a card has.
    Audio,
    /// No player on screen that a Space is a play/pause of, and so nothing to press.
    None,
}

/// The player a play/pause pressed in a window showing a file of this kind is a press on.
///
/// It is asked of the kind rather than of the window, so the whole of the decision is answerable
/// with nothing up: the two players are the same two the transport bar and the sound's card each
/// act on, and a kind that has neither is a kind the key was never about (see
/// `toggle_pinned_by_key`).
pub(super) fn pin_toggle_target(kind: Option<MediaType>) -> PinToggle {
    match kind {
        Some(MediaType::Video) | Some(MediaType::NativeVideo) => PinToggle::Video,
        Some(MediaType::Audio) => PinToggle::Audio,
        // A picture, a page and a document are drawn or laid out rather than played, and none of
        // them has a player behind it for a key to hold. A sound with no player behind its card
        // is a sound all the same and is answered as one — what it has is no playback, which is
        // its own toggle's answer rather than this one's (see `toggle_pinned_audio`).
        _ => PinToggle::None,
    }
}

/// Answer a play/pause pressed in a pinned window, by asking the player the file on screen is
/// actually played by.
///
/// A video is the pin's own transport, so the key is the pause its bar offers — the same action,
/// asked by the keyboard rather than by a press on the bar, and the same state read the same way
/// (see `toggle_pinned_playback`). A sound has no bar, and a card with no way to hold a sound is
/// a sound that can only be listened to from beginning to end, so the key is the card's clock
/// instead (see `toggle_pinned_audio`).
///
/// Everything else is refused, and refused by doing nothing at all: a picture and a page have no
/// playback to hold, and a pin with no media up has nothing a key could be about. A Space in
/// front of one of those is swallowed rather than acted on, which is the answer the key already
/// got in the window procedure — a key this window is in front of is not a key to be dismissed
/// with, and there is no file here for one to mean anything about.
///
/// And the gate on all of it is the window the key arrived at, which is the whole of it: a key
/// is answered only because Windows routed it to a pin the user is in, and while the user is
/// not in the pin no key arrives here at all. Nothing reads the keyboard behind this and
/// nothing can be switched off (see `pinned_key_command` and the window procedure's
/// `WM_KEYDOWN`).
pub(super) fn toggle_pinned_by_key(
    started: &mut Option<Instant>,
    offset: &mut f64,
    paused: &mut Option<f64>,
) {
    // A collapsed pin's bubble answers a press for itself, and what it answers about is the two
    // switches' playback rather than the window's own. A park standing is given back the same road
    // a restore gives one back — and never held over, because a park written a tick after the
    // press that asked for the playback back is the bubble arguing with the hand, which is the
    // whole of what this is (see `settle_bubble_playback`).
    //
    // Nothing parked is nothing to give back, and the press is then the ordinary play/pause of the
    // player behind the window: nothing was playing when the pin went down, so what the press asks
    // for is what the bar would have asked for with the window up. It claims the bubble either
    // way — a bubble answering a press of its own is a window the user is in.
    if pin_is_collapsed() {
        if let Some(park) = pin_bubble_pause() {
            put_back_bubble_playback(park, started, offset);
            update_pin_bubble_pause(None);
        } else {
            toggle_the_player_on_screen(started, offset, paused);
        }
        with_pin(|pin| pin.bubble_run = true);
        return;
    }

    toggle_the_player_on_screen(started, offset, paused);
}

/// Answer a play/pause against the player the file on screen is actually played by.
///
/// It is the same match the pin's own bar and card press through, read out of the same place: a
/// video is the transport and a sound is the card's clock, and which of the two it is comes from
/// what the window is showing rather than from the key (see `pin_toggle_target`).
fn toggle_the_player_on_screen(
    started: &mut Option<Instant>,
    offset: &mut f64,
    paused: &mut Option<f64>,
) {
    match pin_toggle_target(current_media_type()) {
        PinToggle::Video => {
            // The same three the bar's own press reads, out of the same place: a pin with no
            // playback state behind it has no video to hold, and the key is then a key that did
            // nothing, which is what a Space is against a pin with no media up anyway.
            if let Some((path, content, transport, _)) = pinned_playback_state() {
                toggle_pinned_playback(&path, content, transport);
            }
        }
        PinToggle::Audio => toggle_pinned_audio(started, offset, paused),
        PinToggle::None => {}
    }
}

/// Begin another player of this app's for a pinned sound, at a second of its file, in place of
/// whatever was playing it: the whole of both a resume from a hold and a seek taken by a press
/// on the card's bar.
///
/// A player of this app's can be told nothing once it is running, so a sound that goes on from
/// another second is a new player begun there rather than a session moved (see
/// `start_audio_player`). The player that is playing the file now is ended first, which is the
/// question a video's seek asks as well: a handle is dropped rather than killed when the media
/// is replaced, so a sound that is not ended here goes on playing over the one beginning at the
/// second (see `restart_pinned_player`).
///
/// The answer is when the player was started, and it is nothing where none came up — a decoder that
/// will not have the file is the same answer. Which is the answer a caller keeps a sound held on
/// rather than one that begins counting a clock nothing is moving.
pub(super) fn restart_pinned_audio(path: &Path, from: f64) -> Option<Instant> {
    // The level this player is started at is the pin's own rather than the tray's, because the
    // level belongs to the window it was moved on: a knob turned on a card and then a seek would
    // otherwise restart the sound at `Volume → Audio` and undo what the hand asked for. It is read
    // before the media's lock below, because the loop's audio block reaches this with that lock
    // already held (see `pinned_audio_chrome`).
    let volume = pinned_audio_level();

    let mut current = CURRENT_MEDIA.lock().ok()?;
    let media = current.as_mut()?;

    kill_player_process(media);

    if !start_audio_playback_at(path, media, from, volume) || media.video_process.is_none() {
        return None;
    }

    Some(Instant::now())
}

/// Begin a sound a pin has just taken up again where it is, where the two rules for the end of
/// its file disagree.
///
/// A sound a pin comes up on is the sound the hover was already playing, and no player is begun
/// for it at the take-up (see `PreviewMessage::Pin`): so the rule it goes round by is the one its
/// hover's start read, while the `End of File` the card is thereafter drawn against is the pin's
/// own. Where those two answer differently the pin has a sound it cannot govern — a player that
/// takes the file round on its own never exits, so no pass of it ends for the pin's setting to be
/// read at, and the card's clock, which does ask the pin's answer, is held at the whole of the
/// file from the first pass on (see `player_loops_itself` and `audio_clock`).
///
/// Beginning it again is what tells a player of this app's anything at all (see
/// `restart_pinned_audio`), and it is begun at the second it is at rather than at the second the
/// pin's own seek would start it from: this is a sound the user is listening to, and what has to
/// change is the rule it ends by. A sound the two rules *agree* about is left exactly as it is,
/// which is every sound under the pin's default `End of File` — and a sound the hover began part
/// way into its file has no loop of its own to disagree about (see
/// `a_taken_up_sound_is_begun_again`).
///
/// The engine's own session is moved by `play_audio`, which stops whatever is there and begins
/// the sound at a second with the engine's own loop mode: a sound the engine Windows has is
/// looped by that mode and by nothing this side can set later (see `video_player::play_audio`).
pub(super) fn take_up_pinned_audio(
    path: &Path,
    started: &mut Option<Instant>,
    offset: &mut f64,
    paused: Option<f64>,
) {
    // A sound a key is holding has no player going round anything, and a sound nothing will play
    // has no player at all: both are answered with nothing begun.
    let Some(track) = audio_track::playable(path) else {
        return;
    };
    if paused.is_some() {
        return;
    }

    if !a_taken_up_sound_is_begun_again(current_audio_loop(), pin_audio_loops(), *offset) {
        return;
    }

    let played = audio_clock(path, *started, *offset, None).0.unwrap_or(0.0);

    match playing_player(path, &track) {
        Player::Native => {
            video_player::play_audio(path, pinned_audio_level(), played, pin_audio_loops());
            // The engine is taken to the second itself rather than this side counting from one,
            // so what is written down here is the record of where the sound was begun and
            // nothing more: its clock is its own (see `audio_clock`).
            *offset = played;
        }
        // The player is this app's own, so it is ended and begun at the second it is at — and a
        // player that would not come up is a sound nothing is playing, which is written down as
        // the answer that gets rather than as an old clock left counting (see
        // `PinTransport::player_gone`).
        Player::Ffmpeg => {
            *started = restart_pinned_audio(path, played);
            *offset = played;
        }
    }
}

/// Whether a sound a pin has just taken up has to be begun again for the pin's own `End of File`
/// to be able to reach it, asked of the two rules and the second the sound was started at alone —
/// so that the answer can be read without a player, a file or an engine (see
/// `take_up_pinned_audio`, and `pinned_pass_end` for the same shape about the end itself).
///
/// What is compared is not the two switches but what each of them makes of the player: a sound
/// begun anywhere but the beginning of its file is given no loop of its own however the switch is
/// set, because the position `ffplay`'s loop seeks back to is the position it was started at
/// (see `player_loops_itself`), so the two rules agree about it whatever they say.
pub(super) fn a_taken_up_sound_is_begun_again(
    hover_loops: bool,
    pin_loops: bool,
    from: f64,
) -> bool {
    player_loops_itself(hover_loops, from) && !player_loops_itself(pin_loops, from)
}

/// The level a sound this app starts is played at: the pin's own where a pin is showing the
/// sound, and `Volume → Audio` otherwise — the setting a hover's player is started at, read the
/// way it is read everywhere else (see `current_audio_volume`).
pub(super) fn pinned_audio_level() -> u32 {
    pinned_level(true)
}

/// The length a pinned sound's card is drawn with, which is the length its bar is a share of: the
/// engine's own answer where the engine is the one playing the file, and the file's own otherwise
/// — the same pair a card's own clock is measured against, so that a bar drawn to a length and a
/// bar pressed to one are the same bar (see `audio_clock`).
pub(super) fn pinned_audio_duration(path: &Path) -> Option<f64> {
    audio_clock(path, None, 0.0, None).1
}

/// The generation of the pinned player: a newer seek or step begins a new one, and every relaunch
/// is tagged with the one that is current when it is begun.
///
/// Two relaunches can be in the air at once — a seek's and the subtitle copy's that landed on it —
/// and the earlier one still lands: its replacement arrives and something shows it at a stale box,
/// outside the now-current band, playing. A relaunch superseded by a newer road is killed when
/// superseded — never shown, never placed — and the show places a replacement only while its
/// generation is current.
pub(super) static PIN_PLAYER_GENERATION: AtomicU64 = AtomicU64::new(0);

/// A relaunch whose player the picture has not been handed to yet: what the next bump kills if a
/// newer gesture arrives first, and what the next relaunch kills if it is still unpublished when
/// that one begins.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct PendingPinnedRelaunch {
    pub(super) pid: u32,
    pub(super) generation: u64,
}

/// The in-flight relaunch, if one is waiting for its source to take the picture.
pub(super) static PIN_PENDING_RELAUNCH: Lazy<Mutex<Option<PendingPinnedRelaunch>>> =
    Lazy::new(|| Mutex::new(None));

/// The generation relaunches are tagged with.
pub(super) fn current_pinned_generation() -> u64 {
    PIN_PLAYER_GENERATION.load(Ordering::Acquire)
}

/// Tag a relaunch with the current generation. A launch that never came up leaves nothing in
/// flight rather than a pid nothing will ever publish.
pub(super) fn note_pinned_relaunch(pid: u32) {
    let mut pending = PIN_PENDING_RELAUNCH
        .lock()
        .unwrap_or_else(|pending| pending.into_inner());
    if pid == 0 {
        *pending = None;
    } else {
        *pending = Some(PendingPinnedRelaunch {
            pid,
            generation: current_pinned_generation(),
        });
    }
}

/// The in-flight relaunch, if one is still waiting for its source to take the picture.
pub(super) fn pending_pinned_relaunch() -> Option<PendingPinnedRelaunch> {
    PIN_PENDING_RELAUNCH
        .lock()
        .ok()
        .and_then(|pending| *pending)
}

/// Forget the in-flight relaunch: its swap consumed it, or its kill reaped it.
pub(super) fn clear_pending_pinned_relaunch() {
    if let Ok(mut pending) = PIN_PENDING_RELAUNCH.lock() {
        *pending = None;
    }
}

/// Kill the superseded in-flight relaunch where it stands, and forget it where
/// the kill is confirmed — answering which pid was reaped, if one was.
///
/// The one helper every supersede path goes through, so an unconfirmed kill is
/// never cleared from under itself: only a confirmed kill forgets the record,
/// and a kill not yet taken stays pending for the next path to retry. Callers
/// that own no retry — the pin's own teardown — hand what is left to the
/// orphan reaper instead of dropping it (see `retire_orphaned_player`).
///
/// **A player that is the picture is never killed as superseded**, whatever the record still says:
/// its window is the source the host is drawing, and ending it is the black band a relaunch exists to
/// keep out. A record left naming one is what killed the picture at the next relaunch — the player went
/// with its replacement no better off, and the band was the host's black until a decode later — so this
/// asks the picture before it kills and leaves a standing player to the retirement that replaces it (see
/// `retire_replaced_player`).
pub(super) fn reap_superseded_relaunch() -> Option<u32> {
    let stale = pending_pinned_relaunch()?;
    if pin_film_thumbnail::the_source_is_standing(stale.pid) {
        // The record names the picture, so nothing is in flight and nothing is killed: it is forgotten
        // rather than carried, and the player it names is the one its own replacement retires (see
        // `retire_replaced_player`).
        clear_pending_pinned_relaunch();
        return None;
    }
    let killed = kill_superseded_player(stale.pid).then_some(stale.pid);
    if let Some(pid) = killed {
        // **A relaunch that never stood is dropped in silence, and says so here**: nothing of this
        // player ever reached the screen, the one the host draws is untouched, and the next relaunch
        // is already begun (see `relaunch_pinned_player`).
        note_pin_trace!("relaunch superseded before it stood pid {pid}");
        clear_pending_pinned_relaunch();
    }
    killed
}

/// The replacement is the picture: the relaunch that began it is no longer in flight.
///
/// **A record that outlives its swap is the ordering bug this exists to remove.** The in-flight record
/// is what a newer gesture kills before it can publish (see `reap_superseded_relaunch`), so a record
/// still naming a player the host has been handed draws a later kill onto the standing source — the
/// old player dies with its replacement no better off, and the pin shows the host's black for the
/// length of a decode. Written where the swap is made (see `pin_film_thumbnail::hand_the_picture_over`)
/// and where a window of the desktop's is shown (see `show_current_replacement`).
pub(super) fn settle_the_published_relaunch(pid: u32) {
    if pending_pinned_relaunch().is_some_and(|pending| pending.pid == pid) {
        clear_pending_pinned_relaunch();
    }
}

/// Whether there is no in-flight relaunch, or it is of the current generation: the settle's
/// gate for showing or placing anything.
pub(super) fn pending_relaunch_is_current() -> bool {
    pending_pinned_relaunch()
        .map(|pending| pending.generation == current_pinned_generation())
        .unwrap_or(true)
}

/// Place-and-show a replacement iff its generation is still current.
///
/// A superseded generation is never placed nor shown: the gate refuses it and
/// retries the kill instead, so a player a newer gesture condemned cannot
/// land at a stale box.
///
/// The gate and the show are two steps rather than one atomic section, and
/// the residual is stated honestly rather than hidden. `VIDEO_PID` and
/// `VIDEO_HWND` are independent atomics with no common lock — the monitor
/// thread publishes the handle, the loop publishes the pid — so a bump on the
/// window thread landing *between* the gate and the show still shows a
/// condemned window: place and show are one `SetWindowPos` behind the
/// read-only `ensure_video_window_topmost`, and nothing here can split them
/// or hold them off. What bounds the damage is the revalidation below: the
/// window is at most a tick old before it is hidden again and its kill
/// retried. Closing the slice entirely would take the place-and-show split
/// under `PIN_PENDING_RELAUNCH`'s lock, which no caller here can reach
/// through — the show lives behind that read-only call.
pub(super) fn show_current_replacement(pid: u32, content: ScreenRegion) {
    if !pending_relaunch_is_current() {
        let _ = reap_superseded_relaunch();
        return;
    }

    let _ = ensure_video_window_topmost(
        content.0,
        content.1,
        (content.2 - content.0).max(1),
        (content.3 - content.1).max(1),
    );
    // A window of the desktop's that is up is the picture: the relaunch is over. A film drawn through
    // the pin's host is not shown by that call at all — its swap is the thumbnail's, and the record is
    // cleared there (see `pin_film_thumbnail::hand_the_picture_over`).
    if !pin_film_host::the_pinned_player_was_begun_in_the_host() {
        settle_the_published_relaunch(pid);
    }
    revalidate_shown_replacement(pid);
}

/// Re-check a just-shown replacement, and un-show a condemned one.
///
/// This is the half of the gate-to-show race the gate cannot cover (see
/// `show_current_replacement`): a bump landing between the gate and the show
/// leaves a player on screen something else has already condemned. Only the
/// shown pid is ever hidden — where a newer relaunch has already published,
/// `VIDEO_PID` names it and this touches nothing — and the kill is retried
/// rather than dropped.
pub(super) fn revalidate_shown_replacement(pid: u32) {
    if pending_relaunch_is_current() {
        return;
    }
    if pid != 0 && VIDEO_PID.load(Ordering::Acquire) == pid {
        hide_pinned_player_window();
    }
    let _ = reap_superseded_relaunch();
}

/// Whether a relaunch has begun a player the picture has not been handed to yet: the loop's own
/// gate on a film's end and on a failure mark, because a pin across a relaunch is up until the
/// replacement's source takes the picture over (see `pending_relaunch_is_current`).
pub(super) fn a_relaunch_is_in_flight() -> bool {
    pending_pinned_relaunch().is_some()
}

/// Settle what a swap is taking up with: every piece of state that outlives the file it was made
/// against.
///
/// The take-up builds the pin whole, so the old pin's own drag, aim and hold go with it and need
/// nothing said here. Two things are beside the pin rather than in it, and neither is reconciled by
/// a take-up:
///
/// * the in-flight relaunch, whose player is left behind a launch nothing will carry on, and is
///   reaped rather than left for the next run's sweep;
/// * and the box request, because a box request belongs to the file that made it. A resize's end
///   writes the rect the hand settled on and the loop drains it on its next turn, which is a turn
///   away from here — so a step that lands in between found a request standing and replayed it as a
///   `PinBox`, laying the incoming file out at the outgoing one's rect. Nothing else drains it: the
///   loop takes it (see `event_loop`) and there is no other reader, so a request nothing drained
///   before a swap outlived its own file.
pub(super) fn reconcile_swap_take_up() {
    drain_pin_box_request();

    if pending_pinned_relaunch().is_some() {
        let _ = reap_superseded_relaunch();
        if let Some(still) = pending_pinned_relaunch() {
            retire_orphaned_player(still.pid);
            clear_pending_pinned_relaunch();
        }
    }
}

/// Give up the relayout request a resize's end wrote, whoever the file under the pin is by now.
///
/// The loop's own reader takes it on its next turn and there is no other, so a request that
/// outlives the file that made it is replayed against the file that replaced it — the incoming
/// film laid out at the outgoing one's rect (see `reconcile_swap_take_up`).
pub(super) fn drain_pin_box_request() {
    if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
        *request = None;
    }
}

/// Whether a press on the seekbar arms the gesture: only an aimed second does.
///
/// A press with no second to seek to has nothing to move and nothing to take the file to, so the
/// bar claims nothing at all and the pointer is given back rather than held for a release with no
/// aim in it to answer (see `pinned_transport_press`).
pub(super) fn seek_press_arms(seeking: Option<f64>) -> bool {
    seeking.is_some()
}

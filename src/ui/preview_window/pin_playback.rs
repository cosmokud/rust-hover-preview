//! What the pin plays: where the playhead is, seeking it, toggling it, stepping a subtitle,
//! and parking the player's window while a drag is in flight.

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

/// What the transport's own actions need of the pin: the file being played, the box the player's
/// window fills, where its playback is, and the level it is playing at.
pub(super) fn pinned_playback_state() -> Option<(PathBuf, ScreenRegion, PinTransport, PinVolume)> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    Some((pin.path.clone(), pin.content, pin.transport, pin.volume))
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
/// replaced is *taken out of the media* rather than ended here. It is parked — still running, still
/// drawing — until the player begun in its place has a window of its own, which the loop ends it
/// for (see `retire_replaced_player`). Ending it first would leave the pin's band showing the
/// desktop for as long as the new player takes to open the file, seek, and put a window up, which
/// on a 1440p file is long enough to read as a flash, and which is why nothing here goes through
/// `stop_video_playback`: that call is for a player that has ended, and this one has not.
///
/// What *is* stopped is the media's own background work, which belongs to the player being
/// replaced and has nothing to do with whether that player is still there.
/// What is *not* done here is a wait for the new player's window, and the reason is the order of
/// the two players: the one being replaced is parked rather than ended, so that its window is on
/// screen until the one replacing it has a window of its own (see `retire_replaced_player`).
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
) {
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

    // The handle is taken and the process left alone, which is the whole of what "parked" means for
    // the media: nothing else in here can reach the player, so the only thing that can end it is
    // the loop that was told about it.
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.cancel_background_work();
            media.video_process = None;
        }
    }

    let process = start_video_playback(
        path, content.0, content.1, width, height, seconds, volume, subtitle,
    );
    let pid = process.as_ref().map(|child| child.id()).unwrap_or(0);

    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.video_process = process;
        }
    }

    // The player being replaced is parked once the replacement exists, and it is parked *and
    // running* — a relaunch that ended it instead would leave the band empty for the whole of the
    // wait below, and the wait is as long as a player takes to open the file.
    if pid != 0 {
        retire_replaced_player(replaced, pid);
    } else {
        // A replacement that did not come up leaves nothing to arrive, so the player on screen is
        // ended here instead of waiting for an arrival that is never going to happen — which is the
        // road the sweep is still for, and the only road that reaches it.
        kill_stray_video_process();
    }

    // The window the new player puts up is placed by the tick, which re-asserts it every two
    // hundred milliseconds; a seek is one window gone and another arriving, so it is asked for
    // now rather than at the next of those.
    let _ = ensure_video_window_topmost(content.0, content.1, width, height);

    update_pin_transport(|transport| transport.begun(seconds, pid != 0, holding));
    with_pin(|pin| pin.volume.playing_at = volume);

    // A relaunch is also how a drag's parking is undone, because a relaunch brings a player of its
    // own with a window of its own: there is nothing left to unpark, and the flag that would keep
    // the band painted flat has to go with it or the new picture is drawn behind an opaque
    // rectangle. The hold written just above is the relaunch's own, so it is not the one the drag
    // made and there is nothing here to restore (see `park_pinned_player`).
    with_pin(|pin| pin.parked = false);
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
        Some(MediaType::Video) => restart_pinned_player(path, content, seconds, pinned_is_held()),
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
            if transport.pending_hold {
                // The player is playing and no key has reached it, so there is nothing to let go
                // of: what this press does is stop the hold from ever being sent.
                update_pin_transport(|state| state.released(transport.paused_at.unwrap_or(0.0)));
            } else if playing {
                if ffplay_key_pause() {
                    update_pin_transport(|state| state.held(playhead));
                } else {
                    hold_pinned_player_without_a_key(playhead);
                }
            } else if ffplay_key_pause() {
                update_pin_transport(|state| state.released(transport.paused_at.unwrap_or(0.0)));
            } else {
                // A player begun to resume a file is not one to be told to hold, or the press
                // would arrive twice over: once here, and once as a hold the relaunch owed. So
                // the relaunch is told, at its own call site, that this is a file being let go.
                restart_pinned_player(path, content, transport.paused_at.unwrap_or(0.0), false);
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
/// is nothing on screen for a track to change.
pub(super) fn step_pinned_subtitle() {
    if current_media_type() != Some(MediaType::Video) {
        return;
    }

    let Some((path, content, transport, _)) = pinned_playback_state() else {
        return;
    };

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
    );
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
/// so nothing is written down for it; a player this app started cannot be told anything and is
/// ended instead, with the second it had got to kept — on the card, so that the sound is still
/// drawn where the hand left it, and for the player that takes its place when the key is pressed
/// again.
///
/// What this has in common with the bubble's is the mechanics and not the setting:
/// `Pin Mode → Pause Preview` is about what a collapse does, and this is about a key pressed in
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
            // The engine is asked to hold, and to go on from where it is holding. A session that
            // is not there — a file already let go — is asked to play, which begins no session
            // rather than starting a sound nobody asked for.
            video_player::set_paused(video_player::is_playing());
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

/// Put a pinned video's player away for as long as a drag lasts.
///
/// The reason is a measurement rather than a guess. A video pin's picture is FFmpeg's own window
/// standing behind a transparent band, and every pointer move of a resize puts that window to a new
/// size and position (`place_pinned_siblings` -> `ensure_video_window_topmost`). The compositor
/// then has to re-blit a 2560x1440 surface at the pointer's pace, on a 144 Hz screen, while the
/// player is still decoding at whatever rate the machine can manage — which arrives at the hand as
/// a drag that stutters, and arrives at the eye as a picture that is soft and laggy for as long as
/// the hand is on the edge.
///
/// So two things are done about it, by two different callers at two different times, and this is
/// one of them. The window is hidden here, on the drag's own first pointer message, and the film is
/// held by the tick (see `video_drag_hold_apply`), which is the only thing that may pause it: a
/// pause posted from here as well would be a second toggle onto the same film, so a drag begun over
/// a playing film would leave it playing (and its window hidden) and a drag begun over a held one
/// would start it. The split is also the faster of the two orders — the window is hidden on the
/// message the hand sent, not on the tick after it — and the two halves cannot disagree about
/// whether a drag is in flight, because both are asked of `pin.dragging`.
///
/// The band is painted flat while the picture is away rather than left transparent, so the desktop
/// does not show through a window that has been hidden: the paint reads this flag (see
/// `paint.parked`).
pub(super) fn park_pinned_player() -> bool {
    let parked = pin_state().is_some_and(|mut pinned| {
        let Some(pin) = pinned.pin_mut() else {
            return false;
        };

        // A drag that parks twice is one drag, not two: a second call would hide a window that is
        // already hidden and answer a question the first call has already answered.
        if std::mem::replace(&mut pin.parked, true) {
            return false;
        }

        true
    });

    if parked {
        hide_pinned_player_window();
    }

    parked
}

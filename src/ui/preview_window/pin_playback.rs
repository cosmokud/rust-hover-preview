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
    //
    // **But not while a park stands, and that is the second flash this WS is about.** The call is a
    // show as well as a place (see `ensure_video_window_topmost`), and a replacement that has not
    // opened its file yet has a window within milliseconds of being begun — so a relaunch behind a
    // resize's park put an empty window in front of the placeholder for the whole of the decode,
    // and the band was opaque for every one of those milliseconds and did not matter. The settle
    // is what puts a replacement up, in the same tick it takes the flag down (see
    // `settle_pinned_park`), so nothing is left on screen but the frame the drag was holding.
    if !pin_player_is_parked() {
        let _ = ensure_video_window_topmost(content.0, content.1, width, height);
    }

    update_pin_transport(|transport| transport.begun(seconds, pid != 0, holding));
    with_pin(|pin| pin.volume.playing_at = volume);

    // A relaunch is also how a drag's parking is undone, and it is undone by the same question as
    // everywhere else: a player's window has to be there for the band to be a hole again, and the
    // player just begun is running and has no window yet. So this is the settle rather than a write
    // — which is what keeps the band this app's to fill, holding the last frame scaled to the box
    // the drag settled on, for as long as the replacement takes to put a window up
    // (see `settle_pinned_park`).
    settle_pinned_park();
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

/// How long the stale frame may stand in the band before a park that has a window behind it is
/// taken back whatever that window has decoded.
///
/// **It is a bound on a wait, not a delay, and the difference is the whole of what makes it safe.**
/// A park that only ever swapped on a player presenting would hold the placeholder against an
/// empty window for ever — a replacement's window is published within a few milliseconds and stays
/// empty until the file is open and the first frame is decoded — which on a dark scene, a player
/// that died mid-relaunch, or a machine with no decoder is a black band that never becomes a
/// picture. The bound ends that, and it is read from the moment the park was *asked to give the
/// band back* rather than from the park itself, because a hand can rest on an edge for minutes
/// (see `PinParkSwap`).
///
/// **It does not end a park that has no window behind it, because there is nothing there to give
/// the band to.** A band with a player standing in it is transparent and shows that player through;
/// the same band with nothing behind it is the desktop, which is the hole this whole arrangement
/// exists not to leave — so an expired bound with no window holds rather than swaps, and the
/// placeholder it is holding is opaque. What the bound buys there is nothing, and what it would
/// cost is the picture (see `park_swap_arm`).
///
/// 600 ms is the width of what it covers rather than a round number: a player begun on a resize
/// release has to open the file, seek, decode hardware and present, and on a 1440p HEVC that is
/// the few hundred milliseconds the placeholder is held for. Holding the frame the user was already
/// looking at for that long is not a cost — it is the same picture, and it is the alternative to a
/// black band in front of the desktop for exactly as long.
pub(super) const PIN_PARK_SWAP_TIMEOUT: Duration = Duration::from_millis(600);

/// What a park is holding, and since when it was asked to let it go.
///
/// `player` is the process the band was captured for, read once on the pointer message so that a
/// relaunch begun afterwards is a *fact* rather than a guess (see `player_replaced_since`).
/// `replacing` is what a resize's park already knows: a move has no relaunch behind it, so the
/// player standing in the band when the drag ends is the player that was playing all along, while a
/// resize ends in one that has decoded nothing yet. `awaiting_relaunch` is what a seek's cover
/// does not know yet: no relaunch has been begun behind it — the release makes it — so there is
/// nothing to swap to and no bound to spend until one is (see `park_swap_arm_for_the_band`).
/// `since` is left unset until the loop first asks for the band back — armed by the park so a tick
/// in the middle of a drag cannot start the clock, stamped by the settle so a ten-second drag does
/// not spend the budget of a one-tick wait.
#[derive(Clone, Copy)]
pub(super) struct PinParkSwap {
    pub(super) player: u32,
    pub(super) replacing: bool,
    pub(super) awaiting_relaunch: bool,
    pub(super) since: Option<Instant>,
}

/// The one park's bookkeeping, behind a lock because the two ends are not the same thread: it is
/// written on the pointer message that begins a drag and read by every tick of the loop after it.
pub(super) static PIN_PARK_SWAP: Lazy<Mutex<Option<PinParkSwap>>> = Lazy::new(|| Mutex::new(None));

/// Which of the two arms a park's end was taken on.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum ParkSwap {
    /// The player in the band is the one the park captured for: a drag holds its player rather than
    /// replacing it, so that window has the picture the band is holding on it, and there is nothing
    /// to wait for.
    Presented,
    /// The wait is up and there *is* a window to hand the band to, but it has decoded nothing — the
    /// placeholder goes rather than an empty window standing in for a frame that may never come, and
    /// a placeholder held against one for ever is a black band that is never a picture.
    TimedOut,
}

/// What the band is handed back on, from the two facts that can be read of a player and the one
/// that can only be waited for.
///
/// **The `Presented` arm is a fact about the process rather than a guess about the pixels**, and the
/// reason is the one that made this WS necessary: a player that has just been begun publishes its
/// window within a few milliseconds, and that window is visible, correctly sized, and *empty* — so
/// "visible and correctly sized" is true for the whole of the decode and answers nothing. What is
/// not true for the whole of the decode is the process behind it: a move's park is answered by the
/// player that was playing before the pointer went down, which a drag holds at its first pointer
/// message and which therefore has a picture of its own — the one the band is holding, and the one
/// it will show again the instant it is shown — whereas a resize's park is answered by a player
/// that opened the file after the drag began (see `video_drag_hold_apply`).
///
/// **No window at all is not a third arm, and that is what this WS is for.** The band is
/// transparent while a film is playing, so handing it back with nothing behind it is the desktop —
/// and a replacement that has not published a window within the bound is a player whose window
/// `ensure_pinned_sibling_box` cannot find, so the swap would take the parked flag down and show
/// nothing: a hole in the shape of a video, for as long as the user looks at it. So the bound is
/// not a reason to swap on its own; it is the bound *on a wait for a window that has arrived*, and
/// a park with no window keeps the placeholder and is asked again on the next tick. The wait is
/// extended rather than restarted, so a window that turns up after the bound has gone by is handed
/// the band on the first tick that finds it (see `park_swap_arm_for_the_band`).
///
/// **The pixel sample the timeout is paired with is not in here, and cannot be.** It has to be read
/// off the screen inside the player's rect, and the placeholder is what stands there — a layered
/// window's own pixels are what the band is, and the player behind them is not on the screen to be
/// read until the band goes transparent. Reading the player's own DC instead does not work either:
/// FFmpeg's window draws through D3D, whose contents are not in the window's DC at all, which is
/// the same reason `hold_video_window_frame` reads the desktop with `BitBlt` rather than asking the
/// window with `PrintWindow`. So the honest form of "presented or wait" is "the player that was
/// playing is back, or the wait is up and there is a window to show" — and the wait is what
/// bounds a replacement, which is the only case there was ever anything to wait for.
pub(super) fn park_swap_arm(window_up: bool, replaced: bool, waited: Duration) -> Option<ParkSwap> {
    // Asked first and refused on its own, because it is a fact about there being anything to show
    // at all: everything below is an argument about *when* to show it, and every one of them ends
    // with the flag down and a window up.
    if !window_up {
        return None;
    }

    if !replaced {
        return Some(ParkSwap::Presented);
    }

    (waited >= PIN_PARK_SWAP_TIMEOUT).then_some(ParkSwap::TimedOut)
}

/// The arm a park's end was taken on, and how long it was waited for.
///
/// It is written down rather than printed because this app has no log: it is a tray application
/// built as `windows_subsystem`, so a `println!` goes nowhere and `OutputDebugStringW` would want a
/// `Win32_System_Diagnostics_Debug` feature this crate does not enable. What the arm is for is the
/// next person to tune `PIN_PARK_SWAP_TIMEOUT` on a machine with a real film on screen, and this is
/// where they read it from — a test drives a settle through [`settle_pinned_park_where`] and reads
/// it back through [`park_swap_last_arm`], and on a machine with a player it is read in a debugger.
pub(super) static PIN_PARK_LAST_ARM: Lazy<Mutex<Option<(ParkSwap, Duration)>>> =
    Lazy::new(|| Mutex::new(None));

/// Write the arm down (see [`PIN_PARK_LAST_ARM`]).
fn note_park_swap(arm: ParkSwap, waited: Duration) {
    PIN_PARK_LAST_ARM
        .lock()
        .unwrap_or_else(|slot| slot.into_inner())
        .replace((arm, waited));
}

/// The last park's arm and how long it was waited for, for a test that drove the settle itself.
#[cfg(test)]
pub(super) fn park_swap_last_arm() -> Option<(ParkSwap, Duration)> {
    PIN_PARK_LAST_ARM.lock().ok().and_then(|arm| *arm)
}

/// Forget the arm a previous test's swap wrote down.
///
/// It is one slot for the whole process, so a test that asserts *no* arm was taken can be answered
/// by a swap that ran before it, and the assertion is then about the order the runner happened to
/// use. Cleared by the tests that read it rather than left to a race.
#[cfg(test)]
pub(super) fn clear_park_swap_arm() {
    if let Ok(mut arm) = PIN_PARK_LAST_ARM.lock() {
        *arm = None;
    }
}

/// What a park did, in the order it did it.
///
/// It exists because the order is the whole of what this WS fixes and nothing else about a park is
/// observable on a machine with no player of this app's: `hide_pinned_player_window` and the paint
/// before it are both calls out of the preview thread with nothing to stand in for them, so a test
/// that wants to know which came first has to be told (see `park_pinned_player`).
#[cfg(test)]
pub(super) static PARK_TRACE: Mutex<Vec<&'static str>> = Mutex::new(Vec::new());

/// Record one step of a park, in the order it happened.
#[cfg(test)]
pub(super) fn trace_park_step(step: &'static str) {
    if let Ok(mut trace) = PARK_TRACE.lock() {
        trace.push(step);
    }
}

/// The same call, and nothing to do: the trace is a test's seam and this build has no tests in it.
#[cfg(not(test))]
pub(super) fn trace_park_step(_step: &'static str) {}

/// What the last park did, in order.
#[cfg(test)]
pub(super) fn park_trace() -> Vec<&'static str> {
    PARK_TRACE
        .lock()
        .map(|trace| trace.clone())
        .unwrap_or_default()
}

/// Forget the last park's steps, so a park is read from its own first message.
#[cfg(test)]
pub(super) fn clear_park_trace() {
    if let Ok(mut trace) = PARK_TRACE.lock() {
        trace.clear();
    }
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
/// **The hide is the last of three steps here, and the order is the point rather than the tidiness.**
/// The band is painted flat while the picture is away rather than left transparent, so the desktop
/// does not show through a window that has been hidden — but a layered window is only changed by an
/// `UpdateLayeredWindow`, so a hide that came before that paint spent the gap between the two with
/// the band's *last painted* pixels on screen, and for a film playing those are transparent ones. The
/// three are therefore capture, paint, hide, all on the message that begins the drag: the capture is
/// a read of what is on screen and there is nothing to read once the window has gone, the paint is
/// what the compositor is handed next, and only then is the window behind it taken away. The same
/// three answer for a resize, which enters through the same call (see `begin_pin_drag`), which is why
/// the flash at the start of a resize needed no mechanism of its own.
pub(super) fn park_pinned_player(hwnd: HWND, at: (i32, i32), resizing: bool) -> bool {
    park_pinned_player_inner(hwnd, at, resizing, resizing, false)
}

/// Put a pinned video's player away for a seek taken from the transport bar.
///
/// The same cover a drag is given, with the same hold: the press holds the
/// film (see `seek_press_hold`), the aim moves silently under the cover, and
/// the release relaunches once carrying that hold. No resume frame, because
/// the playhead is moving and a frame rendered at the second the hand started
/// from is a frame of the wrong second. The relaunch the release makes runs
/// behind this cover and the settle swaps it once (see `settle_pinned_park`).
pub(super) fn park_pinned_player_for_seek(hwnd: HWND, at: (i32, i32)) -> bool {
    park_pinned_player_inner(hwnd, at, true, false, true)
}

/// Whether a press on the seekbar arms the gesture: only an aimed second does.
/// A press with no second to seek to (an unknown length) parks nothing and
/// holds nothing, or the cover stands over a playing film with no relaunch
/// coming to end it.
pub(super) fn seek_press_arms(seeking: Option<f64>) -> bool {
    seeking.is_some()
}

/// Hold a pinned film for a seek taken from the transport bar: the same
/// gesture hold as a drag, taken at the press rather than on the tick after
/// it.
///
/// A seeking player is a held player: audio and video both, for the whole
/// gesture — the cover the press parks is over a silent player, and the
/// relaunch the release makes carries the hold behind it (see
/// `PinTransport::begun`). A press over a paused film holds nothing, and a
/// second press while the aim is still held posts no second key: the pause is
/// a toggle, and a toggle per press is an unpause (see
/// `video_drag_hold_decision`).
///
/// Scrub steps never come through here — the drag only moves the aim — so one
/// press is one key for the whole gesture.
pub(super) fn seek_press_hold() {
    let Some((_, _, transport, _)) = pinned_playback_state() else {
        return;
    };

    if video_drag_hold_decision(true, pin_is_playing(&transport), transport.drag_held)
        != VideoDragHold::Hold
    {
        return;
    }

    if ffplay_key_pause() {
        seek_press_hold_apply(pin_playhead(&transport).unwrap_or(0.0));
    }
}

/// Record the hold a seek press's key has posted: the same write as a drag's,
/// with the aim the press stored kept standing across it.
///
/// `held` clears `seeking`, and without the restore the scrub's own guard
/// (`pinned_transport_drag` answers only while an aim is standing) refuses
/// every step and the release finds nothing to take the file to — the cover
/// the press parked stranded over a held film. So the aim is read back out of
/// the same write rather than left to whatever `held` does with it, which
/// keeps `held` the same answer for every other caller.
pub(super) fn seek_press_hold_apply(at: f64) {
    update_pin_transport(|state| {
        let aim = state.seeking;
        state.held(at);
        state.seeking = aim;
        state.drag_held = true;
    });
}

/// Spend the stamp a scrub's ticks kept, so the swap bound runs from the
/// release's relaunch rather than from the press.
///
/// The stamp is taken by the first tick that finds the park standing, and a
/// scrub holds the cover past the bound without relaunching anything — so
/// without this the first tick after the release finds the bound already
/// spent and swaps onto whatever window merely exists: published within
/// milliseconds of being begun, and empty until the file is open and the
/// first frame decoded. The wait the bound buys is the new player's decode,
/// and only a stamp taken at the relaunch buys it.
pub(super) fn rearm_seek_cover_for_relaunch() {
    let mut held = PIN_PARK_SWAP
        .lock()
        .unwrap_or_else(|swap| swap.into_inner());
    if let Some(swap) = held.as_mut() {
        if swap.awaiting_relaunch {
            swap.since = None;
        }
    }
}

/// What a park begins: a fresh park over a live window, or an extend over a standing one.
///
/// `replacing` is whether a relaunch is behind the cover, and `prepare` whether a frame for
/// that relaunch is wanted off the gesture's own time: a resize's is, a seek's is not, and a
/// move has no relaunch at all. `awaiting` is whether that relaunch is still to come: a seek
/// parks before its release relaunches, so until one is begun there is nothing to swap to and
/// no bound to spend (see `park_swap_arm_for_the_band`).
fn park_pinned_player_inner(
    hwnd: HWND,
    at: (i32, i32),
    replacing: bool,
    prepare: bool,
    awaiting: bool,
) -> bool {
    let begin = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;

        // A begin over a standing park extends it rather than refusing it: a second gesture
        // arriving before the first one's settle — a release-then-instant-regrab, or a second
        // click on the bar — is the same cover held longer, not a second park. Refusing it leaves
        // the in-flight settle's record superseded with nobody owning the swap, which is the
        // frozen placeholder standing over a live player until the next release. So the record is
        // kept current on the same generation chain, while the capture is kept from the first
        // begin, which read the screen while the window was fully visible.
        if std::mem::replace(&mut pin.parked, true) {
            if let Ok(mut swap) = PIN_PARK_SWAP.lock() {
                match swap.as_mut() {
                    Some(record) => {
                        record.player = VIDEO_PID.load(Ordering::SeqCst);
                        record.replacing = record.replacing || replacing;
                    }
                    None => {
                        *swap = Some(PinParkSwap {
                            player: VIDEO_PID.load(Ordering::SeqCst),
                            replacing,
                            awaiting_relaunch: awaiting,
                            since: None,
                        });
                    }
                }
            }
            return Some(true);
        }

        // **Written with the pin held, which is the whole of what keeps a park from being stranded.**
        // The flag and this record are one fact about one park, and they are read on two threads:
        // the flag by every paint and every raise, this record by the tick that ends the park. Two
        // locks and two writes leave a gap, and a park begun in it has its record taken away under
        // it — a band this app paints flat over a paused film, with nothing left that will ever
        // write that record again. An extend over a standing park refreshes the record rather than
        // replacing it, so the write a second begin would have made is still made (see
        // `park_swap_arm_for_the_band`).
        PIN_PARK_SWAP
            .lock()
            .unwrap_or_else(|swap| swap.into_inner())
            .replace(PinParkSwap {
                player: VIDEO_PID.load(Ordering::SeqCst),
                replacing,
                awaiting_relaunch: awaiting,
                since: None,
            });

        Some(false)
    });

    let Some(extended) = begin else {
        return false;
    };

    // An extend keeps the first begin's capture, paint and hide: the window was fully visible
    // then and there is nothing to read once it has gone. Only the generation the settle drains
    // is refreshed, above — and a resize arriving over a move still asks for the frame the
    // replacement upgrades the band with, after giving up the one the first begin asked for.
    if extended {
        trace_park_step("extend");
        if prepare {
            forget_resume_frame();
            prepare_resume_frame();
        }
    } else {
        forget_resume_frame();

        // Taken before the window is hidden, and only when this really is a park: it is a read of
        // what is on screen, so there is nothing to read once the window has gone (see
        // `hold_video_window_frame`).
        trace_park_step("capture");
        hold_video_window_frame();

        // Asked for here rather than at the end of the drag, because the drag is the only time
        // there is: the release has to hand the band back this tick (see `settle_pinned_park`), and
        // a render begun there is a tenth of a second of the frame the hand let go on being
        // replaced by nothing. Only a relaunch asks, because only a relaunch ends in a player that
        // has decoded nothing (see `spawn_video_resume_frame`).
        if prepare {
            prepare_resume_frame();
        }

        // **Painted before the window is hidden, and this is the whole of the first flash.** A
        // layered window is drawn from its own surface and that surface is only replaced by an
        // `UpdateLayeredWindow`, so a park that hid the player and left the paint to the next
        // repaint spent the gap between the two with the band's last painted pixels still on
        // screen — which, for a film playing, are transparent ones. A transparent band with nothing
        // behind it is the desktop: the window the user was dragging out of is what they see, for
        // as long as the compositor takes to get to the next repaint, and a move answers none of
        // them at all until the transport bar's own (see `compose_parked_band`).
        trace_park_step("paint");
        // SAFETY: the handle is the pin's own window, from the same call the pointer message that
        // began this drag arrived on, and the paint is a layered-window blit out of a surface this
        // thread drew for exactly this window. `GdiFlush` is in the same block because what it is
        // waiting on is the same surface.
        unsafe {
            render_pinned_preview_at(hwnd, at.0, at.1);
            // GDI batches, and what reads the surface next is not a GDI call: the compositor has to
            // have the pixels before the window behind them is taken away, or the band that was
            // opaque a frame ago is transparent again.
            let _ = GdiFlush();
        }

        trace_park_step("hide");
        hide_pinned_player_window();
    }

    true
}

/// Ask for the frame this park gives back, on the drag's own time.
///
/// It is asked of the pin rather than of the media because the media's own background work belongs
/// to the player being parked and was stopped when the park began (see `restart_pinned_player`).
fn prepare_resume_frame() {
    let Some((path, content, transport, _)) = pinned_playback_state() else {
        return;
    };
    let Some(at) = pin_playhead(&transport) else {
        return;
    };

    spawn_video_resume_frame(
        path,
        at,
        (content.2 - content.0).max(0) as u32,
        (content.3 - content.1).max(0) as u32,
    );
}

/// Take a park back, answering whether there was one to take.
///
/// **The park ends when there is a picture to see through the band again, and not one moment
/// before.** The band is transparent while a film is playing because FFmpeg's window stands in it,
/// so a park that ends while that window is not there is a hole in the desktop in every band a
/// player of this app's stands in — which is what a relaunch used to leave for the whole of the
/// wait for its replacement, because the flag was written down the moment the replacement was
/// begun rather than the moment it had a window of its own (see `restart_pinned_player`).
///
/// **And it does not end one moment after either, which is the other half of the flash this settle
/// used to leave.** Being visible is not having presented: a replacement's window is on screen and
/// correctly sized within a few milliseconds of being begun, and empty until the file is open, the
/// seek is done and the first frame is decoded — which on a 1440p HEVC is long enough to read as a
/// black band over the desktop at the exact moment the hand let go. So the flag stays down while
/// there is nothing behind it, and what it is waiting for is the process rather than the window
/// (see `park_swap_arm`).
///
/// **And it does not end when the bound goes by with no window standing in it either**, which is
/// the same hole from the other side: the swap is the flag down *and* a window up, and a
/// replacement that has not published one is answered out of the hand by every place that could put
/// it up (see below), so an expiry with nothing behind it takes the flag down and shows nothing.
/// The bound is therefore a bound on a wait for a window that has arrived, and a park with no window
/// keeps the placeholder it is holding — which is opaque, and is the picture the hand let go of.
///
/// **Nothing here goes looking for a window, and the reason is worth stating because it looks like
/// a missing step.** A window of this app's player is published by a monitor thread that walks the
/// desktop and *skips every window it cannot see* (see `enum_windows_callback`), and it refuses to
/// raise at all while a park stands (see `apply_noactivate_to_hwnd`). So a replacement that is
/// being kept off the screen is a window nothing will ever find — which is the point: it is found
/// and shown by the swap below, in the tick the flag comes down, and every tick before that the
/// band is the frame the drag took rather than a window that has decoded nothing.
///
/// The window is put in the band before the flag goes down rather than after it, which is why this
/// is not `ensure_pinned_sibling_box` on its own: that call is answered out of the hand for as long
/// as the park stands, and the tick's own raise is answered out of the hand for the same reason,
/// so the tick that takes the park back is the one that has to place the window.
pub(super) fn settle_pinned_park() -> bool {
    settle_pinned_park_where(
        &Win32PinWindow,
        video_window_for(VIDEO_PID.load(Ordering::SeqCst)).is_some(),
    )
}

/// The settle with the two facts a machine with no player of this app's in it cannot supply given
/// to it: the window the swap acts on, and whether the player's window is standing in the band.
///
/// It is a second function and not two more parameters because those two are the whole of what the
/// loop's own version reaches outside this app for — `video_window_for` walks the desktop and
/// `Win32PinWindow` is a handle to a window this app created — and a settle tested only through the
/// loop's own version is a settle not tested at all. Everything else it reads is this app's own
/// state, which a test stands (see `stand_pin`).
pub(super) fn settle_pinned_park_where(window: &dyn PinWindow, window_up: bool) -> bool {
    if !pin_player_is_parked() {
        forget_the_park_that_is_not();
        return false;
    }

    // A drag in flight is a park doing its job, not a park waiting to be ended: the clock below is
    // stamped by the first tick that finds the drag gone, and a hand resting on an edge for ten
    // seconds must not spend the budget of a one-tick wait.
    if pin_is_dragging() {
        return false;
    }

    let Some((arm, waited)) = park_swap_arm_for_the_band(window_up) else {
        // Still waiting for a window, so the placeholder stays standing — but a frame a background
        // rendered for this park can have landed in the meantime, and a frame nothing repaints is
        // a frame the band is not showing (see `upgrade_the_parked_band`).
        upgrade_the_parked_band(window);
        return false;
    };

    // The swap itself: the flag goes down and the window goes up in the same tick, and the band is
    // painted through the window in that same tick and not before — so the band is never a moment of
    // nothing with a player behind it, and never a frame of the drag's own picture still composited
    // over one (see `unpark_pinned_player`).
    if !unpark_pinned_player(window) {
        return false;
    }

    // A seek's hold ends with the swap rather than with a key: the replacement
    // has been playing since the release relaunched it, so posting one would
    // pause the film the swap just uncovered. The owed hold is taken back and
    // the gesture's claim dropped with it, and a playing film is playing after
    // (see `settle_seek_hold_after_swap`). A drag's hold never reaches here —
    // its own release lets go of it before any relaunch — so a claim standing
    // at a swap is a seek's.
    settle_seek_hold_after_swap();

    // Given up through the same helper the other end uses, and for the same reason: a park begun in
    // the gap between the flag going down and this write has written a record of its own, and this
    // must not take it away.
    forget_the_park_that_is_not();
    note_park_swap(arm, waited);
    true
}

/// End a seek's hold on the tick its swap uncovered, without posting a key.
///
/// The press held the old player with one key, the release carried that hold
/// onto the replacement as owed, and the wait deferred it while the gesture
/// held the claim (see `pending_hold_delivers`) — so the player behind the
/// band has been playing since the relaunch, silently behind the cover, and
/// the only thing left is the bookkeeping: the hold taken back at the second
/// the player has actually reached, which is what `released` writes. No key
/// is posted because there is nothing to toggle: one more would pause the
/// film the swap just uncovered.
///
/// A paused film has no claim — its press held nothing — so it keeps its hold
/// and its owed flag for the loop to deliver through the new window, and
/// stays paused (see `settle_pending_hold`).
pub(super) fn settle_seek_hold_after_swap() {
    let ours = pin_state()
        .and_then(|pinned| {
            pinned
                .pin()
                .map(|pin| (pin.transport.drag_held, transport_clock(&pin.transport)))
        })
        .filter(|(held, _)| *held);

    let Some((_, at)) = ours else {
        return;
    };

    update_pin_transport(|state| state.released(at.unwrap_or(0.0)));
}

/// Put the frame a background rendered for this park in place of the one the park captured, and
/// repaint the band that is holding it.
///
/// **The repaint is the whole of what is asked of the window here, and it cannot be left to the
/// next one.** A park still standing has no drag in flight to repaint it — the drag's own repaints
/// are its pointer moves, and those are over — and the tick's own raise is answered out of the hand
/// while a park stands, so an upgrade nothing paints is a frame that is held and never shown. The
/// paint is the ordinary one for a window standing where it is, and it reads the band as a parked
/// band, so the frame goes up under the same opaque fill the park painted with it (see
/// `compose_parked_band`).
fn upgrade_the_parked_band(window: &dyn PinWindow) -> bool {
    if !install_resume_frame() {
        return false;
    }

    window.repaint();
    true
}

/// The arm this band's park is to be taken on, and how long it has been waiting for one — or
/// nothing at all while it is still to be waited for.
///
/// **The stamp is taken once and kept, so a park that outlives its bound is waiting on the rest of
/// its wait rather than starting a new one.** That is what makes an expired park with no window
/// cheap rather than a loop: the next tick that finds a window hands it the band at once, whether
/// it arrives a millisecond after the bound or a second after it, and the ticks in between cost
/// one comparison (see `park_swap_arm`).
fn park_swap_arm_for_the_band(window_up: bool) -> Option<(ParkSwap, Duration)> {
    // Read through a poisoned lock rather than refused by it: this is the function that decides
    // whether the band is ever handed back, and a park that no tick will end is a black band for
    // the rest of the pin's life.
    let mut held = PIN_PARK_SWAP
        .lock()
        .unwrap_or_else(|swap| swap.into_inner());

    // A seek's cover is not a wait for a window yet: its release has not relaunched, so there is
    // nothing to swap to — and the bound below runs from the relaunch, so a scrub held past it
    // must neither spend it nor swap to the player the cover still stands over. Held until a
    // relaunch is begun behind it, which the release does (see `seek_pinned_playback`).
    if held
        .as_ref()
        .is_some_and(|swap| swap.awaiting_relaunch && !player_replaced_since(swap.player))
    {
        return None;
    }

    // **A park with no record of itself is a park with no replacement to wait for**, and that is
    // what it is answered as: there is nothing behind it that this app began and nothing that is
    // going to arrive, so it is handed to whatever window is standing in the band on the first tick
    // that finds one. The alternative — the refusal a missing record used to be — is the stuck
    // placeholder: the band stays painted flat over a film nothing is going to be shown behind, and
    // nothing will ever write the record that could have ended it.
    let (replaced, waited) = held
        .as_mut()
        .map(|swap| {
            let since = *swap.since.get_or_insert_with(Instant::now);
            (
                swap.replacing || player_replaced_since(swap.player),
                since.elapsed(),
            )
        })
        .unwrap_or((false, Duration::ZERO));

    park_swap_arm(window_up, replaced, waited).map(|arm| (arm, waited))
}

/// Whether the standing park is a seek's cover still waiting for its release's relaunch.
///
/// A scrub aims without relaunching, so until one is begun behind the cover there is nothing
/// for the settle to swap to; a gesture that ends without relaunching one — a move let go of, a
/// capture stolen mid-aim — hands the band back to the player it covers instead of leaving a
/// cover nothing ends (see `finish_pin_drag` and `pin_capture_lost`).
pub(super) fn seek_cover_is_waiting() -> bool {
    PIN_PARK_SWAP
        .lock()
        .ok()
        .and_then(|held| {
            held.as_ref()
                .map(|swap| swap.awaiting_relaunch && !player_replaced_since(swap.player))
        })
        .unwrap_or(false)
}

/// Give up the record of a park there is no longer, with the pin held.
///
/// **The pin is held across the write, and that is the other half of the park and its record being
/// one fact.** A park writes the flag and the record together under the pin's lock (see
/// `park_pinned_player`), so a park begun while this is deciding has either written both already —
/// and this sees the flag up and leaves the record alone — or writes both after this has released
/// the lock. Read without it, a park begun in the gap between the read and the write loses the
/// record it had just written, and a park that has lost its record is a placeholder nothing ends.
fn forget_the_park_that_is_not() {
    let Some(mut pinned) = pin_state() else {
        return;
    };
    let Some(pin) = pinned.pin_mut() else {
        return;
    };
    if pin.parked {
        return;
    }

    forget_pin_park_swap();
}

/// Put down everything a park was holding: its bookkeeping, and the frame the background was
/// preparing for it.
///
/// **Given up with the pin held wherever it is asked for from the loop**, which is what
/// `forget_the_park_that_is_not` is for; it is this function because a test standing a park needs to
/// give it up without a pin's lock in the way.
pub(super) fn forget_pin_park_swap() {
    PIN_PARK_SWAP
        .lock()
        .unwrap_or_else(|swap| swap.into_inner())
        .take();
    forget_resume_frame();
}

/// The park taken back when `player_up` says the player's own window is standing in the band, or
/// left standing when it does not: the flag and the window are two facts and the flag is only ever
/// this app's own account of the window (see `settle_pinned_park`).
///
/// The frame goes with the park rather than with the drag, because it is the band's picture for as
/// long as the band has no window in it and not one tick longer (see `forget_video_frame`).
pub(super) fn settle_pinned_park_onto(player_up: bool) -> bool {
    let settled = pin_state().is_some_and(|mut pinned| {
        pinned.pin_mut().is_some_and(|pin| {
            pin.parked && player_up && {
                pin.parked = false;
                true
            }
        })
    });

    if settled {
        forget_video_frame();
    }

    settled
}

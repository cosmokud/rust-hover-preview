use super::*;

/// a pinned video laid out again for the box its window was dragged to, which /// is where a resize used to flash a sheared picture for the moment before the engine handed /// the next frame over.
#[test]
fn a_resized_pinned_video_keeps_the_frame_its_pixels_are() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(8, 4);
    video.media_type = MediaType::NativeVideo;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let path = std::env::temp_dir()
        .join("rust-hover-preview-video-tests")
        .join("resized-pin.mp4");

    // A box far larger than the frame, which is what a window dragged out to the screen is.
    relayout_pinned_media(&path, (0, 0, 800, 600), 96, None);

    let frame = {
        let media = CURRENT_MEDIA
            .lock()
            .expect("the media is where it was left");
        let frame = media
            .as_ref()
            .and_then(|media| media.frames.first())
            .expect("the frame is where it was left");
        (frame.width, frame.height, frame.pixels.len())
    };

    assert_eq!(
        frame,
        (8, 4, 8 * 4 * 4),
        "the frame is the size its pixels are, whatever box the window has been dragged to"
    );

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous;
    }
}

/// a seek made while a pinned video is paused.
#[test]
fn a_seek_made_while_a_pinned_video_is_paused_moves_the_second_it_is_drawn_at() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::NativeVideo;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let path = std::env::temp_dir()
        .join("rust-hover-preview-video-tests")
        .join("paused-seek.mp4");

    // Twice: a bar dragged while the file is held, and the same drag while it is playing —
    // only the first of those is drawn from a second this side wrote down.
    let cases: [(Option<f64>, Option<f64>); 2] = [(Some(30.0), Some(90.0)), (None, None)];

    for (paused_at, wanted) in cases {
        stand_pin(Some(PinnedPreview {
            path: path.clone(),
            bound: Some(320),
            content: (0, 0, 320, 240),
            restore: None,
            dpi: 96,
            transport_bar: true,
            transport_live: true,
            frame: PinFrame::Shaped,
            overlay: true,
            hides_chrome: true,
            caption: pinned_caption_height(96, None),
            chrome: PinChrome::on_arrival(Instant::now()),
            collapsed: false,
            bubble_pause: None,
            hovered: None,
            pressed: None,
            tooltip: PinTooltip::default(),
            dragging: None,
            parked: false,
            transport: PinTransport {
                duration: Some(120.0),
                paused_at,
                ..Default::default()
            },
            volume: PinVolume::default(),
            audio_hovered: None,
            audio_pressed: None,
        }));

        seek_pinned_playback(&path, (0, 0, 320, 240), 90.0);

        let transport = pin_state().and_then(|state| {
            state
                .pin()
                .map(|pin| (pin.transport.paused_at, pin.transport.seeking))
        });

        assert_eq!(
            transport,
            Some((wanted, None)),
            "a paused file is drawn at the second it was taken to, and a drag that is over is \
                 a drag that is over"
        );
    }

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A transport bar begun at a second, and the claim it makes about a player.
fn transport_beginning_at(from: f64) -> PinTransport {
    let mut transport = PinTransport::default();
    transport.begun(from, true, false);
    transport
}

#[test]
fn a_player_begun_where_the_film_was_watched_from_is_the_one_the_bar_is_drawn_against() {
    let transport = transport_beginning_at(90.0);

    assert!(
        transport_playing(&transport, true),
        "a player that has been begun is playing until something says otherwise"
    );
    assert_eq!(
        transport.started.map(|(_, from)| from),
        Some(90.0),
        "and the bar's clock starts from the second it was told to begin at rather than from \
             the beginning, so a seek does not silently rewind what is on screen"
    );

    // A relaunch that did not come up is still the bar's reference: nothing is playing, but
    // the second is remembered so that the wait does not throw the position away.
    let mut failed = PinTransport::default();
    failed.begun(90.0, false, false);
    assert!(
        !transport_playing(&failed, true),
        "a player that never came up is not one the bar may claim is playing"
    );
    assert!(
        failed.started.is_none(),
        "and nothing is left claiming a clock for a player that is not there"
    );
}

#[test]
fn a_hold_keeps_the_player_and_its_clock_so_a_resume_is_a_key_rather_than_a_start() {
    let mut transport = transport_beginning_at(0.0);

    transport.held(12.5);
    assert_eq!(
        transport.paused_at,
        Some(12.5),
        "a held file is drawn at the second it was stopped at"
    );
    assert!(
        !transport_playing(&transport, true),
        "and the bar stops claiming that it is playing"
    );
    assert!(
        transport.started.is_some(),
        "the player behind it is still the one that is running: a hold is a key, not a kill, \
             and the clock under it is what a resume continues from"
    );

    transport.released(12.5);
    assert_eq!(
        transport.paused_at, None,
        "letting go of the file clears the hold"
    );
    assert!(
        transport_playing(&transport, true),
        "and the same player is playing on from the second it was held at"
    );
    assert_eq!(
        transport.started.map(|(_, from)| from),
        Some(12.5),
        "with the clock re-based onto that second rather than left where it was: the position \
             of a file this app's player is playing is this app's own clock over a moment it began \
             the player, and a clock left running through a hold would make the playhead jump \
             forward by the length of the pause the instant the file was let go"
    );
}

#[test]
fn a_player_that_died_while_the_bar_believed_it_was_playing_leaves_no_stuck_glyph() {
    let mut transport = transport_beginning_at(0.0);

    transport.player_gone(41.0);
    assert!(
        !transport_playing(&transport, true),
        "nothing is playing a file whose player has gone, whatever this side wrote down"
    );
    assert_eq!(
        transport.paused_at,
        Some(41.0),
        "but the second it had reached is kept, so the bar still shows where the film stopped \
             and a press starts it from there rather than from the beginning"
    );
    assert_eq!(
        transport.started, None,
        "and the clock of a player that is not there is given up rather than left running"
    );
}

#[test]
fn a_file_that_dies_while_it_is_held_stays_a_hold_rather_than_becoming_a_play() {
    let mut transport = transport_beginning_at(0.0);
    transport.held(7.0);

    transport.player_gone(7.0);
    assert!(
        !transport_playing(&transport, true),
        "a hold and a dead player both draw a play button: neither is a file that is playing"
    );
    assert_eq!(
        transport.paused_at,
        Some(7.0),
        "and the second it was held at is the one a press would begin it from, because that is \
             where the film stopped rather than where it was taken"
    );
}

#[test]
fn a_seek_taken_while_a_file_is_held_moves_the_hold_with_it() {
    // Held: the bar is drawn from a second this side wrote down, so it has to move.
    let mut held = transport_beginning_at(0.0);
    held.held(30.0);
    held.sought(90.0);
    assert_eq!(
        held.paused_at,
        Some(90.0),
        "a held file is drawn at the second it was taken to, and a drag that is over is a \
             drag that is over"
    );

    // Playing: the position is the player's own clock, which the seek has already moved by
    // beginning a player at the new second — writing a second down here would freeze the bar
    // at where the hand let go.
    let mut playing = transport_beginning_at(0.0);
    playing.sought(90.0);
    assert_eq!(
        playing.paused_at, None,
        "a seek taken while the file is playing writes no position of its own; the player that \
             was begun at the new second is what the bar is measured against"
    );
}

/// The same seek on a video FFmpeg plays, which is not `sought` at all: there is no player
/// to move, so a seek is a player ended and begun again at the second the bar was let go at,
/// and a hold is something the relaunch has to be *told* about rather than something the
/// seek carries out for itself.
#[test]
fn a_seek_through_ffplays_player_comes_back_held_because_the_relaunch_owed_a_hold() {
    // A held file, seeked to ninety seconds: the bar's own answer to that is the second it
    // was taken to, and the file is still meant to be held there.
    let mut transport = transport_beginning_at(0.0);
    transport.held(30.0);
    transport.begun(90.0, true, true);

    assert_eq!(
        transport.paused_at,
        Some(90.0),
        "a relaunch of a held file is begun at the second it was taken to and held there, or a \
             seek made while the file was paused would start it playing — which is what it did \
             before a hold was carried across the relaunch at all"
    );
    assert!(
        transport.pending_hold,
        "and the hold is left owing, because a player that has only just been begun has no \
             window for the key that holds it to be posted to"
    );
    assert!(
        !transport_playing(&transport, true),
        "so the bar draws a play button from the moment of the seek rather than a pause glyph \
             over a film that is playing for the tick it takes to be told otherwise"
    );

    // And the loop's delivery: the second is written from the player's own clock, not from
    // the one the relaunch began at, so the film is not left drawn a tick behind where the
    // player actually stopped.
    let stopped = transport_clock(&transport).expect("a begun player has a clock");
    transport.held(stopped);
    assert!(
        !transport.pending_hold,
        "a player that has been told to hold owes nothing any more"
    );
    assert_eq!(
        transport.paused_at,
        Some(stopped),
        "and it is held at the second it reached, which is at or a tick past the second the \
             relaunch began at and never behind it"
    );
    assert!(
        transport.started.is_some(),
        "with the player still the one behind the bar, so the resume is a key rather than a \
             second start"
    );

    // A press of play takes the owed hold back rather than racing it: the file goes on from
    // where it was, and a pause that had not been delivered must not arrive afterwards.
    let mut resumed = transport;
    resumed.released(transport.paused_at.unwrap_or(0.0));
    assert!(
        !resumed.pending_hold && resumed.paused_at.is_none(),
        "letting go of a file whose hold was still owed cancels the hold, because the hand has \
             asked for the film to play and a pause arriving afterwards would be a press answered \
             by the opposite of what was asked"
    );
    assert!(
        transport_playing(&resumed, true),
        "and the file is playing from the second it was held at, off a clock rebased onto it"
    );

    // A player that died takes the owed hold with the claim it belonged to, so the flag can
    // never outlive the player it was waiting for.
    let mut gone = transport;
    gone.player_gone(95.0);
    assert!(
        !gone.pending_hold,
        "there is no player left to owe a hold to, so nothing is left owing"
    );
}

#[test]
fn a_bar_never_claims_a_file_is_playing_once_the_player_behind_it_is_gone() {
    let transport = transport_beginning_at(0.0);

    assert!(
        transport_playing(&transport, true),
        "a player that is there and was begun at a second is a file that is playing"
    );
    assert!(
        !transport_playing(&transport, false),
        "liveness wins over every claim this app made about what the player was doing: a \
             player that has gone cannot be playing a file, however recently it was begun"
    );
    assert!(
        !transport_playing(&PinTransport::default(), true),
        "and a bar that never claimed a player cannot claim one is playing"
    );

    // A file this app paused by ending its player is the same case from the other side: the
    // hold is written, so the bar says held rather than playing, and liveness is not even
    // asked — which is what keeps a key that could not be reached from leaving the wrong glyph.
    let mut ended = transport_beginning_at(0.0);
    ended.player_gone(3.0);
    assert!(!transport_playing(&ended, true));
}

#[test]
fn a_relaunch_keeps_the_subtitle_track_it_was_asked_for() {
    let mut transport = PinTransport {
        subtitle: Some(1),
        ..PinTransport::default()
    };

    // Everything a relaunch does to a bar that is not about the track itself: it begins a
    // player, gives up a drag in progress, and — where the file was held — writes the hold
    // down again as owed by the player it began. The track is not among them, because the
    // track is the one thing about the file the user chose and a seek is not a reason to
    // take it away.
    transport.begun(90.0, true, false);
    assert_eq!(
        transport.subtitle,
        Some(1),
        "a player begun again at a new second is told which track to show, so the choice \
             survives the seek that caused it"
    );

    // And it survives a resize and a hold and a player that died, for the same reason.
    transport.held(90.0);
    assert_eq!(transport.subtitle, Some(1), "a hold keeps the track");
    transport.player_gone(95.0);
    assert_eq!(
        transport.subtitle,
        Some(1),
        "and a player that died keeps it too, so the film resumes on the track it was on"
    );
}

#[test]
fn a_replaced_player_is_ended_once_there_is_a_window_to_replace_it_with() {
    let waited = Duration::from_millis(200);

    assert!(
        retire_ready(true, true, waited),
        "a replacement with a window of its own is a replacement that has arrived, and the \
             player under it has served its purpose"
    );
    assert!(
        retire_ready(false, false, waited),
        "a replacement that is gone never will be, so waiting longer leaves the old player \
             playing over a file nothing is going to show"
    );
    assert!(
        retire_ready(false, true, Duration::from_secs(VIDEO_START_WAIT_SECS)),
        "a replacement that has been starting for longer than a start ever takes is one this \
             app has already given up on once today, and the wait is bounded by the same number"
    );
    assert!(
        !retire_ready(false, true, waited),
        "and a replacement still on its way is not an arrival: ending the only picture there \
             is now is the flash of the desktop that a relaunch has to avoid"
    );
}

/// A retired player is the one process this app holds that no other check reaches: the handle
/// went when it was parked and `VIDEO_PID` names the replacement, so the call that ends it is
/// the call that has to stop holding it — once its death is confirmed, and not on the request.
#[test]
fn a_retired_players_id_is_given_up_only_once_the_end_is_confirmed() {
    assert_eq!(
        retire_end(true),
        RetireEnd::Asked,
        "an end that has not taken is not an end, and a player this app has stopped holding is \
             asked for by nothing else: it stays parked, and its id stays held, until it is gone"
    );
    assert_eq!(
        retire_end(false),
        RetireEnd::Gone,
        "and a player confirmed gone is the end of the wait, so the record of it goes with it \
             rather than being left in the state file for the next run to pass over"
    );

    // The half of it that is reachable without a player to end: a retirement settled against a
    // process that is not there is the whole of the leak, once per relaunch.
    assert_eq!(
        end_retired_player(u32::MAX),
        RetireEnd::Gone,
        "a retired player that has already gone is confirmed gone the moment it is asked to \
             end, which is the path every settled retirement takes"
    );
}

/// The knob let go of on a pinned FFmpeg video: whether the level costs a player at all.
///
/// A film that is *playing* owes the level to a player, because FFmpeg's player takes one only by
/// being started at it, and that player is begun playing (see `restart_pinned_player`).
///
/// A film that is *held* owes it to nobody yet, and the reason is a measurement rather than a
/// preference: ffplay 9.0.2 has no way to be started paused — its whole option list was read for
/// one and there is none — so a relaunch of a held film is a player that plays, audibly at the new
/// level, from the second it was begun at, until the hold written beside it reaches it. That is one
/// `P` posted on the first tick that finds a window there, which is hundreds of milliseconds after
/// the process was spawned: a held film answers a knob with a burst of its own soundtrack at the
/// second the hand stopped it at, and it stops again a fraction later. The level is therefore
/// written down against the player that *begins when the file is let go of*, which is the same
/// answer a pin with no player behind it has always had, and the player that is holding now is left
/// alone — which also leaves the app's own loop record alone, so a knob can no longer arm the
/// rewind that is a relaunch at zero (see `note_video_loop`).
#[test]
fn pin_level_settled_by_owes_a_level_to_a_playing_player_only() {
    assert_eq!(
        pin_level_settled_by(true, false, false),
        PinLevelSettling::Relaunch,
        "a player that is playing takes a level only by being begun at one, so it is replaced, \
             and a file that was not held is begun playing"
    );
    assert_eq!(
        pin_level_settled_by(false, true, true),
        PinLevelSettling::OwedToTheNextPlayer,
        "a held file must not be relaunched for a level: the replacement cannot be started paused, \
             so it plays audibly at the new level from the second the hold was taken at until the \
             hold reaches it — the loop-back a knob on a held film used to answer with"
    );
    assert_eq!(
        pin_level_settled_by(false, false, true),
        PinLevelSettling::Recorded,
        "and a pin with neither a claim nor a player behind it is owed nothing, held or not: the \
             level is written down and the player that begins when the file is let go of takes it \
             then"
    );
}

/// The other half of the same fix: a level left owed to the player that begins next is worth
/// nothing unless that player refuses a pause key, because a key is the cheap way to let a film go
/// of its hold and it keeps the player that is at the level the hand moved away from.
///
/// The facts it is refused over are refusals rather than permissions, and each is a different
/// mistake. A file that is playing is a file the key is being used to stop, where ending its player
/// instead would take a picture away to answer a pause. A hold that has not reached its player yet
/// is a key that is *being delivered*, which a level turned in the meantime must not cancel. And a
/// hold that is a gesture's is the end of that gesture rather than a press of the pause button's.
#[test]
fn a_held_player_at_a_level_the_pin_has_moved_on_from_cannot_be_let_go_of_with_a_key() {
    assert!(
        level_is_owed_to_the_next_player(true, false, false, 40, 80),
        "a held file whose player was begun at 40 with the bar drawn at 80 owes the level to the \
             player that begins when the file is let go of — and letting it go of the hold with a \
             key would start the film at 40, which is the level the hand moved away from"
    );
    assert!(
        !level_is_owed_to_the_next_player(true, false, false, 80, 80),
        "a held file whose player was begun at the level the bar is drawn at is owed nothing, and \
             the key is the whole of letting it go of the hold: the player that never stopped stays"
    );
    assert!(
        !level_is_owed_to_the_next_player(false, false, false, 40, 80),
        "a file that is not held has this key pressed on it to stop it, so the answer has to be \
             there: ending the player instead would leave the pin's band showing the desktop"
    );
    assert!(
        !level_is_owed_to_the_next_player(true, false, true, 40, 80),
        "and a hold that has not reached its player yet is not a player holding at any level: it is \
             a player playing, waiting for exactly this key, which is what delivers it"
    );
    assert!(
        !level_is_owed_to_the_next_player(true, true, false, 40, 80),
        "a hold that is a gesture's is the end of that gesture rather than a press of the pause \
             button's: a window dragged for a second and released has to carry on from the second \
             it was watching, and a player begun again is the one way not to do that"
    );
}

/// A loop this app gives is posted short of the end of the file rather than at it, and a file
/// nothing has measured is never rewound at all.
///
/// The margin is the whole of why a posted rewind works: a player that has reached the end of
/// its file has already closed its window, so a rewind posted *at* the end is a no-op on a
/// window on its way out.
#[test]
fn a_loop_this_app_gives_is_rewound_short_of_the_end_and_only_where_it_can_be() {
    assert!(
        video_launch::rewind_due(true, Some(8.0), 7.95),
        "a player within the margin of a measured end is sent back before it reaches it"
    );
    assert!(
        !video_launch::rewind_due(true, Some(8.0), 2.0),
        "a player in the middle of the file is left playing it"
    );
    assert!(
        !video_launch::rewind_due(false, Some(8.0), 7.95),
        "a held film is not rewound: a hold that jumped to the beginning would be a hold that \
             did nothing"
    );
    assert!(
        !video_launch::rewind_due(true, None, 7.95),
        "a file whose length nothing has read has no end to be near, and a player begun at zero \
             of it has nothing to be saved from"
    );
    assert!(
        !video_launch::rewind_due(true, Some(0.05), 0.04),
        "a file shorter than the margin is every second inside it, so rewinding it would send \
             the player back the instant it began"
    );
}

/// Only a player that was seeked has its loop given by this app.
///
/// Measured against FFmpeg 9.0.2 on an eight-second clip: `-loop 0` alone wraps from `0` back
/// to `0`, while `-loop 0` with `-ss 6` wraps from `5.208` back to `5.208` — the last two and
/// a half seconds of the film, for ever, and the first six seconds never seen. The seek is
/// what breaks the loop; writing it after the input rather than before does not fix it.
#[test]
fn only_a_player_that_was_seeked_has_its_loop_given_by_this_app() {
    assert!(
        !video_launch::loop_is_ours(false),
        "a hover begins at zero, and a player wrapping to the beginning by itself is also what \
             keeps a hover from closing its own preview at the last frame of the file"
    );
    assert!(
        video_launch::loop_is_ours(true),
        "a player begun at a second wraps back to that second, so the loop is taken away from \
             it and given by the rewind instead"
    );
}

/// Subtitles are drawn by a filter in the same chain as the crop, and the file is named the
/// one way the filter parses.
///
/// The escaping is measured in both directions and only one spelling survives. The `subtitles`
/// filter separates its filename from its options with a colon, and a Windows path opens with
/// one, so the drive letter is read as a filename and the rest as the filter's first option —
/// which is why FFmpeg's complaint names `original_size`, which nobody asked for. The whole of
/// the escaping, the characters beyond the colon and the one that cannot be escaped at all, is
/// in `video_launch` beside it; what is asserted here is that the track the pin is remembering
/// reaches the filter, which is the half that was missing.
#[test]
fn subtitles_are_named_the_way_the_filter_that_reads_them_parses() {
    assert_eq!(
        video_launch::escape_filter_path(&PathBuf::from(r"D:\video\clip.srt")),
        "D\\:/video/clip.srt",
        "the colon after the drive letter is the filter's own separator, and the separators \
             themselves need no escaping because forward slashes are what a Windows path API takes"
    );

    let dir = std::env::temp_dir().join(format!("preview-subs-{}", std::process::id()));
    std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
    let video = dir.join("film.mkv");
    std::fs::write(&video, b"stand-in").expect("a stand-in file is writable");

    let filter = video_launch::subtitle_filter(&video, 2, Some(1))
        .expect("a file with a subtitle stream of its own has a filter to draw it with");
    assert!(
        filter.starts_with("subtitles='") && filter.contains(":si=1"),
        "the value is quoted so the escape survives the filtergraph parser, and the track is the \
             one the pin is remembering rather than a hard-coded first: `si=2` on a third stream is \
             refused outright with 'Unable to locate subtitle stream', so the index has to be \
             counted among subtitle streams and it has to be the one that was chosen: {filter}"
    );

    let _ = std::fs::remove_dir_all(&dir);
}

/// What the probe found is readable by the reader that asked for it, once the probe has finished.
///
/// This is the regression test for an answer that was `None` for the whole run of the app and
/// looked like a perfectly good answer the entire time. The answer was a `OnceLock` completed by
/// the thread that *reads* it: `get_or_init` stores `None` before the probe thread has run a
/// line, so the probe's own `set` was refused every time and `-hwaccel` was never passed. Every
/// film in the run was software decoded and nothing said so.
///
/// The test that could not see it compared the answer with itself — `None == None` — which is why
/// it passed on the broken arrangement and would pass again. So this asserts the thing that was
/// actually broken: an answer written *after* a reader has already read is still readable by a
/// later reader. Reading before the answer lands gives nothing, reading after gives what the
/// probe found, and reading twice gives the same thing both times.
#[test]
fn what_the_probe_finds_is_readable_once_the_probe_has_finished() {
    static PROBE: HwAccelProbe = HwAccelProbe {
        asked: AtomicU32::new(0),
        current: AtomicU32::new(1),
        found: Mutex::new(None),
    };

    assert_eq!(
        PROBE.read(|| Some("dxva2")),
        None,
        "a reader that arrives before the probe has run must be given nothing rather than made \
             to wait: this runs on the preview thread, and a preview begun before the answer lands \
             is software decoded, which is the answer that works on every machine"
    );

    // Waited for rather than joined, and the difference is the whole of the arrangement being
    // tested: the probe thread finishes *by writing*, so a join would return while the answer
    // was still in flight — which is precisely the window in which the old arrangement dropped
    // it on the floor. The wait is bounded, so a probe that never lands fails here rather than
    // hanging the suite.
    let deadline = Instant::now() + Duration::from_secs(5);
    let mut answered = PROBE.read(|| unreachable!("the question is put once"));
    while answered.is_none() && Instant::now() < deadline {
        std::thread::sleep(Duration::from_millis(10));
        answered = PROBE.read(|| unreachable!("the question is put once"));
    }

    assert_eq!(
        answered,
        Some("dxva2"),
        "the device the probe found has to be reachable by a launch that comes after it, or the \
             probe is a question nobody hears the answer to and every film is software decoded for \
             the whole run of the app"
    );
    assert_eq!(
        PROBE.read(|| unreachable!("the question has been put once and is not put again")),
        Some("dxva2"),
        "and reading again gives the same answer rather than starting a second probe: the \
             question is about this machine's drivers and about one build of FFmpeg"
    );
}

/// A setting that is switched makes the answer found for it an answer to a question no longer
/// being asked.
///
/// The tray row promises that the next hover is the first one decoded differently, and without
/// this it would have been a question about the next *run* of the app: the answer found at the
/// first hover stood for the rest of the session, and a probe run with the setting off names no
/// device at all. So the direction that matters is the one that stops naming a device — a film
/// decoded on a card the user has just switched off is a worse fault than one decoded in
/// software, which is what a preview always can be.
#[test]
fn a_switched_setting_stops_the_answer_found_for_the_one_before_it() {
    static PROBE: HwAccelProbe = HwAccelProbe {
        asked: AtomicU32::new(0),
        current: AtomicU32::new(1),
        found: Mutex::new(None),
    };

    assert_eq!(
        PROBE.read(|| Some("dxva2")),
        None,
        "the premise: with the setting on, the first reader puts the question and is itself \
             given nothing"
    );

    let deadline = Instant::now() + Duration::from_secs(5);
    let mut named = PROBE.read(|| unreachable!("the question is put once"));
    while named.is_none() && Instant::now() < deadline {
        std::thread::sleep(Duration::from_millis(10));
        named = PROBE.read(|| unreachable!("the question is put once"));
    }

    assert_eq!(
        named,
        Some("dxva2"),
        "and the device found under that setting is named by every launch after it"
    );

    PROBE.forget();

    assert_eq!(
        PROBE.read(|| None),
        None,
        "after the row is switched the device found for the setting as it was must stop being \
             named, or the row is a question about the next run of the app rather than about the \
             next preview — and a probe run with the setting off names nothing at all"
    );

    // The question is put again for the setting as it now stands, and this time with the answer
    // the setting asks for.
    assert_eq!(
        PROBE.read(|| None),
        None,
        "and it is asked once more, which is what makes the next hover the first one decided by \
             the new setting rather than by the old one"
    );
    assert_eq!(
        PROBE.read(|| unreachable!("the new setting's question has been put once")),
        None,
        "a setting that names no device is asked for once and answered once, so the walk over \
             the candidates is not repeated on every hover"
    );
}

/// A device is named only when a probe found one that survives, and a name FFmpeg does not know
/// is refused before it decodes a frame.
///
/// These two facts are the whole of why the toggle can be on by default and every preview still
/// plays. `-hwaccel` is fatal on machines where FFmpeg's Vulkan renderer cannot be brought up —
/// measured here: `d3d11va` dies of an access violation after four frames, `auto` and `cuda`
/// draw nothing at all — and FFmpeg's own software fallback does not cover that, because a
/// renderer that crashes while the device was being derived never gets as far as a codec the
/// fallback could be offered. So the fallback is this app's: a probe that names nothing when
/// nothing survives, and a preview that plays in software.
///
/// The branch that ends a probe's player when it outlives its budget is the one thing here with
/// no test, and that is deliberate rather than an omission left for later: reaching it needs a
/// player that survives three seconds, which is a fact about the machine rather than about the
/// code, and asserting it would mean leaving a real `ffplay` running on the machine that ran the
/// tests — which is the fault being fixed, reproduced on purpose.
#[test]
fn a_device_is_named_only_when_a_probe_found_one_that_survives() {
    // The answer is read, never waited for: this runs on the preview thread, which is the one
    // thread in this app that must not be waiting on an external process. So this says nothing
    // about *which* device this machine has — it is the arrangement above that is under test —
    // and only that asking twice is asking once.
    assert_eq!(
        video_hw_accel_device(),
        video_hw_accel_device(),
        "two launches of the same run must get the same answer, because the question is about \
             this machine's drivers and about one build of FFmpeg"
    );

    // A device the player refuses before it decodes anything is refused here too, and that is
    // the whole contract of the probe: it is a list walked against the machine, not a name
    // passed on faith.
    assert!(
        !probe_one_hwaccel("no-such-device"),
        "a name FFmpeg does not know is refused before it decodes a frame, so the walk stops \
             rather than naming something fatal"
    );
}

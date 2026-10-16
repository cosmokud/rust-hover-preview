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
    relayout_pinned_media(
        &path,
        (0, 0, 800, 600),
        96,
        None,
        PinRelayoutRoad::BoxChange,
    );

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
            audio_scale: None,
            transport_bar: true,
            transport_live: true,
            frame: PinFrame::Shaped,
            overlay: true,
            hides_chrome: true,
            caption: pinned_caption_height(96, None),
            chrome: PinChrome::on_arrival(Instant::now()),
            collapsed: false,
            bubble_pause: None,
            bubble_run: false,
            hovered: None,
            pressed: None,
            last_press: None,
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
            audio_window_buttons: false,
            menu: PinMenu::default(),
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
/// The filter the pin is given names the copied track it is remembering, spelled the way the
/// filtergraph parses it.
///
/// The pin remembers a track number across every relaunch — a seek, a resize, a change of track —
/// because a player reports nothing about what it did with the number. The number's whole effect
/// is here: it chooses which copied file the filter names. The spelling half is here for the same
/// reason it was always here: a Windows path opens with a colon and the filter's own separator is
/// a colon, which is why FFmpeg's complaint names `original_size`, which nobody asked for. The
/// whole of the escaping, the characters beyond the colon and the one that cannot be escaped at
/// all, is in `video_launch` beside it; what is asserted here is that the track the pin is
/// remembering reaches the filter, which is the half that was missing.
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
    let first = dir.join("sub0.ass");
    std::fs::write(&first, b"[Script Info]\n").expect("a stand-in copy is writable");
    let chosen = dir.join("sub1.ass");
    std::fs::write(&chosen, b"[Script Info]\n").expect("a stand-in copy is writable");
    let copied = DerivedSubtitles {
        tracks: vec![Some(first), Some(chosen)],
        fonts: None,
    };

    // The track the pin is remembering chooses the copy that is named — not a hard-coded first —
    // and the value is quoted the way the filter's parser needs it (see `escape_filter_path`).
    let filter = video_launch::subtitle_filter(None, Some(&copied), Some(1))
        .expect("a film with a copy has a filter to draw it with");
    assert!(
        filter.starts_with("subtitles='") && filter.contains("sub1.ass"),
        "the value is quoted so the escape survives the filtergraph parser, and the copy the pin \
         is remembering is the one the filter names rather than a hard-coded first: {filter}"
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
///
/// **What a probe answers is not asked here either, and for the same reason in the opposite
/// direction.** A probe *is* a process, and a test run starts none (see `video_playback::spawn_player`),
/// so the test build answers "no device survived" without asking anything — a working answer on a
/// machine where nothing does (see `probe_hwaccel_device`). An assertion that a name FFmpeg does
/// not know is refused used to stand here; it would now be asserting that constant, so it is gone
/// rather than kept as a test that cannot fail. What that leaves uncovered is the *walk* — the
/// first device that survives, or nothing — which wants the probe handed in as a closure the way
/// `video_hw::resolve_video_engine` takes its answers, and is worth doing when the walk is next
/// touched.
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
}

/// A video behind the pin, as the call sites below read it: the kind this app plays with a player
/// of its own, with the route stood in per test.
fn stand_video_media() -> Option<MediaData> {
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    previous
}

/// A video behind the pin on the machine's own engine rather than a player of its own: each shared
/// call site reads the native route off the media's kind rather than off `video_route` (see
/// `seek_pinned_playback`), and this is the arm a test about the other two routes pins down
/// alongside them.
fn stand_native_video_media() -> Option<MediaData> {
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::NativeVideo;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    previous
}

/// A pin standing on `film` at a box, playing from the thirtieth second — what a hover leaves
/// behind when it is pinned. No player is begun here; every question below is about the command
/// the mpv route sends or the relaunch FFmpeg's player takes.
fn a_playing_pin(film: &Path, content: ScreenRegion) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = film.to_path_buf();
    pin.transport.begun(30.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.transport.subtitle = Some(0);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// A seek on a pinned film: mpv is taken to the second while it runs, and FFmpeg's player is ended
/// and begun again there — the gate this ticket turns on, asserted on both routes.
#[test]
fn a_seek_on_a_pinned_mpv_is_told_to_the_player_and_relaunches_ffplays() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-seek-film.mkv");
    clear_restart_count();
    mpv_ipc::clear_noted();

    // FFmpeg's player is told nothing while it runs, so the seek is a player begun again at the
    // second the bar was let go at.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    seek_pinned_playback(&film, content, 90.0);
    assert_eq!(
        restart_count(),
        1,
        "a second asked of FFmpeg's player is a player ended and begun again at it"
    );

    // mpv is taken to the second down the pipe it listens on, and the process does not change.
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    // The ffplay half above began a player, which wrote its own pid down: this half is about the
    // player this app was holding before that.
    VIDEO_PID.store(101, Ordering::SeqCst);
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    stand_pin(Some(a_playing_pin(&film, content)));
    seek_pinned_playback(&film, content, 90.0);
    assert_eq!(restart_count(), 0, "mpv is taken to the second while it runs");
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!("{}\n", r#"{"command":["seek",90.000,"absolute"],"async":true}"#)],
        "and the line names the second the bar was let go at"
    );
    assert_eq!(
        pin_state().and_then(|state| {
            state
                .pin()
                .map(|pin| pin.transport.started.map(|(_, from)| from))
        }),
        Some(Some(90.0)),
        "with the bar's own clock rebased onto that second, since no player was begun at it"
    );

    // The machine's own engine: it is told where to go while it runs, so nothing is begun again
    // either, and a held file's bar moves with the seek as it always has.
    stand_native_video_media();
    clear_restart_count();
    let mut held = a_playing_pin(&film, content);
    held.transport.held(30.0);
    stand_pin(Some(held));
    seek_pinned_playback(&film, content, 90.0);
    assert_eq!(
        restart_count(),
        0,
        "the engine is taken to the second while it runs, which is the arm this pins down"
    );
    assert_eq!(
        pin_state().and_then(|state| state.pin().map(|pin| pin.transport.paused_at)),
        Some(Some(90.0)),
        "and a held file's bar is drawn at the second the engine was taken to"
    );

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// A pause and a resume on a pinned film: mpv is told to hold and to go, and a held FFmpeg film
/// whose player has no window to be posted a key to is begun again (see `toggle_pinned_playback`).
#[test]
fn a_pause_on_a_pinned_mpv_is_told_to_the_player_and_relaunches_ffplays_resume() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-pause-film.mkv");
    clear_restart_count();
    mpv_ipc::clear_noted();

    // FFmpeg's player: a resume with no window to post a key to is a player begun again.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    let mut held = a_playing_pin(&film, content);
    held.transport.held(30.0);
    held.transport.started = None;
    stand_pin(Some(held));
    let transport = pinned_playback_state().expect("a pin is up").2;
    toggle_pinned_playback(&film, content, transport);
    assert_eq!(
        restart_count(),
        1,
        "a resume with no window to post a key to is a player begun again"
    );

    // mpv: the hold and the release are two commands, and the process does not change.
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    // The ffplay half above began a player, which wrote its own pid down: this half is about the
    // player this app was holding before that.
    VIDEO_PID.store(101, Ordering::SeqCst);
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    stand_pin(Some(a_playing_pin(&film, content)));
    let transport = pinned_playback_state().expect("a pin is up").2;
    toggle_pinned_playback(&film, content, transport);
    let transport = pinned_playback_state().expect("a pin is up").2;
    toggle_pinned_playback(&film, content, transport);
    assert_eq!(
        restart_count(),
        0,
        "mpv is told to hold and to go: nothing is begun"
    );
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![
            format!("{}\n", r#"{"command":["set_property","pause",true],"async":true}"#),
            format!("{}\n", r#"{"command":["set_property","pause",false],"async":true}"#),
        ],
        "the hold and the release are mpv's own pause property"
    );

    // The machine's own engine: its pause is a state rather than a player ended and begun again,
    // so nothing is begun here and no command is written down a pipe it is not listening on.
    stand_native_video_media();
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand_pin(Some(a_playing_pin(&film, content)));
    let transport = pinned_playback_state().expect("a pin is up").2;
    toggle_pinned_playback(&film, content, transport);
    assert_eq!(
        restart_count(),
        0,
        "the engine's own pause is a state, not a player begun again"
    );
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "and the engine listens on no pipe for one to be written to"
    );

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// The press over a film standing at its own end starts it again at the top on the route mpv is not.
///
/// FFmpeg's player can be told nothing by a key, so a press that resumed a film held at its own end
/// would play the last frames of the file out of a player standing there and let it close its own
/// window: the band is a hole that window is the picture through, so what is left is the desktop and
/// a play button with nothing behind it (measured on the machine this was written on). The pass the
/// press asks for on that route is therefore a player begun again at the top, which is the same
/// relaunch a seek is for it (see `restart_pinned_player`).
#[test]
fn a_press_over_a_finished_ffplay_film_begins_it_again_at_the_top() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("finished-film.mkv");
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    clear_restart_count();
    ffplay_window::clear_pause_posts();

    // A film held at its own end, which is where the end of file that pauses leaves one the rewind
    // could not take home (see `loop_ended_pinned_player`).
    let mut finished = a_playing_pin(&film, content);
    finished.transport.held(200.0);
    finished.transport.started = None;
    stand_pin(Some(finished));
    let transport = pinned_playback_state().expect("a pin is up").2;
    toggle_pinned_playback(&film, content, transport);

    assert_eq!(
        restart_count(),
        1,
        "the film is begun again rather than resumed where it stands"
    );
    assert_eq!(
        ffplay_window::pause_posts(),
        0,
        "and the player is never told to hold: a resume at the end of the file is the last frames of \
         it played out of a player that then closes its own window"
    );
    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.started.map(|(_, from)| from),
        Some(0.0),
        "and the second it is begun at is the top of the file"
    );
    assert_eq!(transport.paused_at, None);
    assert!(pin_is_playing(&transport));

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// A play pressed on the bubble of a pin the collapse parked, with the `Pin Mode → Pause on
/// Minimize` switches on: the park is given back the road a restore gives it back, and the tick
/// that parked it does not park it again over the hand that asked for it.
///
/// Its old road was two faults. A park standing was answered by the ordinary toggle, which let go
/// of a hold a dead player could not be holding and left the bubble holding a park for a film
/// already going — so the next tick's park was written under a film the hand had just asked for,
/// and a restore gave it back a second time on top. A bubble with nothing parked — a film the bar
/// had paused before the pin went down — was the worse one: the press answered nothing at all,
/// which was the whole of "Play on a collapsed pin does nothing".
#[test]
fn a_play_on_a_collapsed_pins_bubble_gives_the_park_back_and_keeps_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _switches = PausePreviewSettings::stood_at(true, true);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("bubble-play-film.mkv");
    clear_restart_count();
    mpv_ipc::clear_noted();

    // FFmpeg's route, where a park the collapse wrote is a player ended at the second it had
    // reached (see `bubble_playback_to_park`).
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    let mut parked = a_playing_pin(&film, content);
    parked.collapsed = true;
    parked.bubble_pause = Some(BubblePause::Player(30.0));
    parked.transport.held(30.0);
    parked.transport.started = None;
    stand_pin(Some(parked));

    let mut started = None;
    let mut offset = 0.0;
    toggle_pinned_by_key(&mut started, &mut offset, &mut None);

    assert_eq!(
        restart_count(),
        1,
        "the press begins the player again: the park the collapse wrote is given back the road a \
              restore gives one back"
    );
    assert!(
        pin_bubble_pause().is_none(),
        "the bubble holds no park for the film it has just given back"
    );
    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    assert_eq!(
        transport.started.map(|(_, from)| from),
        Some(30.0),
        "and it is begun at the second the park wrote down"
    );
    assert!(pin_is_playing(&transport), "the film is going again");

    // And the tick that parked it stands down over it: a park written a tick after the press that
    // asked for the film back is the bubble arguing with the hand.
    settle_bubble_playback(&mut started, &mut offset);
    assert!(
        pin_bubble_pause().is_none(),
        "the switches do not park the film back over the hand that asked for it"
    );
    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    assert!(
        transport.started.is_some() && transport.paused_at.is_none(),
        "and the film is still going, which is what the press asked for"
    );

    // A second press is the ordinary hold, and the hold outlives the pin coming back up: a hold
    // the hand asked for is not a park the collapse made, and a restore gives nothing back.
    toggle_pinned_by_key(&mut started, &mut offset, &mut None);
    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    assert!(
        transport.paused_at.is_some(),
        "the press again holds the film where it stands"
    );
    assert!(
        pin_bubble_pause().is_none(),
        "and holds it as a hold rather than as a park: there is nothing for a restore to give back"
    );

    restore_pin();
    assert!(
        !pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.collapsed))
            .expect("a pin is up"),
        "the pin is up again"
    );
    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    assert!(
        transport.paused_at.is_some(),
        "and the hold is still there: what is not playing is not parked"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// A play pressed on the bubble of a pin whose park the switches never wrote — a film the bar was
/// used to pause before the pin went down, which a collapse parks nothing for.
///
/// This is the one the user reported: the press played the film once and the tick parked it again,
/// because a park with nothing behind it is a park the bubble writes the moment it sees a film
/// running under a collapsed pin. The press is the hand, and the bubble does not argue with it:
/// the film keeps going until the pin comes up again or the same press holds it (see
/// `toggle_pinned_by_key` and `settle_bubble_playback`).
#[test]
fn a_play_on_a_collapsed_pins_bubble_that_parked_nothing_keeps_the_film_going() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _switches = PausePreviewSettings::stood_at(true, true);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("bubble-play-paused-film.mkv");
    clear_restart_count();
    ffplay_window::clear_pause_posts();

    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

    // A film the bar was used to pause before the pin went down: not playing, so the collapse
    // parks nothing for it — what is not playing is not parked (see `bubble_playback_to_park`).
    let mut paused = a_playing_pin(&film, content);
    paused.collapsed = true;
    paused.transport.held(30.0);
    paused.transport.started = None;
    stand_pin(Some(paused));

    let mut started = None;
    let mut offset = 0.0;
    toggle_pinned_by_key(&mut started, &mut offset, &mut None);

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    assert!(
        pin_is_playing(&transport),
        "the press sets the held film going again"
    );

    // And the tick leaves it going, which is the whole of it: the switches read the pin as down
    // and see a film running under it, and the press that started it is the hand answering them.
    ffplay_window::clear_pause_posts();
    settle_bubble_playback(&mut started, &mut offset);
    assert_eq!(
        ffplay_window::pause_posts(),
        0,
        "the switches do not park the film back: no hold is delivered to any player"
    );
    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    assert!(
        pin_is_playing(&transport),
        "the film the hand asked for is still going"
    );
    assert!(
        pin_bubble_pause().is_none(),
        "and the bubble wrote no park for it"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
    ffplay_window::clear_pause_posts();
}
/// A wrap of a film this app is looping: mpv is sent back to the second down its pipe and set going
/// there, and FFmpeg's film is begun again at it (see `video_loop_tick`).
#[test]
fn a_loop_wrap_on_a_pinned_mpv_is_told_to_the_player_and_restarts_ffplays() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (0, 0, 320, 240);
    let film = PathBuf::from("live-loop-film.mkv");
    stand_pin(Some(a_playing_pin(&film, content)));
    clear_restart_count();
    mpv_ipc::clear_noted();

    // A wrap is what a film whose end of file is a *loop* does, and the mpv half below is about the
    // play that brings such a film back. A film whose end of file is `Pause` is taken to its top and
    // *held* there instead of set going, which is its own test (see
    // `the_tick_takes_a_pausing_film_back_to_its_top_and_holds_it_there`).
    let _jumps = VideoJumps::set(false, PinEndOfFile::Loop);

    // The record is moved back as if the pass had run its length: a wrap is due on the next tick.
    let a_pass_has_run = || {
        if let Ok(mut slot) = VIDEO_LOOP.lock() {
            if let Some(loop_) = slot.as_mut() {
                loop_.at = Instant::now() - Duration::from_secs(9);
            }
        }
    };

    // FFmpeg's player: the rewind is a player begun again at the beginning.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    note_video_loop(101, &film, content, 0.0, Some(8.0));
    a_pass_has_run();
    assert!(video_loop_tick(true, true), "the tick takes the film back");
    assert_eq!(
        restart_count(),
        1,
        "a loop FFmpeg's player cannot be told takes a player begun again"
    );

    // mpv: the same wrap is a command, and the process stays.
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    // The ffplay half above began a player, which wrote its own pid down: this half is about the
    // player this app was holding before that.
    VIDEO_PID.store(101, Ordering::SeqCst);
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    note_video_loop(101, &film, content, 0.0, Some(8.0));
    a_pass_has_run();
    assert!(video_loop_tick(true, true), "the tick sends the film back");
    assert_eq!(
        restart_count(),
        0,
        "mpv is told to go back rather than begun again"
    );
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![
            format!("{}\n", r#"{"command":["seek",0.000,"absolute"],"async":true}"#),
            format!("{}\n", r#"{"command":["set_property","pause",false],"async":true}"#),
        ],
        "the wrap is a seek to the beginning, which is the only second a loop can restart at, and \
         then the play that brings back a film mpv had paused on its last frame: mpv keeps a pause \
         across a seek, so a wrap that arrived after the pass had ended would otherwise leave the \
         film standing at its first frame, silent, for good"
    );

    // The machine's own engine: a loop it cannot be told to give is a player begun again, the same
    // answer FFmpeg's player gives and the one mpv alone was spared.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(false));
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    clear_restart_count();
    mpv_ipc::clear_noted();
    VIDEO_PID.store(101, Ordering::SeqCst);
    note_video_loop(101, &film, content, 0.0, Some(8.0));
    a_pass_has_run();
    assert!(video_loop_tick(true, true), "the tick takes the engine's film back");
    assert_eq!(
        restart_count(),
        1,
        "a loop nothing can be told takes a player begun again on the engine's route too"
    );

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// The clock a pass is measured by starts when the film begins playing rather than when its process
/// was spawned: a player is a good part of a second away from the picture, and a clock that runs
/// ahead of it reaches the end of the pass early and cuts that time off the film's tail on every
/// wrap (see `rebase_video_loop_clock`).
#[test]
fn the_loop_clock_is_moved_onto_the_moment_the_film_began_playing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let content = (0, 0, 320, 240);
    let film = PathBuf::from("slow-to-open-film.mkv");
    note_video_loop(101, &film, content, 0.0, Some(8.0));

    let position = || {
        VIDEO_LOOP
            .lock()
            .ok()
            .and_then(|slot| slot.as_ref().map(|loop_| loop_.position()))
            .unwrap_or_default()
    };
    // The clock as a launch leaves it: started when the process was, so standing however long the
    // player spent opening the file ahead of the picture.
    let ahead_by = |seconds: u64| {
        if let Ok(mut slot) = VIDEO_LOOP.lock() {
            if let Some(loop_) = slot.as_mut() {
                loop_.at = Instant::now() - Duration::from_secs(seconds);
            }
        }
    };

    ahead_by(2);
    assert!(
        position() >= 2.0,
        "a clock started at the spawn stands ahead of the picture by the time the player spent \
         opening the file, which is what the rewind is posted early by"
    );

    rebase_video_loop_clock(101);
    assert!(
        position() < 0.5,
        "the frame the player has just put up is the beginning of the pass, so the clock is moved \
         onto this moment and measures the film from where it actually is"
    );

    // And a window coming up for a player this app does not hold is not this clock's to move: the
    // record names the player it is the clock of.
    ahead_by(2);
    rebase_video_loop_clock(202);
    assert!(
        position() >= 2.0,
        "a clock belonging to another player is left where it stands"
    );

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
}

/// The two ends of file a film on screen is played under — the hover's loop and the pin's end —
/// put back when the guard is dropped.
struct VideoJumps {
    hover: bool,
    pin: PinEndOfFile,
}

impl VideoJumps {
    fn set(hover: bool, pin: PinEndOfFile) -> Self {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = VideoJumps {
            hover: config.video_loop,
            pin: config.pin_mode_video_end_of_file,
        };
        config.video_loop = hover;
        config.pin_mode_video_end_of_file = pin;
        was
    }
}

impl Drop for VideoJumps {
    fn drop(&mut self) {
        if let Ok(mut config) = CONFIG.lock() {
            config.video_loop = self.hover;
            config.pin_mode_video_end_of_file = self.pin;
        }
    }
}

/// A film whose end of file is `Pause` is taken back to its first frame and held there, rather than
/// let go of at the end it has reached.
///
/// A film held at its top is begun again by the play button with one command and no seek at all
/// (see `toggle_pinned_playback`), while a film held at its *end* has to be taken to the top by that
/// press — and a seek written against a player that has stopped at the end of its file is a restart
/// of the film's whole chain, which is the picture going away (see `pin_video_pauses_at_the_top`).
/// The wrap is posted before the film gets there, so it is the same seek a loop's wrap is.
#[test]
fn the_tick_takes_a_pausing_film_back_to_its_top_and_holds_it_there() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (0, 0, 320, 240);
    let film = PathBuf::from("paused-at-the-top-film.mkv");
    stand_pin(Some(a_playing_pin(&film, content)));
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    mpv_ipc::clear_noted();

    let _jumps = VideoJumps::set(false, PinEndOfFile::Pause);
    assert!(
        pin_video_pauses_at_the_top(&film),
        "the film on screen is the pinned one whose end of file pauses"
    );
    note_video_loop(101, &film, content, 0.0, Some(8.0));
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if let Some(loop_) = slot.as_mut() {
            loop_.at = Instant::now() - Duration::from_secs(9);
        }
    }

    assert!(
        video_loop_tick(true, pin_video_loops(&film) || pin_video_pauses_at_the_top(&film)),
        "the tick takes the film back, which is the one road this end of file has for mpv"
    );
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![
            format!("{}\n", r#"{"command":["seek",0.000,"absolute"],"async":true}"#),
            format!(
                "{}\n",
                r#"{"command":["set_property","pause",true],"async":true}"#
            ),
        ],
        "the film is taken to its top and *held* there: the pause is the half a seek cannot do, and \
         it is what leaves the play button a plain resume to make"
    );

    let (from, at) = VIDEO_LOOP
        .lock()
        .ok()
        .and_then(|slot| slot.as_ref().map(|loop_| (loop_.from, loop_.at)))
        .expect("the loop's clock is where it was left");
    assert_eq!(from, 0.0, "the pass the film waits on begins at its top");
    assert!(
        at.elapsed() < Duration::from_secs(1),
        "and from this moment, which is when it was taken there"
    );

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.paused_at,
        Some(0.0),
        "and the bar stands at the top with the picture rather than at the end"
    );
    assert!(
        !pin_is_playing(&transport),
        "so the button over it is the play button a held film is owed"
    );

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A pinned film's loop is the *pin's* end of file rather than the hover's, which is the one thing a
/// take-up has to settle for mpv.
///
/// The clock a pass is measured by is made where a player is begun, and mpv is the only player a
/// take-up does not begin again — so the film was begun by the *hover*, under the hover's loop, and
/// that answer was carried into the pin with it. `End of File → Loop` with the hover's loop off
/// was then a pinned film no rewind ever reached, and a pin's step with the hover's loop on was a
/// rewind that undid the step. What fixes both is that the end of file is asked of the film on screen when
/// the end comes rather than remembered with the clock (see `pin_video_loops` and
/// `video_loop_action`).
#[test]
fn a_pinned_films_loop_is_the_pins_own_end_of_file_and_not_the_hovers() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (0, 0, 320, 240);
    let film = PathBuf::from("pinned-loop-film.mkv");
    stand_pin(Some(a_playing_pin(&film, content)));
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    mpv_ipc::clear_noted();
    clear_restart_count();

    // The record as the hover left it: the player it names was begun for the hover, and a pass has
    // run its length since.
    let a_pass_has_run = || {
        if let Ok(mut slot) = VIDEO_LOOP.lock() {
            if let Some(loop_) = slot.as_mut() {
                loop_.at = Instant::now() - Duration::from_secs(9);
            }
        }
    };

    // A hover whose loop is off and a pin whose end of file is a loop: the film on screen is the
    // pin's, so it goes round.
    let _jumps = VideoJumps::set(false, PinEndOfFile::Loop);
    assert!(
        pin_video_loops(&film),
        "the end of file of the film on screen is the pin's rather than the hover's, which is what a \
         record made for the hover's player has to be read by"
    );
    note_video_loop(101, &film, content, 0.0, Some(8.0));
    a_pass_has_run();
    assert!(
        video_loop_tick(true, pin_video_loops(&film)),
        "and the tick takes it back: for a pinned mpv film this tick is the only road the loop has, \
         the player being one a take-up does not begin again"
    );
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![
            format!("{}\n", r#"{"command":["seek",0.000,"absolute"],"async":true}"#),
            format!("{}\n", r#"{"command":["set_property","pause",false],"async":true}"#),
        ],
        "the wrap is the seek and the play, the same as a hover's"
    );

    // And the other way round: a pin's step is not undone by a clock the hover made while looping.
    drop(_jumps);
    let _jumps = VideoJumps::set(true, PinEndOfFile::Continue);
    assert!(
        !pin_video_loops(&film),
        "a pinned film the pin's end of file steps on rather than loops is not one this app gives a loop to"
    );
    note_video_loop(101, &film, content, 0.0, Some(8.0));
    a_pass_has_run();
    mpv_ipc::clear_noted();
    assert!(
        !video_loop_tick(true, pin_video_loops(&film)),
        "so the tick leaves it alone, and the end is the walk's to answer (see \
         `loop_ended_pinned_player`)"
    );
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "and nothing was told to the player at all"
    );

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// A video behind the pin with a player this app holds *and that is still up*.
///
/// `is_video_process_running` reads the handle the media holds, and a test's noted player never
/// ends — which is what makes this the stand-in for the one player that stays up at the end of its
/// film: mpv, launched with `--keep-open=yes`, holding the last frame while the app's own clock
/// runs on past it. Nothing is spawned here; the id is odd, so it names no real process even if a
/// kill ever asked (see `PlayerProcess`).
fn stand_live_video_media() -> Option<MediaData> {
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    video.video_process = Some(video_playback::PlayerProcess::Noted(101));
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    previous
}

/// The state a live player is left in when the app's clock has counted out the whole file: the
/// transport's clock stands `reached` ago, and the player behind it is still up.
fn a_pass_that_has_run_out(reached: Duration) {
    update_pin_transport(|transport| {
        transport.started = Some((Instant::now() - reached, 0.0));
    });
}

/// A pass that is still being played ends at the length of its file and not five seconds before it.
///
/// The five seconds `player_end` allows are for a player that has *gone*, whose going has to be read
/// as the film finishing; a player that is still up has done no going, and reading an end of file
/// off the same grace began every next file while the last one was still talking.
#[test]
fn a_pass_still_being_played_ends_at_the_length_of_its_file() {
    assert!(
        !pass_ended(195.0, Some(200.0)),
        "five seconds of a film still to play is not its end"
    );
    assert!(
        pass_ended(200.0, Some(200.0)),
        "the length of the file is, and it is the only reading that is"
    );
    assert!(pass_ended(201.0, Some(200.0)));
    assert!(
        !pass_ended(30.0, None),
        "a file whose length nothing has read has no end to be at"
    );
    assert_eq!(
        player_end(true, 195.0, Some(200.0), false),
        PlayerEnd::Finished,
        "while the gone-player reading keeps the grace it was written for"
    );
}

/// The end of file that pauses is left to the tick's own rewind wherever the rewind can take the film
/// home, and settled at its end where it cannot.
///
/// This is the end that never used to arrive, and settling the bar on it is now the one thing that
/// must not happen for a film the rewind *can* take home: a player that exits is reconciled by
/// `settle_pinned_transport`, but mpv holds its last frame instead of going, so the clock counted past
/// the end for as long as the pin was up and the button claimed a film that was playing. The clock is
/// settled on this road by the rewind rather than here — the film is taken back to its first frame and
/// held, which stops the clock and writes the bar down as held in the same breath (see
/// `video_loop_tick`) — and a film written down as held *here* is one that rewind would leave alone,
/// with the press that plays it again left to seek a player stopped at the end of its file.
///
/// A film the rewind cannot take home — one shorter than the tenth of a second a rewind is posted
/// within — is still this reading's, because otherwise nothing would ever stop the clock over it.
#[test]
fn the_end_of_file_that_pauses_is_left_to_the_rewind_where_it_can_take_the_film() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_live_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-end-film.mkv");
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    let _jumps = VideoJumps::set(false, PinEndOfFile::Pause);

    assert!(
        is_video_process_running(),
        "the film's player is the one thing this reading is about, and this test stood one up"
    );

    // The film has run out, and nothing has settled it: mpv is still there, holding its last frame.
    a_pass_that_has_run_out(Duration::from_secs(200));

    let mut wait = None;
    assert!(
        !loop_ended_pinned_player(None, &mut wait, false),
        "the rewind owns this end of file, and what it is owed is a film this app still reads as \
         playing"
    );
    assert!(
        wait.is_none(),
        "an end of file that pauses asks for no walk, and no player begun again"
    );

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert!(
        transport.paused_at.is_none(),
        "nothing is written down as held here: the by-passed reading of the film's end is what the \
         rewind is asked to take it home for"
    );
    assert!(transport.started.is_some());
    assert!(
        pin_is_playing(&transport),
        "so the button goes on saying that something is playing, which is what the rewind is owed"
    );

    // And a film the rewind can never take home — one shorter than the margin a rewind is posted
    // within — is still this reading's: nothing else in the tick would ever stop the clock over it.
    update_pin_transport(|transport| transport.duration = Some(0.05));
    a_pass_that_has_run_out(Duration::from_millis(50));
    let mut wait = None;
    assert!(
        loop_ended_pinned_player(None, &mut wait, false),
        "the end of a film too short to be rewound is answered here"
    );
    assert!(wait.is_none(), "and it asks for no walk either");

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.paused_at,
        Some(0.05),
        "with the bar stopped at the film's own end rather than sweeping past it"
    );
    assert!(transport.started.is_none());
    assert!(!pin_is_playing(&transport));

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The answer to a native film's end of pass: the film taken home and held, or the walk asked for.
///
/// The engine is the one player whose loop is its own — a film begun for a hover is given `Video →
/// Loop` and the engine plays it round itself — so `End of File → Pause` on that route is the road
/// mpv's film goes home by (see `pin_video_pauses_at_the_top`) written for a player with no pipe: the
/// film is taken to its first frame and left standing there, and the press that plays it again is the
/// plain resume of a held film rather than a seek against one that has stopped at the end of its
/// file. `Continue` is the walk's next file, on the same margin, and the walk's own wait is what keeps
/// the pin up while it is out.
#[test]
fn a_native_pass_end_takes_its_film_home_or_asks_for_the_next_one() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();

    let content = (0, 0, 320, 240);
    let film = PathBuf::from("native-end-film.mkv");
    stand_pin(Some(a_playing_pin(&film, content)));

    let mut wait = None;
    assert!(
        answer_native_pass_end(NativePassEnd::Hold, &film, None, &mut wait),
        "the end of a pass that pauses is answered here"
    );
    assert!(wait.is_none(), "a pause asks for no walk");

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.paused_at,
        Some(0.0),
        "with the bar held at the top the film has been taken to rather than at the end of the file"
    );
    assert!(!pin_is_playing(&transport));
    assert!(
        transport.started.is_some(),
        "and the clock a pass is measured by is left where it was: the engine's own position is what \
         this bar is read from while the film is held"
    );

    // The end that steps on is the walk's: the session let go and the next file of the folder asked
    // for, which is the same ask the caption's own **Next** button makes.
    assert!(
        answer_native_pass_end(NativePassEnd::Walk, &film, None, &mut wait),
        "an end of file that steps on asks for the next file"
    );
    assert!(wait.is_some(), "and waits for the answer");
    assert!(
        !answer_native_pass_end(NativePassEnd::Nothing, &film, None, &mut wait),
        "while an end of file that asks for nothing writes nothing at all"
    );

    // The walk the answer asked for is taken back out of the slot it was left in rather than left
    // for another test's run to find.
    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }

    stand_pin(previous_pin);
}

/// What a native film's own pass asks of the pin, from the three facts that decide it.
///
/// The engine is the one player whose loop is its own (`Video → Loop` is what a film begun for a
/// hover is given), so a `Loop` end of file is not this app's to answer at all; the other two are the
/// ask the tick's own rewind makes of a film mpv is drawing, written for the route with no pipe — a
/// film due home a tenth of a second before the end of its pass, while it is still playing, and no
/// film this app is holding asked for anything.
#[test]
fn a_native_films_end_of_file_is_read_off_the_engines_own_clock() {
    assert_eq!(
        native_pass_end(PinEndOfFile::Loop, true, true),
        NativePassEnd::Nothing,
        "a film the engine loops itself is not one this app answers the end of"
    );
    assert_eq!(
        native_pass_end(PinEndOfFile::Continue, true, true),
        NativePassEnd::Walk,
        "an end that steps on is the walk's next file"
    );
    assert_eq!(
        native_pass_end(PinEndOfFile::Pause, true, true),
        NativePassEnd::Hold,
        "and an end that pauses is the film taken home and held there"
    );
    assert_eq!(
        native_pass_end(PinEndOfFile::Continue, true, false),
        NativePassEnd::Nothing,
        "a pass with its margin still to run asks for nothing"
    );
    assert_eq!(native_pass_end(PinEndOfFile::Pause, true, false), NativePassEnd::Nothing);
    assert_eq!(
        native_pass_end(PinEndOfFile::Pause, false, true),
        NativePassEnd::Nothing,
        "and a film this app is holding is not taken home again"
    );
}

/// The end of file that pauses takes an FFmpeg film home before it reaches its own end, on the same
/// margin the loop's wrap and mpv's rewind are posted on.
///
/// The route under mpv is a player nothing of this app's can be told once it is going — a film a
/// hover began carries FFmpeg's own `-loop 0`, because a pin begins no player of its own over one
/// that is already there — so the road mpv's film goes home by is written for it as the one move it
/// has: the player begun again at the film's first frame, *held* there, with the hold owed to the
/// tick that finds the new window up (see `settle_pending_hold`). Left to reach its own end instead,
/// such a film plays the last frames of its file out of a player standing at the end and closes its
/// own window — measured, and the band is a hole that window is the picture through.
#[test]
fn the_end_of_file_that_pauses_takes_a_ffplay_film_home_and_holds_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_live_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-pausing-film.mkv");
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    let _jumps = VideoJumps::set(true, PinEndOfFile::Pause);
    clear_restart_count();

    assert!(
        !pin_video_pauses_at_the_top(&film),
        "a rewind this app can only answer with a player begun again at zero is not this route's"
    );

    // A film the pin took over from the hover is a player this app is looping itself — its clock is
    // a record, and the tick reads it as a wrap that is due — and the same record must be left
    // exactly where it is: the pausing rewind is not this route's to give (see `video_loop_tick`).
    note_video_loop(101, &film, content, 0.0, Some(200.0));
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if let Some(loop_) = slot.as_mut() {
            loop_.at = Instant::now() - Duration::from_secs(201);
        }
    }
    assert!(
        !video_loop_tick(true, pin_video_loops(&film) || pin_video_pauses_at_the_top(&film)),
        "so the tick leaves it where it is rather than beginning it again"
    );
    assert_eq!(
        restart_count(),
        0,
        "which is the whole of the fault: a rewind here is a player begun again, playing"
    );
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }

    // A film with its margin still to run is not due home yet: a tenth of a second before the end of
    // the file is where the move is made, and no earlier.
    a_pass_that_has_run_out(Duration::from_secs(199));
    let mut wait = None;
    assert!(
        !loop_ended_pinned_player(None, &mut wait, false),
        "a film with a second of its pass still to run is left alone"
    );
    assert_eq!(restart_count(), 0, "and nothing is begun for it");

    // And at the margin it is begun again at its first frame, held there.
    a_pass_that_has_run_out(Duration::from_secs(200));
    assert!(
        loop_ended_pinned_player(None, &mut wait, false),
        "the last tenth of a second is where the film is taken home"
    );
    assert!(wait.is_none(), "a pause asks for no walk");
    assert_eq!(
        restart_count(),
        1,
        "FFmpeg's player cannot be taken anywhere while it runs: the film comes home as a player \
         begun again at the top"
    );

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.paused_at,
        Some(0.0),
        "with the bar held at the top the film was taken to rather than at the end it was at"
    );
    assert!(
        transport.pending_hold,
        "and the hold owed to the player whose window is not up yet, which the tick delivers"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// The end of file that pauses holds a film the rewind can never take home at the end it has already
/// reached, rather than leaving the bar sweeping over it.
///
/// The road home is a tenth of a second wide, and a film shorter than that is never due home at all —
/// every second of it is inside the margin, so a rewind asked for it would be posted the instant it
/// began (see `video_launch::rewind_due`). What such a film is left with is `Pause`'s other half, the
/// one this reading answers: the film has run out, the player is still up — FFmpeg's player wraps to
/// the beginning by itself where it was given FFmpeg's own `-loop 0`, which is what a *hover's* film
/// is begun with and the pin keeps, because a pin does not begin a player of its own over a hover's —
/// so the player is held where it stands by the pause key the play button already posts, with the
/// fallback for a player with no window to post one through, and the bar is held at the film's end
/// with the play button a finished film is owed.
#[test]
fn the_end_of_file_that_pauses_stops_a_film_the_rewind_cannot_take_home() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_live_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-pausing-film.mkv");
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    let mut tiny = a_playing_pin(&film, content);
    tiny.transport.duration = Some(0.05);
    stand_pin(Some(tiny));
    let _jumps = VideoJumps::set(true, PinEndOfFile::Pause);
    ffplay_window::clear_pause_posts();
    clear_restart_count();

    assert!(
        !pin_video_pauses_at_the_top(&film),
        "a rewind this app can only answer with a player begun again at zero is not this route's"
    );

    // The pass has run out, and the film is one no rewind can be posted for: the hold is this
    // reading's, and the player is told what to hold by the key the play button posts.
    a_pass_that_has_run_out(Duration::from_millis(50));
    let mut wait = None;
    assert!(
        loop_ended_pinned_player(None, &mut wait, false),
        "the end of a film whose player cannot be taken home is answered here"
    );
    assert!(wait.is_none(), "a pause asks for no walk");
    assert_eq!(
        ffplay_window::pause_posts(),
        1,
        "the player is held where it is by the same key the play button posts"
    );

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.paused_at,
        Some(0.05),
        "the bar stands at the film's own end rather than sweeping past it"
    );
    assert!(!pin_is_playing(&transport));

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// The end of file that steps on waits for the film's own end, and the bar stops there while the
/// walk is out.
///
/// Five seconds read off the tail of every film was a next file begun while the one before it was
/// still talking — and, with the bar drawing a wrapped clock, a next file begun while the bar said
/// the film was nowhere near its end. The clock under the bar is what the walk is asked from, so the
/// walk is asked at the end of the file and not a moment before it.
#[test]
fn the_end_of_file_that_continues_waits_for_the_films_own_end() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_live_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-continue-film.mkv");
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    let _jumps = VideoJumps::set(false, PinEndOfFile::Continue);

    // Inside the grace the gone-player reading allows, and nowhere near the film's end: nothing is
    // asked for, and the bar goes on saying the film has five seconds in it.
    a_pass_that_has_run_out(Duration::from_secs(196));
    let mut wait = None;
    assert!(
        !loop_ended_pinned_player(None, &mut wait, false),
        "a film with four seconds still to play is not a film that has ended"
    );
    assert!(wait.is_none(), "and no walk is asked for yet");

    // And at the end of the file it is asked for, with the bar stopped over the film it belongs to
    // until the walk lands rather than counting on past it.
    a_pass_that_has_run_out(Duration::from_secs(201));
    assert!(
        loop_ended_pinned_player(None, &mut wait, false),
        "the film's own end is where the next file is asked for"
    );
    assert!(wait.is_some(), "and the answer is waited for");

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.paused_at,
        Some(200.0),
        "the bar stands at the film's end while the walk is out, not sweeping past it"
    );
    assert!(transport.started.is_none());
    assert!(!pin_is_playing(&transport));

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// An mpv wrap moves the pin's bar onto the pass it began, as well as the loop's own clock.
///
/// A wrap down the pipe is not a relaunch, so `transport.begun` never runs for it and nothing wrote
/// the second the pass now runs from: the bar went on counting from the launch the pin was taken up
/// at, and after the first wrap it stood a whole pass away from the picture it was drawn over.
#[test]
fn an_mpv_wrap_moves_the_pins_bar_onto_the_pass_it_began() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-wrap-film.mkv");
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    let _jumps = VideoJumps::set(true, PinEndOfFile::Loop);
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    mpv_ipc::clear_noted();

    note_video_loop(101, &film, content, 0.0, Some(8.0));
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if let Some(loop_) = slot.as_mut() {
            loop_.at = Instant::now() - Duration::from_secs(9);
        }
    }

    assert!(video_loop_tick(true, true), "a pass that has run out wraps");
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![
            format!(
                "{}\n",
                r#"{"command":["seek",0.000,"absolute"],"async":true}"#
            ),
            format!(
                "{}\n",
                r#"{"command":["set_property","pause",false],"async":true}"#
            ),
        ],
        "the wrap is still the seek and the play, and nothing else is told to the player"
    );

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(
        transport.started.map(|(_, from)| from),
        Some(0.0),
        "the bar's clock is moved onto the second the wrap began the pass at"
    );
    assert!(
        transport
            .started
            .is_some_and(|(at, _)| at.elapsed() < Duration::from_secs(1)),
        "and from this moment, which is where the picture is"
    );
    assert!(
        transport.paused_at.is_none(),
        "a wrapped pass is a playing one"
    );

    drop(_jumps);
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A seek down mpv's pipe moves the loop's clock with the bar, so the next tick does not wrap a
/// film back to the beginning because the clock a pass is measured by is still where the seek left.
#[test]
fn an_mpv_seek_moves_the_loops_clock_onto_the_second() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-seek-clock-film.mkv");
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    mpv_ipc::clear_noted();

    // The clock's reckoning is a hair from the end of the film, which is the state a seek into the
    // last moments leaves it in.
    note_video_loop(101, &film, content, 0.0, Some(200.0));
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if let Some(loop_) = slot.as_mut() {
            loop_.at = Instant::now() - Duration::from_secs(199);
        }
    }

    seek_pinned_playback(&film, content, 150.0);
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["seek",150.000,"absolute"],"async":true}"#
        )],
        "the seek is a command to the player, and nothing else is told to it"
    );

    let (from, at) = VIDEO_LOOP
        .lock()
        .ok()
        .and_then(|slot| slot.as_ref().map(|loop_| (loop_.from, loop_.at)))
        .expect("the loop's clock is where it was left");
    assert_eq!(
        from, 150.0,
        "the pass is measured from the second it was taken to"
    );
    assert!(
        at.elapsed() < Duration::from_secs(1),
        "and from this moment, which is when the seek was sent"
    );

    // And the pin's own bar has moved there too, which is the clock the seek already rebased.
    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert_eq!(transport.started.map(|(_, from)| from), Some(150.0));

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A press on the play button over a film the tick is holding at its top is one command, and no seek.
///
/// The film is standing at its first frame because the rewind took it there (see
/// `pin_video_pauses_at_the_top` and `video_loop_tick`), so playing it again is the plain resume
/// every held film already has — `set_property pause=false` and nothing else. A seek here would be a
/// seek written against a player that has stopped at the end of its file, which mpv answers as a
/// restart of the film's whole chain; that is the picture blinking out, and it is what the hold is
/// for. The seek is left in the code for the film nothing moved: one a pin was taken up on while it
/// stood at its end, which is the road the test below covers.
#[test]
fn a_film_the_tick_holds_at_its_top_is_resumed_by_one_command() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_live_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("held-at-the-top-film.mkv");
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    mpv_ipc::clear_noted();
    clear_restart_count();
    note_video_loop(101, &film, content, 0.0, Some(200.0));

    // The state the rewind leaves: the film standing at the top of itself.
    update_pin_transport(|transport| transport.held(0.0));

    let (path, content, transport, _) = pinned_playback_state().expect("the pin is up");
    toggle_pinned_playback(&path, content, transport);

    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","pause",false],"async":true}"#
        )],
        "the film is at its top already, so the press only sets it going: a seek here is the chain \
         restart the hold was written to avoid"
    );

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert!(transport.paused_at.is_none(), "the pin is playing again");
    assert_eq!(
        transport.started.map(|(_, from)| from),
        Some(0.0),
        "and its clock is at the top with the picture"
    );

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A press on the play button over a film held at its own end begins the film from the top.
///
/// mpv is standing on the last frame, and a player set going from the last frame plays it and ends
/// the pass again at once — so the film is taken to the beginning *before* it is set going, and both
/// clocks are written at the second it is now played from.
#[test]
fn a_film_held_at_its_end_plays_again_from_the_top() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_live_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("live-replay-film.mkv");
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_pin(Some(a_playing_pin(&film, content)));
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    mpv_ipc::clear_noted();
    note_video_loop(101, &film, content, 0.0, Some(200.0));

    // The state the end of file that pauses leaves: the film held at its own end.
    update_pin_transport(|transport| transport.player_gone(200.0));

    let (path, content, transport, _) = pinned_playback_state().expect("the pin is up");
    toggle_pinned_playback(&path, content, transport);

    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![
            format!(
                "{}\n",
                r#"{"command":["seek",0.000,"absolute"],"async":true}"#
            ),
            format!(
                "{}\n",
                r#"{"command":["set_property","pause",false],"async":true}"#
            ),
        ],
        "the film is taken to the top first and set going there, rather than let go of on its last \
         frame"
    );

    let transport = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("the pin is up");
    assert!(transport.paused_at.is_none(), "the pin is playing again");
    assert_eq!(
        transport.started.map(|(_, from)| from),
        Some(0.0),
        "and its clock is at the top with the picture"
    );

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A pin taken up on a film the hover has been playing starts its bar where the picture is, not at
/// the beginning of the file.
///
/// The player is the one the pin adopts rather than a new one (see `start_mpv_video_playback`), so
/// the pass it is drawing has already been running under the pointer: a bar begun at zero would
/// count the whole hover's stretch of the film again before the end of file it is meant to answer,
/// and every end of file — Pause, Loop and Continue alike — would be that stretch late.
#[test]
fn a_pin_taken_up_on_a_film_already_playing_starts_where_the_picture_is() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("adopted-pass-film.mkv");
    note_video_loop(101, &film, content, 0.0, Some(200.0));
    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        if let Some(loop_) = slot.as_mut() {
            loop_.at = Instant::now() - Duration::from_secs(30);
        }
    }

    let pin = take_up_pinned_window(&film, content);
    let reached = pin.transport.started.map(|(_, from)| from).unwrap_or(-1.0);
    assert!(
        (29.0..=31.0).contains(&reached),
        "the bar starts at the second the hover's player had reached, not at zero (was {reached})"
    );

    if let Ok(mut slot) = VIDEO_LOOP.lock() {
        *slot = None;
    }
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A film let go of behind a bubble goes on from the second it was held at, and what the bubble
/// shows while the read beside it begins again is the frame the film stopped on.**
///
/// The user's report, on a bubble they watch a film in: a film watched to the eleventh second,
/// held with SPACE, and let go of again with SPACE was showing the fifth — the second the pin was
/// collapsed at. Both halves of that are asked here, because either one could put the old second
/// back on the screen: the second the press re-bases the film onto, and the picture the bubble
/// falls back to while the read it has just begun is landing its first frame (see
/// `bubble_stream::settle` and `grab_bubble_frame`).
#[test]
fn a_film_let_go_of_behind_a_bubble_is_shown_the_second_it_was_held_at() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    // The film plays on behind the bubble: a bubble whose film is *held* is a bubble whose second
    // never moves, which is not the report (see `bubble_playback_to_park`).
    let _switches = PausePreviewSettings::stood_at(false, false);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();
    let previous_still = BUBBLE_FRAME.lock().ok().and_then(|mut frame| frame.take());
    bubble_stream::stand_no_stream();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("bubble-let-go-film.mkv");
    clear_restart_count();
    mpv_ipc::clear_noted();

    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

    // A pin collapsed on a film six seconds into a pass begun at the fifth second. The app's own
    // clock over the player's start is the only thing that says where such a film is, and it says
    // the eleventh second (see `pin_playhead`).
    let mut pin = a_playing_pin(&film, content);
    pin.transport.started = Some((Instant::now() - Duration::from_secs(6), 5.0));
    pin.collapsed = true;
    stand_pin(Some(pin));

    let box_ = (300, 300, 364, 364);
    let side = bubble_side(box_);

    // The read beside the player is begun where the film is.
    bubble_stream::settle(true, Some(box_));
    let asked = bubble_stream::take_noted_streams();
    assert_eq!(asked.len(), 1, "a playing film is read beside its player");
    assert!(
        (asked[0].1 - 11.0).abs() < 1.0,
        "from the second the film is at rather than the one the pin was collapsed at (was {})",
        asked[0].1
    );

    // The instant the pin's own player was begun at, which is what makes a read the one a tick is
    // looking for — read before the press below, because a film held without a key is a player
    // ended and a start given up (see `hold_pinned_player_without_a_key`).
    let begun = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport.started))
        .flatten()
        .expect("a player was begun")
        .0;

    // SPACE holds the film where it is.
    let mut started = None;
    let mut offset = 0.0;
    toggle_pinned_by_key(&mut started, &mut offset, &mut None);
    let held = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    assert!(
        (held.paused_at.unwrap_or(-1.0) - 11.0).abs() < 1.0,
        "the press holds the film at the second it is at (was {:?})",
        held.paused_at
    );

    // The read this bubble had, stood in with the frame it had just landed: a test build starts no
    // process, and the instant the pin's own player was begun at is what makes it the read the tick
    // is looking for (see `stand_a_running_stream`).
    let frame = [90u8, 90, 90, 255]
        .iter()
        .cycle()
        .take(side as usize * side as usize * 4)
        .copied()
        .collect::<Vec<u8>>();
    bubble_stream::stand_a_running_stream(film.clone(), begun, side, frame.clone());
    bubble_stream::settle(false, Some(box_));
    assert_eq!(
        bubble_stream::latest_frame(),
        Some((frame.clone(), side)),
        "a held film's bubble shows the frame it stopped on"
    );

    // SPACE again lets the film go from the second it was held at.
    toggle_pinned_by_key(&mut started, &mut offset, &mut None);
    let let_go = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport))
        .expect("a pin is up");
    let from = let_go.started.map(|(_, from)| from).unwrap_or(-1.0);
    assert!(
        (from - 11.0).abs() < 1.0,
        "the player is let go at the second it was held at (was {from})"
    );

    bubble_stream::settle(true, Some(box_));
    let asked = bubble_stream::take_noted_streams();
    assert_eq!(
        asked.len(),
        1,
        "and the film is read beside its player again"
    );
    assert!(
        (asked[0].1 - 11.0).abs() < 1.0,
        "from the second the film went on at (was {})",
        asked[0].1
    );

    // And what the bubble shows while that read lands its first frame is the frame the film
    // stopped on: a read that has just begun has no frame of its own yet, and the still the
    // collapse grabbed is the second the pin went down.
    match bubble_art(side).0 {
        BubbleArtOwned::Picture(pixels, _, _, _) => assert_eq!(
            pixels, frame,
            "a film let go of is shown the frame it stopped on rather than the still the collapse \
             grabbed"
        ),
        _ => panic!("a film let go of has a picture in its bubble"),
    }

    bubble_stream::stand_no_stream();
    if let Ok(mut still) = BUBBLE_FRAME.lock() {
        *still = previous_still;
    }
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

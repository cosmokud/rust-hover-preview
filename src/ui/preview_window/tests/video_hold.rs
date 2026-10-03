use super::*;

/// The wait for a video's player ends one of two ways and never by itself: a player
/// whose window is up has arrived, a player that is gone is not coming, and a start
/// that has run past the cap is given up on — while a player that is alive with no
/// window yet is a start that is still going.
#[test]
fn waits_for_a_player_only_until_it_is_there_or_gone() {
    let cap = Duration::from_secs(VIDEO_START_WAIT_SECS);

    assert_eq!(
        player_wait(true, true, Duration::ZERO),
        Some(PlayerWait::Arrived),
        "a player with a window up is the video"
    );
    assert_eq!(
        player_wait(false, true, Duration::from_secs(1)),
        None,
        "a player still starting is a wait that goes on"
    );
    assert_eq!(
        player_wait(true, false, Duration::ZERO),
        Some(PlayerWait::Abandoned),
        "what a dead player leaves behind is a handle, not a preview"
    );
    assert_eq!(
        player_wait(false, false, Duration::ZERO),
        Some(PlayerWait::Abandoned),
        "a player that is gone is not coming back to put a window up"
    );
    assert_eq!(
        player_wait(false, true, cap),
        Some(PlayerWait::Abandoned),
        "a start past the cap is not watched any longer"
    );
    assert_eq!(
        player_wait(true, true, cap),
        Some(PlayerWait::Arrived),
        "and a window that is up has arrived, cap or no cap"
    );
}

/// A held-back video's first frame comes up on one tick, and on that tick exactly one
/// painter is pointed at it: the reveal's, at the box the hover was laid out at. The tick
/// that would paint the same frame into the same surface at the window's own rectangle
/// stands down, which at the size of a 4K display is thirty megabytes of copying and a
/// hand-off to the compositor once a hover.
///
/// Every way of not being that tick is named on both sides, because the cost of getting it
/// wrong is a window put up with nothing in it rather than a wasted copy: a newer hover, a
/// hide since, a pin that has taken the window over, and a frame that has not landed are
/// four different reasons to leave the tick's own repaint standing.
#[test]
fn a_held_back_first_frame_is_painted_once_and_by_the_reveal() {
    let wait = FirstFrameWait {
        generation: 7,
        epoch: 3,
        pos: (120, 240),
    };

    assert_eq!(
        first_frame_lands(&wait, 7, 3, false, true),
        Some((120, 240)),
        "a frame that landed on this hover's own wait comes up at the box the hover laid \
             the preview out at"
    );

    assert_eq!(
        first_frame_lands(&wait, 8, 3, false, true),
        None,
        "a newer hover installed its own media, and this tick's repaint is what draws it"
    );
    assert_eq!(
        first_frame_lands(&wait, 7, 4, false, true),
        None,
        "a hide since is the pointer having left the file, so the frame is not put up at all \
             — but the tick's repaint is not stood down for a frame nobody is going to be shown"
    );
    assert_eq!(
        first_frame_lands(&wait, 7, 3, true, true),
        None,
        "a pin up owns the window, and its own painter is the tick's repaint"
    );
    assert_eq!(
        first_frame_lands(&wait, 7, 3, false, false),
        None,
        "a wait whose frame has not landed is a wait that goes on, and nothing is painted for \
             it by either painter"
    );
}

/// A swap of the pin's file ends one of two ways and never by itself: a frame the engine
/// has drawn is the video, and an engine that is gone or that has said the file is one it
/// cannot play are the end of it — while an engine that is playing and has drawn nothing
/// yet is a wait that goes on, for as long as it takes.
///
/// Both ends install the same file, so what the test is really pinning down is which of
/// them the pin is left frozen on the previous film for: an engine that is merely slow
/// costs as much time on the old picture as the engine takes, and each of the other two
/// costs a backdrop flash the hold was written to remove. Elapsed time is not an answer
/// and there is no longer any way to ask the question: the give-up it used to be measured
/// against fired on a cold read of a large file before the engine had opened it, and the
/// install it caused put the placeholder back on screen.
#[test]
fn a_swap_is_held_for_a_first_frame_until_it_arrives_or_the_engine_gives_it_up() {
    assert_eq!(
        pin_swap_wait(true, true, false),
        Some(PinSwapWait::Arrived),
        "a frame of the file in hand is the video, engine or no engine"
    );
    assert_eq!(
        pin_swap_wait(false, true, false),
        None,
        "an engine that has drawn nothing yet is a wait that goes on"
    );
    assert_eq!(
        pin_swap_wait(false, false, false),
        Some(PinSwapWait::Abandoned),
        "an engine that is gone is not going to draw a frame now"
    );
    assert_eq!(
        pin_swap_wait(false, true, true),
        Some(PinSwapWait::Abandoned),
        "an engine that has said it cannot play the file has said so on the first tick"
    );
    assert_eq!(
        pin_swap_wait(true, true, true),
        Some(PinSwapWait::Arrived),
        "and a frame that did arrive has arrived, whatever else is true of the engine"
    );
    assert_eq!(
        pin_swap_wait(false, false, true),
        Some(PinSwapWait::Abandoned),
        "a gone engine that also gave up is still just one abandoned swap, not two answers"
    );
}

/// A hold that has come to one of its ends is the tick's answer, and it is the *only* answer
/// that tick gives: a load that answered into it waits in its slot for the next one, which is
/// a tick away and has cost nothing.
///
/// Written over, the install is thrown away with the frame the engine had in hand — the film
/// is never installed, the pin stands on the file it was frozen on, and because the walk behind
/// what is on screen was never rewritten either, every caption step after it lands on the same
/// file. That is a pin whose **Next** and **Previous** do nothing at all.
#[test]
fn a_hold_that_has_come_to_its_end_is_the_answer_the_tick_gives() {
    let mut hold = Some(PinSwapHold {
        file: PinInstallable {
            path: PathBuf::from("the-film-reached-by-pressing-next.mp4"),
            update: PinUpdate {
                content: (10, 20, 210, 380),
                dpi: 96,
                volume: 100,
            },
            media: create_loading_media(200, 360),
            audio: None,
            walk: None,
        },
        arc: PinArc::new(),
    });
    let mut load = Some(PinLoad::answered(
        Path::new("later.png"),
        PinUpdate {
            content: (10, 20, 210, 380),
            dpi: 96,
            volume: 100,
        },
    ));

    // Nothing is playing behind this hold — a test has no engine — so the wait is given up on
    // at once, which is one of the two ends and installs the same file the other one does.
    let swap = settle_pin_swap(&mut hold, &mut load);

    assert!(
        matches!(swap, Some(PinSwap::Ready(_))),
        "the tick a hold comes to an end on installs the file it was holding"
    );
    assert!(
        load.is_some(),
        "a load that answered into a tick already installing a hold waits for the next one"
    );
}

/// A player that has stopped is read as one that reached the end of the file it was handed or
/// as one that never played it, and the difference is time: a stop well into the pass the
/// player was given is the end of the file, and a stop within a moment of starting is a
/// player that failed rather than a file that ended. What the file says about its length
/// bounds the pass and never overrules the moment, so the last of a short file — a pass
/// shorter than a moment — is asked only to have been played.
#[test]
fn a_player_that_has_stopped_is_read_as_the_end_of_its_file_or_not() {
    let length = Some(200.0);

    assert!(
        reached_the_end(Duration::from_secs(200), length, 0.0),
        "a player that played the whole file reached the end of it"
    );
    assert!(
        reached_the_end(Duration::from_secs(31), length, 170.0),
        "and so did one given only the last half minute of it"
    );
    assert!(
        !reached_the_end(Duration::from_millis(80), length, 0.0),
        "a player that stopped the moment it started never played the file, long or not"
    );
    assert!(
        !reached_the_end(Duration::from_millis(80), Some(2_000.0), 0.0),
        "and a file of any length is not read from a stop that short"
    );
    assert!(
        reached_the_end(Duration::from_millis(900), Some(0.5), 0.0),
        "a pass shorter than the moment is read by its own length: it cannot be lived past"
    );
    assert!(
        reached_the_end(Duration::from_secs(2), None, 0.0),
        "a file that says nothing about its length is read by the moment alone"
    );
    assert!(
        !reached_the_end(Duration::from_millis(200), None, 0.0),
        "which is still a moment a player that failed has stopped inside of"
    );
}

/// A probe's child is waited for with a deadline, and what a child that has outrun it
/// leaves behind is an answer of its own rather than a wait that goes on: the process is
/// ended where it stands and the caller is told there is nothing from this one, which is
/// the arm both probes already have for a child that answered nothing.
///
/// What stands in for the two children is an ordinary console program, since what the
/// helper does with a child is the same whatever the child is and a test is not going to
/// start ffmpeg. Neither stand-in reaches the network: a command that prints a line
/// finishes however loaded the machine is, and a ping of thirty replies is still there a
/// moment later whatever it is told.
#[test]
fn waits_for_a_probes_child_only_until_the_deadline() {
    let quick = engine_processes::hidden_command("cmd")
        .args(["/C", "echo", "a line from the child"])
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .expect("a child that finishes");

    let answered = wait_bounded(quick, Duration::from_secs(VIDEO_PROBE_TIMEOUT_SECS))
        .expect("a child that ends inside the deadline answers with its output");
    assert!(answered.status.success());
    assert!(
        String::from_utf8_lossy(&answered.stdout).contains("a line from the child"),
        "and what it wrote is in the answer, which is the pipe that was drained"
    );

    let slow = engine_processes::hidden_command("ping")
        .args(["-n", "30", "127.0.0.1"])
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .expect("a child that does not finish");
    let pid = slow.id();

    assert!(
        wait_bounded(slow, Duration::from_millis(50)).is_none(),
        "a child that has outrun the deadline is not waited for any longer"
    );
    assert!(
        !engine_processes::is_running(pid),
        "and it is ended rather than left reading a file nobody is waiting for"
    );
}

/// Two relaunches inside one wait — two presses of the track key, or a resize settling twice
/// — which is the only way a parked player can be displaced, and the only way one can be left
/// alive with nothing left to look at it.
#[test]
fn a_relaunch_inside_a_wait_never_leaves_a_player_nothing_is_waiting_on() {
    let first = Instant::now();
    let second = first + Duration::from_millis(40);

    // The ordinary relaunch: one wait, and no player ended by the relaunch itself.
    let (parked, ended) = retirement_after_relaunch(None, 100, 200, true, first);
    assert_eq!(
        (parked.map(|wait| (wait.retiring, wait.replacement)), ended),
        (Some((100, 200)), None),
        "a player that has put a window up is parked for its replacement and ended by the loop, \
             which is the whole of what parking it is for"
    );

    // A second relaunch while the first replacement is still on its way: the middle player
    // has no window of its own, so it is the one that goes and the oldest stays on screen.
    let waiting = VideoRetirement {
        retiring: 100,
        replacement: 200,
        started: first,
    };
    let (parked, ended) = retirement_after_relaunch(Some(waiting), 200, 300, false, second);
    assert_eq!(
        ended,
        Some(200),
        "the player in between has put no window up to lose, so it is ended at once rather than \
             left playing for as long as the app runs with nothing waiting on it"
    );
    assert_eq!(
        parked.map(|wait| (wait.retiring, wait.replacement)),
        Some((100, 300)),
        "and what stays parked is the oldest player — still the only picture on screen — with \
             the newest player waited for, so the band is never empty for a whole start"
    );
    assert_eq!(
        parked.map(|wait| wait.started),
        Some(first),
        "with the wait still bounded by the first of the starts rather than re-based on this \
             one, so a chain of relaunches cannot leave a player on screen for ever"
    );

    // And the other half of the same question: a second relaunch whose player *has* arrived
    // covers the one parked under it, so that one is what goes.
    let (parked, ended) = retirement_after_relaunch(Some(waiting), 200, 300, true, second);
    assert_eq!(
        ended,
        Some(100),
        "a player that has arrived covers the one under it, which has nothing left to be kept \
             on screen for"
    );
    assert_eq!(
        parked.map(|wait| (wait.retiring, wait.replacement)),
        Some((200, 300)),
        "and the wait is now for the player that is on screen, against the one replacing it"
    );

    // The same player parked again — which is the wait this record was built for — leaves both
    // ends where they are and ends nobody.
    let (parked, ended) = retirement_after_relaunch(Some(waiting), 100, 300, true, second);
    assert_eq!(
        (parked.map(|wait| (wait.retiring, wait.replacement)), ended),
        (Some((100, 300)), None),
        "the player that is already parked stays the one waited over, because it is the one on \
             screen, and the newer player is what has arrived to replace it"
    );
}

#[test]
fn a_next_track_press_walks_the_files_own_tracks_and_stops_where_there_are_none() {
    assert_eq!(
        next_subtitle(Some(0), 3),
        Some(1),
        "the second of three tracks is the one after the first"
    );
    assert_eq!(
        next_subtitle(Some(2), 3),
        Some(0),
        "and the press wraps at the last of them rather than running off the end"
    );
    assert_eq!(
        next_subtitle(Some(0), 1),
        Some(0),
        "a file with one track has a key that does nothing, which is the same answer a key \
             against a sound's card gives"
    );
    assert_eq!(
        next_subtitle(None, 3),
        Some(0),
        "a file nothing has chosen a track of yet steps onto its first"
    );
    assert_eq!(
        next_subtitle(Some(0), 0),
        None,
        "and a file with no subtitle streams is refused rather than sent a track that is not \
             there, because a refused stream specifier is a player that exits"
    );
}

#[test]
fn a_files_first_subtitle_track_is_the_one_the_player_would_pick_for_itself() {
    // The three answers below are ffprobe 9.0.2's verbatim, captured from real files rather
    // than written out by hand: a MatVuka with two subtitle tracks whose *second* is the
    // container's default, an MP4 of the same pair whose first is, and a Matroska with the
    // default cleared off both. The shape matters as much as the values — one `index=`, one
    // `codec_type=` and one `DISPOSITION:default=` line per stream, in that order, with the
    // disposition belonging to the stream above it — because the parser reads the disposition
    // as the stream it follows and would silently number the tracks against the wrong one if
    // the player ever answered in another order.
    let second_default = "index=0\n\
                             codec_type=video\n\
                             DISPOSITION:default=0\n\
                             index=1\n\
                             codec_type=audio\n\
                             DISPOSITION:default=0\n\
                             index=2\n\
                             codec_type=subtitle\n\
                             DISPOSITION:default=0\n\
                             index=3\n\
                             codec_type=subtitle\n\
                             DISPOSITION:default=1\n";
    let first_default = "index=0\n\
                            codec_type=video\n\
                            DISPOSITION:default=1\n\
                            index=1\n\
                            codec_type=audio\n\
                            DISPOSITION:default=1\n\
                            index=2\n\
                            codec_type=subtitle\n\
                            DISPOSITION:default=1\n\
                            index=3\n\
                            codec_type=subtitle\n\
                            DISPOSITION:default=0\n";
    let none_default = "index=0\n\
                           codec_type=video\n\
                           DISPOSITION:default=0\n\
                           index=1\n\
                           codec_type=audio\n\
                           DISPOSITION:default=0\n\
                           index=2\n\
                           codec_type=subtitle\n\
                           DISPOSITION:default=0\n\
                           index=3\n\
                           codec_type=subtitle\n\
                           DISPOSITION:default=0\n";

    // The two subtitle streams sit at absolute indices 2 and 3, so what is being read here is
    // FFmpeg's rule reproduced: the container's default if it marks one, the first of them if
    // it does not — and never the file's own stream numbering, which `-sst s:` does not use.
    let streams = parse_subtitle_streams(second_default);
    assert_eq!(
        streams.count, 2,
        "the two subtitle streams are counted among themselves"
    );
    assert_eq!(
        streams.first, 1,
        "and where the container's default is the *second* of them, that is the track the \
             player picks for itself — taking the first because it is first would show a different \
             language than the one the file asks for"
    );
    assert_eq!(
        streams.chosen(),
        Some(1),
        "so a relaunch before anything has been chosen names exactly what the player would \
             have picked anyway"
    );

    assert_eq!(
        parse_subtitle_streams(first_default),
        SubtitleStreams { count: 2, first: 0 },
        "the same two tracks in a file whose default is the first of them resolve to the \
             first, and the video and audio defaults above them are not mistaken for tracks: \
             `default = default.or(..)` only ever takes the *first* default marked, and only \
             while the stream it belongs to is a subtitle one"
    );

    assert_eq!(
        parse_subtitle_streams(none_default),
        SubtitleStreams { count: 2, first: 0 },
        "and a file that marks no stream as the default still has a first subtitle track, so \
             an unmarked file resolves to a track rather than to nothing"
    );

    // What the capture above cannot produce, and what a probe can still answer: nothing.
    assert_eq!(
        parse_subtitle_streams(""),
        SubtitleStreams::default(),
        "and a probe that answered nothing is a file with no known tracks, which is answered \
             as a file with none rather than guessed at"
    );
}

#[test]
fn a_key_posted_to_ffplays_player_carries_the_scan_code_its_symbol_is_read_from() {
    let lparam = ffplay_key_lparam(FFPLAY_PAUSE_KEY.1).0;

    assert_eq!(
        lparam >> 16,
        FFPLAY_PAUSE_KEY.1 as isize,
        "the scan code is what the posted message has to carry: the player is an SDL program \
             and SDL reads a Win32 key message's scan code to work out which key it names, so a \
             message posted without one names no key at all and pauses nothing"
    );
    assert_eq!(lparam & 1, 1, "and the repeat count of one press is one");
    assert_eq!(
        lparam & (1 << 30),
        0,
        "Windows' own auto-repeat bit is deliberately left clear, because a repeated pause \
             would toggle back to playing and a hand resting on the button would flicker instead \
             of holding"
    );
}

/// A loop is given by beginning the file again from the beginning, and by carrying enough about
/// the film to do it — which is what a key could never have done.
///
/// This is the regression test for a loop that was right on a short clip and wrong on every other.
/// The rewind used to be a posted `Down`, on the belief that the player's own binding for *seek
/// to beginning* was `Down`. Read out of `ffplay.c` for the exact build on this machine (tag
/// `n9.0.2`, `event_loop`), `Down` is `incr = -60.0` — sixty seconds back — and every other key
/// bound to a seek is a fixed increment too: `Left`/`Right` ten, `Up` sixty, `PageUp`/`PageDown`
/// six hundred or a chapter. There is no key that names a second, so there is no key that goes
/// to the beginning. On a clip shorter than a step the seek clamps to the start and looks exactly
/// like the go-to-start it was believed to be; on a film of ten minutes it is a seek to nine
/// minutes, then to eight, while this app's own clock — zeroed by the tick — says it is at the
/// beginning throughout.
///
/// So the second is asserted for a film of every length, and so is the fact that the record
/// carries the file and the box, because a relaunch cannot be made out of a pid and a clock: the
/// first version of this took a `path` and threw it away.
#[test]
fn a_loop_is_begun_again_from_the_beginning_and_carries_what_the_relaunch_needs() {
    // Begun inside the margin the tick acts in, at every length: the record's clock is what says
    // how near the end the film is, and it is the same arithmetic whatever the file's length is.
    let passed = Duration::from_millis(95);

    for duration in [8.0, 61.0, 600.0, 7_200.0] {
        let looped = VideoLoop {
            pid: 1,
            path: PathBuf::from("film.mkv"),
            content: (10, 20, 410, 320),
            from: duration - 0.15,
            duration: Some(duration),
            at: Instant::now() - passed,
        };

        let VideoLoopAction::Rewind {
            path,
            content,
            seconds,
        } = video_loop_action(&looped, true)
        else {
            panic!(
                "a film of {duration}s within a tenth of a second of its end has to be begun \
                     again, or it reaches the end of the file and exits"
            );
        };

        assert_eq!(
            seconds, 0.0,
            "a loop begun again at {seconds} of a {duration}s film is not a loop: this player \
                 has no key that goes to the beginning, every seek key it binds is a fixed increment, \
                 and `Down` in particular is sixty seconds back — which is the whole film on a \
                 sixty-second clip and nine minutes of walking backwards on a ten-minute one"
        );
        assert_eq!(
            path,
            PathBuf::from("film.mkv"),
            "the rewind is a relaunch, so the file has to survive in the record: taking a path \
                 and dropping it is a loop that can be timed but never taken"
        );
        assert_eq!(
            content,
            (10, 20, 410, 320),
            "and the box the player's window fills, because the window is another program's and \
                 the replacement has to be told where to be put rather than laid out by this app"
        );
    }
}

/// A held film is neither begun again nor left with a clock that has run on underneath the hand.
///
/// A drag holds the player with a pause key (see `video_drag_hold_apply`), so a film a second
/// from its end would otherwise be begun again *underneath the hand holding it* — and the second
/// it was held at would keep counting up in the clock underneath the record, so that a film held
/// for a minute near its end was past its end the moment it was let go of and was rewound on the
/// next tick. That is the same arithmetic a release does to the transport's own clock, and for
/// the same reason.
#[test]
fn a_held_film_is_neither_begun_again_nor_left_with_a_clock_that_ran_on() {
    let looped = VideoLoop {
        pid: 1,
        path: PathBuf::from("film.mkv"),
        content: (10, 20, 410, 320),
        from: 7.95,
        duration: Some(8.0),
        at: Instant::now() - Duration::from_millis(95),
    };

    assert!(
        matches!(
            video_loop_action(&looped, false),
            VideoLoopAction::Held
        ),
        "a film a twentieth of a second from its end, held, must not be begun again underneath \
             the hand holding it: the hold is the user's, and a loop that ignores it is a loop that \
             undoes it"
    );

    assert!(
        matches!(
            video_loop_action(&looped, true),
            VideoLoopAction::Rewind { .. }
        ),
        "and the very same film playing *is* due, so the refusal above is about the hold and not \
             about a decision that always says no"
    );
}

//! The frames a bubble shows of a film a player of somebody else's is drawing: one decode beside
//! that player, read as it comes.
//!
//! **A film another process draws is on the screen and nowhere else.** That window is a player's
//! rather than this app's, so unlike a still a frame cannot be looked up in memory — and neither
//! player can hand one over: `ffplay` has no interface at all, and mpv's own answer cannot cross
//! its pipe (see `mpv_ipc`). The picture a bubble can show of such a film *moving* is therefore a
//! second read of the file, and the whole of this module is making that read cheap: **one** decode,
//! begun where the player is, throwing away everything but the newest frame, scaled and cropped to
//! the bubble's own square and capped at the rate a bubble is looked at.
//!
//! **And the two players are read the same way, which is the whole of why this is not two
//! modules.** What a bubble is around is a window of somebody else's, and moving that window into a
//! forty-four pixel circle is a fight with its own region, its own idea of being topmost and its own
//! resizes — the picture a bubble paints itself has none of those, and is the one road that looks
//! the same on either player. The second it is begun at is the transport's own clock rather than a
//! question put to the player, which on this route means the frame follows the clock the bar is
//! drawn from: an mpv whose own opening lag is a few hundred milliseconds is a bubble following it
//! by that much, and asking mpv for its position is the change that would close that gap (see
//! `wanted_stream`).
//!
//! **The player's own start is the key a stream is kept for.** A stream is a decode of *a play* —
//! this film, from this second, in this square — so what it is compared against is the transport's
//! own record of when the player was begun: a level change or a seek on FFmpeg's route kills the
//! player and begins another at a new second, and on mpv's the transport is rebased onto the second
//! the seek went to (see `restart_pinned_player` and `seek_pinned_playback`). Either way it is a new
//! start, and so a stream that has to be begun again. A stream that has gone is not tried again
//! under the same key, because a command the build refuses would otherwise be a process every tick;
//! a player begun afresh is a new key and is streamed again.
//!
//! **Nothing here is a frame the bubble must have.** The still grabbed at the collapse stands until
//! a frame lands, and stands again wherever there is no frame of the film to hold — a stream that
//! cannot start, a player whose film has gone, a build whose decoder refuses the command (see
//! `grab_bubble_frame` and `bubble_art`): each of them is a bubble showing what every bubble showed
//! before this module existed. A pause is the one thing that puts a frame in that place: the frame
//! the film stopped on, which is the newest frame of the film this app has, so that a film let go of
//! again is not shown the second the pin went down (see `keep_bubble_frame`).

use super::*;
use std::process::ChildStdout;

/// The rate the frames are asked for at, and the whole of what keeps a second decode cheap: a bubble
/// is forty-four pixels across, so a film's own rate is more frames than a circle can show, and the
/// pipe, the reads and the paints are all paid for per frame.
const BUBBLE_STREAM_FPS: u32 = 25;

/// The film a stream can be kept for: where, which of the player's own runs it belongs to, and the
/// square it is scaled into.
///
/// The second it *began* at is deliberately not here — that is where a stream starts, and a stream
/// that is running has left it behind (see `Wanted`).
#[derive(Clone, PartialEq, Eq)]
struct StreamKey {
    path: PathBuf,
    started: Instant,
    side: u32,
}

/// A stream the bubble wants: the key it is kept for, and the second it starts at.
struct Wanted {
    key: StreamKey,
    at: f64,
}

/// A stream that is running: the child, the key it was begun for, and the newest frame it has
/// landed.
struct Running {
    /// The decode writing the frames, or nothing for a stream a test build has stood in: the suite
    /// starts no process, and what a test needs is a stream whose frames are coming (see
    /// `stand_a_running_stream`).
    child: Option<Child>,
    key: StreamKey,
    /// The newest frame, written by the reader thread and read by the bubble's paint.
    ///
    /// One slot rather than a queue, because a bubble shows the newest thing there is and a frame
    /// that has been passed by is a frame nobody will look at: a decoder that gets ahead of the paint
    /// costs nothing beyond the frames it drops (see `read_frames`).
    latest: Arc<Mutex<Option<Vec<u8>>>>,
}

/// The stream as it stands.
enum Stream {
    /// A stream whose frames are coming.
    Running(Running),
    /// A stream whose film has *stopped*, and the frame it stopped on: the picture the bubble was
    /// painting a moment ago, held so that a paused bubble shows the second the film was paused at
    /// rather than the still the collapse grabbed (see `settle` and `latest_frame`).
    Held {
        key: StreamKey,
        latest: Option<Vec<u8>>,
    },
    /// A stream that was asked for and did not come: the key it was refused for, so that it is asked
    /// for once and not once a tick (see the module documentation).
    Refused(StreamKey),
}

/// The stream, or nothing while none is.
static STREAM: Lazy<Mutex<Option<Stream>>> = Lazy::new(|| Mutex::new(None));

/// The newest frame read out of the film, and the side it was cut to, or nothing while the bubble
/// has no such picture: what the paint asks, and what it answers with the still when there is none.
///
/// **The frame a film *stopped* on is this bubble's picture, and a frame nothing has replaced is
/// not.** A stream whose child has gone leaves nothing here rather than its last frame, or a bubble
/// would show a film that ran out as though it were still playing (see `end` and `read_frames`) —
/// while a stream whose film was *paused* holds the frame it stopped on, which is where the picture
/// is rather than what the collapse grabbed (see `Stream::Held`).
pub(super) fn latest_frame() -> Option<(Vec<u8>, u32)> {
    // A frame stood in for the one a real child would have written, for the tests that are about what
    // a bubble paints rather than about the reading of it (see `stand_live_frame`).
    #[cfg(test)]
    if let Some(stood_in) = stood_in_frame() {
        return Some(stood_in);
    }

    let stream = STREAM.lock().ok()?;
    match stream.as_ref()? {
        Stream::Running(running) => {
            let frame = running.latest.lock().ok()?.clone()?;
            Some((frame, running.key.side))
        }
        Stream::Held { key, latest } => latest.clone().map(|frame| (frame, key.side)),
        Stream::Refused(_) => None,
    }
}

/// Make the stream the one the bubble needs, which is the whole of what a tick has to do about it.
///
/// **The answer is the state of the world, not an instruction**: whether a pin is a bubble whose film
/// is playing, where the bubble is standing and which film the player is at are all read here, and
/// what is running is made to match them. That is what keeps a level change, a seek, a loop, a film
/// swapped in behind the bubble and a bounce between two displays from each needing their own hook:
/// each of them is a different answer to a question this asks every tick (see the bubble repaint in
/// `run_preview_window`).
pub(super) fn settle(playing: bool, box_: Option<ScreenRegion>) {
    let wanted = box_.filter(|_| playing).and_then(wanted_stream);

    // **A film that has *stopped* is a frame of its own**, and that is what separates the two ways a
    // stream ends: a film the user paused — or one the player has stopped at its own end — is where
    // the picture is, and the frame it stopped on is the one the bubble was painting a moment ago;
    // a bubble showing the still instead would be showing the second the pin was collapsed at, which
    // for a film watched for half an hour is half an hour behind the key the user pressed (see
    // `Stream::Held`). A stream that is ending for any other reason — the film swapped under the
    // bubble, the bubble gone, the player's window closed — keeps nothing.
    let keep_a_stopped_films_frame = !playing && box_.is_some();

    let Ok(mut stream) = STREAM.lock() else {
        return;
    };
    let held = stream.take();
    drop(stream);

    // What is already there is answered first, and answered without a window call, a process or a
    // pipe: a stream that is still the one wanted and whose child is still running is left exactly as
    // it is, and so is a refusal for this very key — a command the build would not take is not worth a
    // process a tick (see the module documentation).
    match (held, &wanted) {
        (Some(Stream::Running(mut running)), Some(wanted)) if running.key == wanted.key => {
            if alive(&mut running) {
                return hold(Stream::Running(running));
            }
            // The child this key was begun for has gone: nothing is running, and this film at this
            // point of this player's run is not asked for again (see the module documentation). The
            // frames it had landed go with it, so that the bubble paints its still.
            let key = running.key.clone();
            end(running);
            return refuse(key);
        }
        (Some(Stream::Held { key, .. }), Some(wanted)) if key == wanted.key => {
            // A film going on again at the same point of the same player's run: the frame that has
            // stopped is dropped rather than handed on, because a picture that has stopped is not a
            // picture that is moving, and the read begins again below.
        }
        (Some(Stream::Held { key, latest }), _)
            if keep_a_stopped_films_frame && the_film_the_pin_shows_is(&key.path) =>
        {
            // And a film that is *still* stopped is still where it stopped: this is asked of every
            // tick the bubble stands through the pause and not of the one the pause happened on, and
            // a frame kept for that one tick is a bubble back on the still the collapse grabbed the
            // moment after (see `Stream::Held`).
            return hold(Stream::Held { key, latest });
        }
        (Some(Stream::Refused(key)), Some(wanted)) if key == wanted.key => {
            return hold(Stream::Refused(key));
        }
        (held, _) => {
            // Everything else is over: a child that is still there is ended. Its newest frame is
            // kept as the bubble's picture where the film has *stopped* on it — a child of this
            // app's own that the pause above is about to end, of the film the pin is showing now —
            // and dropped where it has not (see the note over `keep_a_stopped_films_frame`).
            if let Some(Stream::Running(mut running)) = held {
                let stopped = keep_a_stopped_films_frame
                    && alive(&mut running)
                    && the_film_the_pin_shows_is(&running.key.path);
                let stopped_on = stopped
                    .then(|| running.latest.lock().ok().and_then(|mut slot| slot.take()))
                    .flatten();
                let key = running.key.clone();
                end(running);

                if let Some(frame) = stopped_on {
                    // And the bubble falls back to it as well as showing it: a film let go of again
                    // begins a read whose first frame is a moment away, and what the bubble shows
                    // until it lands has to be the second the film is at rather than the one the pin
                    // was collapsed at (see `keep_bubble_frame`).
                    keep_bubble_frame(&frame, key.side);
                    return hold(Stream::Held {
                        key,
                        latest: Some(frame),
                    });
                }
            }
        }
    }

    let Some(wanted) = wanted else {
        return;
    };

    hold(start(wanted));
}

/// What the bubble's stream should be, where there is one to have: a film a player of somebody
/// else's is drawing, at the second that player is at, in the square the bubble is standing in.
///
/// **Both players are read this way and the app's own engine is not.** A film this app's engine
/// draws is in this app's own window and its frame is in this app's hands already, so a second
/// decode of it would be a picture nobody is looking at (see `bubble_art`); a film either external
/// player draws has no frame here at all, and mpv's being *askable* changes nothing about that —
/// what it can be asked is a position, not a picture. Nothing is asked of the machine here — no
/// probe, no process — which is what lets this sit on a tick.
///
/// The second is the *playhead* rather than the second the player began at, because a film pinned
/// half an hour ago and collapsed now is a bubble that has to show the film where it is: the stream
/// is a decode of the same play, begun at the same point in it, and the two read at the same rate
/// from there.
///
/// **The playhead is this app's clock over the player's start and never the player's own answer.**
/// It is the same second the pin's bar is drawn from, so what a collapse leaves is a bubble in step
/// with the transport the user is looking at rather than with a second read of anything — and the
/// one lag it cannot know about is the player's own: mpv takes a moment to open a file and begin,
/// which this clock has already started counting. shortcut: mpv's own `time-pos` over its pipe is
/// what would close that gap where it is ever visible, at the cost of a read on a pipe that is
/// write-only by decision (see `mpv_ipc`).
fn wanted_stream(box_: ScreenRegion) -> Option<Wanted> {
    if current_media_type() != Some(MediaType::Video) {
        return None;
    }

    let (path, _) = pinned_media_owner()?;

    let started =
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.transport.started))??;
    let at = pinned_playhead()?;

    Some(Wanted {
        key: StreamKey {
            path,
            started: started.0,
            side: bubble_side(box_),
        },
        at,
    })
}

/// Whether a stream's decodes are still coming: what tells a stream that is a picture from one
/// whose decodes have stopped.
///
/// A stream a test build has stood in has no child behind it and is one whose frames are coming
/// (see `Running::child`).
fn alive(running: &mut Running) -> bool {
    match running.child.as_mut() {
        Some(child) => matches!(child.try_wait(), Ok(None)),
        None => true,
    }
}

/// Whether the film a stream was reading is the film the pin is showing *now*: what says a frame
/// read beside a player may stand as this bubble's picture at all.
///
/// It is asked of the *file*, because the file is what changes under a bubble: a walk landing on
/// another film behind a collapsed pin is a bubble that must not be left showing a frame of the one
/// before it (see `settle`).
fn the_film_the_pin_shows_is(path: &Path) -> bool {
    pinned_media_owner().is_some_and(|(pinned, _)| pinned.as_path() == path)
}

/// Begin a stream, and answer what is running — which is a refusal where the command did not start
/// at all, since a child that is not there is a stream the bubble must not ask for again and again.
fn start(wanted: Wanted) -> Stream {
    // **The suite starts no process at all**, and a decode is a process: what a test build answers is
    // the ask — the film, the second and the side — and the bubble keeps its still (see
    // `NOTED_STREAMS` and `mpv_ipc::deliver`, where the same rule is kept for the pipe).
    #[cfg(test)]
    {
        if let Ok(mut noted) = NOTED_STREAMS.lock() {
            noted.push((wanted.key.path.clone(), wanted.at, wanted.key.side));
        }
        return Stream::Refused(wanted.key);
    }

    #[cfg(not(test))]
    {
        // No `-hwaccel` on this child either: the flag is the same one that kills ffplay's own
        // renderer, and a second decode of a film on a core costs the bubble a frame rate nobody
        // looks at (see `video_hw_accel_wanted`). A pipe this cannot open is a bubble left on its
        // still, which is the picture the bubble always has.
        let args = stream_args(&wanted.key.path, wanted.at, wanted.key.side);

        let spawned = engine_processes::hidden_command("ffmpeg")
            .args(&args)
            .stdin(Stdio::null())
            .stdout(Stdio::piped())
            .stderr(Stdio::null())
            .spawn();

        let Ok(mut child) = spawned else {
            return Stream::Refused(wanted.key);
        };
        let _ = engine_processes::adopt(child.id());

        let Some(stdout) = child.stdout.take() else {
            let _ = child.kill();
            return Stream::Refused(wanted.key);
        };

        let latest = Arc::new(Mutex::new(None));
        let reader = Arc::clone(&latest);
        let side = wanted.key.side;
        std::thread::spawn(move || read_frames(stdout, side, reader));

        Stream::Running(Running {
            child: Some(child),
            key: wanted.key,
            latest,
        })
    }
}

/// The command a stream is begun with: the film from the second the player is at, read at the rate a
/// bubble is looked at, cut to the bubble's own square.
///
/// **The seek comes before the file and the rate with it**, because both are about the reading rather
/// than about the writing: `-ss` before `-i` is a seek into the index instead of a decode from the
/// start, and `-re` is the input read at the rate it was recorded at — without it this would be a
/// decode as fast as the disk allows, throwing away all but one frame in twenty-five, which is a core
/// burnt for a picture nobody is looking at any sooner.
///
/// The scale and the crop are the same pair every other bubble's picture is cut to (see
/// `grab_video_frame`): the square the circle is inscribed in, filled rather than letterboxed, so
/// that a bubble shows the middle of the film rather than a stripe across the middle of the circle.
fn stream_args(path: &Path, at: f64, side: u32) -> Vec<String> {
    let mut args: Vec<String> = Vec::new();

    args.extend(
        [
            "-ss".to_string(),
            format!("{at:.3}"),
            "-re".to_string(),
            "-i".to_string(),
            path.to_string_lossy().into_owned(),
            // No sound: the film's own player is making it, and a second soundtrack would be two
            // copies of one film's audio heard at once.
            "-an".to_string(),
            "-vf".to_string(),
            format!(
                "scale={side}:{side}:force_original_aspect_ratio=increase,crop={side}:{side},fps={BUBBLE_STREAM_FPS}"
            ),
            // Raw BGRA on the pipe, which is what the bubble paints with: a picture is not encoded
            // and decoded again across a process boundary when the bytes are what is wanted either
            // side of it (see `grab_video_frame`).
            "-f".to_string(),
            "rawvideo".to_string(),
            "-pix_fmt".to_string(),
            "bgra".to_string(),
            "-".to_string(),
        ],
    );

    args
}

/// Read the frames the child writes, keeping the newest and leaving the rest behind.
///
/// **A frame is only ever stored whole**, and the newest is the only one kept: what a bubble shows is
/// where the film is now, so a frame the paint has not reached is a frame that has been overtaken by
/// the very next one. The read is of exactly one frame's worth of bytes, which is what makes the
/// frames separable at all — rawvideo has no framing of its own (see `stream_args`).
///
/// On the way out — the child gone, or the film ended — the slot is emptied: a last frame left
/// standing is a bubble showing a film that has stopped as though it were still playing, and the
/// still the bubble keeps for exactly this is a frame of the film's own with a mark on it.
#[cfg_attr(test, allow(dead_code))] // reached only from a `start` a test build refuses (see `Stream::Running`)
fn read_frames(stdout: ChildStdout, side: u32, latest: Arc<Mutex<Option<Vec<u8>>>>) {
    use std::io::Read;

    let mut reader = stdout;
    let size = side as usize * side as usize * 4;
    let mut frame = vec![0u8; size];

    loop {
        if reader.read_exact(&mut frame).is_err() {
            break;
        }

        if let Ok(mut slot) = latest.lock() {
            *slot = Some(frame.clone());
        }

        // And the loop is told to paint again: what the bubble shows is a new frame, and nothing else
        // about a bubble over a film at a level of nothing is moving (see `BUBBLE_FRAME_READY`).
        BUBBLE_FRAME_READY.store(true, Ordering::Release);
    }

    if let Ok(mut slot) = latest.lock() {
        *slot = None;
    }
    BUBBLE_FRAME_READY.store(true, Ordering::Release);
}

/// End a stream: the child is killed and waited for, because a handle dropped is not a process
/// ended, and the frames it had landed are dropped with it (see `latest_frame`).
fn end(mut running: Running) {
    if let Some(mut child) = running.child.take() {
        let _ = child.kill();
        let _ = child.wait();
    }

    if let Ok(mut slot) = running.latest.lock() {
        *slot = None;
    }
    BUBBLE_FRAME_READY.store(true, Ordering::Release);
}

/// Put a stream back where it is read from, and tell the loop what the bubble shows has changed —
/// which is what a stream that has just been ended has done (see `end`).
fn hold(stream: Stream) {
    if let Ok(mut slot) = STREAM.lock() {
        *slot = Some(stream);
    }
}

/// Refuse a key: nothing is running, and this film at this point of this player's run is not asked
/// for again (see the module documentation).
fn refuse(key: StreamKey) {
    BUBBLE_FRAME_READY.store(true, Ordering::Release);
    hold(Stream::Refused(key));
}

/// The streams a test build has been asked for, with the second and the side each was asked at (see
/// `start`).
#[cfg(test)]
static NOTED_STREAMS: Mutex<Vec<(PathBuf, f64, u32)>> = Mutex::new(Vec::new());

/// The streams noted so far (see `NOTED_STREAMS`).
#[cfg(test)]
pub(super) fn take_noted_streams() -> Vec<(PathBuf, f64, u32)> {
    NOTED_STREAMS
        .lock()
        .map(|mut noted| std::mem::take(&mut *noted))
        .unwrap_or_default()
}

/// Put the stream back to nothing, as a run that has never had one: what a test of the routes either
/// side of this module stands in (see `stand_no_stream`).
#[cfg(test)]
pub(super) fn stand_no_stream() {
    if let Ok(mut slot) = STREAM.lock() {
        if let Some(Stream::Running(running)) = slot.take() {
            end(running);
        }
    }
    take_noted_streams();
    stand_live_frame(None);
}

/// Stand a stream in as one whose frames are coming, with one of them landed: what a test needs to
/// ask what a bubble does when a stream *stops*, since a test build starts no process and so never
/// has one of its own (see `start`).
///
/// The key is the caller's, and it is the whole of what makes the stood-in stream the film a tick
/// is looking for: `started` is the instant the pin's own transport was begun at, so a test can
/// stand in the very stream a tick would have wanted (see `StreamKey`).
#[cfg(test)]
pub(super) fn stand_a_running_stream(path: PathBuf, started: Instant, side: u32, frame: Vec<u8>) {
    if let Ok(mut slot) = STREAM.lock() {
        if let Some(Stream::Running(running)) = slot.take() {
            end(running);
        }
        *slot = Some(Stream::Running(Running {
            child: None,
            key: StreamKey {
                path,
                started,
                side,
            },
            latest: Arc::new(Mutex::new(Some(frame))),
        }));
    }
}

/// Stand frames in for the ones a real child would have written, for the tests that are about what a
/// bubble paints rather than about the reading of it: a child is a process, and the suite starts none
/// (see `latest_frame`).
#[cfg(test)]
pub(super) fn stand_live_frame(frame: Option<(Vec<u8>, u32)>) {
    if let Ok(mut stood_in) = STOOD_IN_FRAME.lock() {
        *stood_in = frame;
    }
}

/// The frames [`latest_frame`] answers with while a test stands them in (see `stand_live_frame`).
#[cfg(test)]
fn stood_in_frame() -> Option<(Vec<u8>, u32)> {
    STOOD_IN_FRAME.lock().ok().and_then(|frame| frame.clone())
}

/// The frames a test has stood in (see `stand_live_frame`).
#[cfg(test)]
static STOOD_IN_FRAME: Mutex<Option<(Vec<u8>, u32)>> = Mutex::new(None);

#[cfg(test)]
mod tests {
    use super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME;
    use super::*;

    /// The command a stream is begun with: the film from the second the player is at, at the rate a
    /// bubble is looked at, cut to the bubble's own square.
    #[test]
    fn the_command_reads_the_film_from_the_second_the_player_is_at_cut_to_the_bubbles_square() {
        let film = PathBuf::from("a-film-a-bubble-shows-playing.mkv");

        assert_eq!(
            stream_args(&film, 92.5, 44),
            [
                "-ss",
                "92.500",
                // The rate with the seek, because both are about the reading: `-re` is the film read
                // as it was recorded rather than as fast as the disk allows (see `stream_args`).
                "-re",
                "-i",
                "a-film-a-bubble-shows-playing.mkv",
                // No sound: the film's own player is making it.
                "-an",
                "-vf",
                "scale=44:44:force_original_aspect_ratio=increase,crop=44:44,fps=25",
                "-f",
                "rawvideo",
                "-pix_fmt",
                "bgra",
                "-",
            ]
            .into_iter()
            .map(String::from)
            .collect::<Vec<String>>(),
            "the command is the one the spec settles on, and every part of it is what makes a \
             second decode of a film cheap enough to sit behind a bubble"
        );
    }

    /// **What the bubble's stream is, is the state of the world every tick**: a film a player of
    /// somebody else's is drawing is read beside it while the pin is a bubble and the film is
    /// playing, on either player, and nothing is asked for any other kind, any other playback or any
    /// other place.
    ///
    /// The ask is what a test build answers — no child is ever started (see `start`) — so what is
    /// asserted is the decision: which film, at which second, at which side, and how often.
    #[test]
    fn a_film_a_player_is_drawing_is_read_beside_it_while_it_plays() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
        let previous_pin = take_pin_for_a_test();
        stand_no_stream();
        stand_ffplay_in(Some(true));
        stand_mpv_in(Some(false));
        stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

        let a_pin_playing_a_film = |at: f64| {
            let mut pin = PinnedPreview {
                bound: Some(64),
                content: (100, 100, 164, 164),
                chrome: PinChrome::always(),
                ..PinnedPreview::for_test()
            };
            pin.path = PathBuf::from("a-film-a-bubble-shows-playing.mkv");
            pin.transport.begun(at, true, false);
            pin
        };
        stand_pin(Some(a_pin_playing_a_film(30.0)));

        let mut media = create_loading_media(320, 240);
        media.media_type = MediaType::Video;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }

        let box_ = (300, 300, 364, 364);

        settle(true, Some(box_));
        let asked = take_noted_streams();
        assert_eq!(asked.len(), 1, "a playing film is read beside its player");
        let (film, at, side) = &asked[0];
        assert_eq!(film, &PathBuf::from("a-film-a-bubble-shows-playing.mkv"));
        assert!(
            *at >= 30.0,
            "from the second the player is at ({at}) rather than from the top of the film: a \
             bubble shows where the film is now"
        );
        assert_eq!(
            *side, 44,
            "and cut to the circle's own side on this display"
        );

        // The same tick again, and the same answer: a stream that is already running is not begun
        // again, which is the whole of what keeps a tick from being a process.
        settle(true, Some(box_));
        assert!(
            take_noted_streams().is_empty(),
            "a stream that is already the one wanted is left exactly as it is"
        );

        // A film nothing is playing has no picture to keep up with: its bubble is the still.
        settle(false, Some(box_));
        settle(false, None);
        assert!(
            take_noted_streams().is_empty(),
            "a paused film is not read: its picture is where it stopped, and that is the still's"
        );

        // And a player begun afresh — a level change, a seek, a loop — is a new stream, because it is
        // another play of the film even when it is the same film.
        stand_pin(Some(a_pin_playing_a_film(31.0)));
        settle(true, Some(box_));
        assert_eq!(
            take_noted_streams().len(),
            1,
            "a player begun again is read again, from the second it was begun at"
        );

        // A film mpv is drawing is read the same way: the window this used to stand in the bubble is
        // not carried anywhere, and what a bubble shows of that film is a picture it paints itself
        // like any other.
        stand_mpv_in(Some(true));
        stand_video_engine_in(Some((VideoEngine::Mpv, true)));
        stand_pin(Some(a_pin_playing_a_film(32.0)));
        settle(true, Some(box_));
        assert_eq!(
            take_noted_streams().len(),
            1,
            "a film mpv is drawing is read beside it, because mpv can be told nothing and asked \
             for no picture"
        );
        stand_mpv_in(Some(false));
        stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

        // Nor is anything that is not a film: a picture's bubble is the picture.
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            if let Some(media) = current.as_mut() {
                media.media_type = MediaType::StaticImage;
            }
        }
        settle(true, Some(box_));
        assert!(
            take_noted_streams().is_empty(),
            "a still is not read beside anything"
        );

        stand_no_stream();
        stand_pin(previous_pin);
        stand_video_engine_in(None);
        stand_mpv_in(None);
        stand_ffplay_in(None);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = previous_media;
        }
    }

    /// **A film that has stopped keeps the frame it stopped on.** The picture a paused bubble shows
    /// is the one it was painting a moment ago, not the still the collapse grabbed: a film watched
    /// for half an hour and then paused is half an hour past the frame that still was taken from, and
    /// what a hand that pressed the key is looking at is where the film stopped (see `Stream::Held`).
    ///
    /// The three ways out of that rule are here beside it, because each of them is a way for the
    /// answer to be nothing: a film that is *playing* again is read again rather than left standing
    /// at the frame it stopped on, another film under the bubble is not left showing this one's
    /// frame, and a bubble that is not on screen is owed no picture at all. The rule is also asked
    /// of every tick the pause lasts rather than of the tick the pause happened on, because a frame
    /// kept for one tick is a bubble back on its still the moment after.
    #[test]
    fn a_film_paused_while_collapsed_is_left_on_the_frame_it_stopped_on() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
        let previous_pin = take_pin_for_a_test();
        // The pause below hands the frame it keeps to the bubble as well as showing it, and a test
        // that leaves a still in that slot is a test every test after it is standing inside (see
        // `keep_bubble_frame`).
        let previous_still = BUBBLE_FRAME.lock().ok().and_then(|mut frame| frame.take());
        stand_no_stream();
        stand_ffplay_in(Some(true));
        stand_mpv_in(Some(false));
        stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

        let film = PathBuf::from("a-film-a-bubble-shows-playing.mkv");
        let mut pin = PinnedPreview {
            bound: Some(64),
            content: (100, 100, 164, 164),
            chrome: PinChrome::always(),
            ..PinnedPreview::for_test()
        };
        pin.path = film.clone();
        pin.transport.begun(30.0, true, false);
        stand_pin(Some(pin));

        let mut media = create_loading_media(320, 240);
        media.media_type = MediaType::Video;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }

        // The stream this tick *would* want, stood in with a frame that has just landed: the instant
        // the pin's own player was begun at is what makes these two the same stream (see
        // `stand_a_running_stream`).
        let box_ = (300, 300, 364, 364);
        let side = bubble_side(box_);
        let started = pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.transport.started))
            .flatten()
            .expect("the pin's player was begun")
            .0;
        let frame = vec![0x5a; side as usize * side as usize * 4];
        stand_a_running_stream(film.clone(), started, side, frame.clone());

        // While the film is playing the stream is the one wanted, and a stream that is already there
        // is not begun again.
        settle(true, Some(box_));
        assert_eq!(latest_frame(), Some((frame.clone(), side)));
        assert!(
            take_noted_streams().is_empty(),
            "a stream that is already the one wanted is left exactly as it is"
        );

        // And the pause, which is what this test is: the bubble is still standing where it was, and
        // the film is not playing — so the frame it stopped on is what the bubble shows.
        settle(false, Some(box_));
        assert_eq!(
            latest_frame(),
            Some((frame.clone(), side)),
            "a paused film's bubble shows the frame the film was at, not the still the collapse \
             grabbed half an hour before it"
        );

        // And the pause lasts more than one tick: the loop asks this of every tick the bubble is up,
        // and a frame kept for only the tick the pause happened on is a bubble back on the still the
        // collapse grabbed the moment after (see `Stream::Held`).
        settle(false, Some(box_));
        settle(false, Some(box_));
        assert_eq!(
            latest_frame(),
            Some((frame.clone(), side)),
            "the frame a film stopped on is held for as long as the film is stopped"
        );

        // The film going on again: a decode begun again rather than the frame left standing, because
        // a picture that has stopped is not a picture that is moving.
        settle(true, Some(box_));
        assert_eq!(
            take_noted_streams().len(),
            1,
            "a paused film let go of is read again, from the second it is let go at"
        );
        assert!(
            latest_frame().is_none(),
            "and the held frame goes with the pause: what is playing is a stream, not a still"
        );

        // Another film under the bubble is not shown this one's frame, which is the rule a walk
        // landing on a file behind a collapsed pin is answered by.
        stand_a_running_stream(
            PathBuf::from("another-film.mkv"),
            started,
            side,
            frame.clone(),
        );
        settle(false, Some(box_));
        assert!(
            latest_frame().is_none(),
            "a frame read beside another film is not this bubble's picture"
        );

        // And a bubble that is not on screen — one put away into the taskbar — is owed no picture at
        // all, so nothing is kept for it.
        stand_a_running_stream(film.clone(), started, side, frame.clone());
        settle(false, None);
        assert!(
            latest_frame().is_none(),
            "a circle nobody is looking at keeps no frame"
        );

        // And a frame that is *held* goes the way of the film it was read from: a walk landing on
        // another file behind a paused bubble is not left on the frame of the film before it.
        stand_a_running_stream(film, started, side, frame.clone());
        settle(false, Some(box_));
        assert_eq!(latest_frame(), Some((frame, side)), "the pause holds it");
        let mut walked = take_pin_for_a_test().expect("the pin is up");
        walked.path = PathBuf::from("another-film.mkv");
        stand_pin(Some(walked));
        settle(false, Some(box_));
        assert!(
            latest_frame().is_none(),
            "a held frame of another film is not this bubble's picture"
        );

        stand_no_stream();
        if let Ok(mut still) = BUBBLE_FRAME.lock() {
            *still = previous_still;
        }
        stand_pin(previous_pin);
        stand_video_engine_in(None);
        stand_mpv_in(None);
        stand_ffplay_in(None);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = previous_media;
        }
    }
}

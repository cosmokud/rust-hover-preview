//! The frames a bubble shows of a film FFmpeg's player is drawing: one decode beside the player,
//! read as it comes.
//!
//! **A film FFmpeg's player is drawing is on the screen and nowhere else.** That window is a
//! player's rather than this app's — so unlike a still, a frame cannot be looked up in memory, and
//! unlike mpv there is nothing to ask it with: `ffplay` has no interface at all (see `bubble_film`
//! for the route that has one). The picture a bubble can show of such a film *moving* is therefore a
//! second read of the file, and the whole of this module is making that read cheap: **one** decode,
//! begun where the player is, throwing away everything but the newest frame, scaled and cropped to
//! the bubble's own square and capped at the rate a bubble is looked at.
//!
//! **The player's own start is the key a stream is kept for.** A stream is a decode of *a play* —
//! this film, from this second, in this square — so what it is compared against is the transport's
//! own record of when the player was begun: a level change or a seek on this route kills the player
//! and begins another at a new second (see `restart_pinned_player`), which is a new start and so a
//! stream that has to be begun again. A stream that has gone is not tried again under the same key,
//! because a command the build refuses would otherwise be a process every tick; a player begun
//! afresh is a new key and is streamed again.
//!
//! **Nothing here is a frame the bubble must have.** The still grabbed at the collapse stands until
//! a frame lands, stands again whenever the stream ends, and is the ring's colour throughout (see
//! `grab_bubble_frame` and `bubble_art`): a stream that cannot start, a film that ends, a build
//! whose decoder refuses the command — each of them is a bubble showing what every bubble showed
//! before this module existed.

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
    child: Child,
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
    /// A stream whose frames are coming. A test build never has one — every stream it is asked for is
    /// answered with the ask and a refusal (see `start`) — so the variant is allowed to stand unbuilt
    /// there rather than being folded away with `cfg`, which would take the branches that read one
    /// out of the build the suite runs on.
    #[cfg_attr(test, allow(dead_code))]
    Running(Running),
    /// A stream that was asked for and did not come: the key it was refused for, so that it is asked
    /// for once and not once a tick (see the module documentation).
    Refused(StreamKey),
}

/// The stream, or nothing while none is.
static STREAM: Lazy<Mutex<Option<Stream>>> = Lazy::new(|| Mutex::new(None));

/// The newest frame read out of the film, and the side it was cut to, or nothing while the bubble
/// has no live picture: what the paint asks, and what it answers with the still when there is none.
///
/// A frame nothing has replaced is not a live picture: a stream whose child has gone leaves nothing
/// here rather than its last frame, or a bubble would show a film standing still as though it were
/// playing (see `end` and `read_frames`).
pub(super) fn latest_frame() -> Option<(Vec<u8>, u32)> {
    // A frame stood in for the one a real child would have written, for the tests that are about what
    // a bubble paints rather than about the reading of it (see `stand_live_frame`).
    #[cfg(test)]
    if let Some(stood_in) = stood_in_frame() {
        return Some(stood_in);
    }

    let stream = STREAM.lock().ok()?;
    let Stream::Running(running) = stream.as_ref()? else {
        return None;
    };

    let frame = running.latest.lock().ok()?.clone()?;
    Some((frame, running.key.side))
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
            return refuse(running.key);
        }
        (Some(Stream::Refused(key)), Some(wanted)) if key == wanted.key => {
            return hold(Stream::Refused(key));
        }
        (held, _) => {
            // Everything else is over: a child that is still there is ended, and the frames it had
            // landed are dropped so that the bubble paints its still again.
            if let Some(Stream::Running(running)) = held {
                end(running);
            }
        }
    }

    let Some(wanted) = wanted else {
        return;
    };

    hold(start(wanted));
}

/// What the bubble's stream should be, where there is one to have: a film FFmpeg's player is
/// drawing, at the second that player is at, in the square the bubble is standing in.
///
/// **The route is the whole of the choice.** A film mpv draws has its own window standing in the
/// bubble already (see `bubble_film`) and a film this app's engine draws holds its frame here, so
/// neither has anything a second decode would add. Nothing is asked of the machine here — no probe,
/// no process — which is what lets this sit on a tick.
///
/// The second is the *playhead* rather than the second the player began at, because a film pinned
/// half an hour ago and collapsed now is a bubble that has to show the film where it is: the stream
/// is a decode of the same play, begun at the same point in it, and the two read at the same rate
/// from there.
fn wanted_stream(box_: ScreenRegion) -> Option<Wanted> {
    if current_media_type() != Some(MediaType::Video) {
        return None;
    }

    let (path, _) = pinned_media_owner()?;
    if video_route(&path) != VideoRoute::Ffplay {
        return None;
    }

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

/// Whether a stream's child is still running: what tells a stream that is a picture from one whose
/// decodes have stopped coming.
fn alive(running: &mut Running) -> bool {
    matches!(running.child.try_wait(), Ok(None))
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
        // The device the film is already being decoded on, where the machine has one: the child is a
        // second decode of the same film, and a card can do it twice as easily as a core can do it
        // once. shortcut: a build that refuses the device for a pipe output cannot be told apart from
        // one that cannot open the file, so both leave the bubble on its still.
        let device = video_hw_accel_device();
        let args = stream_args(&wanted.key.path, wanted.at, wanted.key.side, device);

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
            child,
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
fn stream_args(path: &Path, at: f64, side: u32, device: Option<&str>) -> Vec<String> {
    let mut args: Vec<String> = Vec::new();
    if let Some(device) = device {
        args.push("-hwaccel".to_string());
        args.push(device.to_string());
    }

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
    let _ = running.child.kill();
    let _ = running.child.wait();

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
    /// bubble is looked at, cut to the bubble's own square — and the device the film is already
    /// decoded on where the machine has one.
    #[test]
    fn the_command_reads_the_film_from_the_second_the_player_is_at_cut_to_the_bubbles_square() {
        let film = PathBuf::from("a-film-a-bubble-shows-playing.mkv");

        assert_eq!(
            stream_args(&film, 92.5, 44, None),
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

        // And the device first, where the machine has one: the child is a second decode of the film
        // the player is already decoding, and a card can do that twice as easily as a core can do it
        // once (see `video_hw_accel_device`).
        let with_device = stream_args(&film, 0.0, 88, Some("d3d11va"));
        assert_eq!(
            with_device[..2],
            ["-hwaccel".to_string(), "d3d11va".to_string()]
        );
        assert_eq!(
            with_device.len(),
            stream_args(&film, 0.0, 88, None).len() + 2
        );
    }

    /// **What the bubble's stream is, is the state of the world every tick**: a film FFmpeg's player
    /// is drawing is read beside it while the pin is a bubble and the film is playing, and nothing is
    /// asked for any other kind, any other route, any other playback or any other place.
    ///
    /// The ask is what a test build answers — no child is ever started (see `start`) — so what is
    /// asserted is the decision: which film, at which second, at which side, and how often.
    #[test]
    fn a_film_ffmpegs_player_is_drawing_is_read_beside_it_while_it_plays() {
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

        // A film mpv is drawing is not read at all: its own window stands in the bubble already (see
        // `bubble_film`).
        stand_mpv_in(Some(true));
        settle(true, Some(box_));
        assert!(
            take_noted_streams().is_empty(),
            "the mpv route is the window in the ring, and a second decode of it would be a picture \
             nobody is looking at"
        );
        stand_mpv_in(Some(false));

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
}

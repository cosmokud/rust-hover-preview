//! Every frame of a picture that plays, decoded once at one size and held, so that a hand on its
//! track is shown whatever frame it names at once rather than walked to it.
//!
//! A GIF's frames and an animated WebP's are delta-coded and carry no frame index, so a decode of
//! frame *n* is a decode of every frame before it (see `animation_frames`). Shown on a drag, that
//! is a walk per column of the track: a few milliseconds for a GIF, whose frames are patches
//! composited into a canvas that is scaled once, and tens of them for a WebP, whose every frame is
//! a `VP8` image decoded whole and copied out of libwebp. What makes micro-seeking feel wrong is
//! therefore the file, not the seek: going back four frames costs the walk to the beginning again.
//!
//! So the run is read once, in the background, and every frame of it is kept — at the size the
//! pin draws at when that fits the room `Performance → Memory → Animated Image` names, and
//! the largest size that does when it does not. The frames are what is *kept*, so the walk this
//! file needs is paid once per pin rather than once per press, and a hand that goes back and forth
//! reads frames of memory both ways.
//!
//! **It is a fallback for the runs the pin cannot hold whole, and nothing else.** A pinned picture's
//! own run is held up to the same budget (see `retained_room`), so a run that fits it is decoded
//! once by the player and every frame of it stays in hand — a second run of the same frames would be
//! the same picture held twice, and a frame a hand named is already a playhead put down on the
//! player's own run (see `show_the_frame_of`).
//!
//! **The walk is started once per pin and never per press.** A run already in hand for the file on
//! screen and the size the pin draws at is the run a press reads, and a start for it answers
//! nothing; so a drag across a long picture asks for no decode of its own at all, and the memory it
//! costs is the run and nothing else.
//! It is deliberately only the two kinds whose walks are expensive and which are decoded here:
//! a APNG is composited by the `image` crate from its own frame records and needs no index at all,
//! and a HEIF or a JXL sequence is held whole by the engine in the first place (see
//! `load_animated_heif` and `load_animated_jxl`).
//!
//! **A frame of a smaller size is a frame that is drawn scaled up**, so it is the shape of the
//! picture that is true rather than its detail, which is what a hand reading where in the file it
//! is wants: a run this file holds is a run too long for the player to keep, so its frames are the
//! picture until the player's own run reaches that place.
//!
//! **And nothing is read for a picture the player already holds whole.** A short animation is
//! decoded from start to finish by the player itself and every frame of it stays in hand, so a
//! frame a hand names is already a playhead put down on that run (see `show_the_frame_of`) — a
//! second run of the same frames would be the same picture held twice.

use super::*;

/// How much of a picture's run is held at once, which is what decides whether the frames are the
/// pin's own size or a size that fits: a file of a few hundred frames at a thumbnail's box is held
/// whole and exactly, and one of a few thousand is held whole and small.
///
/// It is the `Performance → Memory → Animated Image` setting rather than a constant: the run
/// is what a hand on a picture's track is shown a frame of, so how sharp a dragged picture is, is
/// a question a machine with less memory to spare answers for itself (see
/// `animated_webp_budget_bytes`).
pub(super) fn retained_bytes() -> usize {
    crate::config::config::animated_webp_budget_bytes()
}

/// The narrowest a proxy frame is drawn at: below this a frame is a smear rather than a place in
/// the file. A run that would be drawn smaller is the frames from the beginning that fit at this
/// width instead — the run is cut rather than made unreadable, and the frames it holds are the ones
/// a hand starts from.
const INDEX_PROXY_MIN_WIDTH: u32 = 96;

/// A held run of one file at one size.
struct Index {
    path: PathBuf,
    version: FileVersion,
    /// The size the frames are held at, which is the size a box change asks for again: a picture
    /// laid out at another box is a run at another size, and the frames of the old one are frames
    /// the window is no longer drawn at.
    width: u32,
    height: u32,
    /// The walk that is filling this run, which is what says whether a frame arriving belongs in
    /// it: a box change asks for the same file at another size, and a frame of the size it was is a
    /// frame of a run nobody holds any more.
    walk: u64,
    /// The room the run was read into, which is what the setting said when it was asked for: a run
    /// read into one room is not the run a smaller one would have held, so a changed setting is a
    /// run to read again.
    room: usize,
    /// How many frames the run may hold, which is the budget divided by the size of one frame of
    /// it: a file of a few thousand frames is read as far as the budget reaches rather than to the
    /// end of it, so that an index of one picture never becomes the reason the app is using memory.
    max_frames: usize,
    /// The frames in hand, by the place they are in the file. A frame the walk has not reached is
    /// nothing yet, and a frame that could not be decoded keeps its hole.
    frames: Vec<Option<Arc<ImageFrame>>>,
}

static INDEX: Lazy<Mutex<Option<Index>>> = Lazy::new(|| Mutex::new(None));

/// Which walk is the current one. A new pin, or a take-down, is a walk that is no longer wanted,
/// and its thread stops where it stands on its next frame rather than filling a run nobody asked
/// for (see `start` and `forget`).
static INDEX_WALK: AtomicU64 = AtomicU64::new(0);

/// Read the picture at `path` once, in the background, and hold every frame of it.
///
/// `display` is the size the pin draws the picture at, which is the size the run is held at where
/// the budget allows it. `total` is how many frames the file has, which is what says whether the
/// whole of it fits — a file whose container the walk could not read has no run to hold, and its
/// track has no frames a hand could name either (see `pin_frame_share`).
///
/// Nothing is held for a picture whose kind needs no index (see the module documentation), and a
/// second call for the same file and size answers nothing at all: what is already being read is
/// what was going to be read.
pub(super) fn start(path: &Path, kind: MediaType, total: usize, display: (u32, u32)) {
    if total == 0 || display.0 == 0 || display.1 == 0 {
        return;
    }
    let webp = matches!(kind, MediaType::AnimatedWebP);
    let gif = matches!(kind, MediaType::AnimatedGif);
    if !webp && !gif {
        return;
    }

    // **A picture the media already holds whole is a picture with nothing to index** (see the
    // module documentation): its frames are in hand, and a second run of them would be the same
    // picture held twice.
    let held_whole = CURRENT_MEDIA.lock().ok().is_some_and(|media| {
        media.as_ref().is_some_and(|media| {
            media.media_type == kind && media.is_fully_loaded() && media.frames.len() >= total
        })
    });
    if held_whole {
        return;
    }

    // **And a run that fits the budget is one the player will hold whole**, decoded once as it
    // plays: this file's own walk would be the same frames kept a second time (see `retained_bytes`).
    // A frame a hand names is a playhead put down on the run the player has by the time it has
    // reached that place, which is the answer a hand on the track wants anyway.
    if total as u64 * display.0 as u64 * display.1 as u64 * 4 <= retained_bytes() as u64 {
        return;
    }

    let version = file_version(path);
    let (width, height) = proxy_size(total, display);
    if width == 0 || height == 0 {
        return;
    }
    let room = retained_bytes();

    // What is already being read is what was going to be read: a second ask for the same file at
    // the same size and the same room is a walk already on its way, and one for a run already in
    // hand is nothing asked for at all. That is the whole of what keeps a drag from walking the
    // file per press (see `show_the_frame_of`).
    if let Ok(index) = INDEX.lock() {
        if let Some(index) = index.as_ref() {
            if index.path == path
                && index.version == version
                && index.width == width
                && index.height == height
                && index.room == room
            {
                return;
            }
        }
    }

    let walk = INDEX_WALK.fetch_add(1, Ordering::AcqRel) + 1;
    let max_frames = max_frames(width, height, room);

    if let Ok(mut index) = INDEX.lock() {
        *index = Some(Index {
            path: path.to_path_buf(),
            version,
            width,
            height,
            room,
            walk,
            max_frames,
            frames: Vec::new(),
        });
    }

    let path = path.to_path_buf();
    std::thread::spawn(move || {
        // Every frame the walk answers is pushed as it arrives rather than at the end of the file:
        // a hand that starts dragging a moment after the pin came up is already served by the
        // frames that have been read, and one that drags while the rest is still being read follows
        // the walk instead of waiting for it.
        let mut push = |index_in_file: usize, frame: ImageFrame| -> bool {
            if INDEX_WALK.load(Ordering::Acquire) != walk {
                return false;
            }
            let Ok(mut held) = INDEX.lock() else {
                return false;
            };
            let Some(held) = held.as_mut() else {
                return false;
            };
            if held.walk != walk {
                return false;
            }
            if index_in_file >= held.max_frames {
                return false;
            }
            if held.frames.len() <= index_in_file {
                held.frames.resize(index_in_file + 1, None);
            }
            held.frames[index_in_file] = Some(Arc::new(frame));
            true
        };

        if webp {
            walk_animated_webp(&path, (width, height), &mut push);
        } else {
            walk_animated_gif(&path, (width, height), &mut push);
        }
    });
}

/// How many frames of a run held at `width` by `height` fit `room`, and never none.
fn max_frames(width: u32, height: u32, room: usize) -> usize {
    let frame_bytes = (width as u64 * height as u64 * 4).max(4);

    (room as u64 / frame_bytes).max(1) as usize
}

/// The size a run of `total` frames is held at: the pin's own where that fits the budget, and the
/// largest that does where it does not.
fn proxy_size(total: usize, display: (u32, u32)) -> (u32, u32) {
    let frame_bytes = display.0 as u64 * display.1 as u64 * 4;
    let retained = retained_bytes() as u64;
    let wanted = frame_bytes.saturating_mul(total as u64);
    if wanted <= retained {
        return display;
    }

    let room = retained / total.max(1) as u64;
    // The scale that keeps the frame's own shape inside the room, taken from its area.
    let scale = (room as f64 / frame_bytes as f64).sqrt();
    let width = ((display.0 as f64 * scale).floor() as u32).max(INDEX_PROXY_MIN_WIDTH);
    if width >= display.0 {
        return display;
    }
    let height = ((display.1 as u64 * width as u64) / display.0 as u64).max(1) as u32;

    (width, height)
}

/// The frame of `path` at `index`, with whether it is the size the pin draws at: a frame that is
/// not is worth showing at once and still owes the decode of the real one (see
/// `show_the_frame_of_the_file`).
pub(super) fn at(path: &Path, index: usize) -> Option<Arc<ImageFrame>> {
    let index_held = INDEX.lock().ok()?;
    let index_held = index_held.as_ref()?;
    if index_held.path != path {
        return None;
    }

    index_held
        .frames
        .get(index)
        .and_then(Option::as_ref)
        .map(Arc::clone)
}

/// The frame of `path` nearest `index` that is held, which is what a hand is shown while the frame
/// it named is being read.
pub(super) fn nearest(path: &Path, index: usize) -> Option<(usize, Arc<ImageFrame>)> {
    let index_held = INDEX.lock().ok()?;
    let index_held = index_held.as_ref()?;
    if index_held.path != path {
        return None;
    }

    let (held, frame) = index_held
        .frames
        .iter()
        .enumerate()
        .filter_map(|(at, frame)| frame.as_ref().map(|frame| (at, frame)))
        .map(|(at, frame)| (at.abs_diff(index), at, frame))
        .min_by_key(|(distance, at, _)| (*distance, *at))
        .map(|(_, at, frame)| (at, frame))?;

    Some((held, Arc::clone(frame)))
}

/// Give a held run back, and stop the walk that is still filling it: what is held belongs to the
/// file that was on screen when it was asked for.
pub(super) fn forget() {
    INDEX_WALK.fetch_add(1, Ordering::AcqRel);
    if let Ok(mut index) = INDEX.lock() {
        *index = None;
    }
}

/// Read a GIF whole, compositing its frames the way the player does, and hand each one over at
/// `target`.
///
/// The canvas is the file's own and is filled once per frame, so what each frame costs is the
/// patch of it the file draws plus one scale — which is the reason a GIF is the cheap half of this
/// module's work, and the reason its frames are already most of what a scrub wants.
fn walk_animated_gif(
    path: &Path,
    target: (u32, u32),
    push: &mut dyn FnMut(usize, ImageFrame) -> bool,
) {
    let Ok(file) = File::open(path) else {
        return;
    };
    let mut decoder = DecodeOptions::new();
    decoder.set_color_output(gif::ColorOutput::RGBA);
    let Ok(mut decoder) = decoder.read_info(BufReader::new(file)) else {
        return;
    };

    let (gif_width, gif_height) = (decoder.width() as u32, decoder.height() as u32);
    let Some(canvas_bytes) = frame_bytes_within_budget(gif_width, gif_height, 4) else {
        return;
    };
    let mut canvas = vec![0u8; canvas_bytes];

    let mut index = 0usize;
    while let Ok(Some(frame)) = decoder.read_next_frame() {
        composite_gif_frame(&mut canvas, frame, gif_width, gif_height);
        let delay_ms = (frame.delay as u32 * 10).max(MIN_ANIMATION_FRAME_DELAY_MS);
        if let Some(mut img) =
            decode_gif_frame_to_image(&canvas, gif_width, gif_height, target.0, target.1, delay_ms)
        {
            // The place in the file travels with the frame, the way it does in every other run of
            // one: what the bar says the picture is on is read of it (see `current_frame_of_file`).
            img.file_index = index;
            if !push(index, img) {
                return;
            }
        }
        index += 1;
    }
}

/// Read an animated WebP whole, handing over every frame libwebp composites.
///
/// Each frame here is a `VP8` image decoded whole and painted into the canvas libwebp keeps, so
/// what the walk costs is the file's own decode times the number of frames it holds — which is the
/// cost an index is the only way out of, since nothing can begin at a frame.
fn walk_animated_webp(
    path: &Path,
    target: (u32, u32),
    push: &mut dyn FnMut(usize, ImageFrame) -> bool,
) {
    let Some(buffer) = read_within_budget(path) else {
        return;
    };
    let options = webp_animation::DecoderOptions {
        use_threads: true,
        color_mode: webp_animation::ColorMode::Bgra,
    };
    let Ok(decoder) = webp_animation::Decoder::new_with_options(buffer.as_slice(), options) else {
        return;
    };
    let (orig_width, orig_height) = decoder.dimensions();
    if orig_width == 0 || orig_height == 0 {
        return;
    }
    if frame_bytes_within_budget(orig_width, orig_height, 4).is_none() {
        return;
    }

    let mut previous_timestamp = 0i32;
    for (index, frame) in decoder.into_iter().enumerate() {
        let timestamp = frame.timestamp();
        let delay_ms = (timestamp - previous_timestamp).max(0) as u32;
        previous_timestamp = timestamp;

        if let Some(mut img) = decode_webp_animation_frame_to_image(
            frame.data(),
            orig_width,
            orig_height,
            target.0,
            target.1,
            delay_ms,
        ) {
            img.file_index = index;
            if !push(index, img) {
                return;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The index is one per process, so the tests that hold runs and the one that reads a file off
    /// a thread of its own share it: a walk put down by another test's `forget` is a walk this one
    /// is waiting on for ever. One test at a time is what a state of this shape needs.
    static ONE_TEST: Mutex<()> = Mutex::new(());

    fn one_test_at_a_time() -> MutexGuard<'static, ()> {
        ONE_TEST
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn a_frame(file_index: usize, side: u32) -> ImageFrame {
        let mut frame = ImageFrame::new(vec![0u8; (side * side * 4) as usize], side, side, 100);
        frame.file_index = file_index;
        frame
    }

    fn hold(path: &Path, room: usize, frames: &[(usize, u32)]) {
        let (width, height) = (320, 240);
        let mut index = INDEX.lock().expect("the index's lock");
        let mut frames_held = Vec::new();
        for (at, side) in frames {
            if frames_held.len() <= *at {
                frames_held.resize(*at + 1, None);
            }
            frames_held[*at] = Some(Arc::new(a_frame(*at, *side)));
        }
        *index = Some(Index {
            path: path.to_path_buf(),
            version: file_version(path),
            width,
            height,
            room,
            walk: INDEX_WALK.fetch_add(1, Ordering::AcqRel) + 1,
            max_frames: max_frames(width, height, room),
            frames: frames_held,
        });
    }

    fn a_file() -> PathBuf {
        let path = std::env::temp_dir()
            .join("pin-index-tests")
            .join("run.webp");
        std::fs::create_dir_all(path.parent().expect("a folder")).expect("a fixture folder");
        std::fs::write(&path, b"RIFF").expect("a fixture file");
        path
    }

    /// **Every frame of the run is a frame a hand can be shown at once**, which is the whole of
    /// what the index is for: a frame that has been read is answered for by its place in the file,
    /// and one that has not is answered by the nearest that has rather than by a walk.
    #[test]
    fn a_frame_of_the_run_is_answered_at_once_and_a_frame_it_does_not_hold_is_answered_nearest() {
        let _one = one_test_at_a_time();
        forget();
        let path = a_file();
        hold(&path, retained_bytes(), &[(3, 320), (4, 320), (5, 320)]);

        let frame = at(&path, 4).expect("a frame of the run");
        assert_eq!(frame.file_index, 4, "the place in the file is the key");
        assert!(at(&path, 99).is_none(), "one not read yet is nothing yet");

        let (index, frame) = nearest(&path, 8).expect("the nearest frame held");
        assert_eq!((index, frame.file_index), (5, 5), "the nearest of the run");

        // A run is held for one file and one size: another file is another run.
        let elsewhere = path.with_file_name("other.webp");
        std::fs::write(&elsewhere, b"RIFF").expect("a fixture file");
        assert!(at(&elsewhere, 4).is_none(), "a run belongs to one file");
        forget();
    }

    /// **A run is read off the file's own frames, in the order they play, each frame carrying the
    /// place it has in the file** — which is what says that the frame a hand named is the frame it
    /// got, and what the bar is drawn on while the picture is standing there.
    #[test]
    fn a_run_is_read_off_the_file_with_a_frame_for_every_place_in_it() {
        let _one = one_test_at_a_time();
        let folder = std::env::temp_dir().join("pin-index-walk-tests");
        std::fs::create_dir_all(&folder).expect("a fixture folder");
        let path = folder.join("run.webp");

        // A three-frame animation written by libwebp, which is the same library that reads it.
        let dimensions = (32u32, 24u32);
        let mut encoder = webp_animation::Encoder::new(dimensions).expect("an encoder");
        for (index, shade) in [0u8, 90, 200].into_iter().enumerate() {
            let pixels: Vec<u8> = [shade, shade, shade, 255]
                .into_iter()
                .cycle()
                .take((dimensions.0 * dimensions.1 * 4) as usize)
                .collect();
            encoder
                .add_frame(&pixels, index as i32 * 100)
                .expect("a frame");
        }
        let bytes = encoder.finalize(300).expect("an animation");
        std::fs::write(&path, bytes.as_ref()).expect("a fixture file");

        // A count no budget could hold whole, which is what makes the file's own walk owed (see
        // `start`): the fixture holds three frames because the three are what is being walked, and
        // what the count decides here is only that the player will not be holding them itself.
        forget();
        start(&path, MediaType::AnimatedWebP, 200_000, dimensions);

        // The walk runs on a thread of its own, so the wait is on the condition rather than on a
        // stretch of time (see the terminal rules this app is developed under).
        let deadline = Instant::now() + Duration::from_secs(10);
        while Instant::now() < deadline {
            if (0..3).all(|index| at(&path, index).is_some()) {
                break;
            }
            std::thread::sleep(Duration::from_millis(5));
        }

        for index in 0..3 {
            let frame = at(&path, index).unwrap_or_else(|| panic!("frame {index} of the run"));
            assert_eq!(
                frame.file_index, index,
                "and it says where it is in the file"
            );
            assert_eq!(
                (frame.width, frame.height),
                dimensions,
                "at the size asked for"
            );
        }
        assert!(at(&path, 3).is_none(), "and nothing past the last frame");
        forget();
    }

    /// **A run is held at the pin's own size while that fits its budget, and at the largest that
    /// does when it does not** — and a run too long for even its narrowest is held at the narrowest,
    /// so that a hand is still shown where in the file it is.
    #[test]
    fn a_run_is_held_at_the_pin_own_size_while_it_fits_and_scaled_down_when_it_does_not() {
        let display = (320, 240);
        let frame_bytes = 320u64 * 240 * 4;
        let retained = retained_bytes() as u64;
        let fits = (retained / frame_bytes) as usize;

        assert_eq!(
            proxy_size(fits, display),
            display,
            "a run that fits the budget is held exactly"
        );
        assert_eq!(
            proxy_size(1, display),
            display,
            "and so is a file of one frame"
        );

        let (width, height) = proxy_size(fits + 1, display);
        assert!(
            width < display.0 && height < display.1,
            "and one that does not is scaled down"
        );
        assert_eq!(
            (width as u64 * height as u64) as f64 / (display.0 as u64 * display.1 as u64) as f64,
            (width as f64 / display.0 as f64) * (height as f64 / display.1 as f64),
            "a proxy keeps the frame's own shape"
        );
        assert!(
            width as u64 * height as u64 * 4 <= retained / (fits as u64 + 1),
            "and one frame of it is inside its own room in the budget"
        );

        // The narrowest a proxy is drawn at is the floor under it, whatever the file holds.
        let (width, _) = proxy_size(usize::MAX, display);
        assert_eq!(width, INDEX_PROXY_MIN_WIDTH);
    }
}

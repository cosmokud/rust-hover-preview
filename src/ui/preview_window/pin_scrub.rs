//! The frames of a picture that plays that this pin has already decoded, held so that a hand
//! dragging the bar's track can be shown the frame under it without asking the file for it again.
//!
//! A drag of a picture is a walk through its frames, and the walk is the expensive part: GIF, APNG
//! and WebP are delta-coded and carry no frame index, so landing on frame `n` means decoding the
//! ones before it (see `animation_frames`). What a decode answers is therefore worth holding on to:
//! the frames a load or a box change lands are remembered here, and a hand that comes back to a
//! place it has already been is shown that place at once rather than walked to it again (see
//! `show_the_frame_of_the_file`).
//!
//! It is a budget rather than a file: the frames are held decoded *and scaled*, which is what makes
//! them worth holding, and a picture of a few thousand frames would be hundreds of megabytes of them.
//! The oldest go first, so what a hand is doing now and what it did a moment ago is what is kept.
//!
//! It is not the frame cache (see `image_cache`): that one holds the one frame a still is drawn
//! from, keyed by file and size, and answers the next hover. This one holds a *run* of a picture's
//! frames for the length of one pin's drag.

use super::*;

/// How much of a picture's decoded run is held at once: enough for a few thousand frames of a
/// thumbnail's own size, and small enough that a drag across a feature film's worth of animation
/// does not become the reason the app is using memory.
pub(super) const SCRUB_RETAINED_BYTES: usize = 64 * 1024 * 1024;

/// What a held run is only valid for: the file, its version, and the pixel size the frames were
/// decoded at — a frame decoded for one box is not a frame for another (see `ImageCacheKey`, which
/// is the same three answers for the one frame a still is drawn from).
struct ScrubKey {
    path: PathBuf,
    version: FileVersion,
    width: u32,
    height: u32,
}

struct Scrub {
    key: ScrubKey,
    /// The frames in hand, by where they are in the file.
    held: HashMap<usize, Arc<ImageFrame>>,
    /// The order they were taken in, so that the oldest are the ones given back.
    order: VecDeque<usize>,
    bytes: usize,
}

static SCRUB: Lazy<Mutex<Option<Scrub>>> = Lazy::new(|| Mutex::new(None));

/// Hold the frames a decode has just answered for, so a hand that comes back to them is shown them
/// at once.
///
/// Called where an answer is installed rather than where it is asked for: what lands in the window
/// is what the picture has, and a run of frames asked for and then thrown away is a run nobody can
/// be shown.
pub(super) fn remember(path: &Path, frames: &[Arc<ImageFrame>]) {
    let Some(first) = frames.first() else {
        return;
    };
    let (width, height) = (first.width, first.height);
    if width == 0 || height == 0 {
        return;
    }
    let key = ScrubKey {
        path: path.to_path_buf(),
        version: file_version(path),
        width,
        height,
    };

    let Ok(mut scrub) = SCRUB.lock() else {
        return;
    };
    let scrub = scrub.get_or_insert_with(|| Scrub {
        key: ScrubKey {
            path: key.path.clone(),
            version: key.version.clone(),
            width: key.width,
            height: key.height,
        },
        held: HashMap::new(),
        order: VecDeque::new(),
        bytes: 0,
    });

    // A run of a file, or of a size, that is not the one held is a run given back whole: half a
    // picture at one size and half at another is a picture whose frames do not line up.
    if scrub.key.path != key.path
        || scrub.key.version != key.version
        || scrub.key.width != key.width
        || scrub.key.height != key.height
    {
        *scrub = Scrub {
            key,
            held: HashMap::new(),
            order: VecDeque::new(),
            bytes: 0,
        };
    }

    for frame in frames {
        if scrub.held.contains_key(&frame.file_index) {
            continue;
        }
        scrub.bytes += frame.pixels.len();
        scrub.held.insert(frame.file_index, Arc::clone(frame));
        scrub.order.push_back(frame.file_index);
    }

    while scrub.bytes > SCRUB_RETAINED_BYTES {
        let Some(oldest) = scrub.order.pop_front() else {
            break;
        };
        if let Some(frame) = scrub.held.remove(&oldest) {
            scrub.bytes = scrub.bytes.saturating_sub(frame.pixels.len());
        }
    }
}

/// The frame of `path` at `index`, if this pin has decoded it already.
pub(super) fn at(path: &Path, index: usize) -> Option<Arc<ImageFrame>> {
    let scrub = SCRUB.lock().ok()?;
    let scrub = scrub.as_ref()?;
    if scrub.key.path != path {
        return None;
    }

    scrub.held.get(&index).map(Arc::clone)
}

/// The frame of `path` nearest `index` that is held, which is what a hand is shown while the frame
/// it actually named is still being decoded.
pub(super) fn nearest(path: &Path, index: usize) -> Option<(usize, Arc<ImageFrame>)> {
    let scrub = SCRUB.lock().ok()?;
    let scrub = scrub.as_ref()?;
    if scrub.key.path != path {
        return None;
    }

    scrub
        .held
        .iter()
        .map(|(held, frame)| (held.abs_diff(index), *held, frame))
        .min_by_key(|(distance, held, _)| (*distance, *held))
        .map(|(_, held, frame)| (held, Arc::clone(frame)))
}

/// Give a held run back: what is held belongs to the file that was on screen when it was decoded.
pub(super) fn forget() {
    if let Ok(mut scrub) = SCRUB.lock() {
        *scrub = None;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn a_frame(file_index: usize, side: u32) -> Arc<ImageFrame> {
        let mut frame = ImageFrame::new(vec![0u8; (side * side * 4) as usize], side, side, 100);
        frame.file_index = file_index;
        Arc::new(frame)
    }

    fn a_file() -> PathBuf {
        let path = std::env::temp_dir().join("pin-scrub-tests").join("run.gif");
        std::fs::create_dir_all(path.parent().expect("a folder")).expect("a fixture folder");
        std::fs::write(&path, b"GIF89a").expect("a fixture file");
        path
    }

    /// **A run that has been decoded is a run that can be shown again**, which is the whole of what
    /// the cache is for: a hand that comes back to a place it has already been is not walked there
    /// again, and a hand on a frame still being decoded is shown the nearest one that is held.
    #[test]
    fn a_decoded_run_is_shown_again_and_the_nearest_one_is_offered_while_the_rest_decodes() {
        forget();
        let path = a_file();
        let frames: Vec<Arc<ImageFrame>> = [10, 11, 12]
            .into_iter()
            .map(|index| a_frame(index, 4))
            .collect();
        remember(&path, &frames);

        assert_eq!(
            at(&path, 11).map(|frame| frame.file_index),
            Some(11),
            "a frame that has been decoded is a frame that can be shown"
        );
        assert!(at(&path, 99).is_none(), "and one that has not, is not");

        let (index, frame) = nearest(&path, 40).expect("the nearest frame that is held");
        assert_eq!(
            (index, frame.file_index),
            (12, 12),
            "the nearest of the run"
        );

        // And it is the run of a file and a size: another file, or another size, is another run.
        let elsewhere = path.with_file_name("other.gif");
        std::fs::write(&elsewhere, b"GIF89a").expect("a fixture file");
        assert!(at(&elsewhere, 11).is_none(), "a run is held for one file");

        forget();
        assert!(at(&path, 11).is_none(), "and given back whole");
    }

    /// **A run is given back from the oldest end once it is over its budget**, so that a drag across
    /// a long picture is the recent frames held rather than the whole file: what a hand is doing and
    /// what it did a moment ago is what it comes back to.
    #[test]
    fn a_run_longer_than_its_budget_gives_the_oldest_frames_back() {
        forget();
        let path = a_file();
        let side = 256;
        let bytes = (side * side * 4) as usize;
        let how_many = SCRUB_RETAINED_BYTES / bytes + 4;

        let frames: Vec<Arc<ImageFrame>> =
            (0..how_many).map(|index| a_frame(index, side)).collect();
        remember(&path, &frames);

        assert!(
            at(&path, 0).is_none(),
            "the oldest frames are the ones given back"
        );
        assert!(
            at(&path, how_many - 1).is_some(),
            "and the newest are the ones kept"
        );
        forget();
    }
}

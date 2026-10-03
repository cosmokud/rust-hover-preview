//! What a film is asked of FFprobe: its size and length, the subtitle streams it holds, and
//! the crop the picture is really drawn in.

use super::*;

pub(super) const VIDEO_CROPDETECT_LIMIT: &str = "24";
pub(super) const VIDEO_CROPDETECT_ROUND: &str = "16";
pub(super) const VIDEO_CROPDETECT_FRAMES: &str = "48";
pub(super) const VIDEO_CROP_MAX_AXIS_TRIM_RATIO: f32 = 0.10;
pub(super) const VIDEO_CROP_MAX_ASYMMETRY_PX: i32 = 12;

/// How long either of a video's two probes is given before it is killed and the file is
/// answered as one that could not be measured.
///
/// The wait this bounds is the one wait in the preview that nothing else ends: the hover is
/// in a `PendingLoad` that is not waiting on an engine, so the cap a page is waited for under
/// is not standing behind it, and a file no probe ever answers for is a spinner that stands
/// at the pointer until the pointer moves. Ten seconds is what a player's own start is given
/// (`VIDEO_START_WAIT_SECS`) and the same order as the work — forty-eight decoded frames of a
/// 4K file is the slow case — and what a file slower than this is answered with is a video
/// placed without its crop, which is what a file `ffprobe` could not read has always been
/// given. If a file of the user's turns out to be slower than this, it is one constant.
pub(super) const VIDEO_PROBE_TIMEOUT_SECS: u64 = 10;

/// Wait for one of a probe's children, giving up after `timeout`, and answer with what it
/// wrote or with nothing at all.
///
/// A child that is still running at the deadline is killed *before* anything is read of it,
/// because what is being stopped is not the answer but the work: a cropdetect pass over a
/// file whose frames it cannot keep up with runs until the file ends, and a hover that has
/// given up on it is a hover that must not leave it running. The wait that follows the kill
/// is what reaps it, and the output a killed child leaves behind is a partial line — which
/// is the answer arm both callers already have for a probe that answered nothing.
///
/// A child that has finished needs no kill and no second wait: `wait_with_output` drains the
/// pipes and is what closes them, and the handle being signalled is what says there is
/// something to drain. Nothing here is recorded to strike off — the probes are adopted and
/// never written down, a process of a few dozen milliseconds having no file to leave.
pub(super) fn wait_bounded(mut child: Child, timeout: Duration) -> Option<Output> {
    let handle = HANDLE(child.as_raw_handle());
    let waited = unsafe { WaitForSingleObject(handle, timeout.as_millis() as u32) };

    if waited == WAIT_TIMEOUT {
        let _ = child.kill();
        let _ = child.wait();
        return None;
    }

    child.wait_with_output().ok()
}

/// Get video dimensions using ffprobe
pub(super) fn get_video_dimensions(path: &PathBuf) -> Option<(u32, u32, Option<f64>)> {
    // Spawned rather than run through `Command::output`, which is these two calls
    // under one name, so that the probe is in the job before it is waited on: a
    // probe left behind by a crash would otherwise go on reading a file that nobody
    // is waiting for. The wait that follows is bounded rather than the plain one, so
    // that a file this probe cannot finish with is answered rather than waited on
    // (see `wait_bounded`).
    let child = engine_processes::hidden_command("ffprobe")
        .args([
            "-v",
            "error",
            "-err_detect",
            "ignore_err",
            "-fflags",
            "+genpts+discardcorrupt+igndts",
            "-select_streams",
            "v:0",
            "-show_entries",
            "stream=width,height:format=duration",
            "-of",
            "default=noprint_wrappers=1",
        ])
        .arg(path)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    engine_processes::adopt(child.id());

    let output = wait_bounded(child, Duration::from_secs(VIDEO_PROBE_TIMEOUT_SECS))?;
    let output_str = String::from_utf8_lossy(&output.stdout);

    let mut width = None;
    let mut height = None;
    let mut duration = None;

    for line in output_str.lines() {
        let Some((key, value)) = line.trim().split_once('=') else {
            continue;
        };
        let value = value.trim();

        match key.trim() {
            "width" => width = value.parse::<u32>().ok(),
            "height" => height = value.parse::<u32>().ok(),
            // A stream with no length of its own — a live capture, a container that does not
            // say — answers `N/A`, which is the same answer as nothing at all here.
            "duration" => {
                duration = value
                    .parse::<f64>()
                    .ok()
                    .filter(|value| *value > 0.0 && value.is_finite())
            }
            _ => {}
        }
    }

    Some((width?, height?, duration))
}

/// Read a file's subtitle streams, counted and ordered the way FFmpeg numbers them for `-sst`.
///
/// It is asked of the same probe that measures the shape and the length, and run beside the two
/// others rather than after them, because none of the three needs another's answer — a file's
/// subtitle streams are in its header whether or not anything has been decoded out of it yet, so
/// adding this read costs a third process in parallel rather than a round trip in series.
///
/// `ffprobe` is asked for every stream rather than for `v:0`, because a selection is the one
/// thing that would hide the answer: the count and the order of a file's subtitle streams are
/// counted among themselves, so selecting the video alone would report no subtitles at all for a
/// file that has three. What comes back is one `key=value` per line in stream order, and the two
/// facts wanted are read from it the same way the shape is read from the other probe — by
/// splitting on the first `=`, with the disposition's own `key:value` name left whole.
pub(super) fn probe_subtitle_streams(path: &Path) -> SubtitleStreams {
    let child = engine_processes::hidden_command("ffprobe")
        .args([
            "-v",
            "error",
            "-err_detect",
            "ignore_err",
            "-fflags",
            "+genpts+discardcorrupt+igndts",
            "-show_entries",
            "stream=index,codec_type:stream_disposition=default",
            "-of",
            "default=noprint_wrappers=1",
        ])
        .arg(path)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .ok();

    let Some(child) = child else {
        return SubtitleStreams { count: 0, first: 0 };
    };
    engine_processes::adopt(child.id());

    let Some(output) = wait_bounded(child, Duration::from_secs(VIDEO_PROBE_TIMEOUT_SECS)) else {
        return SubtitleStreams { count: 0, first: 0 };
    };

    parse_subtitle_streams(&String::from_utf8_lossy(&output.stdout))
}

/// The subtitle streams of a probe's flat answer: how many there are and which the player would
/// pick by itself.
///
/// The player's own choice is the first stream the container marks as the default one, and the
/// first subtitle stream at all where the container marks none — which is how FFmpeg resolves
/// it, and reproducing the rule here is the whole of what makes an unchosen track survive a
/// relaunch unchanged rather than becoming whatever this app happened to guess was first.
///
/// The count is taken before the default is settled rather than by counting the parsed list
/// afterwards, so that a container which marks no default still has its first subtitle stream as
/// `first` rather than as a stream that was never looked for.
pub(super) fn parse_subtitle_streams(probe: &str) -> SubtitleStreams {
    let mut count = 0usize;
    let mut default = None;
    let mut last_is_subtitle = false;

    for line in probe.lines() {
        let Some((key, value)) = line.trim().split_once('=') else {
            continue;
        };
        match key.trim() {
            "codec_type" => {
                // A stream begins here, so the disposition read for the previous one is finished
                // with. Keying off the codec type rather than off the index is what keeps the two
                // in step for a container that has chosen not to number its streams.
                last_is_subtitle = value.trim() == "subtitle";
                if last_is_subtitle {
                    count += 1;
                }
            }
            // The disposition belongs to the stream just above it, so it is only allowed to name
            // a default while that stream is a subtitle one — otherwise the default *video*
            // stream would be taken for a default subtitle track, which is the mistake a file
            // with no subtitles at all would otherwise produce.
            "DISPOSITION:default" if last_is_subtitle && value.trim() == "1" => {
                default = default.or(Some(count - 1));
            }
            _ => {}
        }
    }

    SubtitleStreams {
        count,
        first: default.unwrap_or(0).min(count.saturating_sub(1)),
    }
}

pub(super) fn parse_cropdetect_line(line: &str) -> Option<VideoCrop> {
    let idx = line.rfind("crop=")?;
    let token = line[idx + 5..]
        .split_whitespace()
        .next()
        .unwrap_or_default();
    let mut parts = token.split(':');
    let width: u32 = parts.next()?.parse().ok()?;
    let height: u32 = parts.next()?.parse().ok()?;
    let x: u32 = parts.next()?.parse().ok()?;
    let y: u32 = parts.next()?.parse().ok()?;
    if parts.next().is_some() {
        return None;
    }

    Some(VideoCrop {
        width,
        height,
        x,
        y,
    })
}

pub(super) fn validate_detected_crop(crop: VideoCrop, src_w: u32, src_h: u32) -> bool {
    if crop.width == 0 || crop.height == 0 || crop.width > src_w || crop.height > src_h {
        return false;
    }

    let right = crop.x.saturating_add(crop.width);
    let bottom = crop.y.saturating_add(crop.height);
    if right > src_w || bottom > src_h {
        return false;
    }

    let trim_left = crop.x as i32;
    let trim_top = crop.y as i32;
    let trim_right = src_w.saturating_sub(right) as i32;
    let trim_bottom = src_h.saturating_sub(bottom) as i32;
    let trim_x = src_w.saturating_sub(crop.width);
    let trim_y = src_h.saturating_sub(crop.height);

    if trim_x == 0 && trim_y == 0 {
        return false;
    }

    let trim_x_ratio = trim_x as f32 / src_w as f32;
    let trim_y_ratio = trim_y as f32 / src_h as f32;
    if trim_x_ratio > VIDEO_CROP_MAX_AXIS_TRIM_RATIO
        || trim_y_ratio > VIDEO_CROP_MAX_AXIS_TRIM_RATIO
    {
        return false;
    }

    (trim_left - trim_right).abs() <= VIDEO_CROP_MAX_ASYMMETRY_PX
        && (trim_top - trim_bottom).abs() <= VIDEO_CROP_MAX_ASYMMETRY_PX
}

/// Every crop rectangle ffmpeg's detector reported for the file, with the number
/// of frames that reported it. The source dimensions are not needed to collect
/// them, which is what lets this run alongside the probe that reads them.
pub(super) fn collect_video_crop_candidates(path: &PathBuf) -> HashMap<(u32, u32, u32, u32), u32> {
    let filter = format!(
        "cropdetect={}:{}:0",
        VIDEO_CROPDETECT_LIMIT, VIDEO_CROPDETECT_ROUND
    );

    // Spawned rather than run through `Command::output` for the same reason the
    // ffprobe next to it is: a probe that is in the job is one a crash cannot leave
    // reading a file with nobody waiting for it.
    let child = match engine_processes::hidden_command("ffmpeg")
        .args([
            "-v",
            "info",
            "-err_detect",
            "ignore_err",
            "-fflags",
            "+genpts+discardcorrupt+igndts",
            "-i",
        ])
        .arg(path)
        .args([
            "-frames:v",
            VIDEO_CROPDETECT_FRAMES,
            "-vf",
            &filter,
            "-f",
            "null",
            "-",
        ])
        .stdout(Stdio::null())
        .stderr(Stdio::piped())
        .spawn()
    {
        Ok(child) => child,
        Err(_) => return HashMap::new(),
    };

    engine_processes::adopt(child.id());

    let output = match wait_bounded(child, Duration::from_secs(VIDEO_PROBE_TIMEOUT_SECS)) {
        Some(output) => output,
        None => return HashMap::new(),
    };

    let stderr = String::from_utf8_lossy(&output.stderr);
    let mut counts: HashMap<(u32, u32, u32, u32), u32> = HashMap::new();
    for line in stderr.lines() {
        if let Some(crop) = parse_cropdetect_line(line) {
            *counts
                .entry((crop.width, crop.height, crop.x, crop.y))
                .or_insert(0) += 1;
        }
    }

    counts
}

/// The crop the detector was most sure of, among those that hold up against the
/// source dimensions.
pub(super) fn best_valid_crop(
    counts: HashMap<(u32, u32, u32, u32), u32>,
    src_w: u32,
    src_h: u32,
) -> Option<VideoCrop> {
    let mut best: Option<(VideoCrop, u32)> = None;
    for ((width, height, x, y), count) in counts {
        let crop = VideoCrop {
            width,
            height,
            x,
            y,
        };
        if !validate_detected_crop(crop, src_w, src_h) {
            continue;
        }

        match best {
            Some((existing, existing_count)) => {
                let existing_area = (existing.width as u64) * (existing.height as u64);
                let candidate_area = (crop.width as u64) * (crop.height as u64);
                if count > existing_count
                    || (count == existing_count && candidate_area > existing_area)
                {
                    best = Some((crop, count));
                }
            }
            None => best = Some((crop, count)),
        }
    }

    best.map(|(crop, _)| crop)
}

/// The geometry a video has already been probed for, when it has been probed at all: the
/// answer the probe gave, read from the cache and nothing else.
///
/// This is the lookup the preview thread is allowed to make — it is a mutex and a hash,
/// not two external processes — and it is what tells a hover whether its file has been
/// measured yet (see `video_probe_due`).
pub(super) fn cached_video_geometry(path: &Path) -> Option<ProbedGeometry> {
    let key = VideoGeometryKey {
        path: path.to_path_buf(),
        version: file_version(path),
    };

    video_geometry_cache().get(&key).copied()
}

/// Probe a video's geometry, from the cache when the file and its version have been
/// probed before.
///
/// This is the one caller that may run the two external processes, so it is only ever
/// called from a thread whose waiting does not matter: the load worker, and the probe a
/// hover is waiting on (see `video_probe_due` in the preview loop). What it answers is
/// held — the failure included — so the next hover of the file is a lookup.
pub(super) fn probe_video_geometry(path: &PathBuf) -> ProbedGeometry {
    let key = VideoGeometryKey {
        path: path.clone(),
        version: file_version(path),
    };

    if let Some(cached) = video_geometry_cache().get(&key) {
        return *cached;
    }

    // Reading the dimensions and detecting the crop are two external processes,
    // and the detector is the one that decodes frames: neither needs the other's
    // answer until the crop is validated, so they run at once and the hover waits
    // for the slower one rather than for both in turn. The subtitle streams are a
    // third, for the same reason and because they are in the file's header rather
    // than in its pictures — a header read does not wait on a decode.
    let (dimensions, candidates, subtitles) = std::thread::scope(|scope| {
        let dimensions = scope.spawn(|| get_video_dimensions(path));
        let candidates = scope.spawn(|| collect_video_crop_candidates(path));
        let subtitles = scope.spawn(|| probe_subtitle_streams(path));

        (
            dimensions.join().unwrap_or(None),
            candidates.join().unwrap_or_default(),
            subtitles
                .join()
                .unwrap_or(SubtitleStreams { count: 0, first: 0 }),
        )
    });

    // A file FFmpeg is not there for — or one its own probe could not read — is asked of
    // the media engine Windows has, which is also the engine that would play it. That is
    // the whole of the fallback's geometry: there is no crop to detect, because cropdetect
    // is an FFmpeg filter and the engine is handed the frame as the file holds it.
    let Some((src_w, src_h, src_duration)) = dimensions
        .or_else(|| video_player::dimensions(path).map(|(width, height)| (width, height, None)))
    else {
        // No picture in the file at all — which leaves two answers, and the one that matters
        // is asked first. A container of a video's name whose streams hold a sound and no
        // picture is a song: the sound is probed for, and a machine that can play it is
        // answered as one from here on, because the router asks this verdict before it asks
        // the video list — so the hover that is replayed for this probe is laid out as the
        // card it is rather than dropped as a video with no shape (see
        // `audio_formats::probed_audio_only`). What is left is a file nothing here can read,
        // which is the answer this arm always gave it.
        let playable = match audio_track::probed(path) {
            Probed::Track(_) => true,
            Probed::Nothing => false,
            Probed::NotAsked => {
                let probed = probe_audio_track(path);
                audio_track::remember(
                    path,
                    match &probed {
                        Some(track) => Probed::Track(track.clone()),
                        None => Probed::Nothing,
                    },
                );
                probed.is_some()
            }
        };

        if playable {
            audio_formats::remember_audio_only(path);
        }

        let mut cache = video_geometry_cache();
        if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, ProbedGeometry::Unmeasurable);

        return ProbedGeometry::Unmeasurable;
    };

    // The film's soundtrack is measured beside its geometry where `Normalize` is on for videos —
    // but on a thread of its own, and that is the whole of the change.
    //
    // **A gain measurement decodes every sample of the file, so its cost is the length of the
    // film rather than anything about the film.** Measured: 828ms for a twenty-second 2160p HEVC
    // clip, against 238ms for the cropdetect pass beside it and 31ms for the header read — and the
    // clip is the short case, because the decoder is what is being waited on and a full-length
    // episode is two orders of magnitude longer than that. It was being waited on *here*, on the
    // thread the hover is waiting on, so a user who had turned `Normalize` on for videos paid a
    // decode of the whole film before the first frame of it was asked for.
    //
    // Nothing is lost by moving it. `start_video_playback` already asks for the gain, and where
    // nothing has measured it yet it starts the scan itself and plays the film as the file holds
    // it — so this was a second measurement of the same thing, on the one thread that must not be
    // waiting, for a gain the very next launch already knows how to do without.
    //
    // At the setting's own start (`Normalize` is off for videos) this is nothing at all, and it
    // stays nothing at every setting — which is what "no full-file decode on the way to a first
    // frame" has to mean to be worth anything.
    spawn_gain_scan(path);

    let crop = best_valid_crop(candidates, src_w, src_h);

    // The frame is kept beside the crop rather than only in the shape, because the two players
    // are told about a crop in different terms: FFmpeg's player in pixels, and the media engine
    // as a share of the frame the rectangle was cut from (see `video_player::Crop`).
    let geometry = if let Some(crop) = crop {
        VideoGeometry {
            width: crop.width,
            height: crop.height,
            frame_width: src_w,
            frame_height: src_h,
            crop: Some(crop),
            duration: src_duration,
            subtitles,
        }
    } else {
        VideoGeometry {
            width: src_w,
            height: src_h,
            frame_width: src_w,
            frame_height: src_h,
            crop: None,
            duration: src_duration,
            subtitles,
        }
    };

    let mut cache = video_geometry_cache();
    if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
        cache.clear();
    }
    cache.insert(key, ProbedGeometry::Measured(geometry));

    ProbedGeometry::Measured(geometry)
}

//! What a film is asked of FFprobe: its size and length, the subtitle streams it holds, and
//! the crop the picture is really drawn in.
//!
//! The first three of those are one pass over one header rather than a pass each, which is what
//! `probe_video_header` is for; the crop is a pass of its own because it decodes frames, and is
//! the leg a hover waits for.

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

/// What one read of a file's header answers: the shape of its first picture, how long it plays,
/// and what subtitle streams it holds.
///
/// The three are one read rather than three because they are all in the same header, and a read
/// of a header is paid for twice over once it has a process of its own: a spawn to begin with,
/// the file opened and its streams walked again, and a deadline of its own standing on the hover
/// that waits for it. **Measured on this machine, two `ffprobe` passes run beside each other cost
/// 34–61 ms against 31–51 ms for one pass carrying all three answers** — and the reason the two
/// were ever apart was not one that survives being written down: each was asking for every
/// stream, so neither could be narrowed, and two passes over one header is two walks of it.
///
/// So one `ffprobe`, no stream selection, asking for every field any of the three wants.
pub(super) struct VideoHeader {
    pub(super) width: u32,
    pub(super) height: u32,
    pub(super) duration: Option<f64>,
    pub(super) subtitles: SubtitleStreams,
    /// The codec name of every subtitle stream, by subtitle-relative index, which is what the
    /// extraction command is built from: one output per codec that has a small form, at the
    /// index `-map 0:s:<i>` masks (see `subtitle_files::extraction_args`).
    pub(super) subtitle_codecs: Vec<String>,
    /// The codec name of every attachment stream, in the order the container holds them, which
    /// is the order the extraction's dump specifier counts in — the container's fonts are only
    /// dumped, and a cover image is not, so the index has to be the container's own rather than
    /// one this app renumbered.
    pub(super) attachment_codecs: Vec<String>,
}

/// Read all three of a file's header answers with one pass over it: the shape of its first
/// picture, its length, and its subtitle streams.
pub(super) fn probe_video_header(path: &PathBuf) -> Option<VideoHeader> {
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
            // No `-select_streams`, because the streams wanted are two kinds of stream and a
            // selection names one of them: `v:0` is the picture on its own and `s` would be the
            // subtitles on their own, so the one selection that could carry both does not exist.
            // Asking for the whole header and reading the two kinds out of it is what makes this
            // one pass rather than two (see `VideoHeader`).
            //
            // `width` and `height` are asked of every stream rather than only of a picture, which
            // is what puts a `width=N/A` in the output for every subtitle stream — see
            // `parse_video_picture`, where that line is the reason the picture is read off the
            // first video stream rather than off whichever line came last. `codec_name` rides
            // along for the same shape of reason: it is the one field the extraction needs and
            // the header already holds, so asking for it here saves the second read of the file
            // the extraction's command would otherwise be built from (see `parse_stream_codecs`).
            "-show_entries",
            "stream=index,codec_type,codec_name,width,height:stream_disposition=default:format=duration",
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

    let (width, height, duration) = parse_video_picture(&output_str)?;
    let (subtitle_codecs, attachment_codecs) = parse_stream_codecs(&output_str);

    Some(VideoHeader {
        width,
        height,
        duration,
        subtitles: parse_subtitle_streams(&output_str),
        subtitle_codecs,
        attachment_codecs,
    })
}

/// The shape of the first picture a probe's flat answer describes, and how long the file is.
///
/// **Only the first video stream is read, and that is the whole of what this half of the parser
/// is for.** Asking for the whole header means every stream prints its own `width` and `height`,
/// and a stream that is not a picture prints `N/A` for both — so a parser that took the last pair
/// it saw, which is all the shape needed when `-select_streams v:0` meant no other stream ever
/// printed one, reads a subtitled file as a container with no picture in it and answers it
/// `Unmeasurable` (see `probe_video_geometry`). Keying off `codec_type` and taking the first
/// video stream instead is what makes the two answers the same answer.
///
/// The length is the format's own rather than a stream's, because a stream with no length of its
/// own — a live capture, a container that does not say — answers `N/A`, which is the same answer
/// as nothing at all here.
pub(super) fn parse_video_picture(probe: &str) -> Option<(u32, u32, Option<f64>)> {
    let mut picture = None;
    let mut duration = None;

    // Whether the stream being read is the picture, and whether one has been taken already — the two
    // kept apart because a file may hold more than one video stream and only the first is its
    // shape. A stream that is not a picture is read for neither, which is what keeps a subtitle
    // stream's `N/A` pair out of the answer.
    let mut wants_picture = false;
    let mut seen_video = false;
    let mut width = None;
    let mut height = None;

    for line in probe.lines() {
        let Some((key, value)) = line.trim().split_once('=') else {
            continue;
        };
        let value = value.trim();

        match key.trim() {
            // A stream begins here, so what the one before it said is finished with — and only a
            // video stream is one whose numbers were taken, so a subtitle stream's `N/A` pair can
            // neither become the shape nor erase one.
            "codec_type" => {
                if wants_picture && width.is_some() && height.is_some() {
                    picture = Some((width?, height?));
                }
                width = None;
                height = None;

                wants_picture = value == "video" && !seen_video;
                seen_video |= value == "video";
            }
            "width" if wants_picture => width = value.parse::<u32>().ok(),
            "height" if wants_picture => height = value.parse::<u32>().ok(),
            "duration" => {
                duration = value
                    .parse::<f64>()
                    .ok()
                    .filter(|value| *value > 0.0 && value.is_finite())
            }
            _ => {}
        }
    }

    // The last stream is never flushed by the line after it, so it is flushed here.
    if wants_picture && width.is_some() && height.is_some() {
        picture = Some((width?, height?));
    }

    let (width, height) = picture?;
    Some((width, height, duration))
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

/// The codec name of every subtitle stream and every attachment in a probe's flat answer.
///
/// The subtitle names are held by subtitle-relative index, because that is the index every other
/// subtitle question here counts in — `-sst s:<i>`, the filter's `si=`, and the extraction's
/// `-map 0:s:<i>` — while the attachment names are held in the container's own stream order,
/// because that is what the extraction's `-dump_attachment:t:<n>` specifier counts in. Both are
/// read out of the same pass that answered the shape and the streams, rather than out of a
/// second look at the file: what each of them is for is building one command, and the command
/// is built once per film (see `subtitle_files::extraction_args`).
///
/// The name belongs to the stream *above* it in FFprobe's flat output — `codec_name` is printed
/// before `codec_type` — so it is held until the type line says which list it goes into, the
/// same shape the disposition above is read with. A stream of any other kind is not a name
/// either list wants.
pub(super) fn parse_stream_codecs(probe: &str) -> (Vec<String>, Vec<String>) {
    let mut subtitles = Vec::new();
    let mut attachments = Vec::new();
    let mut codec = None;

    for line in probe.lines() {
        let Some((key, value)) = line.trim().split_once('=') else {
            continue;
        };

        match key.trim() {
            "codec_name" => codec = Some(value.trim().to_string()),
            "codec_type" => {
                match value.trim() {
                    "subtitle" => subtitles.push(codec.take().unwrap_or_default()),
                    "attachment" => attachments.push(codec.take().unwrap_or_default()),
                    _ => {}
                }
                codec = None;
            }
            _ => {}
        }
    }

    (subtitles, attachments)
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

    video_geometry_cache().get(&key).cloned()
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
        return cached.clone();
    }

    // Two external processes, and the detector is the one that decodes frames: neither needs
    // the other's answer until the crop is validated, so they run at once and the hover waits
    // for the slower one rather than for both in turn.
    //
    // **A file's subtitle streams used to be a third of them**, run beside these two on the
    // reasoning that a header read does not wait on a decode — which was true, and left a hover
    // paying a third spawn and a third deadline for a read that shares a header with the shape
    // and the length anyway. It is the same pass now (see `probe_video_header`), so the count is
    // still in hand before the hover's player is launched and the hover waits for two processes
    // rather than three.
    //
    // **A read of the film's own folder is a third leg, and it is here rather than at the launch
    // for the same reason the gain scan above was moved off the other thread.** Finding the
    // sidecar beside a film is a `read_dir` of that folder (see `video_launch::sidecar_for`), and
    // it was being walked from inside `start_video_playback` — which is the preview thread, the one
    // thread that must not wait, and which every seek, resize, volume change and track change goes
    // back through. It was also the only cold-sensitive piece of the whole path: the subtitle
    // filter itself costs ~40-120ms even on a 4K file carrying 120 000 subtitle events, so a hover
    // is slow cold and fast warm because of the walk and not because of libass.
    //
    // A third leg is nearly free here, which is the whole of what moving it buys. It is a
    // `read_dir` rather than a process, so it costs no spawn and adds no deadline, and the hover
    // waits for the slowest of the three legs rather than for all of them in turn — so the walk is
    // now behind the two processes it was very likely to be slower than anyway. And it is asked
    // once per file and version rather than once per launch, so a resize no longer pays it again.
    let (header, candidates, sidecar) = std::thread::scope(|scope| {
        let header = scope.spawn(|| probe_video_header(path));
        let candidates = scope.spawn(|| collect_video_crop_candidates(path));
        let sidecar = scope.spawn(|| video_launch::sidecar_for(path));

        (
            header.join().unwrap_or(None),
            candidates.join().unwrap_or_default(),
            sidecar.join().unwrap_or(None),
        )
    });

    // A file FFmpeg is not there for — or one its own probe could not read — is asked of
    // the media engine Windows has, which is also the engine that would play it. That is
    // the whole of the fallback's geometry: there is no crop to detect, because cropdetect
    // is an FFmpeg filter and the engine is handed the frame as the file holds it. It is
    // also the whole of what is known of such a file's subtitles, which is nothing: the
    // engine answers with a shape, and the read that would have said what else the file
    // holds is the one that could not read it (see `video_subtitles`).
    let header = header.or_else(|| {
        video_player::dimensions(path).map(|(width, height)| VideoHeader {
            width,
            height,
            duration: None,
            subtitles: SubtitleStreams::default(),
            // Nothing read the header, so nothing knows a codec name — which is the same answer
            // the extraction takes as "no track to copy" (see `extraction_due`).
            subtitle_codecs: Vec::new(),
            attachment_codecs: Vec::new(),
        })
    });

    let Some(header) = header else {
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

    let VideoHeader {
        width: src_w,
        height: src_h,
        duration: src_duration,
        subtitles,
        subtitle_codecs,
        attachment_codecs,
    } = header;

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
    // frame" has to mean to be worth anything. The gate is the setting's own — the one the
    // launch reads too (see `normalizing_video` and `start_video_playback`) — and not ffmpeg's
    // presence, which is all `spawn_gain_scan` asks of its own: a machine that could measure
    // but a user who has not asked it to is a machine that measures nothing here. Nothing
    // tests the gate: the spawn sits behind probes that need FFmpeg, and the one observable
    // that tells a spawned scan from an unspawed one is the gain it leaves behind — readable
    // only once the whole-file decode the scan is has run, which is a wall-clock assertion
    // rather than a seam.
    if normalizing_video() {
        spawn_gain_scan(path);
    }

    // **The film's own subtitle tracks are copied out of it here, once, and that read is what a
    // hover of an embedded-subtitle film stops paying.** The filter that draws an embedded track
    // makes FFmpeg open the film a second time and stream the container to its first subtitle:
    // measured, a cold 1.4 GB MKV took 14 904 ms to a first frame and read 1 423 MB, against
    // 492 ms and 41 MB for the same film drawn from the 30 KB `.ass` this extraction writes. The
    // pass is started on this thread because it is the thread that already runs a film's slow
    // work, and nothing waits for it: what a later hover reads is the geometry entry the thread
    // updates (see `subtitle_files::finish`).
    //
    // The order of the two questions is the order they can be answered: what is already in the
    // cache folder is a `read_dir` of one small folder, and whether this machine would draw the
    // film with FFmpeg's player at all is a question for the engine — asked last, because it is
    // the one that can be expensive and it cannot change the first answer. The engine's files are
    // not FFmpeg's to draw, so a file the media engine plays is never extracted: its subtitles
    // are the engine's business. On a machine with FFmpeg installed that question is answered
    // from the install alone without the file being opened, and where it is not, the answer is
    // memoised per file and version (see `media_engine_plays`).
    let derived = subtitle_files::resolve(path, &subtitle_codecs);
    if !media_engine_plays(path)
        && subtitle_files::extraction_due(
            sidecar.as_deref(),
            derived.as_ref(),
            subtitles.count,
            &subtitle_codecs,
        )
    {
        subtitle_files::spawn_subtitle_extraction(path, &subtitle_codecs, &attachment_codecs);
    }

    let crop = best_valid_crop(candidates, src_w, src_h);

    // The frame is kept beside the crop rather than only in the shape, because the two players
    // are told about a crop in different terms: FFmpeg's player in pixels, and the media engine
    // as a share of the frame the rectangle was cut from (see `video_player::Crop`).
    //
    // The sidecar is held beside the shape rather than looked for again at the launch, which is
    // the whole of what it is here for: `video_sidecar` reads it from this cache entry, so the
    // launch draws a film beside a subtitle file without ever touching the folder again. The
    // derived files are here for the same reason, and they are read off the disk *now* rather
    // than when the extraction finishes, because both answers are one lookup for the launch and
    // the folder is what the next probe would ask again anyway.
    let geometry = if let Some(crop) = crop {
        VideoGeometry {
            width: crop.width,
            height: crop.height,
            frame_width: src_w,
            frame_height: src_h,
            crop: Some(crop),
            duration: src_duration,
            subtitles,
            sidecar,
            derived,
            // Nothing has failed yet as far as this probe knows: a failure is written by the
            // extraction thread into this very entry, and a probe that finds an entry answers
            // from it instead of running again (see the cache lookup at the top).
            subtitle_extraction_failed: false,
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
            sidecar,
            derived,
            subtitle_extraction_failed: false,
        }
    };

    let mut cache = video_geometry_cache();
    if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
        cache.clear();
    }
    // Cloned rather than moved, which is the whole of what dropping `Copy` costs: one string
    // copy per file, once, on the thread that was going to wait for two processes anyway.
    cache.insert(key, ProbedGeometry::Measured(geometry.clone()));

    ProbedGeometry::Measured(geometry)
}

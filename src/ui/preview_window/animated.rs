//! The animated decoders: GIF, APNG, WebP, HEIF and JXL played as a stream of frames,
//! each with its own pacing and its own budget for what may stay decoded.

use super::*;

/// Waits while the player is far enough behind that decoding should pause.
/// Returns false when the preview was cancelled while waiting.
pub(super) fn await_frame_queue_room(
    shared: &Arc<Mutex<StreamedFrames>>,
    cancel: &Arc<AtomicBool>,
) -> bool {
    while !cancel.load(Ordering::Acquire) {
        let queued = shared
            .lock()
            .map(|streamed| streamed.queue.len())
            .unwrap_or(0);
        if queued < ANIMATION_QUEUE_FRAMES {
            return true;
        }
        std::thread::sleep(Duration::from_millis(5));
    }

    false
}

pub(super) fn load_animated_gif(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let file = File::open(path).ok()?;
    let mut decoder = DecodeOptions::new();
    decoder.set_color_output(gif::ColorOutput::RGBA);
    let mut decoder = decoder.read_info(BufReader::new(file)).ok()?;

    let (gif_width, gif_height) = (decoder.width() as u32, decoder.height() as u32);
    let (target_width, target_height) =
        scale_dimensions(gif_width, gif_height, max_width, max_height, preview_scale);

    // The canvas every frame is composited into is the size of the GIF itself, so
    // it is the one allocation here that a file chooses, and it is asked for under
    // the same budget as a decoded picture's.
    let canvas_bytes = frame_bytes_within_budget(gif_width, gif_height, 4)?;

    let mut canvas = vec![0u8; canvas_bytes];
    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut reached_end = false;

    while initial_frames.len() < ANIMATION_STARTUP_FRAMES {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let frame = match decoder.read_next_frame() {
            Ok(Some(frame)) => frame,
            Ok(None) => {
                reached_end = true;
                break;
            }
            Err(_) => return None,
        };

        composite_gif_frame(&mut canvas, frame, gif_width, gif_height);
        let delay_ms = (frame.delay as u32 * 10).max(MIN_ANIMATION_FRAME_DELAY_MS);
        let img = decode_gif_frame_to_image(
            &canvas,
            gif_width,
            gif_height,
            target_width,
            target_height,
            delay_ms,
        )?;
        initial_bytes = initial_bytes.saturating_add(img.pixels.len());
        if initial_bytes > ANIMATION_RETAINED_BYTES {
            return None;
        }
        initial_frames.push(Arc::new(img));
    }

    if initial_frames.is_empty() || (reached_end && initial_frames.len() <= 1) {
        return None;
    }

    if reached_end {
        return Some(MediaData {
            frames: initial_frames,
            shared_frames: None,
            all_frames_loaded: None,
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type: MediaType::AnimatedGif,
            stream_cancel: Some(cancel),
            video_process: None,
            loading_start: None,
            text_state: None,
        });
    }

    let shared = Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    }));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();

    let path_clone = path.clone();
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let mut skip = skip_frames;

        loop {
            let file = match File::open(&path_clone) {
                Ok(f) => f,
                Err(_) => break,
            };
            let mut dec = DecodeOptions::new();
            dec.set_color_output(gif::ColorOutput::RGBA);
            let mut dec = match dec.read_info(BufReader::new(file)) {
                Ok(d) => d,
                Err(_) => break,
            };

            let mut canvas = vec![0u8; (gif_width * gif_height * 4) as usize];
            let mut frame_idx = 0usize;
            let mut cancelled = false;

            while let Ok(Some(frame)) = dec.read_next_frame() {
                if cancel_clone.load(Ordering::Acquire)
                    || !await_frame_queue_room(&shared_clone, &cancel_clone)
                {
                    cancelled = true;
                    break;
                }

                composite_gif_frame(&mut canvas, frame, gif_width, gif_height);
                if frame_idx < skip {
                    frame_idx += 1;
                    continue;
                }

                let delay_ms = (frame.delay as u32 * 10).max(MIN_ANIMATION_FRAME_DELAY_MS);
                if let Some(img) = decode_gif_frame_to_image(
                    &canvas,
                    gif_width,
                    gif_height,
                    target_width,
                    target_height,
                    delay_ms,
                ) {
                    if let Ok(mut streamed) = shared_clone.lock() {
                        streamed.queue.push_back(img);
                    }
                }
                frame_idx += 1;
            }

            // Whether the file has to be decoded again is settled under the same
            // lock the player takes before it gives frames back, so that a release
            // and the end of a pass cannot pass each other: whichever happens
            // first, the other sees it. A pass that ends with nothing given back
            // leaves every frame in the player's hands — the file is its own
            // memory from there, and no frame of it is ever decoded again — while a
            // pass that ends after frames were given back is one to run again.
            let replay = match shared_clone.lock() {
                Ok(mut streamed) => {
                    if streamed.released {
                        true
                    } else {
                        streamed.decoded = true;
                        false
                    }
                }
                // A lock that cannot be taken is nothing left to decode for.
                Err(_) => false,
            };

            if cancelled || !replay {
                break;
            }
            skip = 0;
        }

        loaded_flag_clone.store(true, Ordering::Release);
    });

    Some(MediaData {
        frames: initial_frames,
        shared_frames: Some(shared),
        all_frames_loaded: Some(loaded_flag),
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedGif,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    })
}

/// Open an animated PNG frame iterator for the given path.
pub(super) fn apng_frames(path: &PathBuf) -> Option<image::Frames<'static>> {
    let file = File::open(path).ok()?;
    // Limits are taken at construction rather than set afterwards: a PNG reader
    // holds the ones it was built with, and every frame of the animation is read
    // through it.
    let decoder =
        image::codecs::png::PngDecoder::with_limits(BufReader::new(file), image_decode_limits())
            .ok()?;
    Some(decoder.apng().ok()?.into_frames())
}

/// APNG delays are exact ratios; clamp to the shared floor so a zero-delay
/// animation cannot spin the render loop.
pub(super) fn apng_frame_delay_ms(frame: &image::Frame) -> u32 {
    let (numerator, denominator) = frame.delay().numer_denom_ms();
    if denominator == 0 {
        return MIN_ANIMATION_FRAME_DELAY_MS;
    }
    (numerator / denominator).max(MIN_ANIMATION_FRAME_DELAY_MS)
}

/// Convert an APNG frame into an ImageFrame. The decoder already composites
/// blend and dispose operations, so every frame arrives as the full canvas.
pub(super) fn decode_apng_frame_to_image(
    source: &image::RgbaImage,
    target_width: u32,
    target_height: u32,
    delay_ms: u32,
) -> ImageFrame {
    let (source_width, source_height) = source.dimensions();
    let rgba = if target_width != source_width || target_height != source_height {
        image::imageops::resize(
            source,
            target_width,
            target_height,
            frame_resize_filter(source_width, source_height, target_width, target_height),
        )
        .into_raw()
    } else {
        source.as_raw().clone()
    };

    ImageFrame::new(rgba_to_bgra(&rgba), target_width, target_height, delay_ms)
}

pub(super) fn load_animated_apng(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let mut frames_iter = apng_frames(path)?;

    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut reached_end = false;
    let mut target_size: Option<(u32, u32)> = None;

    while initial_frames.len() < ANIMATION_STARTUP_FRAMES {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let frame = match frames_iter.next() {
            Some(Ok(frame)) => frame,
            Some(Err(_)) => return None,
            None => {
                reached_end = true;
                break;
            }
        };

        let (target_width, target_height) = match target_size {
            Some(size) => size,
            None => {
                let size = scale_dimensions(
                    frame.buffer().width(),
                    frame.buffer().height(),
                    max_width,
                    max_height,
                    preview_scale,
                );
                target_size = Some(size);
                size
            }
        };

        let delay_ms = apng_frame_delay_ms(&frame);
        let img = decode_apng_frame_to_image(frame.buffer(), target_width, target_height, delay_ms);
        initial_bytes = initial_bytes.saturating_add(img.pixels.len());
        if initial_bytes > ANIMATION_RETAINED_BYTES {
            return None;
        }
        initial_frames.push(Arc::new(img));
    }

    if initial_frames.is_empty() || (reached_end && initial_frames.len() <= 1) {
        return None;
    }

    if reached_end {
        return Some(MediaData {
            frames: initial_frames,
            shared_frames: None,
            all_frames_loaded: None,
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type: MediaType::AnimatedApng,
            stream_cancel: Some(cancel),
            video_process: None,
            loading_start: None,
            text_state: None,
        });
    }

    let shared = Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    }));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();
    let (target_width, target_height) = target_size?;

    let path_clone = path.clone();
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let mut skip = skip_frames;

        loop {
            let frames = match apng_frames(&path_clone) {
                Some(frames) => frames,
                None => break,
            };

            let mut cancelled = false;
            for frame in frames.skip(skip) {
                let frame = match frame {
                    Ok(frame) => frame,
                    Err(_) => break,
                };

                if cancel_clone.load(Ordering::Acquire)
                    || !await_frame_queue_room(&shared_clone, &cancel_clone)
                {
                    cancelled = true;
                    break;
                }

                let delay_ms = apng_frame_delay_ms(&frame);
                let img = decode_apng_frame_to_image(
                    frame.buffer(),
                    target_width,
                    target_height,
                    delay_ms,
                );
                if let Ok(mut streamed) = shared_clone.lock() {
                    streamed.queue.push_back(img);
                }
            }

            // Whether the file has to be decoded again is settled under the same
            // lock the player takes before it gives frames back, so that a release
            // and the end of a pass cannot pass each other: whichever happens
            // first, the other sees it. A pass that ends with nothing given back
            // leaves every frame in the player's hands — the file is its own
            // memory from there, and no frame of it is ever decoded again — while a
            // pass that ends after frames were given back is one to run again.
            let replay = match shared_clone.lock() {
                Ok(mut streamed) => {
                    if streamed.released {
                        true
                    } else {
                        streamed.decoded = true;
                        false
                    }
                }
                // A lock that cannot be taken is nothing left to decode for.
                Err(_) => false,
            };

            if cancelled || !replay {
                break;
            }
            skip = 0;
        }

        loaded_flag_clone.store(true, Ordering::Release);
    });

    Some(MediaData {
        frames: initial_frames,
        shared_frames: Some(shared),
        all_frames_loaded: Some(loaded_flag),
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedApng,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    })
}

pub(super) fn decode_webp_animation_frame_to_image(
    bgra: &[u8],
    orig_width: u32,
    orig_height: u32,
    target_width: u32,
    target_height: u32,
    delay_ms: u32,
) -> Option<ImageFrame> {
    let expected_bgra = orig_width as usize * orig_height as usize * 4;
    if bgra.len() != expected_bgra {
        return None;
    }

    let pixels = if target_width == orig_width && target_height == orig_height {
        bgra.to_vec()
    } else {
        let mut rgba = Vec::with_capacity(expected_bgra);
        for chunk in bgra.as_chunks::<4>().0.iter() {
            rgba.push(chunk[2]);
            rgba.push(chunk[1]);
            rgba.push(chunk[0]);
            rgba.push(chunk[3]);
        }
        let img = image::RgbaImage::from_raw(orig_width, orig_height, rgba)?;
        let resized = image::imageops::resize(
            &img,
            target_width,
            target_height,
            frame_resize_filter(orig_width, orig_height, target_width, target_height),
        );
        rgba_to_bgra(&resized.into_raw())
    };

    Some(ImageFrame::new(
        pixels,
        target_width,
        target_height,
        delay_ms,
    ))
}

pub(super) fn load_animated_webp(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    // The file is read whole — libwebp decodes from bytes rather than from a reader
    // of ours — so it is read under the budget like every other file a hover opens.
    let buffer = Arc::new(read_within_budget(path)?);
    let options = webp_animation::DecoderOptions {
        use_threads: true,
        color_mode: webp_animation::ColorMode::Bgra,
    };
    let decoder = webp_animation::Decoder::new_with_options(buffer.as_slice(), options).ok()?;

    // Frames are decoded at the animation's own size, and this reader is libwebp's
    // rather than the `image` crate's, so the budget is asked for by hand where every
    // decoder above is handed it.
    let (orig_width, orig_height) = decoder.dimensions();
    if orig_width == 0 || orig_height == 0 {
        return None;
    }

    frame_bytes_within_budget(orig_width, orig_height, 4)?;

    let (target_width, target_height) = scale_dimensions(
        orig_width,
        orig_height,
        max_width,
        max_height,
        preview_scale,
    );
    if target_width == 0 || target_height == 0 {
        return None;
    }

    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut previous_timestamp = 0i32;
    let mut reached_end = false;
    let mut iterator = decoder.into_iter();

    while initial_frames.len() < ANIMATION_STARTUP_FRAMES {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let frame = match iterator.next() {
            Some(frame) => frame,
            None => {
                reached_end = true;
                break;
            }
        };

        let timestamp = frame.timestamp();
        let delay_ms = (timestamp - previous_timestamp).max(0) as u32;
        previous_timestamp = timestamp;

        let img = decode_webp_animation_frame_to_image(
            frame.data(),
            orig_width,
            orig_height,
            target_width,
            target_height,
            delay_ms,
        )?;
        initial_bytes = initial_bytes.saturating_add(img.pixels.len());
        if initial_bytes > ANIMATION_RETAINED_BYTES {
            return None;
        }
        initial_frames.push(Arc::new(img));
    }

    if initial_frames.is_empty() || (reached_end && initial_frames.len() <= 1) {
        return None;
    }

    if reached_end {
        return Some(MediaData {
            frames: initial_frames,
            shared_frames: None,
            all_frames_loaded: None,
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type: MediaType::AnimatedWebP,
            stream_cancel: Some(cancel),
            video_process: None,
            loading_start: None,
            text_state: None,
        });
    }

    let shared = Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    }));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();

    drop(iterator);
    let buffer_clone = Arc::clone(&buffer);
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let mut skip = skip_frames;

        loop {
            let options = webp_animation::DecoderOptions {
                use_threads: true,
                color_mode: webp_animation::ColorMode::Bgra,
            };
            let decoder =
                match webp_animation::Decoder::new_with_options(buffer_clone.as_slice(), options) {
                    Ok(decoder) => decoder,
                    Err(_) => break,
                };

            let mut previous_timestamp = 0i32;
            let mut cancelled = false;

            for (frame_idx, frame) in decoder.into_iter().enumerate() {
                if cancel_clone.load(Ordering::Acquire)
                    || !await_frame_queue_room(&shared_clone, &cancel_clone)
                {
                    cancelled = true;
                    break;
                }

                let timestamp = frame.timestamp();
                let delay_ms = (timestamp - previous_timestamp).max(0) as u32;
                previous_timestamp = timestamp;

                if frame_idx < skip {
                    continue;
                }

                if let Some(img) = decode_webp_animation_frame_to_image(
                    frame.data(),
                    orig_width,
                    orig_height,
                    target_width,
                    target_height,
                    delay_ms,
                ) {
                    if let Ok(mut streamed) = shared_clone.lock() {
                        streamed.queue.push_back(img);
                    }
                }
            }

            // Whether the file has to be decoded again is settled under the same
            // lock the player takes before it gives frames back, so that a release
            // and the end of a pass cannot pass each other: whichever happens
            // first, the other sees it. A pass that ends with nothing given back
            // leaves every frame in the player's hands — the file is its own
            // memory from there, and no frame of it is ever decoded again — while a
            // pass that ends after frames were given back is one to run again.
            let replay = match shared_clone.lock() {
                Ok(mut streamed) => {
                    if streamed.released {
                        true
                    } else {
                        streamed.decoded = true;
                        false
                    }
                }
                // A lock that cannot be taken is nothing left to decode for.
                Err(_) => false,
            };

            if cancelled || !replay {
                break;
            }
            skip = 0;
        }
        loaded_flag_clone.store(true, Ordering::Release);
    });

    Some(MediaData {
        frames: initial_frames,
        shared_frames: Some(shared),
        all_frames_loaded: Some(loaded_flag),
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedWebP,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    })
}

/// An AVIF or HEIF image sequence, played through the media engine Windows has.
///
/// This is a *decode* of the sequence, not the video path: frames are pulled synchronously
/// and handed to the same queue every other animation is, so a moving `.avif` behaves like
/// a moving `.gif` — it keeps the animated scale, the sliding window, and the picture's
/// gate — rather than becoming a video the moment it moves. That is a deliberate choice:
/// the alternative is a preview that changes kind the moment the file turns out to
/// animate, and the user is looking at a picture.
///
/// The media engine is asked for a sequence it may not have a demuxer for, and the AV1 and
/// HEVC codecs may not be installed, so every failure here is `None` — and `None` means the
/// still path draws the first frame, which is what this app did before any of it existed.
/// Nothing about a moving picture is a reason to show nothing at all (see
/// `heif_sequence`).
pub(super) fn load_animated_heif(
    path: &Path,
    max_width: u32,
    max_height: u32,
    _preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let sequence = heif_sequence::decode_sequence(path, max_width, max_height, &cancel)?;

    if sequence.frames.len() <= 1 {
        // One frame is a picture, not an animation: the animated reader is turned down for
        // it and the still path draws it, exactly as it does for a `.gif` with one frame.
        return None;
    }

    let mut frames = Vec::with_capacity(sequence.frames.len());
    for (pixels, delay_ms) in sequence.frames {
        frames.push(Arc::new(ImageFrame::new(
            pixels,
            sequence.width,
            sequence.height,
            delay_ms,
        )));
    }

    Some(MediaData {
        frames,
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedHeif,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// An animated JPEG XL, read by `jxl-oxide`.
///
/// Where every other animation here is streamed, this one is decoded whole: the delays
/// live in the codestream's image header, so the frame count and every frame's hold are
/// known before a single pixel is rendered, and there is nothing to learn after the first
/// frame that would be worth opening the preview early for. So there is no queue and no
/// decoder thread, and the bounds that hold here are its own rather than the sliding
/// window's — which does not reach this path at all, it is the streaming path's, and a
/// `MediaData` with no queue behind it never releases anything.
pub(super) fn load_animated_jxl(
    path: &Path,
    max_width: u32,
    max_height: u32,
    _preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    // The timing first, and separately from the pixels: it is cheap beside a decode, and a
    // file that is not a sequence is settled here without a single frame being rendered.
    let (delays_ms, (orig_width, orig_height)) = jxl_image::animation(path)?;

    if delays_ms.len() <= 1 {
        return None;
    }
    if orig_width == 0 || orig_height == 0 {
        return None;
    }
    frame_bytes_within_budget(orig_width, orig_height, 4)?;

    // `jxl-oxide` has no scaler of its own, so the frame is decoded whole and the bound
    // is a refusal rather than a target: a picture larger than the box is not shown at a
    // smaller size here, it is not shown. What the other decoders spend a filter on, this
    // one spends a return of `None` on.
    if orig_width > max_width || orig_height > max_height {
        return None;
    }

    // Everything is held at once, so the total is what has to be bounded rather than any
    // one frame. A thousand-frame sequence at a size that fits a display is gigabytes, and
    // there is no window to release them through: the preview holds them until the next
    // hover replaces the whole `MediaData`. So this is refused, and refused the same way
    // a file over the decode budget is — a first frame is the right answer for it, and the
    // still path gives exactly that.
    let frame_bytes = usize::try_from(orig_width)
        .ok()
        .and_then(|width| width.checked_mul(usize::try_from(orig_height).ok()?))
        .and_then(|pixels| pixels.checked_mul(4))?;
    let retained = frame_bytes.checked_mul(delays_ms.len())?;
    if retained > ANIMATION_RETAINED_BYTES {
        return None;
    }

    let mut frames = Vec::with_capacity(delays_ms.len());
    let mut retained = 0usize;
    for (index, delay_ms) in delays_ms.iter().copied().enumerate() {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let pixels = jxl_image::decode_frame(path, index, max_width, max_height)?;
        retained = retained.saturating_add(pixels.len());
        if retained > ANIMATION_RETAINED_BYTES {
            return None;
        }

        // The same floor every other animation here applies, and for the same reason: a
        // frame's own duration is whatever the file says, and a file saying zero — or a
        // still timebase, which is zero too — is a frame the playhead would advance
        // through as fast as the message pump allows, spinning the render loop on a
        // picture that never appears to change.
        frames.push(Arc::new(ImageFrame::new(
            pixels,
            orig_width,
            orig_height,
            delay_ms.max(MIN_ANIMATION_FRAME_DELAY_MS),
        )));
    }

    Some(MediaData {
        frames,
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedJxl,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

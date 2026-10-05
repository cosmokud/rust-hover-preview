//! What the rest of the crate asks of the preview window, and what the renderers answer it
//! with: the calls that put a preview up, take it down and refresh it, and the notices a render
//! sends back.

use super::*;

/// Whether a preview is pinned — a window of its own with a caption, which stays until it
/// is closed. The Explorer hook asks this every tick: while it is true nothing is spawned
/// and nothing is taken down, which is the whole of what "the preview mode is paused"
/// means (see `pin_window::pin_is_up`).
pub fn pinned() -> bool {
    pin_window::pin_is_up()
}

/// Ask for the pinned preview to come down, from any thread. What a pin is belongs to the
/// preview loop, so the loop is what ends it — on its next tick, by the same path its own
/// close button takes. Asking twice is asking once.
///
/// Named for what it is rather than `end_pin`, because it is not an end: it is a request, and
/// the end is the loop's own `end_pin_state(Reason::Asked)`, which is the same call the
/// caption's cross makes and the only place a pin actually ends on this road.
pub fn request_pin_end() {
    PIN_END_REQUESTED.store(true, Ordering::Release);
}

/// Whether a pin has been closed since this was last asked. The Explorer hook reads it once
/// per tick to know that what is under the pointer is a hover it has not answered yet: the
/// file it was on when the pin went up is not a file the pointer has left and come back to,
/// and treating it as one would hold the next preview back for the re-hover delay (see
/// `pin_window::take_pin_resumed`).
pub fn take_pin_resumed() -> bool {
    pin_window::take_pin_resumed()
}

/// The file the pinned window is showing, if there is one. It is what the Explorer hook reads
/// to know whether the file the pointer picks is a file the pin is already showing, and it is
/// the file a *new* pin is told apart from an old one by (see `PinState`).
pub fn pinned_path() -> Option<PathBuf> {
    pin_state().and_then(|state| state.pin().map(|pin| pin.path.clone()))
}

/// Ask for the pinned window to be shown another file, from any thread: the Explorer hook's
/// answer to the user clicking a file, selecting one with the keyboard, or — where
/// `Pin Mode → Update Preview → On Hover` asks for it — settling on one.
///
/// What a pin is belongs to the preview loop, so the loop is what takes the swap up, on its
/// next tick, by the path the tray's own rows take. A pin that is not up is not a window to
/// show anything in, and a file the pin is already showing is not another file: both are
/// dropped here rather than sent, since the hook has no other answer to give them.
pub fn update_pinned_preview(path: &Path) {
    if !pinned() || !pin_update_enabled() {
        return;
    }

    if pinned_path().as_deref() == Some(path) {
        return;
    }

    send_preview(PreviewMessage::PinUpdate(path.to_path_buf()));
}

/// Show a preview of the file the pointer hovers, opened from the cursor it was hovered
/// at.
///
/// The region that comes with it is the one the Explorer hook measured off the item the
/// file is — left, top, right and bottom of it, with whether it is a column of the
/// view rather than the item's own text — and it is what the placement is kept off (see
/// `AvoidRegion`). A hover the hook could not measure a region for comes as `None` and
/// is placed by its position mode alone.
pub fn show_preview(path: &Path, x: i32, y: i32, avoid: Option<((i32, i32, i32, i32), bool)>) {
    // A pinned preview is the whole of what this app is showing: a hover raised while one
    // is up would be a second thing on screen, and the pin exists to stop exactly that
    // (see `PIN_ACTIVE`).
    if pinned() {
        return;
    }

    // The hook answers with the region and whether it is a column of the view rather
    // than the item's own text (see `AvoidRegion`).
    let avoid = avoid.map(|(region, column)| {
        if column {
            AvoidRegion::column(region)
        } else {
            AvoidRegion::text(region)
        }
    });

    send_preview(PreviewMessage::Show(path.to_path_buf(), x, y, avoid));
}

pub fn show_preview_keyboard(
    path: &Path,
    item_left: i32,
    item_top: i32,
    item_right: i32,
    item_bottom: i32,
    avoid: Option<ScreenRegion>,
    draws_columns: bool,
) {
    // A keyboard hover is a hover like any other, and a pinned preview is what is on
    // screen instead of one.
    if pinned() {
        return;
    }

    // A keyboard preview is not the pointer's, so the item the pointer was last read
    // on has nothing to say about it: the box goes rather than gating a preview the
    // keyboard asked for (see `HOVER_POINTER_BOX`).
    clear_pointer_item_box();

    send_preview(PreviewMessage::ShowKeyboard(
        path.to_path_buf(),
        item_left,
        item_top,
        item_right,
        item_bottom,
        avoid,
        draws_columns,
    ));
}

pub fn hide_preview() {
    // A pinned preview is not a hover, and the twenty-odd dismissals the Explorer hook
    // sends while a pointer moves are not about it: coming down is what its close button
    // is for, and a pin that went away because the pointer crossed the window would be no
    // better than the hover it came from. The process behind it is left alone here for
    // the same reason — what ends it is the take-down the pin's own end performs (see
    // `PIN_ACTIVE` and `end_pin`).
    if pinned() {
        return;
    }

    // The window coming down, the count of its coming down and the item the pointer
    // was last read on are written under one lock, so a load that is still running —
    // started under the count from before this — cannot put the preview back up after
    // it, and nothing is revealed against a hover that has already gone (see
    // `HIDDEN_EPOCH` and `HOVER_POINTER_BOX`).
    {
        let mut hidden = HIDDEN_EPOCH.lock().ok();
        if let Some(hidden) = hidden.as_mut() {
            **hidden += 1;
        }
        clear_pointer_item_box();
        unsafe {
            let hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
            if !hwnd.is_invalid() {
                // Posted and not sent: the window belongs to the preview loop, and a
                // send to a loop that is busy is this thread — the hook's — held until
                // it pumps again, which is the one wait a dismissal must not have. What
                // the hide is *for* is the count above, and that is already moved, so
                // the window comes down a moment later rather than now; the loop takes
                // it down itself on the `Hide` below as well, and either of the two is
                // the same window state (see `HIDDEN_EPOCH` and `preview_stall_ms`).
                let _ = ShowWindowAsync(hwnd, SW_HIDE);
            }
        }
    }

    if let Ok(mut current) = CURRENT_MEDIA.try_lock() {
        if let Some(ref mut media) = *current {
            media.cancel_background_work();
            // The player process is ended here and the media engine is not: a session
            // belongs to the thread that made it, which is the preview loop's and not this
            // one (see `video_player`), so a stop asked for from this thread would touch
            // nothing while the sound went on playing. What ends the engine is the loop's
            // own take-down, and the `Hide` below is what calls for it — the media is left
            // where it is for that take-down to find rather than taken away, because a
            // media taken here is one nothing is left to stop.
            kill_player_process(media);
        }
    } else {
        // The media state is locked elsewhere; kill the recorded ffplay by PID
        // instead (verified to still be ffplay before terminating it).
        kill_stray_video_process();
    }

    send_preview(PreviewMessage::Hide);
}

pub fn refresh_preview() {
    send_preview(PreviewMessage::Refresh);
}

/// The tray's `Render HTML` row has been switched off: a page the engine is drawing for a
/// hover comes down with the switch it was asked for by, while a page a pin is showing is
/// left exactly as it is — a pin is not a hover, and what a setting asks is owed to the next
/// hover rather than to a window that stands (see `PreviewMessage::Refresh`, which leaves a
/// pin alone for the same reason, and `webview_preview::hide_html_preview`).
pub fn refresh_render_html() {
    if pinned() {
        return;
    }

    webview_preview::hide_html_preview();
}

/// The tray's `Pin Mode → Enable` row was clicked, or the configuration that decides whether
/// the key is watched was reloaded. Whether there is a pin to take down is a question
/// only the preview thread can answer — the window and the media under it are its own
/// — so the answer is left to it.
pub fn refresh_pin() {
    send_preview(PreviewMessage::PinChanged);
}

/// A preview type was switched on or off in the tray, which is a question only
/// the preview thread can answer: whether what is on screen is of that kind.
pub fn refresh_preview_types() {
    send_preview(PreviewMessage::RefreshTypes);
}

/// The render tier is done with a document. Sent from the engine thread through
/// the same channel every other message arrives on, so the preview loop learns
/// about a page the moment it exists.
pub fn notify_office_render(path: &Path, generation: u64, ok: bool) {
    send_preview(PreviewMessage::OfficeRenderReady {
        path: path.to_path_buf(),
        generation,
        ok,
    });
}

/// A video's probe is done. Sent from the thread the probe ran on, through the same
/// channel every other message arrives on, so the hover that was waiting for it is
/// replayed the moment there is an answer to place it with.
pub(super) fn notify_video_probed(path: &Path, generation: u64) {
    send_preview(PreviewMessage::VideoProbed {
        path: path.to_path_buf(),
        generation,
    });
}

/// The small subtitle files a film's own tracks were copied into are in the cache. Sent from
/// the extraction thread through the same channel every other answer arrives on, because the
/// one thing waiting on it is a pinned window: a pin showing that film is begun again so the
/// frame it draws next is drawn with the copy, and beginning a player belongs to the preview
/// loop (see `reload_pinned_subtitles`). A hover is told nothing — it was answered without
/// subtitles by the user's own choice, and the hover after it reads what this wrote (see
/// `video_launch::subtitle_filter`).
pub(super) fn notify_video_subtitles_ready(path: &Path) {
    send_preview(PreviewMessage::VideoSubtitlesReady(path.to_path_buf()));
}

/// The ImageMagick engine is done with a file. Sent from the engine's own thread, through the
/// same channel every other answer arrives on, so the hover that was waiting is replayed the
/// moment there is a picture to place it with — or taken down, where the answer is that the
/// file is not one the engine can read.
pub fn notify_magick_ready(path: &Path, generation: u64, ok: bool) {
    send_preview(PreviewMessage::MagickReady {
        path: path.to_path_buf(),
        generation,
        ok,
    });
}

/// And the PeaZip engine, on the same terms: the archive has been listed, or it is one the engine
/// will not list. What is waiting on it is a hover that has already asked and is showing the
/// spinner — the page an archive is drawn as cannot be measured before its listing exists — and
/// what the answer does is replay that hover, with the listing in hand or with nothing at all.
pub fn notify_peazip_ready(path: &Path, generation: u64, ok: bool) {
    send_preview(PreviewMessage::PeazipReady {
        path: path.to_path_buf(),
        generation,
        ok,
    });
}

/// Probe a video's geometry on a thread of its own, and tell the preview loop.
///
/// The probe is two external processes and the hover waits for the slower of them, so it
/// is done here rather than on the preview thread: what is on screen while it runs is the
/// waiting spinner, and the hover it belongs to is replayed when the answer lands (see
/// `video_probe_due` and `video_probe` in the preview loop). A probe whose hover has moved
/// on is not wasted — what it answers is held for the next hover of the file — so nothing
/// here is cancelled or waited for.
///
/// What is not left to the probe is whether the hover is answered at all: the wait is not
/// one an engine or a cap will ever end, so the answer is sent whatever the probe did —
/// including a probe that panicked, which is a thread that would otherwise unwind past the
/// notify and leave the spinner standing over a file it has finished with (see
/// `video_probe_due` and `awaiting_engine`).
pub(super) fn spawn_video_probe(path: PathBuf, generation: u64) {
    std::thread::spawn(move || {
        let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            probe_video_geometry(&path);

            // Which of the two engines plays the file is a question for the engine itself, and
            // the hover this probe is for is replayed on the thread that draws it — where the
            // answer has to be in hand rather than asked for. What the ask is, is a source reader
            // over the file with a decoder chain built for it, which is the shape of work this
            // thread exists to keep off that one; and the answer is held per file and version, so
            // what the replay, the load and the pin after it pay is a lookup (see
            // `media_engine_plays`). It is asked of every video the probe runs for, because the
            // machine does not get a vote in which of the two engines is started for a file: where
            // FFmpeg is installed the answer is no without the file being opened at all, and where
            // it is not, this is the ask that decides whether the file has a preview at all —
            // including in the layout, which reads the same answer out of an unmeasurable video's
            // fall-back box.
            let _ = media_engine_plays(&path);
        }));

        notify_video_probed(&path, generation);
    });
}

/// Which park the frame being rendered is for, counted rather than flagged.
///
/// It is the one thing the thread rendering that frame cannot be told by the loop, because the loop
/// is what answers the park — so the answer is written down here and the render reads it between
/// its reads, which is the only moment it can be read at without holding a pipe open past the end
/// (see `spawn_video_resume_frame`).
///
/// **It is a count of parks and not a flag, because a park is not the only one that can be in
/// flight.** A flag says whether *some* park has been given back, and a render is up to
/// `RESUME_FRAME_WAIT` of FFmpeg away from answering: a second park begun inside that window took
/// the flag back down, so the first render read a park that had ended as a park that had not, and
/// installed a frame taken at the second the *previous* drag let go at into a band the new park
/// was holding for a picture of its own. A render is therefore given the park it was opened for
/// and asked whether that is still the one this app is waiting for, which two overlapping parks
/// can never both be (see `resume_frame_asked_for`).
pub(super) static RESUME_FRAME_PARK: AtomicU64 = AtomicU64::new(0);

/// Open a park's frame, answering the park it is for: a render begun here is wanted only while
/// that park is the current one.
pub(super) fn begin_resume_frame() -> u64 {
    RESUME_FRAME_PARK
        .fetch_add(1, Ordering::SeqCst)
        .wrapping_add(1)
}

/// Close every park's frame — a park that has been given back has a player in the band and no use
/// for a decode, and a render still inside its bound is left to find out for itself (see
/// [`RESUME_FRAME_PARK`]).
pub(super) fn abandon_resume_frame() {
    RESUME_FRAME_PARK.fetch_add(1, Ordering::SeqCst);
}

/// Whether the render opened for the park named `park` is still the one this app is waiting for.
///
/// It is the render's own refusal, asked twice: once between its reads, so a park that ended while
/// it was still working is not waited out to the bound, and once with the frame in hand, so a
/// render the loop has stopped wanting cannot put its pixels in a slot a later park will install
/// from.
pub(super) fn resume_frame_asked_for(park: u64) -> bool {
    RESUME_FRAME_PARK.load(Ordering::Acquire) == park
}

/// How long a resume frame may take before it is not worth the drag it is for.
///
/// A drag is over in about a second, and this is a whole extra FFmpeg pass over a file — open, seek,
/// decode one frame — so a render slower than the drag it is preparing for has nothing to prepare.
/// The bound is what makes the wait safe to leave on a thread of its own: a render that overruns is
/// killed and reaped rather than left reading a file for a park that ended long ago.
pub(super) const RESUME_FRAME_WAIT: Duration = Duration::from_millis(1200);

/// Render the frame a resize's park will give back, on a thread of its own.
///
/// **It is spawned at the park rather than asked for at the release, because the release is the one
/// moment there is no time.** A drag holds the pointer's own thread for its whole length, and the
/// work is two external processes — open the file, seek to the second the film will go on from,
/// decode one frame — which is the hundred-millisecond class on any file a preview can be played
/// for. Spawned at the park it overlaps the drag the hand is already spending; asked for at the
/// release it would be the release that waits, which is the frame the hand let go on being replaced
/// by nothing at all for a tenth of a second.
///
/// It is asked for on a **resize** and not on a move, which is the whole of what makes it worth
/// doing at all: a move has no relaunch behind it, so the player standing in the band when the drag
/// ends is the player that was playing all along and has a decoded frame of its own to show (see
/// `park_swap_arm`). A resize ends in a relaunch, and the placeholder is standing in for a window
/// that has decoded nothing — which is the whole of the case this frame is for.
///
/// Nothing is waited for and nothing is asked of the loop: the answer is left in its own slot and
/// taken up by the next settle (see `install_resume_frame`), so an overrun is a park that keeps its
/// stale frame rather than a band that goes blank. And the process is ended on every path out of
/// here — the park's end, the bound, or a panic — because it is this app's own child and Windows
/// does not end it for us.
///
/// **It is opened for a named park and asked again on the way out, so a render that outlives its
/// park cannot answer a later one.** A park's end closes it, and a later park opens the next — so
/// the second is told apart from the first even while the first is still decoding, which is what a
/// single flag could not do (see `RESUME_FRAME_PARK`).
pub(super) fn spawn_video_resume_frame(path: PathBuf, at: f64, width: u32, height: u32) {
    // FFmpeg's chroma planes want even dimensions, and a band of one pixel is not a frame anyway.
    let (width, height) = (width & !1, height & !1);
    if width < 2 || height < 2 || !at.is_finite() || at < 0.0 {
        return;
    }

    let park = begin_resume_frame();

    std::thread::spawn(move || {
        // A panic in here would take the thread down before it can see it is no longer wanted and
        // leave the child running, and this is a thread nobody is waiting on: the same reason the
        // probe above catches its own unwind (see `spawn_video_probe`).
        let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            render_video_resume_frame(&path, at, width, height, park);
        }));
    });
}

/// The render itself: one frame of the file at one second, into the slot the settle takes it from.
///
/// The child is polled rather than waited on because the pipe it writes to is this thread's to
/// drain, and a `wait_with_output` has no moment in it at which a park that has ended can be
/// noticed — so the drain is a thread of its own and the wait is this loop, which is where the
/// park and the bound are both read. `park` is the one it was opened for, and a render that is no
/// longer that park's stops where it stands rather than to the bound (see `resume_frame_asked_for`).
fn render_video_resume_frame(path: &Path, at: f64, width: u32, height: u32, park: u64) {
    use std::io::Read;

    let seconds = format!("{at:.3}");
    let size = format!("{width}x{height}");

    let Ok(mut child) = engine_processes::hidden_command("ffmpeg")
        .args(["-v", "error", "-nostdin", "-ss"])
        .arg(&seconds)
        .arg("-i")
        .arg(path)
        .args([
            "-frames:v",
            "1",
            "-s",
            &size,
            "-f",
            "rawvideo",
            "-pix_fmt",
            "bgra",
            "-",
        ])
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
    else {
        return;
    };
    let pid = child.id();
    engine_processes::adopt(pid);

    let wanted = width as usize * height as usize * 4;
    let Some(pipe) = child.stdout.take() else {
        let _ = child.kill();
        let _ = child.wait();
        engine_processes::forget(pid);
        return;
    };

    let collected: Arc<Mutex<Vec<u8>>> = Arc::new(Mutex::new(Vec::new()));
    let sink = Arc::clone(&collected);
    let drain = std::thread::spawn(move || {
        let mut all = Vec::new();
        let mut pipe = pipe;
        let _ = pipe.read_to_end(&mut all);
        if let Ok(mut sink) = sink.lock() {
            *sink = all;
        }
    });

    let deadline = Instant::now() + RESUME_FRAME_WAIT;
    while !drain.is_finished() {
        if !resume_frame_asked_for(park) || Instant::now() >= deadline {
            break;
        }
        std::thread::sleep(Duration::from_millis(8));
    }

    // Reaped on every path out, and bounded even on the path where nothing went wrong: the same
    // helper the probes use, for the same reason — a child still writing into a pipe nobody is
    // draining is a child that never finishes, and a child that is killed is waited for rather
    // than left for the next drag to find (see `wait_bounded`).
    let _ = wait_bounded(child, Duration::from_millis(250));
    let _ = drain.join();
    engine_processes::forget(pid);

    if !resume_frame_asked_for(park) {
        return;
    }

    // A short read is a frame that was not written whole, which is answered with nothing rather
    // than with the part of it that was: a band filled from half a frame is a band of garbage.
    let mut pixels = match collected.lock() {
        Ok(all) if all.len() == wanted => all.clone(),
        _ => return,
    };
    // GDI is not what wrote this, but a layered window is composited from premultiplied coverage,
    // and this frame is opaque everywhere for the same reason the captured one is (see
    // `hold_video_window_frame`).
    for pixel in pixels.as_chunks_mut::<4>().0 {
        pixel[3] = 255;
    }

    if let Ok(mut prepared) = RESUME_VIDEO_FRAME.lock() {
        *prepared = Some(HeldVideoFrame {
            pixels,
            width,
            height,
        });
    }
}

/// Screen-space box of the preview surface that is on screen right now, if any.
/// The Explorer hook uses it to decide whether a keyboard preview was placed
/// over the parked pointer, so a pointer sitting under the preview cannot drive
/// previews or dismiss them.
///
/// An SVG document's preview is the engine's window rather than this app's, so its box
/// is asked for as well — a document is a preview of this app's in every way but the
/// window it is drawn in.
/// Whether `handle` is this app's own preview window.
///
/// The Explorer hook asks it so that a click landing on a pinned window is read out of the
/// listing that window is standing on, rather than being taken as a click on something that
/// is not a listing at all (see `click_is_over_a_listing`).
pub fn is_preview_window(handle: isize) -> bool {
    handle != 0 && handle == PREVIEW_HWND.load(Ordering::Acquire)
}

pub fn preview_screen_rect() -> Option<(i32, i32, i32, i32)> {
    if let Some(rect) = webview_preview::screen_rect() {
        return Some(rect);
    }

    unsafe {
        let candidates = [
            PREVIEW_HWND.load(Ordering::SeqCst),
            VIDEO_HWND.load(Ordering::SeqCst),
        ];

        for hwnd_value in candidates {
            if hwnd_value == 0 {
                continue;
            }

            let hwnd = HWND(hwnd_value as *mut _);
            if !IsWindowVisible(hwnd).as_bool() {
                continue;
            }

            let mut rect = RECT::default();
            if GetWindowRect(hwnd, &mut rect).is_ok()
                && rect.right > rect.left
                && rect.bottom > rect.top
            {
                return Some((rect.left, rect.top, rect.right, rect.bottom));
            }
        }
    }

    None
}

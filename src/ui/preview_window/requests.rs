//! What the rest of the crate asks of the preview window, and what the renderers answer it
//! with: the calls that put a preview up, take it down and refresh it, and the notices a render
//! sends back.

use super::pin_trace::note_pin_trace;
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

/// Ask for the card's menu to be put away, from any thread — the Explorer hook's answer to a
/// press that landed on a window that is not the pin's own. The pin window is
/// `WS_EX_NOACTIVATE` and holds no capture, so a press on another window never reaches its
/// procedure: the hook's poll is the only side that sees every press, and this is the seam a
/// pin's own end rides (see `request_pin_end`). Asking twice is asking once, and only the
/// menu goes — the pin itself is left standing.
pub fn request_pin_menu_dismiss() {
    PIN_MENU_DISMISS_REQUESTED.store(true, Ordering::Release);
}

/// One of the tray's two `Volume` rows was moved, and which half of it: the film's, or the
/// sound's.
///
/// A pin is playing at a level of its own — the one it was taken up at, and whatever a hand
/// has done on its knob since — so a row moved in the tray is the one road that has to reach
/// past that and say the level out loud to whatever is already up. It is said to the loop,
/// which is where the pin's player and its clock are, and a pin showing a file of the other
/// half is left exactly where it is (see `tray_volume_moved`).
pub fn publish_tray_volume(audio: bool) {
    send_preview(PreviewMessage::TrayVolume { audio });
}

/// The tray's `Text → Font Size` row was picked, and the size it names.
///
/// A page a pin is showing is drawn at a size of its own — the configuration's at the take-up, and
/// whatever a `Ctrl`+wheel has put it to since — so the row is a second road that has to reach past
/// that and say the size out loud to a pin already up. It is said to the loop, which is where the
/// page is drawn, and it is the row's own: a size hand-edited into the file, or the page a reader
/// has wheeled, is read where it is rather than written over (see `redraw_pinned_page`).
pub fn publish_text_font_size(percent: u32) {
    send_preview(PreviewMessage::TrayTextSize { percent });
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
///
/// **A file the pin is already *on its way to* is not another file either**, and dropping it is the
/// whole of what keeps one click from restarting another's swap: the pin's own `path` does not move
/// until the incoming file is installed, so a press on the file being loaded or held used to be
/// read as a new pick — the load dropped, the hold abandoned, its player killed — and the picture
/// that was being kept went with them (see `pin_install::THE_FILE_IN_FLIGHT`).
pub fn update_pinned_preview(path: &Path) {
    if !pinned() || !pin_update_enabled() {
        return;
    }

    if pinned_path().as_deref() == Some(path) {
        return;
    }

    if pin_install::pin_file_in_flight().as_deref() == Some(path) {
        note_pin_trace!(
            "pick dropped: {} is already the file in flight",
            pin_trace::path_name(path)
        );
        return;
    }

    send_preview(PreviewMessage::PinUpdate(path.to_path_buf()));
}

/// The file a pin is being shown next while it is still in flight — a load that has not answered, a
/// box nothing has measured yet, a swap held for its picture — if there is one (see
/// `update_pinned_preview`).
pub fn pin_file_in_flight() -> Option<PathBuf> {
    pin_install::pin_file_in_flight()
}

/// Note that the hand pressed the listing behind the pin, before the file the press picked is
/// known.
///
/// The click activates the listing, which is that window being put over the pin, and the shell's
/// answer for the file can be a tick or ten away — so the pin is put back in front of the listing
/// here, where the press is, rather than after the answer (see `Pin Mode → Update Preview → Move
/// To Top`). `window` is the frame the press was made in, carried for the reason
/// [`pin_press_was_not_a_pick`] carries it: the activation is still to land when this loop reads
/// the press, so the window the pin is to be answered against is the one the click landed in and
/// not the foreground — which is the window the user came from at that moment, and may be another
/// program or another listing by the time the raise is asked for (see `PinFrontOwed`). A press that
/// names no file this app previews — a folder, a file no kind previews, a gap between the rows — is
/// put back behind the listing by [`pin_press_was_not_a_pick`] instead, and a pick that lands
/// raises the pin again.
pub fn press_pinned_preview(window: isize) {
    if !pinned() {
        return;
    }

    send_preview(PreviewMessage::PinPress { window });
}

/// Note that the press above named no file, and put the pin back behind the window the press was
/// made in — `window`, the frame of the one under the point — which is what a press that picked a
/// folder, a file no kind previews, or a gap owes, and what undoes the raise the press began. The
/// frame rather than the foreground, because the click that made the press activates the listing
/// after this loop reads it: the foreground is still the window the user came from, and that is a
/// window the pin is already behind (see `PinUpdateWatch::press_frame`).
///
/// **A press that picked a file is never noted this way**, and that is the whole of what a mouse click
/// taught this road: the file may be named by either of the two lookups a press is answered out of
/// (see `the_press_named_no_file`), and a lowering sent for the first lookup alone landed *after* the
/// pick's own raise — the pin went behind the listing on the very click that picked the file, and
/// nothing raised it again. So the answer is carried to the end of the tick, where both lookups have
/// spoken.
pub fn pin_press_was_not_a_pick(window: isize) {
    if !pinned() {
        return;
    }

    send_preview(PreviewMessage::PinNoPick { window });
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

    // The window this hover puts up is the film its subtitle copy belongs to, so the copy of
    // any *other* film is dropped: the pass reads the whole film, and only the film being shown
    // is worth that read (see `subtitle_files`).
    subtitle_files::keep_extraction_for(Some(path));

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

    // And it is a hover like any other for the subtitle copy too: the film being shown is the
    // one whose copy is kept, and any other's is dropped (see `show_preview`).
    subtitle_files::keep_extraction_for(Some(path));

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

    // A hover that is going away takes the copy of its film's subtitles with it: the pass reads
    // the whole film, and a read outliving the window that asked for it is a read nobody is
    // waiting for (see `subtitle_files`). A pin is left alone above, because a pin that is up is
    // a window still showing its film.
    subtitle_files::keep_extraction_for(None);

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

    // A hover's film host goes down with it. This thread may not own that window — the loop's
    // `Hide` below takes it down there — so a take-down here is only ever the same thread's case,
    // and the `Hide` covers every other (see `hover_film_host`).
    unsafe { hover_film_host::take_down_the_hover_film_host() };

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
/// one thing waiting on it is a pinned window: a pin showing that film takes the copy as a
/// preload for its next relaunch, and what a player is handed belongs to the preview loop (see
/// `reload_pinned_subtitles`). A hover is told nothing — it was answered without subtitles by
/// the user's own choice, and the hover after it reads what this wrote (see
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
    if handle == 0 {
        return false;
    }

    // A picture pin stands in a window of its own, and a press on it is a press on the pin: the hook
    // asks this to tell a click on something that is not a listing from a click on this app's own
    // preview (see `click_is_over_a_listing`), and a window of the app's own that it did not know
    // would be read as the desktop behind it.
    if handle == PREVIEW_HWND.load(Ordering::Acquire) {
        return true;
    }

    pin_real_window::real_pin_window().is_some_and(|hwnd| hwnd.0 as isize == handle)
}

pub fn preview_screen_rect() -> Option<(i32, i32, i32, i32)> {
    if let Some(rect) = webview_preview::screen_rect() {
        return Some(rect);
    }

    unsafe {
        let candidates = [
            pin_real_window::real_pin_window()
                .map(|hwnd| hwnd.0 as isize)
                .unwrap_or(0),
            // A hosted hover's picture is the hover's own host: `PREVIEW_HWND` is hidden behind it
            // and a hosted player has no window of the desktop's at all, yet the host is the visible
            // preview surface, in the hover's box and where a pin takes it up (see `hover_film_host`).
            hover_film_host::hover_film_host()
                .map(|hwnd| hwnd.0 as isize)
                .unwrap_or(0),
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

/// The box of this app's own preview window — the pinned one where a pin is up — read from the
/// window itself rather than from whatever was last published for a surface on screen.
///
/// It is deliberately not `preview_screen_rect`'s question. A document the engine draws has a
/// window of its own standing *inside* the pin's, and that one is the surface on screen, while
/// what the Explorer hook watches for is this app's own window moving under a hand that has not
/// moved (see `PinUpdateWatch::note_pin_box`). The window is read whether or not it is visible,
/// because a pin that has collapsed into its bubble has hidden its window rather than moved it.
pub fn preview_window_rect() -> Option<(i32, i32, i32, i32)> {
    // A picture pin stands in a window of its own, and that is the window whose box the hook watches
    // for a pin moving under a hand that has not moved.
    if let Some(hwnd) = pin_real_window::real_pin_window() {
        if pin_is_its_own_window() {
            return window_rect_of(hwnd.0 as isize);
        }
    }

    let value = PREVIEW_HWND.load(Ordering::SeqCst);
    if value == 0 {
        return None;
    }
    window_rect_of(value)
}

/// The box of a window of this app's own, or nothing where it has no area.
fn window_rect_of(value: isize) -> Option<(i32, i32, i32, i32)> {
    unsafe {
        let mut rect = RECT::default();
        if GetWindowRect(HWND(value as *mut _), &mut rect).is_ok()
            && rect.right > rect.left
            && rect.bottom > rect.top
        {
            return Some((rect.left, rect.top, rect.right, rect.bottom));
        }
    }

    None
}

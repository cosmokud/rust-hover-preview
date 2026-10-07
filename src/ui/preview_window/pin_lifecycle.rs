//! A pin's own lifetime: the state it is installed and ended from, what ends it, and the
//! window it stands in beside the player's.

use super::*;

/// A pin was asked to come down, by the Explorer hook (previews were turned off, or the</path>
/// trigger key is holding them back), by the tray, or by the resumption of the machine
/// from sleep. What a pin *is* — a window and the media under it — belongs to the
/// preview loop, so this is a request rather than a take-down: the loop ends it on its
/// next tick, through the same path its own close button takes.
pub(super) static PIN_END_REQUESTED: AtomicBool = AtomicBool::new(false);

impl PinnedPreview {
    /// The box the window occupies: the media's own box with the caption above it and the transport
    /// bar below it — or the media's box itself, for a kind whose chrome is drawn over it, since a
    /// strip of chrome over a picture needs no room beside it.
    pub(super) fn window_box(&self) -> ScreenRegion {
        pinned_window_box_of(
            self.content,
            self.dpi,
            self.transport_bar,
            self.overlay,
            self.caption,
        )
    }

    /// The size of the window, as the renderer and the hit tests want it.
    pub(super) fn window_size(&self) -> (i32, i32) {
        let window = self.window_box();
        ((window.2 - window.0).max(1), (window.3 - window.1).max(1))
    }

    /// Which edge or corner of the window a point is on, if it is on one: what a resize is begun
    /// by, and what the shape of the pointer is decided from (see `pin_resize_edge`, which is the
    /// question itself, asked of this pin's own geometry).
    pub(super) fn resize_edge(&self, x: i32, y: i32) -> Option<PinResize> {
        pin_resize_edge(self.window_size(), self.dpi, x, y)
    }
}

/// Which edge or corner of a window of this size a point is on, if it is on one.
///
/// Four bands, one to a side, and the corners where two of them meet taken first so that a
/// point in one is answered with the corner rather than with whichever side happens to be
/// asked about first. A corner's band is the wider of the two: it is the smallest target the
/// window has, and the top-right one is under the caption's buttons — which is the same
/// bargain Windows makes on a window whose frame is drawn for it.
///
/// It is a question about a size and a scale rather than about a pin, because two callers ask
/// it: the pinned window's own procedure, of a press that landed on it (see `pinned_press`),
/// and the loop, of a press that landed on the engine's window standing in the pin's media band
/// (see `pinned_engine_press_action`). One geometry answers both, so a press is answered the
/// same way whichever window it happened to reach.
pub(super) fn pin_resize_edge(size: (i32, i32), dpi: u32, x: i32, y: i32) -> Option<PinResize> {
    let (width, height) = size;
    let border = logical_px(dpi, PIN_RESIZE_BORDER_PIXELS).max(2);
    let corner = border + logical_px(dpi, PIN_RESIZE_CORNER_EXTRA_PIXELS).max(1);

    // A window too small to tell its edges apart is not resized at all: every band would
    // overlap the others and a press would land on whichever was asked for first.
    if width <= corner + border || height <= corner + border {
        return None;
    }

    let mut edge = PinResize {
        left: x < border,
        top: y < border,
        right: x >= width - border,
        bottom: y >= height - border,
    };

    // A corner is the two bands that meet there, and its own is wider than either: a point
    // inside one is on both of the sides it is between, whether or not it is deep enough into
    // them to have been read as a side on its own.
    if y < corner {
        if x < corner {
            edge.left = true;
            edge.top = true;
        } else if x >= width - corner {
            edge.right = true;
            edge.top = true;
        }
    } else if y >= height - corner {
        if x < corner {
            edge.left = true;
            edge.bottom = true;
        } else if x >= width - corner {
            edge.right = true;
            edge.bottom = true;
        }
    }

    (edge.left || edge.top || edge.right || edge.bottom).then_some(edge)
}

/// A drag of a pinned window: where the pointer was when it began, the box the window had
/// then, and what the drag is doing with it.
#[derive(Clone, Copy)]
pub(super) struct PinDrag {
    pub(super) from: (i32, i32),
    pub(super) window: ScreenRegion,
    pub(super) action: PinDragAction,
    /// Whether the press that began this drag was delivered to this window as a message, or
    /// read from what the Explorer hook publishes about a press taken by the engine's window.
    ///
    /// It decides where the *end* of the drag is read from, and the distinction is the whole of
    /// it. A press this window was given brings its release with it, so a drag begun that way is
    /// ended by `pinned_release` and needs nothing else. A press read from the hook was taken by
    /// a window of somebody else's, and the release of a press somebody else took is that
    /// window's to answer: nothing guarantees it is ever delivered here, and a drag whose end is
    /// a message that never arrives is a drag that never ends (see `settle_pinned_engine_drag`).
    pub(super) delivered: bool,
    /// Where the pointer was when this drag was last carried out to, so that a drag carried on
    /// from the tick and a drag carried on from a pointer message are the same work asked twice
    /// rather than twice the work (see `apply_pin_drag`).
    pub(super) carried: (i32, i32),
}

#[derive(Clone, Copy)]
pub(super) enum PinDragAction {
    Move,
    Resize(PinResize),
}

/// Which edges of a pinned window a resize is dragging.
///
/// A side is one of them set and a corner is two, so the two questions a drag asks of it — which
/// way the box grows, and which way the hand is facing — are asked of the flags rather than kept
/// beside them.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) struct PinResize {
    pub(super) left: bool,
    pub(super) top: bool,
    pub(super) right: bool,
    pub(super) bottom: bool,
}

impl PinResize {
    /// Whether the drag moves the window's width, which a left or a right edge does.
    pub(super) fn horizontal(self) -> bool {
        self.left || self.right
    }

    /// Whether the drag moves its height, which a top or a bottom edge does.
    pub(super) fn vertical(self) -> bool {
        self.top || self.bottom
    }
}

/// A command the chrome has asked for: a button that was clicked, a window that was dragged or
/// resized, or a collapse. The window procedure cannot do any of it — the media, the player and
/// the browser are the preview loop's — so it is written here and drained by the loop.
///
/// Public because a key asks for the same two as a click does, and a key that walks a pin
/// arrives in the window procedure rather than being read on the hook thread (see
/// `pinned_key_command`).
///
/// One door, whichever side of the app the gesture came from.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum PinCommand {
    /// The file before the one pinned, in the order the folder it was taken up in is showing
    /// them. It goes through the loop like the rest, because the file it names is a walk of
    /// the folder and the loop is what owns the configuration the walk is read under (see
    /// `shell::pin_navigation`).
    Previous,
    /// The file after the one pinned, the same walk the other way.
    Next,
    Minimize,
    Maximize,
    Close,
    /// The bubble a collapsed pin left was clicked: the window goes back up.
    Restore,
    /// Hold what is on screen where it stands, or set it going again: what a Space in a pinned
    /// sound or a pinned video is, and which of the two it is a question the file in the window
    /// answers rather than the key (see `pin_toggle_target`).
    ///
    /// It is the only pause a sound's card has — the transport bar is a video's, and a card is
    /// drawn with no buttons on it — and for a video it is that bar's own pause, asked by the
    /// keyboard rather than by a press on the bar (see `toggle_pinned_audio` and
    /// `toggle_pinned_playback`).
    TogglePlayback,
    /// The next subtitle track, wrapping within the file's own.
    ///
    /// It is a relaunch and not a key sent to the player, and the reason is the same as it is for
    /// a seek: the choice has to be *remembered*, because every relaunch — this one, a seek, a
    /// resize, a change of level — begins a player again, and a player begun without being told
    /// which track to show falls back on its own default. So the number is written down first and
    /// the relaunch is told it, which is what makes the choice this one makes survive the next
    /// three (see `next_subtitle` and `PinTransport::subtitle`).
    ///
    /// Refused for every kind but a video FFmpeg plays, because it is the only player this app
    /// begins again on its own and the only one whose track this app therefore has to know.
    NextSubtitle,
}

// The commands the window procedure has left for the preview loop are a field of the pin's own
// state, beside the window they are about (see `pin_window::ask_pin` and
// `pin_window::take_pin_command`).

/// Whether the pin key is watched, as the configuration has it.
pub(super) fn pin_enabled() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_enabled)
        .unwrap_or(true)
}

/// Whether a pin that is up is shown the file the user picks next, as the configuration has it
/// (see `Pin Mode → Update Preview`). It is read on both sides of the swap: by the Explorer hook
/// before it watches for anything at all, and here before one is taken up.
pub(super) fn pin_update_enabled() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_update_enabled)
        .unwrap_or(DEFAULT_PIN_UPDATE_ENABLED)
}

/// Whether a pin collapsed into its bubble holds the video it is playing where it is, as the
/// configuration has it (see `Pin Mode → Pause Preview → Video`). It is read by the tick that
/// keeps a collapsed pin's playback in step with the bubble, so a switch thrown while the pin
/// is a bubble is answered on the next tick (see `settle_bubble_playback`).
pub(super) fn pin_pause_video() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_pause_video)
        .unwrap_or(DEFAULT_PIN_PAUSE_VIDEO)
}

/// And the same question about a sound, asked in the same place (see
/// `Pin Mode → Pause Preview → Audio`).
pub(super) fn pin_pause_audio() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_pause_audio)
        .unwrap_or(DEFAULT_PIN_PAUSE_AUDIO)
}

/// The pin a key press asks for, if there is anything on screen to pin: the file the preview
/// is of, and the box it occupies at this moment.
///
/// Nothing is pinned while a load is still running. What is on screen then is a spinner, and a
/// spinner is a promise rather than a preview — pinning one would leave a window standing over
/// a file that had not been read yet, and nothing in the pin's own machinery would ever take
/// it down.
pub(super) fn pin_what_is_on_screen(
    current_show: &Option<PreviewMessage>,
    pending: Option<&PendingLoad>,
) -> Option<PreviewMessage> {
    if pending.is_some() {
        return None;
    }

    let path = current_show.as_ref().and_then(show_path)?.clone();

    // A document, a specimen and a page of HTML are settled by the engine rather than by
    // the media, and the wait above is what stands in for the engine while it is not (see
    // `pin_screen_is_settled`).
    let engine_draws = engine_kind_of(&path).is_some()
        && webview_preview::showing_path().as_deref() == Some(path.as_path());
    if !pin_screen_is_settled(current_media_type(), engine_draws) {
        return None;
    }

    let rect = preview_screen_rect()?;
    Some(PreviewMessage::Pin { path, rect })
}

/// Whether what is on screen is a preview rather than a promise: either a frame this app has
/// finished making, or a page the engine has finished drawing.
///
/// The media answers it for everything this app draws, and cannot answer it for the three
/// kinds the engine draws: nothing of this app's goes on screen behind a document, a specimen
/// or a page of HTML, so what is left in the media slot is either nothing at all or the spinner
/// that stood in for the page until it landed — and the page's landing takes the wait without
/// touching the slot (see the loop's `webview_preview::showing_path`). Asking the media is
/// therefore what left the pin key doing nothing at all over every file the browser draws.
///
/// The engine says it by naming the file its window is showing, which is published only once a
/// document has actually been put up: a browser still coming up, a page that has not landed and
/// a navigation that failed all leave the name of something else, or of nothing (see
/// `webview_preview::showing_path`).
pub(super) fn pin_screen_is_settled(media: Option<MediaType>, engine_draws: bool) -> bool {
    if engine_draws {
        return true;
    }

    media.is_some_and(|kind| !kind.is_loading())
}

/// Take the pin down from the loop's own tick, answering with the message that does it.
///
/// The state goes first, and that is the whole of what this function is: what follows is the
/// ordinary take-down a hover's dismissal goes through — the window comes down, the player is
/// ended, the media goes — and it must not be refused by the pin's own guards. The bubble a
/// collapsed pin left goes with it, and so does the record of a pin having been up at all,
/// which is what lets the next hover through (see `pin_window::end_pin`).
///
/// `reason` is said rather than chosen as a road, because the road used to be a free choice and
/// the watchdog's copy of this list is what that cost: a pin it cleared kept the keyboard it had
/// claimed, kept a focusable window, and left a queued walk to be answered into a pin that no
/// longer existed. Every reason but `Reason::Hung` is this road, and the bubble is taken down
/// here because the loop's own take-down does not know about it.
pub(super) fn end_pin_state(reason: Reason) -> PreviewMessage {
    let owed = end_pin(reason, &Win32PinWindow);

    debug_assert_eq!(
        owed,
        PinHide::WithTheTakeDown,
        "the loop's own tick brings the window down as part of the message below"
    );

    PreviewMessage::Hide
}

/// The window a road out of a pin acts on: this app's own, and the only one there is.
///
/// The first of the two adapters. Every operation is a handle read from the slot its own
/// window was created into, or a call on the thread that window belongs to, which is why the
/// whole trait is `isize` rather than `HWND` (see `pin_window::PinWindow`).
#[derive(Clone, Copy)]
pub(super) struct Win32PinWindow;

impl PinWindow for Win32PinWindow {
    fn hwnd(&self) -> isize {
        PREVIEW_HWND.load(Ordering::SeqCst)
    }

    fn pointer(&self) -> Option<(i32, i32)> {
        cursor_screen_point()
    }

    fn window_box(&self, hwnd: isize) -> Option<ScreenRegion> {
        let (left, top, width, height) = window_origin(HWND(hwnd as *mut _))?;
        Some((left, top, left + width, top + height))
    }

    fn capture(&self, hwnd: isize) {
        // Safety: the handle is this app's own preview window, and a capture is held per thread —
        // only the thread the press was delivered on can take one, and every caller of this is a
        // message on this thread's own window procedure. The call is refused rather than acted on
        // for a window that has since gone.
        let _ = unsafe { SetCapture(HWND(hwnd as *mut _)) };
    }

    fn release_capture(&self, hwnd: isize) {
        // Safety: the handle is read from the slot this app's own preview window was created
        // into and is only asked to give up a capture it may hold, which it checks before
        // releasing. A window that has since gone is refused by the call rather than acted on.
        unsafe { release_pin_capture(HWND(hwnd as *mut _)) }
    }

    fn set_focusable(&self, hwnd: isize, focusable: bool) {
        // Safety: the handle is this app's own preview window, and only its extended style is
        // touched. A window that has since gone is refused by the calls rather than acted on.
        unsafe { pin_set_focusable(HWND(hwnd as *mut _), focusable) }
    }

    fn set_focus(&self, hwnd: isize) {
        // Safety: asks the keyboard for a window, and names this thread's own. Neither is a
        // dereference and neither can fail.
        let _ = unsafe { SetFocus(HWND(hwnd as *mut _)) };
    }

    fn set_foreground(&self, hwnd: isize) {
        // Safety: names a window to bring to the front. The handle is either this app's own
        // preview window or the one that was in front a moment ago, which by then is a live
        // window or a handle to one that has since gone — in which case the call is refused
        // and nothing else is disturbed.
        let _ = unsafe { SetForegroundWindow(HWND(hwnd as *mut _)) };
    }

    fn hide_pin_windows(&self) {
        // Safety: each handle is read from the slot its own window was created into, and each
        // is only asked to hide. `hide_pin_bubble` does the same for the bubble, and a handle
        // to a window that has since gone is refused by the call rather than acted on.
        unsafe {
            hide_pinned_windows();
            hide_pin_bubble();
        }
    }

    fn hide_pin_bubble(&self) {
        // The bubble is a window of this app's own and hiding it is a plain call on a handle
        // read from the slot it was created into, with nothing this fn has to be unsafe about.
        hide_pin_bubble();
    }

    fn repaint(&self) {
        // Safety: the handle is read from the slot this app's own preview window was created
        // into, and the paint is a layered-window blit of a surface this app drew for exactly
        // this window. A window that has since gone is refused by the call rather than acted on.
        let hwnd = HWND(self.hwnd() as *mut _);
        if !hwnd.is_invalid() {
            unsafe { render_layered_preview(hwnd) };
        }
    }

    fn unpark_player_window(&self, band: Option<ScreenRegion>) {
        // Nothing here is unsafe: the player's window is found by the process this app started it
        // for rather than held as a handle, so there is no dereference to justify and a window
        // that has since gone is found not to be there.
        match band {
            // The place is the show — the window goes up in the same `SetWindowPos` that moves
            // it, so a band answered here is answered on screen at once. Placed directly rather
            // than through `ensure_pinned_sibling_box`: that call is refused while a park stands,
            // and the swap places the window BEFORE taking the flag down, so the cover is still up
            // when the window goes back (see `unpark_pinned_player`).
            Some(band) => {
                let _ = ensure_video_window_topmost(
                    band.0,
                    band.1,
                    (band.2 - band.0).max(1),
                    (band.3 - band.1).max(1),
                );
            }
            // A pin with no band to go back into has no rect to be told, so only its hiddenness is
            // taken back (see `unpark_pinned_player`).
            None => show_pinned_player_window(),
        }
    }

    fn post(&self, hwnd: isize, message: u32) {
        // Safety: the handle is read from the slot the window was created into, and the message
        // asks only for something this window does to itself. A window that has since gone is
        // refused by the call rather than acted on.
        let _ = unsafe { PostMessageW(HWND(hwnd as *mut _), message, WPARAM(0), LPARAM(0)) };
    }
}

/// The rest of what a pin's end has to settle, which is not the pin.
///
/// The pin's own window, its keyboard claim and the commands its chrome had queued are fields
/// of its state, and `pin_window::end_pin` settles those with the window work in one list. The
/// three things here are each somebody else's: the walk the planner is working on, the bubble's
/// drag latch and the box a drag had left the window at. They stay where they are because their
/// owners are elsewhere — the planner holds its lock across a `Condvar::wait_timeout`, and the
/// bubble and the box are windows of their own — but they are settled from the same one exit, so
/// a fourth road out of a pin cannot leave them standing.
///
/// It is a separate function and not part of the state for the same reason it is separate from
/// the teardown's own list: it holds three locks in turn, none of them the pin's, and a reader
/// that wanted one lock for a pin's end would be holding the planner's.
pub(super) fn end_pin_beside_the_state() {
    // A walk is the same answer to the same question a queued command is — a caption button
    // asking for the next file, worked out for a pin that has gone.
    if let Ok(mut jobs) = PIN_JOBS.0.lock() {
        *jobs = None;
    }

    // A drag on the bubble is a drag on a window that is going away, and the latch that says
    // one is in progress is what tells a click on it from a drag; left standing, the next
    // bubble treats a press as the end of a drag that began before it existed.
    PIN_BUBBLE_MOVED.store(false, Ordering::Release);
    if let Ok(mut drag) = PIN_BUBBLE_DRAG.lock() {
        *drag = None;
    }

    // A box a drag has left a pinned window at is a layout for a window that is going away, and
    // a take-down that kept it laid the next pin out into a box the user dragged a window they
    // can no longer see into.
    if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
        *request = None;
    }

    // And the frame a drag of this pin held for its media band. A pin taken down with a drag still
    // in flight — a pick from the bubble under the hand, the watchdog — never reaches the other end
    // of its park, and what it was holding is a picture the size of a display (see
    // `compose_parked_band`).
    forget_video_frame();

    // And the gesture that killed with its end still to come. A pin taken
    // down mid-dead-band never reaches its relaunch, and the snapshot left
    // behind is a resume for a pin that has gone: no end will ever take it,
    // and the next gesture would relaunch at a stale second. Given up with
    // the pin — the player is already dead or riding the orphan reaper, so
    // there is nothing here left to kill.
    take_gesture_snapshot();

    // And the relaunch that never landed: a pin taken down with one in flight behind its cover —
    // a walk stepping off, the watchdog — leaves a player nothing will ever show. It dies
    // unpublished rather than lingering, and reaped rather than orphaned. Only a confirmed kill
    // forgets the record; a kill not yet taken stays pending — and the pin is going away, so no
    // settle will ever retry it. That half is handed to the orphan reaper that outlives the pin,
    // which asks again on every tick until the death confirms: never dropped, never waited on
    // (see `retire_orphaned_player`).
    if pending_pinned_relaunch().is_some() {
        let _ = reap_superseded_relaunch();
        if let Some(still) = pending_pinned_relaunch() {
            retire_orphaned_player(still.pid);
            clear_pending_pinned_relaunch();
        }
    }

    // And the cover a seek of this pin was aiming under. A pin taken down
    // mid-aim — a walk stepping off, the watchdog — never reaches its swap,
    // and the record left behind is a cover over a pin that has gone: no park
    // flag to end it, and the next park overwrites it rather than extends it,
    // but a settle asked before that would answer a dead player. Given up
    // with the frame it was holding.
    forget_pin_park_swap();
}

/// Whether the thing a pin is a window onto is still there.
///
/// Two kinds of media can go away on their own, because something outside this thread is
/// drawing them: a video FFmpeg's player is playing, and a document or a specimen the browser
/// is drawing. Each is a process this app started and does not own, and a pin whose process
/// has died is a window onto nothing — so it comes down and previews resume, which is the half
/// of "until it is closed, or it comes apart" that is not a button.
///
/// A sound is deliberately not one of them, whichever player is behind it. A card is text this
/// app draws, and its player plays the pass it was given and stops, so a player that is not
/// running is the moment between two passes rather than a pin onto nothing — asking about it took
/// the window down at the end of every pass (see `wrap_audio_player`). A card whose player never
/// came back at all is a card with a still clock, which is what a machine with no output device
/// gives, and the window stays up over it.
///
/// And the question is whether the thing is *there*, not whether it has drawn anything yet: a
/// browser still coming up, and a document it has been asked for that has not landed, are a pin
/// with something to be a window onto (see `webview_preview::is_behind`).
///
/// What `navigating` is for is the case where the player is not running *because* this app is
/// asking for a different file, which is not a pin that came apart at all. A swap empties the
/// slot this is read out of before it knows whether the file it is swapping in will play (see
/// `swap_pinned_media`), so a refused swap leaves a pin with nothing behind it — and read as a
/// pin that came apart, that took the window down a tick before the walk could step onto the next
/// file. That is why a corrupted film closed a pinned window while a corrupted picture did not: a
/// picture is refused by the plan, which empties nothing (see `pin_step_off`).
pub(super) fn pin_media_is_alive(navigating: bool) -> bool {
    // A pin that is already being shown another file is a pin with something to be a window onto:
    // whether the player behind the file it is showing right now has gone is answered by what is
    // being shown instead, and not by this.
    if navigating {
        return true;
    }

    // What kind it is, is taken in one look and the lock is let go of before anything is asked
    // *about* the answer. What a player's liveness is asked through — `is_video_process_running`
    // — reads the media itself, so asking it with the media already locked by this thread is
    // asking for a lock this thread owns, which is not a wait but a stop: the preview loop would
    // never draw another frame, and the pinned window would stand there answering nothing for the
    // rest of the run.
    let kind = {
        let Ok(media) = CURRENT_MEDIA.lock() else {
            return true;
        };
        let Some(media) = media.as_ref() else {
            // Nothing is on screen: a pin with no media behind it has already come apart.
            return false;
        };

        media.media_type
    };

    // What the bubble parked is not something that came apart: a player this app ended, or an
    // engine it paused, is the pin's media still — what the pin is a window onto has not gone
    // anywhere, and the restore that follows is what puts it back (see `BubblePause`). Asked
    // before the kinds below, one of which would answer "gone" about a player that is
    // deliberately not running.
    if pin_bubble_pause().is_some() {
        return true;
    }

    match kind {
        MediaType::Video => {
            // A gesture that killed leaves the player gone with its end
            // relaunch still to come: the cover stands over a dead band, and
            // that is not a pin that came apart.
            if gesture_snapshot_active() {
                return true;
            }
            // A supersede-kill leaves the player gone with its replacement on its way — the
            // cover standing over a relaunch still to come, or one still in flight. That is not
            // a pin that came apart: the release relaunches, the in-flight one lands, and the
            // cover holds until one of them does.
            if pin_park_covers_a_relaunch() {
                return true;
            }
            is_video_process_running()
        }
        // A card is this app's own text and a player this app started, and the player is
        // between passes rather than gone — the whole of what is asked about here is above.
        MediaType::Audio => true,
        MediaType::NativeVideo => video_player::is_playing(),
        // A document or a specimen is the browser's, and what says it is still there is the
        // engine standing behind the file rather than a window with pixels in it: a window is
        // what a browser being *started* has none of yet, and reading that as "gone" closed a
        // pin the moment it was shown the first document of a run — the browser has to come up
        // before it can be checked for, so the check has to know what coming up looks like (see
        // `webview_preview::is_behind`). Which file that is, is the pin's to say rather than the
        // media's, and it is read here because the media's own lock has been let go of by now:
        // the two are taken one after the other, never across each other.
        MediaType::EngineSvg | MediaType::EngineFont => {
            pinned_media_owner().is_some_and(|(path, _)| webview_preview::is_behind(&path))
        }
        // A frame this app holds is a frame nothing outside this thread can take away, and the
        // mark a failed file stands as is exactly that: a window with the cross in it is a window
        // with a file's own shape in it, and reading it as a window onto nothing is what closed
        // the pin of a file put straight onto a corrupted one. The arm is written out rather than
        // left to the catch-all below because it is load-bearing — the catch-all happens to answer
        // the same way today, and nothing would say so if it stopped (see `show_pin_failure`).
        MediaType::Unplayable => true,
        // A frame this app holds is a frame nothing outside this thread can take away.
        _ => true,
    }
}

/// How long a player behind a file a pin has just been started is given to draw something of it
/// before a player that is not there is read as one that died rather than as a pin that came
/// apart.
///
/// It is a give-up rather than a wait anything is expected to reach, and it is well under
/// `VIDEO_START_WAIT_SECS` on purpose: that is how long a start may run while a *window* is on its
/// way up, where a player that is gone has plainly died and only the window has not appeared yet.
/// What is asked about here is the other way round — a player that is not there — and FFmpeg's
/// player opens a file it cannot read, prints its complaint and exits well inside a second, so a
/// third of that is a hundred times over the answer's own arrival. Past the give-up a player that
/// is gone is a film watched to its end or a window the user closed, and neither is a file to step
/// over.
pub(super) const PIN_PLAYER_GIVE_UP: Duration = Duration::from_secs(3);

/// The file a pin is showing whose player died without ever drawing anything of it, if what is on
/// screen now is such a file.
///
/// It is the question `pin_media_is_alive` asks from the other side: that one is whether what the
/// pin is a window onto is still there, and this is whether what was there ever arrived. The two
/// players fail in different ways and each is asked in its own terms. The engine takes a file it
/// cannot draw and then hands over no frame of it — or reports an error and hands over nothing at
/// all — which `failing_before_a_frame` is the whole of, and which already excludes a file that
/// played and died afterwards. FFmpeg's player leaves no such record: it is either running or it
/// is not, and all that says a file was refused rather than watched is that it is not running and
/// has not been running long.
///
/// Which file it is is the pin's to say rather than the media's, for the reason written down where
/// the kind is taken: the media's own lock is let go of before anything is asked about the player,
/// because what a player's liveness is asked through reads the media. `player_started` is when
/// the player behind what is on screen now was started, read here rather than asked of the player
/// because a player that does not exist cannot say when it began.
pub(super) fn pin_media_failed_before_a_frame(player_started: Option<Instant>) -> Option<PathBuf> {
    // The kind is taken in one look and the lock is let go of with the statement, before anything
    // is asked *about* the answer — the reason `pin_media_is_alive` writes down at length.
    let kind = CURRENT_MEDIA
        .lock()
        .ok()?
        .as_ref()
        .map(|media| media.media_type)?;

    match kind {
        // The file the engine is failing at, and it is this file: a pin is a window onto the file
        // it is showing and not onto whatever session happens to be running.
        MediaType::NativeVideo => video_player::failing_before_a_frame()
            .filter(|path| pinned_path().as_deref() == Some(path.as_path())),
        MediaType::Video => {
            // A player ended for a gesture or a park is not a file that failed: it was killed on
            // purpose and its replacement is on its way, which is exactly how `pin_media_is_alive`
            // reads the same two facts. Without this the tick after a kill road's own player is
            // read as a dead file — inside the give-up window — and the pin falls to the failure
            // mark before the relaunch that road owes can arrive (the "window breaks permanently"
            // after a Next: a film's player is killed and the pin is stuck on the cross).
            if gesture_snapshot_active() || pin_park_covers_a_relaunch() {
                return None;
            }

            if is_video_process_running() {
                return None;
            }

            player_started
                .filter(|started| started.elapsed() < PIN_PLAYER_GIVE_UP)
                .and_then(|_| pinned_path())
        }
        // Everything else a pin can be showing is either this app's own drawing or something with
        // a wait of its own on being up, and a player that refused the file is not what those
        // come apart over.
        _ => None,
    }
}

/// The clock a sound's card is drawn from, which belongs to the preview loop and not to the
/// media: when the player this app started was started, the second of the file it was started
/// at, where a key has held it, how far a name the card has no room for has been scrolled, and
/// what the card's own controls are saying (see `audio_clock`).
#[derive(Clone, Copy)]
pub(super) struct AudioCardClock {
    pub(super) started: Option<Instant>,
    pub(super) from: f64,
    /// The second of the file a sound is being held at, where a key in the window has paused
    /// one: a player of this app's was ended to pause it, and this is what the card is drawn
    /// at in its place.
    pub(super) paused: Option<f64>,
    pub(super) name_offset: i32,
    pub(super) dpi: u32,
    /// What the card's own row of buttons is saying, read in one look before the media's own lock
    /// is taken rather than from under it (see `pinned_audio_chrome`).
    pub(super) chrome: Option<CardChrome>,
}

/// The file and the display scale a pin is showing, for the work that has to lay its media out
/// again — read in one look so that the lock is not held across it.
pub(super) fn pinned_media_owner() -> Option<(PathBuf, u32)> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    Some((pin.path.clone(), pin.dpi))
}

/// The box a pinned window is standing at, if one is up.
pub(super) fn pinned_window_box() -> Option<(ScreenRegion, i32, i32)> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    let window = pin.window_box();
    Some((
        window,
        (window.2 - window.0).max(1),
        (window.3 - window.1).max(1),
    ))
}

/// The pin's own window, for a road that has no message in hand to name it.
///
/// The handle the window was created into, which is the same one
/// [`PinWindow::hwnd`] reads and the same one `raise_pinned_window` is asked
/// with from the loop's own tick. It is here for the two covers that run with no
/// pointer message of their own — a box change and a file step are both the
/// loop's work, and both owe a paint before a player is hidden — so the handle
/// they paint through is read rather than carried down from a window procedure
/// that is not on either road.
pub(super) fn pinned_window() -> HWND {
    HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _)
}

/// Put a pinned window up at the box its state says, painted before it is shown: what a layered
/// window shows between one paint and the next is the surface it already has, stretched into
/// whatever box the window has (see `show_loading_spinner`).
pub(super) unsafe fn show_pinned_window(hwnd: HWND) {
    let Some((window, width, height)) = pinned_window_box() else {
        return;
    };

    if !IsWindowVisible(hwnd).as_bool() {
        let _ = MoveWindow(hwnd, window.0, window.1, width, height, false);
    }

    render_pinned_preview_at(hwnd, window.0, window.1);
    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        window.0,
        window.1,
        width,
        height,
        SWP_NOACTIVATE | SWP_SHOWWINDOW,
    );
    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
}

/// Put the thing a pin leaves a band open for where that band is: the player's own window, or
/// the browser's, each of which is a window standing in the hole the pinned window leaves — so
/// that the picture is theirs and the caption is this app's (see `render_pinned_preview_at`).
///
/// It is asked where a box changes — a pin taken up, a window maximized, restored, dragged or
/// resized — and not once a tick: the tick has its own, slower re-assertion for the player's
/// window, and a browser that is told the same bounds it already has every sixteen milliseconds
/// is work nobody asked for.
pub(super) fn place_pinned_siblings() {
    let Some(content) = pinned_content() else {
        return;
    };

    let kind = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.media_type));

    match kind {
        Some(MediaType::Video) => ensure_pinned_sibling_box(content),
        // A document is the same kind of thing — a window of somebody else's standing in the pin's
        // band — and it is *moved* rather than told again where it is, because a browser holding a
        // page has no reason to put it anywhere else by itself (see `webview_preview::place`).
        Some(MediaType::EngineSvg) | Some(MediaType::EngineFont) => {
            if let Some((path, _)) = pinned_media_owner() {
                webview_preview::place(
                    &path,
                    webview_preview::Area {
                        x: content.0,
                        y: content.1,
                        width: (content.2 - content.0).max(1),
                        height: (content.3 - content.1).max(1),
                    },
                    engine_background(&path),
                );
            }
        }
        _ => {}
    }
}

/// The media box of the pin that is up, or nothing when there is no pin or it is collapsed: a
/// collapsed pin has no bands for anybody else's window to stand in.
pub(super) fn pinned_content() -> Option<ScreenRegion> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    (!pin.collapsed).then_some(pin.content)
}

/// The card's own controls as they stand, read in one look and handed over rather than asked for
/// under the media's own lock — see `MediaData::refresh_audio_card`.
///
/// It is a function rather than a read of the pin inside that paint because both the paint and the
/// relayout are reached with `CURRENT_MEDIA` already held, and a pin's lock asked from under it is
/// a lock this thread may already own: `toggle_pinned_playback` reaches `CURRENT_MEDIA.lock()`
/// from paths that have been near the pin's, so the two are taken in this order everywhere and
/// never the other way round.
pub(super) fn pinned_audio_chrome(playing: bool) -> Option<CardChrome> {
    // The published flag rather than the state: this is asked on every tick that does a full
    // media look, and a run with nothing pinned pays one atomic read for it.
    if !pinned() {
        return None;
    }

    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    // A collapsed pin has no card on screen, and a card with no window to press is a card with no
    // controls — the same answer `audio_preview::Card::controls` gives a hover.
    if pin.collapsed {
        return None;
    }

    Some(CardChrome {
        playing,
        volume: pin.volume.level,
        hovered: pin.audio_hovered,
        pressed: pin.audio_pressed,
        window_buttons: pin.audio_window_buttons,
    })
}

/// Whether this pin is showing a sound's card, which is the one kind whose card carries its own
/// controls — and the answer every question about one of them is gated on.
///
/// It is asked of the pin's own three facts rather than of the media's kind, and that is the whole
/// of why it is a function at all: a window procedure that had to take the media's lock to know
/// what it was showing would be one that could be asked while that lock is held (see
/// `MediaData::refresh_audio_card`), and the three facts are all settled at the take-up from the
/// kind anyway. A sound is the only kind with no transport strip, no overlay chrome and no frame
/// of its own to be resized within.
pub(super) fn pin_shows_an_audio_card(pin: &PinnedPreview) -> bool {
    !pin.collapsed && !pin.transport_bar && !pin.overlay && pin.frame == PinFrame::None
}

/// The box a pin of a sound is given, which is the box its hover was.
///
/// A pin's card carries its controls where a hover's does not, but it is not a different card:
/// the buttons stand in the bar's own row and are carved out of the bar's width, so the card a pin
/// measures is the card a hover measured — at the audio room, the same room the hover's own
/// card is measured at (see `audio_box_room`) — and the window is the one the hover put up (see
/// `audio_preview::bar_row`). The card is measured again all the same, with the controls it is
/// going to carry on it and a clock at nothing — which button is lit is not part of the layout —
/// and the hover's own top left corner is kept: where a window is put is the hover's place, and
/// the clamp the take-up runs afterwards pulls a box the card would not fit back onto the display
/// (see `pinned_caption_height` for the kind that has no caption above this one at all).
pub(super) fn pinned_audio_card_box(rect: ScreenRegion, path: &Path, dpi: u32) -> ScreenRegion {
    let chrome = Some(CardChrome {
        playing: false,
        volume: current_audio_volume(),
        hovered: None,
        pressed: None,
        window_buttons: false,
    });
    let Some(card) = audio_card(path, None, None, 0, chrome) else {
        return rect;
    };

    let room = audio_box_room(work_area_at(rect.0, rect.1), current_audio_scale(), dpi);
    let Some((width, height)) = audio_preview::measure(
        &card,
        room.0.max(1),
        room.1.max(1),
        dpi,
        current_audio_options(),
    ) else {
        return rect;
    };

    (
        rect.0,
        rect.1,
        rect.0 + width as i32,
        rect.1 + height as i32,
    )
}

/// The card's own control the pointer is over, asked of the card's own layout and answered in the
/// window's own coordinates: the card fills the pin's media box and is drawn at the media band's
/// own row, so the point is taken to the card the way a press on it is (see `media_point` and
/// `pinned_band_rows`).
pub(super) fn pin_audio_control_at(pin: &PinnedPreview, x: i32, y: i32) -> Option<CardControl> {
    if !pin_shows_an_audio_card(pin) {
        return None;
    }

    let (_, height) = pin.window_size();
    let (top, _) = pinned_band_rows(
        height,
        pin.caption,
        pinned_transport_height(pin.dpi, pin.transport_bar),
        pin.overlay,
    );

    audio_preview::control_at(
        x,
        y - top,
        (pin.content.2 - pin.content.0).max(1) as u32,
        pin.dpi,
        current_audio_options(),
        true,
    )
}

/// Whether the card's two window buttons are showing at `y`, the pointer's row in
/// the window's own coordinates: the top band as tall as the buttons' own hit boxes,
/// which is the one band of the window a hand in is a hand near the window's top
/// border or near the buttons themselves — the one thing on a sound's card that
/// comes and goes, because a button twice the margin's own side stands in the name's
/// own rows (see `audio_preview::window_button_band`).
pub(super) fn pin_audio_window_buttons(pin: &PinnedPreview, y: i32) -> bool {
    y < audio_preview::window_button_band(pin.dpi, current_audio_options())
}

/// Put a window of somebody else's — the player's — where a pinned window's media band is.
///
/// **A band whose player is parked is left alone, which is what this is for rather than an
/// afterthought of it.** The park is a hide on the drag's first pointer message, and a raise here
/// is a show as well as a place (see `ensure_video_window_topmost`): a resize drag puts the player
/// here on every pointer move, and a maximized or restored window, a relayout and a take-up each put
/// it here once more, so the film would be back on screen for as long as the drag lasted. What the
/// park holds is the window's *absence*, and there is nothing about being out of date to say the
/// absence can be ended (see `pin_player_is_parked`).
pub(super) fn ensure_pinned_sibling_box(content: ScreenRegion) {
    if pin_player_is_parked() {
        return;
    }

    let _ = ensure_video_window_topmost(
        content.0,
        content.1,
        (content.2 - content.0).max(1),
        (content.3 - content.1).max(1),
    );
}

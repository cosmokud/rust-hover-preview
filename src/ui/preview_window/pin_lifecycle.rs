//! A pin's own lifetime: the state it is installed and ended from, what ends it, and the
//! window it stands in beside the player's.

use super::*;

/// A pin was asked to come down, by the Explorer hook (previews were turned off, or the</path>
/// trigger key is holding them back), by the tray, or by the resumption of the machine
/// from sleep. What a pin *is* — a window and the media under it — belongs to the
/// preview loop, so this is a request rather than a take-down: the loop ends it on its
/// next tick, through the same path its own close button takes.
pub(super) static PIN_END_REQUESTED: AtomicBool = AtomicBool::new(false);

/// The card's menu was asked to come down, by the Explorer hook, because a press landed on a
/// window that is not the pin's own. The pin window is `WS_EX_NOACTIVATE` and holds no
/// capture, so a press on another window never reaches its procedure and the hook's poll is
/// the only side that sees it (see `requests::request_pin_menu_dismiss`). Only the menu goes:
/// the pin itself is left standing, and the loop takes the ask on its next tick.
pub(super) static PIN_MENU_DISMISS_REQUESTED: AtomicBool = AtomicBool::new(false);

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
    /// The picture inside the window moving rather than the window: what a press on the media of a
    /// picture the wheel has taken inside its window begins, and where the pan stood when it began —
    /// which is what the pointer's own travel is added to, so a drag is measured from where it began
    /// rather than accumulated move by move (see `PinZoom` and `apply_pin_drag`).
    Pan {
        from: (f32, f32),
    },
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

/// A command the chrome has asked for: a button that was clicked, a window that was dragged,
/// resized or double-clicked, or a collapse. The window procedure cannot do any of it — the media,
/// the player and the browser are the preview loop's — so it is written here and drained by the
/// loop.
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
    /// Collapse the pin into a place of the asker's own choosing, rather than into the one
    /// `Pin Mode → Minimize to` has it put: what the card's menu's `To Bubble` and
    /// `To Taskbar` rows ask for, which say where the pin goes rather than where the next
    /// minimize would take it (see `collapse_pin_to`).
    MinimizeTo(PinMinimizeTo),
    Maximize,
    /// Put the window back to the size the scaling settings ask for the kind it is showing, in the
    /// middle of the screen it stands on: what the second press of a double-click asks for where
    /// the configuration has the size reset on. It is a box rather than a place — a media has to be
    /// laid out at its size rather than carried to it — so it goes through the loop like a maximize
    /// does (see `reset_pinned_size`).
    ResetSize,
    Close,
    /// The bubble a collapsed pin left was clicked: the window goes back up.
    Restore,
    /// The Shell brought a pin's window back on its own: a pin put away as a window is minimized rather
    /// than hidden, and the button that restores it belongs to the window rather than to this app, so no
    /// press of the app's arrives to ask for the pin (see `PinRestore::TheShell`). The state follows the
    /// window — what a collapse parks is given back, and the pin stops answering as one that is away.
    CameBack,
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
    /// Take a pinned page of text back a screenful — and, where the window holds no page, the
    /// previous file, which is what an up arrow has always meant to every other kind of pin.
    ///
    /// Which of the two it is, is the file in the window's answer rather than the key's, the way a
    /// Space is a hold of a video, a sound, or nothing at all (see `pin_command_request`).
    PageBack,
    /// The same, a screenful on, and the next file where there is no page to move.
    PageOn,
    /// The beginning of the page in the window, for a window that holds one, and nothing at all for
    /// one that does not: a picture has no beginning to go back to.
    PageTop,
    /// The last screenful of it, the same way.
    PageBottom,
    /// Take a pinned page to the next place its own selection appears, and select it: what a
    /// reader does with a word they are looking for again (see `find_next`).
    ///
    /// Nothing at all to a pin that holds no page, or that holds one with nothing selected.
    FindNext,
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

/// Whether a pinned window stays always on top of other windows, as the configuration has it.
pub fn pin_always_on_top() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_always_on_top)
        .unwrap_or(DEFAULT_PIN_ALWAYS_ON_TOP)
}

/// Whether the pin's window belongs in the topmost band: the band `Always On Top` asks for, and
/// no other.
///
/// **There is no other band for it to stand in, and nothing to remember.** A pin that is not
/// Always On Top is an ordinary window: Windows raises the window a click activated, and that
/// window is then in front of the pin, which is the whole of what "not on top" means on this
/// desktop. This app used to keep a record of a pin that "owed the front" so a preview it had just
/// updated stayed over the listing, and every release rule that record needed was one more way to
/// get it wrong — a pin still holding the topmost band after a resize, a restore, or `Win+D` is a
/// window over every program the user works in, with the setting that says it must not be.
///
/// The band is the setting's. What this app asks for on the user's own press is a raise within the
/// ordinary band, and it asks for it one press at a time (see `put_pin_at_the_front`).
pub(crate) fn pin_holds_the_front() -> bool {
    pinned() && pin_always_on_top()
}

/// The band a window of this app's is placed in when it is placed for reasons of its own — a
/// spinner drawn into it, a box re-asserted under a hand, a window put back on the screen: the
/// topmost one for a hover, and for a pin that is Always On Top, holds a band a pick's press took, or
/// owes the front; the ordinary one for a pin that owes neither.
///
/// **Every placement is this question rather than a band of its own.** Placed any other way, a pin
/// with `Always On Top` off was put in the topmost band by a resize or a spinner and stayed there:
/// the band's own flag was never set by those roads, so nothing that gives the band back knew to
/// look, and the window stood over every program the user worked in until a press happened to raise
/// it — the one road that takes the band off a window (see `pin_holds_the_front` and
/// `lower_pinned_preview`). The band is a decision, and a window is placed in the one the decision
/// names.
///
/// **A held band is asked here too, and it is the road that mattered most.** A pick is installed in the
/// pin's window between the press and the click, and that install places the window: answered from the
/// setting alone, every pick put its own window back in the ordinary band a few milliseconds before the
/// click was activated over it, which is the band given back ahead of its clock (see `PIN_BAND_HELD`).
pub(crate) fn band_place() -> HWND {
    if pin_holds_the_front() || !pinned() || PIN_BAND_HELD.load(Ordering::SeqCst) {
        HWND_TOPMOST
    } else {
        HWND_NOTOPMOST
    }
}

/// Whether a pin being put up takes the topmost band for the moment it appears, and gives it back.
///
/// A pin that is a window of its own is put up over everything, whatever the setting says: a window
/// that appears behind the listing it was asked for from is a window the user does not see, and the
/// press that asked for it (SPACE) is one the foreground window got, so `HWND_TOP` would do nothing
/// (see `put_pin_at_the_front`). The band is held for two calls and given back, so nothing of the
/// setting is taken away from it: a window put up this way is one the next click puts behind.
///
/// Asked only of a pin that is *not* already on the screen — a repaint, a drag and a resize are not
/// appearances, and putting the pin in front on every one of them would be a window that could not be
/// put behind anything.
pub(super) fn showing_takes_the_band_to_appear(on_the_screen: bool) -> bool {
    pin_is_its_own_window() && !pin_always_on_top() && !on_the_screen
}

/// Put a window over every other one, and give the band straight back: what a window of this app's own
/// is owed when it has to come forward over a window that holds the foreground.
///
/// **`HWND_TOP` is no way to the front for a process that is not in the foreground.** Windows refuses
/// the place to it — the call answers success and the window stays exactly where it was — so the
/// topmost band is the only road there is, and `HWND_NOTOPMOST` is what puts the window over every
/// ordinary one in one call: the front of exactly the band a pin that is not
/// `Pin Mode → Always On Top` belongs in, reached in a way the foreground window cannot refuse.
/// Nothing of the band is left held, so the next click on another window puts the pin behind it,
/// exactly as any window behaves (see `pin_holds_the_front`).
unsafe fn take_the_band_and_give_it_back(hwnd: HWND) {
    // Nobody holds a band this road gives back, whatever it was held for: the window is in the ordinary
    // band from the second call on, and a flag left saying otherwise would have the next placement put it
    // back over every program for a hold that is over (see `PIN_BAND_HELD`).
    PIN_BAND_HELD.store(false, Ordering::SeqCst);
    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
    );
    let _ = SetWindowPos(
        hwnd,
        HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
    );
}

/// The pin owes the front: a pick, a take-up or a restore asked for it, and no press that named no
/// file and no window change away from the listing has given it back since.
///
/// It is a flag rather than a question asked of the window because the band is asked for *before*
/// the pick is read (see `arm_pin_to_the_front`), and the window's band at that moment says what the
/// last click left it in rather than what this one is owed.
pub(super) static PIN_FRONT_OWED: AtomicBool = AtomicBool::new(false);

/// The topmost band is being held for a pick's press: taken with the press, and given back when the
/// grace the press began has run out (see `raise_pinned_preview` and `give_the_band_back`).
///
/// **It is a flag rather than a look at the window for the same reason `PIN_FRONT_OWED` is one**: the
/// band outlives the call that took it by several milliseconds, and every placement of that window in
/// between has to know that the band is not the setting's to give back yet. A pick is installed in that
/// window — the file is read, decoded and shown — and the show places the window, so a placement that
/// answered from the setting alone put the pin back in the ordinary band *before* the click that picked
/// the file had been activated over it: Explorer flashed over the pin on every click, and on none of
/// the arrows, which walk the same files without a press to hold anything (see `band_place`).
pub(super) static PIN_BAND_HELD: AtomicBool = AtomicBool::new(false);

/// Say that a pin that is up owes the front: what a take-up, a pick and a restore each ask for.
///
/// A pin that is Always On Top owes nothing of its own — its band is the setting's — and neither
/// does one whose `Move To Top` row is off, because that row is the whole of what says an update
/// puts the pin in front of what it was behind. A collapsed pin has no window in any band until it
/// is restored, and the restore is one of the doors that asks again — except where the update is
/// itself what restores it, which is the same row's other answer (see
/// `bring_a_put_away_pin_back`).
pub(super) fn arm_pin_to_the_front() {
    if !pinned() || pin_always_on_top() || !pin_update_move_to_top() || pin_is_collapsed() {
        return;
    }

    PIN_FRONT_OWED.store(true, Ordering::SeqCst);
}

/// Whether the window itself is in the topmost band right now, which is what a demotion has to take
/// away and what a raise into the band is asked to keep.
///
/// It is read off the window rather than from `PIN_FRONT_OWED` because the band can be entered by a
/// road that is not the flag's — a resize re-asserts the place a carried window has — and a demotion
/// that trusted the flag alone would leave that window over the listing it just went behind.
pub(super) unsafe fn window_holds_the_topmost_band(hwnd: HWND) -> bool {
    !hwnd.is_invalid() && GetWindowLongPtrW(hwnd, GWL_EXSTYLE) & WS_EX_TOPMOST.0 as isize != 0
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

/// Whether an update also brings the pin in front of what it was behind, as the configuration has
/// it (see `Pin Mode → Update Preview → Move To Top` and `raise_pinned_preview`).
pub(super) fn pin_update_move_to_top() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(DEFAULT_PIN_UPDATE_MOVE_TO_TOP)
}

/// Whether a pin collapsed into its bubble holds the video it is playing where it is, as the
/// configuration has it (see `Pin Mode → Pause on Minimize`). It is read by the tick that
/// keeps a collapsed pin's playback in step with the bubble, so a switch thrown while the pin
/// is a bubble is answered on the next tick (see `settle_bubble_playback`).
pub(super) fn pin_pause_video() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_pause_video)
        .unwrap_or(DEFAULT_PIN_PAUSE_VIDEO)
}

/// And the same question about a sound, asked in the same place (see
/// `Pin Mode → Pause on Minimize`).
pub(super) fn pin_pause_audio() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_pause_audio)
        .unwrap_or(DEFAULT_PIN_PAUSE_AUDIO)
}

/// Whether a double-click on a pinned window puts it in the middle of the screen it stands on, as
/// the configuration has it. It is read as the second press of a pair arrives, which is the only
/// moment the question is asked: a switch thrown while a pin is up is answered by the next
/// double-click rather than by the next pin.
pub(super) fn pin_double_click_center() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_double_click_center)
        .unwrap_or(DEFAULT_PIN_DOUBLE_CLICK_CENTER)
}

/// Whether a double-click on a pinned window puts it back to the size the scaling settings ask for
/// the kind it is showing, as the configuration has it — the other half of what the pair asks for,
/// read beside the first and at the same moment (see `pin_double_click_center`).
pub(super) fn pin_double_click_reset_size() -> bool {
    CONFIG
        .lock()
        .map(|config| config.pin_double_click_reset_size)
        .unwrap_or(DEFAULT_PIN_DOUBLE_CLICK_RESET_SIZE)
}

/// Where a pin the minimize button puts away goes — the round bubble on the desktop, or the app's
/// own button in the taskbar, as the configuration has it (see `Pin Mode → Minimize`).
///
/// It is read the moment the minimize button is pressed rather than every tick: which of the two
/// a pin is away as is settled when it is put away, and a switch thrown while one is already away
/// is the next minimize's (see `collapse_pin`).
pub(super) fn pin_minimize_to() -> PinMinimizeTo {
    CONFIG
        .lock()
        .map(|config| config.pin_minimize_to)
        .unwrap_or(DEFAULT_PIN_MINIMIZE_TO)
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

    // **Cleared after the teardown, not before it.** The teardown asks which window the pin stands in
    // — it lets go of the capture, gives the keyboard back and takes the window down — and a pin
    // whose flag was cleared first answered with the preview's window: a hidden window of the hover's
    // was released and hidden a second time while the pin's own was left standing on the screen with
    // a capture nothing would let go of (see `pin_stands_in_its_own_window`, `pinned_window`).
    pin_stands_in_its_own_window(false);

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
        pinned_window().0 as isize
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
        // is only asked to hide. The stand-ins a collapse can leave — the bubble and the taskbar
        // one — go with them (`hide_minimized_pin`), and a handle to a window that has since gone
        // is refused by the call rather than acted on.
        unsafe {
            hide_pinned_windows();
            hide_pin_taskbar();
        }
        hide_minimized_pin();
    }

    fn hide_pin_bubble(&self) {
        // The bubble and the taskbar stand-in are this app's own windows and hiding them is a
        // plain call on a handle read from the slot each was created into, with nothing this fn
        // has to be unsafe about. Both are taken down because a pin that ends must leave neither
        // behind, whichever of the two it was put away as.
        //
        // **Both, and not through `hide_minimized_pin` alone.** That call leaves the taskbar
        // stand-in up for a pin that is not Always On Top, because the pin on the screen wears one
        // too and a restore must not take it away — but this road is a pin that *ends*, and a
        // button left behind it is a taskbar entry for a window with nothing in it, which is what
        // a hover over it then finds (see `pin_taskbar`).
        hide_minimized_pin();
        unsafe { hide_pin_taskbar() };
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
/// And the window the card's menu was drawn in, which goes with the pin for the same reason the
/// bubble's does: it is a window this app put up standing for a pin that has gone, and a panel
/// nothing will ever paint again is a window the desktop keeps for nothing (see
/// `pin_menu_window`).
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
    // bubble treats a press as the end of a drag that began before it existed. And the latch
    // that says the bubble has been carried somewhere, which the next bubble starts without
    // (see `restore_pin`).
    PIN_BUBBLE_MOVED.store(false, Ordering::Release);
    PIN_BUBBLE_CARRIED.store(false, Ordering::Release);
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

    // And the window a pinned mpv film was drawn in — a host this app made a child of the pin's —
    // with the player it held: it is a window of this app's own, and one left standing for a pin
    // that has gone is a black band in the next pin's picture (see `pin_film_host`). The player is
    // the loop's own kill road's, which is asked for by the take-down the message below is.
    unsafe {
        pin_film_host::take_down_the_pin_film_host();
    }

    // And the window the card's menu was drawn in, if one is up: it is a window of this
    // app's own standing for a pin that has gone, exactly as the bubble is, and a panel
    // nothing will ever paint again is a window the desktop keeps for nothing.
    unsafe {
        hide_pin_menu_window();
    }

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

    // **And neither is a pin whose place on the screen is the round bubble.** What a bubble shows
    // is a picture this app holds — a frame read beside the film, a frame the media engine drew, or
    // the still the collapse grabbed — so a film whose player has gone leaves the circle standing on
    // that picture rather than taking the whole pin down: a bubble is not a window onto nothing, and
    // one that vanished over a player nobody is watching would take with it the only way back to the
    // file it was showing. What a *walk* leaves behind is the whole of what this is for: a step's
    // file whose player never came up — one the engine refused, one a swap for another film ended on
    // its way — is read here as the pin having come apart, which closed a bubble over a press that
    // was only asking for the next file (see `PinStep`, `abandon_pin_swap`).
    //
    // The pin is what comes apart instead where it is *restored*: a window put back up on a film
    // whose player is still gone is read without this exemption, which is the same answer the film
    // was owed all along. A pin collapsed into the taskbar stand-in is answered the same way, and
    // the cost of it is a stand-in that outlives its film rather than a pin that comes apart — an
    // icon the user has to click or close, against a bubble lost for a press nobody asked to end it.
    if pin_is_collapsed() {
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
    if pin_is_its_own_window() {
        if let Some(hwnd) = pin_real_window::real_pin_window() {
            return hwnd;
        }
    }

    HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _)
}

/// Whether a pin's window is down — not on the screen at all, or minimized, which is what the Shell
/// leaves of it under the desktop. It is asked by the show itself, which tells a pin standing where the
/// user put it from one being brought back from away, and by the restore's own question about what counts
/// as away for a kind that has no button of Windows to come back by (see `show_pinned_window` and
/// `the_pin_owes_a_show`).
///
/// **Which window is asked is the caller's to get right, and it is the pin's own one**: a picture pin
/// stands in a window of its own and has put the preview's down, so the hover's window answers "down"
/// for the whole life of the pin — and a restore that read it would put the pin up again on every
/// tick, on the one road that exists for a pin the user cannot see (see `pinned_window`). What that
/// costs a wheel is the whole of what it draws: the frames it presents went to the window that is
/// down, and the pin on the screen was moved to the box its state holds and painted there by this
/// restore instead (see `real_pin_window_at`), so each frame of a travel was the frame before it
/// stretched into the box the frame after it asked for (see `render_layered_preview` and
/// `step_pin_zoom`).
pub(super) fn the_pin_window_is_down(hwnd: HWND) -> bool {
    !unsafe { IsWindowVisible(hwnd) }.as_bool() || unsafe { IsIconic(hwnd) }.as_bool()
}

/// Whether the tick is to put a pin's window back on the screen: the pin is up, its window is away, and
/// the user is back in a window rather than looking at the desktop.
///
/// It is the whole of the question the loop's restore asks, gathered here so that the answer has one
/// home: which window is asked (the pin's own, see `pinned_window`), what counts as away, and whether
/// the desktop is still what the user is looking at.
///
/// **And what counts as away is not the same question for the two kinds of window a pin stands in.**
/// A pin that stands in the hover's window is a window this app put on the screen and this app has to
/// put back, so a `Win`+`D` — the Shell minimizing the preview's window with everything else — is the
/// road this restore exists for, and a minimized window is a window to bring up.
///
/// A pin that stands in a window of its own is a window **Windows** keeps: its own taskbar button puts
/// it down and brings it back, `Show Desktop` hands it back with everything else it took, and a press on
/// it is handled by the Shell rather than here. A tick that brought such a window up was a tick that
/// undid every minimize the user asked for — the button could restore and never minimize — and it is
/// asked about the one state Windows has no button for: a window that was *hidden*. Nothing in this app
/// hides a pin's own window while the pin is up (a collapse hides it with the pin's own state, which is
/// read above), so what this answers is nothing at all for the ordinary pin, and the safety net is kept
/// for the day the Shell takes a window off the screen rather than minimizing it.
pub(super) fn the_pin_owes_a_show(hwnd: HWND, foreground: HWND) -> bool {
    if !pinned() || pin_is_collapsed() || unsafe { foreground_is_the_desktop(foreground) } {
        return false;
    }

    // A pin waiting on its hosted film's first frame is hidden on purpose: its window goes up at
    // the handover and on no earlier tick, so the restore leaves it down (see
    // `pin_film_host::settle_pin_handover`).
    if pin_film_host::pin_handover_armed() {
        return false;
    }

    if pin_is_its_own_window() {
        return !unsafe { IsWindowVisible(hwnd) }.as_bool();
    }

    the_pin_window_is_down(hwnd)
}

/// Whether the pin taken up on a file stands in the pin's own window, decided once at the
/// take-up and written by [`pin_stands_in_its_own_window`].
///
/// Asked of the media in hand where it has a kind: what fills the pin's band is the kind's
/// answer, and every kind this app draws the whole way stands in the pin's own window (see
/// `MediaType::draws_in_a_window_of_its_own`). A film's answer there is its live route, so a pin
/// shown a film of another route after the engine row changed is answered for the film it is
/// showing now. The two placeholder kinds — a spinner and the cross a failed file leaves — are not
/// the file, so they ask the file's route instead, as a pin taken up with no kind at all does (see
/// [`a_pin_of_this_file_stands_in_its_own_window`]).
pub(super) fn pin_should_stand_in_its_own_window(kind: Option<MediaType>, path: &Path) -> bool {
    match kind {
        Some(kind) if !matches!(kind, MediaType::Loading | MediaType::Unplayable) => {
            !kind.draws_in_a_window_of_its_own(path)
        }
        _ => a_pin_of_this_file_stands_in_its_own_window(path),
    }
}

/// Whether a file whose media is not in hand belongs in the pin's own window.
///
/// A placeholder is still a pin of the file under it, so the file's own route is asked rather
/// than the spinner or the cross. Every route this app draws the whole way — a reader of its
/// own, a card, or the page an engine writes into this app's frame — belongs here; the browser's
/// three do not. A film is the pin's own window where mpv draws it into the pin's own host, and
/// the preview's where FFmpeg's player draws a window of the desktop's; the media engine's film is
/// drawn into this app's own surface like any other frame (see `engine_kind_of`,
/// `MediaType::draws_in_a_window_of_its_own` and `pin_film_host`).
fn a_pin_of_this_file_stands_in_its_own_window(path: &Path) -> bool {
    use crate::formats::routing::DrawnBy;

    // The engine's three are the browser's window, and the engine's own answer is what says it
    // is really drawing this file: the route names a page of HTML whatever the tray's switch
    // says, while the engine with the switch off draws the text preview instead.
    if engine_kind_of(path).is_some() {
        return false;
    }

    match HoverFacts::read(path).route.drawn_by {
        DrawnBy::MediaEngine => {
            matches!(video_route(path), VideoRoute::MediaEngine | VideoRoute::Mpv)
        }
        DrawnBy::WebView(_) | DrawnBy::Native(_) | DrawnBy::Card | DrawnBy::Nothing => true,
    }
}

/// Whether the pin on the screen is drawn the whole way by this app, and so stands in a window of its
/// own: a window Windows activates on a press, lists in Alt+Tab and gives a taskbar button to.
///
/// Asked of the kind rather than of the setting, because it is what the kind can be: the preview's
/// window is a tool window that never takes the focus, and every kind this app draws the whole way
/// has nothing of anybody else's standing in its band (see
/// `MediaType::draws_in_a_window_of_its_own`).
pub(super) fn pin_is_its_own_window() -> bool {
    pinned() && PIN_IS_ITS_OWN_WINDOW.load(Ordering::SeqCst)
}

/// Whether the pin that is up stands in a window of its own, written where the pin is taken up and
/// cleared where it ends.
///
/// It is a decision taken once, at the take-up, rather than a question asked of the media every time
/// a road wants the window: what is under the pin when it ends is not what it was taken up on, and the
/// window a pin stood in is the pin's own fact rather than the media's.
pub(super) fn pin_stands_in_its_own_window(its_own: bool) {
    PIN_IS_ITS_OWN_WINDOW.store(its_own, Ordering::SeqCst);
}

/// The pin that is up stands in a window of its own: written at the take-up from the kind the pin was
/// taken up on, or from the file's route where the media is a placeholder (see
/// `pin_should_stand_in_its_own_window`).
static PIN_IS_ITS_OWN_WINDOW: AtomicBool = AtomicBool::new(false);

/// Put a pinned window up at the box its state says, painted before it is shown: what a layered
/// window shows between one paint and the next is the surface it already has, stretched into
/// whatever box the window has (see `show_loading_spinner`).
///
/// A pin that is still a bubble has no window to put up. What is behind the bubble is the file the
/// walk reached, installed there *because* the window is away: raising it here would stand a second
/// window beside the bubble standing in for it, and the restore that takes the collapse down is what
/// puts the window up on that file (see `take_up_pinned_window`).
pub(super) unsafe fn show_pinned_window(hwnd: HWND) {
    if pin_is_collapsed() {
        return;
    }

    let Some((window, width, height)) = pinned_window_box() else {
        return;
    };

    let started = Instant::now();
    let mut created_now = false;

    // **The window is asked for, not taken from the caller.** A pin this app draws itself stands in a
    // window of its own — Windows activates it on a press and lists it in Alt+Tab — and the preview's
    // window is put down, because two windows holding the same picture is one picture too many.
    //
    // **But not yet.** The preview's window is what the user is looking at, and it is the only thing of
    // ours on the screen until the pin's has been painted *and* shown: put down first, the row behind
    // them shows through for as long as the paint takes — a scale of the file, a fresh surface and a
    // compositor round trip, one frame to four — which is the blink the hand sees taking a hover into a
    // pin. It is put down at the end of this function instead, once nothing of the file is left for it
    // to explain.
    //
    // Every caller passes the window its own message arrived in, which is the preview's for most of
    // them and the pin's own for the loop's restore — and it is no longer read for a pin that stands in
    // its own window: the answer is the state's, and it is the same answer every other road asks for
    // (see `pinned_window`).
    let hwnd = if pin_is_its_own_window() {
        created_now = pin_real_window::real_pin_window().is_none();
        let Some(hwnd) = pin_real_window::real_pin_window_at(window) else {
            return;
        };

        // The name on the button the window brings is the file's own, so that the Shell's label
        // beside the hover says which file the window is of (see `pin_taskbar_title`).
        pin_taskbar::set_pin_taskbar_title(hwnd);
        hwnd
    } else {
        pin_real_window::hide_real_pin_window();
        hwnd
    };

    // Whether the pin's window is on the screen as this is asked, which is what tells a pin standing
    // where the user put it from one being brought back from away — and it is read before the showing
    // below, since that is the call which puts it on the screen. A window the Shell minimized is a
    // window being brought back as well, which is why the question is `the_pin_window_is_down` rather
    // than the visibility alone (see `the_pin_window_is_down`): a pin minimized away is one to put in
    // front of wherever the Shell left it, not one already standing where the user left it.
    let on_the_screen = !the_pin_window_is_down(hwnd);
    if !on_the_screen {
        let _ = MoveWindow(hwnd, window.0, window.1, width, height, false);
    }

    // **The film's window goes where the band it is seen through went — before the paint and the
    // showing, not after.** A pin standing in a window of its own gets a host to put a pinned
    // mpv's picture in, at the band this window is about to be painted around; a pin standing in
    // the preview's has no such band, because the film it shows is a window of FFmpeg's player
    // and it is placed from the outside (see `pin_film_host` and `ensure_video_window_topmost`).
    // A host left by a pin that has just left this window is taken down here rather than waiting
    // for the hide that ends the pin: the file being installed may not be a film at all.
    //
    // **And a film the pin took up from a hover is begun again in that host behind the hover**,
    // rather than left drawing in the top-level window the hover gave it: the relaunch starts the
    // hosted player in the background while the hover stays visible and playing as the placeholder,
    // the pin's window and its strips stay down with the chrome hidden, and the tick's handover is
    // the only visible swap — the pin appearing with the film already drawing, and only then the
    // hover coming down (see `pin_film_host::settle_pin_handover`). Where the pin is already
    // standing over a park, the film is adopted the synchronous way instead: the cover that park
    // raised is the placeholder, and the show below hands the band over through it.
    // **And a film the pin took up from a hover whose player is already in the app's host is
    // moved rather than begun again.** The host the hover was drawing in is re-parented into this
    // window once the pin is on the screen, and its player keeps the same child and the same
    // playhead — no freeze, no read back of `time-pos`, no second decode. Every other take-up from
    // a hover keeps the freeze/relaunch handover below as the fallback.
    let adopts_the_hover_host = pin_is_its_own_window()
        && pin_state()
            .and_then(|pinned| {
                pinned
                    .pin()
                    .map(|pin| hover_film_host::the_hover_film_to_adopt(&pin.path).is_some())
            })
            .unwrap_or(false);
    unsafe {
        if pin_is_its_own_window() {
            if !adopts_the_hover_host {
                pin_film_host::take_up_the_pin_film_host(None);
                if begin_the_adopted_film_in_the_host() && !pin_player_is_parked() {
                    with_pin(|pin| pin.chrome = PinChrome::hidden());
                    pin_film_host::arm_pin_handover();
                    render_pinned_preview_at(hwnd, window.0, window.1);
                    return;
                }
            }
        } else {
            pin_film_host::take_down_the_pin_film_host();
        }
    }

    render_pinned_preview_at(hwnd, window.0, window.1);
    // Where the pin is owed: the topmost band for a pin that is Always On Top or one that owes the
    // front while it follows the listing (see `pin_holds_the_front`), and the ordinary one for one
    // that owes neither. The second is no raise — `HWND_NOTOPMOST` asks for the place a window
    // already has, and a pin the user can see is standing in it — but it *is* the one thing that
    // takes the topmost band off a window the Shell had been drawing as a hover's.
    let place = band_place();
    let _ = SetWindowPos(
        hwnd,
        place,
        window.0,
        window.1,
        width,
        height,
        SWP_NOACTIVATE | SWP_SHOWWINDOW,
    );

    // **And a window that was not on the screen is put in front, because the call above leaves it
    // where it is.** `HWND_NOTOPMOST` asks for the place a window already has, and the place a
    // hidden window has is the bottom of the z-order — behind every window opened over it since it
    // went down. A pin restored from the taskbar button or from its bubble came back *behind*
    // whatever the user had been working in, which is not the place a click on that button or on
    // that bubble ever asked for (see `restore_pin`). The raise is its own call rather than the
    // flag above, because the demotion is what a pin that is not Always On Top owes whatever it
    // stands over; and it is asked for of a window that was *off* the screen only, since one that
    // is on it is to be left exactly where the user's own clicks put it.
    if !pin_always_on_top() && !on_the_screen {
        // **A pin that is a window is asked for the front and left there.** What a window of the
        // Shell's own gets for being put up is being activated — that is how Windows puts a window in
        // front of the one the user was in, and it is the whole of what this app asks for: no band, no
        // place, no second call to undo the first. Whether the front arrives is Windows' to decide, and
        // a window shown behind the one that holds the foreground is Windows' answer rather than a
        // fault this road can repair without the band `Always On Top` exists to keep (see
        // `pin_real_window`).
        if showing_takes_the_band_to_appear(on_the_screen) {
            // **On top of everything the moment it appears, and a normal window from then on:** what
            // is left after the band is given back is a window at the front of the ordinary one, which
            // the next click on another window puts behind it, exactly as any window behaves.
            // `Pin Mode → Always On Top` is still the only thing that keeps a pin over other programs.
            take_the_band_and_give_it_back(hwnd);
        } else {
            let _ = SetWindowPos(
                hwnd,
                HWND_TOP,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );
        }
    }
    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);

    // The pin is painted and on the screen over the hover's still-top-level host, and only now is
    // that host moved into this window: the move is the visible handover, and the picture never
    // sits in a hidden parent or behind a pin that has not appeared yet.
    if adopts_the_hover_host {
        adopt_the_hovered_film_in_the_host();
    }

    // Traced for the take-up that shows the pin: a step's cover was painted before this line, so
    // the bars appear over the held frame, and a take-up from a hover reaches it through the
    // handover with the film already drawing (see `settle_pin_handover`).
    trace_park_step("show");

    // **And now the hover goes**, painted, shown and settled: from here the pin is what stands in the
    // box, and the window that stood in for it while it was only a hover comes down with nothing left
    // to explain (see the top of this function, which is where it used to go).
    let preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let preview_up = unsafe { preview != 0 && IsWindowVisible(HWND(preview as *mut _)).as_bool() };
    if pin_is_its_own_window() {
        if preview != 0 && preview != hwnd.0 as isize {
            let _ = ShowWindow(HWND(preview as *mut _), SW_HIDE);
        }
    }

    if !pin_always_on_top() {
        show_pin_taskbar();
    }

    // What the hand's blink would be measured as, where the trace is asked for: the whole road from the
    // box being read to the pin being on the screen, and whether this take-up is the one that made the
    // window (the first of a run is the slowest). `preview_up` says whether there was a hover under the
    // pin to hand over from at all.
    crate::shell::explorer_hook::note_pin_click!(
        "SHOW own {} created {} span {}us box {:?} preview_up {}",
        pin_is_its_own_window() as u8,
        created_now as u8,
        started.elapsed().as_micros(),
        window,
        preview_up as u8,
    );
}

/// How long after a pick a pin keeps answering for the front of its band.
///
/// The pick reaches the loop from the low-level hook *before* the listing's window has seen the
/// click, and a window is activated after the click reaches it — activation being Windows putting
/// that window at the front of the band the pin is in. So the raise a pick makes can be undone by
/// the click that made it, and the pin owes the front for as long as that can still happen: every
/// tick of the grace asks whether it still stands where it was put and raises it again if it does
/// not (see `keep_pin_at_the_front`).
///
/// **The grace is also the clock the hold a press takes runs on.** A press holds the band rather than
/// taking it and giving it straight back — no activation can be put over a topmost window, so the
/// tick the grace would answer on is a tick the user does not see — and the band goes back when this
/// much time has passed since the press, which is the last moment that click can undo anything. Asked
/// of the foreground instead, the release would be made on an answer Windows gives a moment before
/// the window it names has been raised, and the raise it was waiting for would put the listing over
/// the pin (see `PinFrontOwed` and `give_the_band_back`).
///
/// The tick a static pin wakes on is `STATIC_PIN_WAIT_MS`, so the grace has to be several of those
/// rather than the few milliseconds an activation takes.
///
/// shortcut: a window activated later than this after a pick stays over the pin; upgrade if a
/// listing is ever met that takes longer than this to answer a click.
pub(super) const PIN_RAISE_GRACE: Duration = Duration::from_millis(250);

/// Whether an update is what brings a put-away pin back up: `Pin Mode → Update Preview → Move To Top`
/// says an update brings the pin in front of the listing it follows, and a pin that is away has no
/// window in any band to bring forward — so the row asks for the pin itself to come back up, and the
/// file the update names is the window coming back rather than one held for the key (see
/// `pin_bubble_pick` and `bring_a_put_away_pin_back`).
///
/// Asked of the pin's own state rather than of the way it was put away, because the two places a
/// collapse can leave a pin — the round bubble on the desktop, and the window Windows minimizes
/// into its own taskbar button — are one fact about the pin (`pin.collapsed`), and a row that
/// reached only one of the two would be a setting the user cannot tell from one that does nothing.
/// Asked of the pin rather than of the kind it is showing for the same reason: the row is about
/// updates, and a pin put away is a pin the user cannot see, whichever kind it is showing.
///
/// The update itself is the two rows read together: the first says the pin follows the file the user
/// picks next, and the second only what it does about being behind — the tray greys the second out
/// while the first is off, and a pick the switch refuses is no update, so a bubble keeps holding its
/// file the way it always has (see `pin_update_enabled`).
pub(super) fn an_update_brings_a_put_away_pin_back() -> bool {
    pin_is_collapsed() && pin_update_enabled() && pin_update_move_to_top()
}

/// Bring a put-away pin back up on the file an update names, and put it in front of the window the
/// file was picked in: the whole of what `Pin Mode → Update Preview → Move To Top` asks of an update
/// for a pin that is away (see `raise_pinned_preview` for the same update's answer to a pin that is
/// up).
///
/// **Asked where the file is on the window, never where it is picked.** An update names a file that
/// nothing has read yet — the read and the decode are a thread's, and a window is painted from the
/// file it holds before it is put up (see `show_pinned_window`) — so a restore answered at the pick
/// puts the window up on the file *before* it and swaps the new one in a tick or more later, which is
/// the flash the user sees between the two images. So `named` is the file the pin must already be
/// showing: the loop owes the return at the pick and answers it in the take-up of the file, where the
/// install has made the named file the pin's own and the media behind it the window's (see
/// `install_pinned_media` and the loop's `PreviewMessage::Pin` arm). A file the pin cannot be shown
/// puts nothing on the window, so it is no file to come back on: the pin stays away until a file it
/// can show arrives.
///
/// **The front comes second, and it is asked for of both kinds of pin.** The restore has to be
/// first: a window that is off the screen has no place in the z-order to be raised from, and what
/// the show leaves is the place the pin went down in — behind everything opened over it while it
/// was away. And the raise is this door's own rather than the show's, because the show brings one
/// of the two kinds of pin forward by itself: a pin that is a window of its own is taken into the
/// topmost band for the moment it appears and the band is given straight back, while the kinds
/// standing in the preview's window are put up with `HWND_TOP` — a place Windows refuses to a
/// process the foreground window is not in, which is every update that arrives while the user is
/// picking in the listing (see `show_pinned_window` and `put_pin_at_the_front`). For a pin of its
/// own the ask is the same band twice in one tick, with no paint between the two calls, and that is
/// the price of one road for both kinds rather than one that quietly comes back behind the listing.
///
/// The keyboard stays in the listing the file was picked in: a pick is one step of a walk with the
/// arrows as often as it is a click, and a restore that took the focus would end that walk at the
/// first file. Nothing of the band is kept either — `Always On Top` is the setting's, and this is an
/// ordinary window from here (see `PinRestore::Update`).
///
/// Answers whether the pin was away and is up now, on `named`: a pin that is up is an update's own to
/// raise and not this road's, one whose row is off is the user's to bring back, and one that is
/// showing another file is a window that would come up on the file before it.
pub(super) fn bring_a_put_away_pin_back(named: &Path) -> bool {
    if !an_update_brings_a_put_away_pin_back() {
        return false;
    }

    // **The file the update names is the file the window must already be showing.** What a window put
    // up before the file is on it shows is the file before it, for as long as the read and the decode
    // take — the flash the user reported between the two images (see the top of this function). The
    // pin's own record is what answers, rather than the caller's word: the install writes the file and
    // the media in one breath, so a pin naming this file is a pin showing it, and a pick that arrives
    // ahead of its own load has nothing here to come back on.
    if pinned_path().as_deref() != Some(named) {
        return false;
    }

    restore_pin(&Win32PinWindow, PinRestore::Update);

    // Safety: asks the desktop about a window of this app's own — where it stands in the z-order,
    // and for the band and straight back out of it — which is the pair of calls every update that
    // finds the pin on the screen already makes.
    unsafe { raise_pinned_preview(pinned_window(), false) };

    !pin_is_collapsed()
}

/// Put a pinned window at the top of the band it is allowed to stand in, without taking the
/// keyboard and without moving it: what a preview update owes the user when the pin was behind the
/// listing they picked the file in (see `Pin Mode → Update Preview → Move To Top`).
///
/// The switch is read here rather than in the raise itself, because a press on the pin owes the
/// raise whatever the switch says about updates (see `put_pin_at_the_front`). A pin that is put away
/// has no window to raise and is left alone, and so is one whose row is switched off — the row's own
/// answer for a pin that is away is to bring it back up first, which is the other door and not this
/// one (see `bring_a_put_away_pin_back`).
///
/// **`keep_the_band` is the pick's press asking for the band to be held rather than given back**, and
/// it is the one thing a raise can be asked for that a later window can undo. A press is read off the
/// hook *before* the click reaches the listing, and the activation that click ends in is the listing
/// put in front of the ordinary band — over the pin the press just raised, for as long as it takes this
/// loop to answer, which is the flash the user sees between picking a file and the pin following it. A
/// window in the topmost band is not one an activation can be put over, so the band is held across that
/// one click and given back when the grace the press begins has run out — never on the foreground's word,
/// which Windows gives before it raises the window it names (see `keep_pin_at_the_front`). Nothing of it
/// is held past that grace, and no road that ends the grace leaves it held (see `give_the_band_back`).
///
/// **Only a pin that is a window of its own has a band to hold**, so a stand-in kind's raise is the
/// band taken and given straight back whatever `keep_the_band` says: the stand-in kinds have no window
/// of their own to be put over the listing, and their front is this app's place for them rather than a
/// band Windows could keep (see `put_pin_at_the_front`).
pub(super) unsafe fn raise_pinned_preview(hwnd: HWND, keep_the_band: bool) {
    if pin_is_collapsed() || !pin_update_move_to_top() {
        return;
    }

    // **A pin that is a window of its own is brought forward the one way Windows lets it be.** The
    // update arrives while the user is working in the listing, so the listing is the foreground window
    // and `HWND_TOP` is refused to this process — the pin would stay exactly where it was, behind the
    // listing it follows. The band is not refused, and it is given straight back: the pin comes over
    // everything the update would have left it behind and is an ordinary window from then on (see
    // `take_the_band_and_give_it_back`). Asked only where Windows has not already answered — a press
    // that activated the pin is the pin already in front, and a band taken over it is a band the user
    // sees change for nothing.
    if pin_is_its_own_window() {
        if !pin_always_on_top() {
            if keep_the_band {
                // **Taken and not given back**, because the press that asks for this has a window still
                // to come over it: the listing's activation, which no ordinary band can be kept in front
                // of. The window is already where the user's pick is to be seen, so this is not a raise
                // a band is owed for — it is the front kept for the one moment it can be taken away.
                if !window_holds_the_topmost_band(hwnd) {
                    let _ = SetWindowPos(
                        hwnd,
                        HWND_TOPMOST,
                        0,
                        0,
                        0,
                        0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                    );

                    // **A band the desktop is not giving is asked for again, from outside it.** The
                    // desktop answers a band assignment as made and can still not make it: a window it
                    // has not laid out yet — one shown this moment — and one whose process has no
                    // foreground rights are both windows it leaves in the band they were in, which for a
                    // hold is a pin the click it is held for goes over. Leaving the band and taking it is
                    // the same front asked a second time, and it is the way the take-and-give raise asks
                    // for the front on every update (see `take_the_band_and_give_it_back`): one call more
                    // on a road the first ask did not answer, and none at all where it did.
                    if !window_holds_the_topmost_band(hwnd) {
                        let _ = SetWindowPos(
                            hwnd,
                            HWND_NOTOPMOST,
                            0,
                            0,
                            0,
                            0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                        );
                        let _ = SetWindowPos(
                            hwnd,
                            HWND_TOPMOST,
                            0,
                            0,
                            0,
                            0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                        );
                    }
                }

                // Said out loud as well as done, because the window is placed again before the click it
                // is held for has landed: the pick's own show asks `band_place`, and an answer from the
                // setting alone is the band given back a few milliseconds too early (see
                // `PIN_BAND_HELD`).
                PIN_BAND_HELD.store(true, Ordering::SeqCst);
            } else if GetForegroundWindow() != hwnd {
                take_the_band_and_give_it_back(hwnd);
            }
        }
        return;
    }

    put_pin_at_the_front(hwnd);
}

/// Put a pinned window at the front of the band it is allowed to stand in, without taking the
/// keyboard and without moving it.
///
/// **`HWND_TOP` cannot do this, and asking for it is a call that answers nothing.** Windows will not
/// let a window of a process that is not in the foreground stand over the window that is: the pick
/// arrives while the user is clicking in the listing, so the listing's window is the foreground one
/// and the place asked for is refused, leaving the pin exactly where it was. The topmost band is
/// not subject to that — a topmost window stands over the foreground window by definition — so the
/// raise is made there and given straight back. `HWND_NOTOPMOST` puts the window above every
/// non-topmost window and behind the topmost ones, which is the front of exactly the band a pin
/// that is not Always On Top belongs in, reached in a way the foreground window cannot refuse. A
/// pin that *is* Always On Top keeps the band it is in and its raise is the one call.
///
/// **A press is the other caller, and the refusal above is the whole of why it is.** A press on the
/// pin asks Windows for the foreground, and Windows refuses that to a process whose windows another
/// window in front has put it behind — which is exactly the press that matters: the first one after
/// the user has clicked anything else. The keyboard still arrives (`SetFocus` is refused only while
/// the window is `WS_EX_NOACTIVATE`, see `pin_take_focus`), so the pin reads as focused while the
/// window it is in stays behind the one that was pressed before it, and the press looks unanswered.
/// The band-and-back raise is not refused, so it is what answers a press.
///
/// The keyboard is left where the user is working, which is the half of a keyboard walk a raise that
/// took the focus would break, and the window is left where they put it: neither call moves it.
///
/// What stands in the pin's band is somebody else's window and is brought along with it, so the
/// raise is a whole preview rather than a caption over whatever happened to be behind it.
pub(super) unsafe fn put_pin_at_the_front(hwnd: HWND) {
    // **A pin that is a window and already holds the foreground is in front.** The press that asks
    // for this has activated the window it landed on, and a window Windows has activated is the one
    // in front of the ordinary band: taking the topmost band and giving it back would be a second
    // answer to a question already answered, and the band it took is a flicker the user sees.
    //
    // A pin that is a window and does *not* hold the foreground is the case this whole road exists
    // for, and it is the ordinary one: a pick or a preview update arrives while the user is clicking
    // in the listing, so the listing is the foreground window and the pin is what is to be brought in
    // front of it (see `raise_pinned_preview`). Asking about the foreground rather than about the kind
    // of window is what makes both roads right with one question.
    // **A pin that is a window is left to Windows** - activated by the press that lands on it, and
    // with no band to take here for it: a press that lands on the pin is the pin already in front, and a
    // press that names no file is answered by the lowering rather than by this road (see
    // `lower_pinned_preview`). Everything below is the stand-in the kinds still standing in the preview's
    // window need, because a window the Shell does not own is never raised by a click.
    if pin_is_its_own_window() {
        return;
    }

    // A pin that owes the front *holds* the topmost band rather than taking it and giving it back:
    // the band is the half of this that the listing's activation cannot undo, and it is kept until
    // the user's own press gives it back (see `pin_holds_the_front`, `lower_pinned_preview`). What
    // stands in the pin's band is put in it too, in the order `place_pinned_siblings` asks for.
    if pin_holds_the_front() {
        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        place_pinned_siblings();
        settle_pinned_mpv_ontop();
        return;
    }

    // Taken and given straight back: the band is the only way to the front of the ordinary one for a
    // window this process cannot ask the foreground for (see `take_the_band_and_give_it_back`).
    take_the_band_and_give_it_back(hwnd);
    place_pinned_siblings();
}

/// Put a pinned window back where the user's own clicks had it: behind the window the press was
/// made in, and out of the topmost band a pin that owes the front stands in.
///
/// The window the press was made in is handed in as the place to go behind, rather than a band:
/// `HWND_NOTOPMOST` would put the pin at the top of the ordinary band, which is where the window it
/// is to go behind already is, so the place is the one directly below the window the press landed
/// in. That window and not the foreground one, which is still the window the user came from while
/// the click that made the press has not been dispatched yet — and a pin put behind the window the
/// user came from is a pin left over the one they are looking at. `window` is handed in rather than
/// read here for the same reason `keep_pin_at_the_front` takes it: the caller knows the moment the
/// press was read, and a test can name a window of its own. An invalid window is no place at all:
/// the band is still taken away.
///
/// **The band is left before the place is asked for.** A topmost window cannot be put under an
/// ordinary one, so a pin that holds the topmost band is first put back in the ordinary band and
/// only then placed below the window. `HWND_NOTOPMOST` alone would be the pin coming back to the
/// front of the ordinary band, which is over the very window it is to go behind.
///
/// A pin that is Always On Top has no ordinary band to go back to and keeps the one it is in, and
/// one whose `Move To Top` row is off was never raised by an update: both are left alone. So is a
/// pin that already stands behind the window, and one whose window is the pin itself — a press on
/// the pin is not a press in the listing.
///
/// What stands in the pin's band is somebody else's window — a film, or a page the engine draws
/// in — and it goes under the pin rather than staying where it was: a window left over the listing
/// is the pin's picture floating over it with no caption, and one the band is being taken from is
/// the same window over the listing the pin just went behind. A film is moved by the placement that
/// already inserts it below the pin (see `player_insert_after`), and is taken out of the topmost
/// band first where it holds one; a page is put there here as well, because the engine's own
/// placement is a message away and a page left in front of the listing for the length of it is the
/// picture floating over the window the pin went behind. **A page is put in front of the pin rather
/// than behind it**, so the window it goes under is the same one the pin went under and not the pin
/// itself (see `Host::put_window`). The next raise puts both where the band asks for them again
/// (see `place_pinned_siblings`).
pub(super) unsafe fn lower_pinned_preview(hwnd: HWND, window: HWND) {
    // **A pin that is a window of its own is put behind by this road too, and here is why the Shell
    // cannot be the one to do it.** The Shell puts an ordinary window behind the one a click activates,
    // but a click in the window that is *already* active reorders nothing — and Explorer is the
    // foreground window while the user works in it, so a press on its toolbar, its tree or the gap
    // between its rows left the pin standing over the listing it follows. The press names no file, and
    // the pin goes behind the window it was made in. That is recoverable rather than final: the pin's
    // own window is one a click may activate, so a click on the pin brings it back (see
    // `pin_set_focusable`).
    if pin_is_collapsed() || pin_always_on_top() || !pin_update_move_to_top() {
        return;
    }

    if window == hwnd {
        return;
    }

    // A band held for a press is given back by this road whatever comes of the lowering, because a press
    // that named no file has ended the answer the band was kept for — and every placement after it asks
    // the setting again (see `PIN_BAND_HELD`).
    PIN_BAND_HELD.store(false, Ordering::SeqCst);

    let topmost = window_holds_the_topmost_band(hwnd);

    // A topmost window — the taskbar, a tooltip, another app's own band window — is not a place to
    // go behind: a pin put under the taskbar is a pin over every ordinary window, which is the
    // listing it was meant to go behind. The window to go behind is then the front-most ordinary
    // one, which is the window the user is working in.
    let pressed_topmost = !window.is_invalid() && window_holds_the_topmost_band(window);
    let place = if pressed_topmost {
        front_ordinary_window(hwnd)
    } else {
        window
    };

    if !topmost && !pressed_topmost && (place.is_invalid() || !pin_stands_over(hwnd, place)) {
        // Already behind it, or with no window to go behind at all: there is no band to give back,
        // and asking for the front of the ordinary one would be the pin coming back over the
        // listing by the road this is meant to take it away by.
        //
        // **No window is the desktop's, and it is no place.** The tick hands it in as no window rather
        // than as the desktop's own handle, whose bottom is below the wallpaper: a pin put there is a pin
        // nobody can click back (see `event_loop`). So the press asks for nothing that is about a place,
        // and the front a pick owed ends here instead — a front left owed is a grace ticking to undo a
        // press that has already been answered.
        PIN_FRONT_OWED.store(false, Ordering::SeqCst);
        return;
    }

    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    if topmost {
        let _ = SetWindowPos(
            hwnd,
            HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }

    if !place.is_invalid() {
        let _ = SetWindowPos(
            hwnd,
            place,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }

    let kind = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.media_type));
    match kind {
        Some(MediaType::Video) => {
            // A film that held the topmost band — the player's own `--ontop`, or the band the pin
            // owed — is taken out of it before it is put under the pin: the place below an ordinary
            // pin has no topmost window in it, and a film left in the band is the picture floating
            // over the listing the pin just went behind.
            let film = VIDEO_HWND.load(Ordering::SeqCst);
            if film != 0 {
                let film = HWND(film as *mut _);
                if window_holds_the_topmost_band(film) {
                    let _ = SetWindowPos(
                        film,
                        HWND_NOTOPMOST,
                        0,
                        0,
                        0,
                        0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                    );
                }
            }
            place_pinned_siblings();
            settle_pinned_mpv_ontop();
        }
        Some(MediaType::EngineSvg) | Some(MediaType::EngineFont) => {
            let page = webview_preview::showing_hwnd();
            if page != 0 {
                let page = HWND(page as *mut _);
                if window_holds_the_topmost_band(page) {
                    let _ = SetWindowPos(
                        page,
                        HWND_NOTOPMOST,
                        0,
                        0,
                        0,
                        0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                    );
                }
                // Below the same window the pin went below, and asked for after it: inserting a
                // window directly below `place` puts it directly above the pin, which is where a
                // page stands. Below the *pin* would be the page blended through the band's hole —
                // the black a browser's surface pays for standing behind a layered window (see
                // `Host::put_window`).
                if !place.is_invalid() {
                    let _ = SetWindowPos(
                        page,
                        place,
                        0,
                        0,
                        0,
                        0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                    );
                }
            }
        }
        _ => {}
    }
}

/// The pick a pin still owes the front for: when its grace runs out, and the listing the click that
/// made it was in.
///
/// The listing travels with the grace rather than being read off the foreground on the tick, because
/// the two are different windows and the answer is about the first one. The foreground is whatever
/// the user has just moved to: another program where they went elsewhere, another folder window
/// where they picked a second listing — and a pin raised over either is a pin with `Always On Top`
/// off standing over a window the user chose, which is user story 10 exactly. Read where the press
/// is, the window is the one the click landed in, and the click is what the raise answers for (see
/// `PIN_RAISE_GRACE` and `keep_pin_at_the_front`).
///
/// **The band the press holds while this is owed is the grace's to give back, not the answer's.**
/// What is owed is a front, and with the listing's activation still on its way the band is the only
/// place a front can be kept; so the pick holds it, every road that ends the answer gives it back,
/// and no road that ends it leaves it held (see `give_the_band_back`). The band's own deadline is the
/// press's and is never extended, so a file that outlasts this grace is answered by a pin that is no
/// longer over every program the user works in.
#[derive(Clone, Copy)]
pub(super) struct PinFrontOwed {
    /// When the answer is no longer owed: set by the press, and moved on by every update that reaches
    /// the window late (see `extend_pin_front_owed`).
    pub(super) until: Instant,
    /// When the band the press took goes back: set by the press too, and **never moved on**. A file's
    /// read can outlast the grace it was picked in, and the answer owes the pin a front for as long as
    /// that takes — but the band is the answer's answer, not the file's, and a hold that followed the
    /// hand through a folder would be the pin over every program the user works in for as long as they
    /// kept stepping (see `keep_pin_at_the_front` and `give_the_band_back`).
    pub(super) band_until: Instant,
    pub(super) listing: HWND,
}

/// Ask the grace a pick began to last another `PIN_RAISE_GRACE` from now, where the file it was
/// raised for has only just reached the window: the read and the decode can outlast the grace the
/// press began, and the click that answers for the listing may still be on its way (see
/// `keep_pin_at_the_front`).
///
/// A grace that is not owed is not begun here. An update with no press behind it — a file walked to
/// with the keys, a hover where `Pin Mode → Update Preview → On Hover` asks for one — is a pick no
/// click is to undo, and the raise the update makes stands on its own.
///
/// **The answer is extended and the band the press held is not.** What a late file asks for is another
/// grace to be answered in; the front the press kept belongs to the click, and letting an update move it
/// would make the pin over every program the user works in for as long as files kept being stepped
/// through (see `PinFrontOwed`).
pub(super) fn extend_pin_front_owed(owed: &mut Option<PinFrontOwed>) {
    if let Some(owed) = owed.as_mut() {
        owed.until = Instant::now() + PIN_RAISE_GRACE;
    }
}

/// Answer, on each tick of the grace a pick began, for the front of the pin's band: hold the band a pick's
/// press took until the grace it began has run out, and leave the pin where the answer belongs when it does
/// — over the listing the click was made in, or behind the window the user moved to instead (see
/// `PIN_RAISE_GRACE`, `where_the_band_goes_back` and `give_the_band_back`).
///
/// Nothing is asked once the grace has run out, and nothing of a pin that already stands where it was put — a
/// tick costs the walk from that listing up to the pin and no more, which is a handful of calls rather than a
/// reorder of the desktop. The listing is the one the owed answer carries rather than one read here, so the
/// walk is asked of windows a test can make (see `PinFrontOwed`) — and `foreground` is the window the user is
/// in, handed in for the same reason: a test process cannot make its own window the foreground, and the
/// question the release is made on is whether that window is the listing.
pub(super) unsafe fn keep_pin_at_the_front(
    hwnd: HWND,
    owed: &mut Option<PinFrontOwed>,
    now: Instant,
    foreground: HWND,
) {
    let Some(PinFrontOwed {
        until,
        band_until,
        listing,
    }) = *owed
    else {
        return;
    };

    // **Nothing of the answer outlives the click.** What is owed here is the front of one pick: the band
    // is held across the click that made it and given back on the press's own clock, never past the grace,
    // and once the grace has run out — or a press that named no file has given the front back — nothing
    // asks again (see `PIN_FRONT_OWED`, `raise_pinned_preview` and `give_the_band_back`).
    //
    // **The chase is what is left of the grace once the band is out of it.** A raise is made where the
    // pick arrives and the click that made it reaches the listing *after* that: the window the click
    // landed in is activated a moment later, which is that window put over the pin again. A pin that is a
    // window of its own is over it from that raise until `band_until` — the band leaves no window above it
    // to be activated over (see `PIN_BAND_HELD`) — so the chase is the stand-in kinds' answer, and the
    // answer of a real pin whose band was given back ahead of its clock (see
    // `a_pick_chases_the_listing_when_the_activation_lands_after_its_raise`). Every raise the chase makes
    // keeps the band only if one is still held, and takes and gives it back otherwise.
    //
    // A pin that no longer owes the front has given it back: a press that named no file is the
    // pin's own answer, and a raise here would be the grace putting it back over the listing the
    // press just took it behind — the pin fighting the hand that moved it. That road's raise goes behind
    // the window the user is in whether or not it is the listing, which is the one thing it differs in
    // from every other road out of the grace (see `where_the_band_goes_back`).
    //
    // **The debt is what is asked, not the band.** A pick arms `PIN_FRONT_OWED` and a press that names
    // no file clears it, and that flag is the whole of the promise this grace is; the band is the
    // setting's and answers "no" with `Always On Top` off, which is every pin this grace exists for.
    // Asked of the band, the grace ended on its first tick and the click that made the pick — reaching
    // the listing *after* the raise, and activating it — was left to undo the pin's answer: the pin
    // went back behind the listing it had just been raised over, which is the fault the user reported
    // twice ("we have the same problem just like in fake windows").
    if !PIN_FRONT_OWED.load(Ordering::SeqCst) {
        *owed = None;
        give_the_band_back(hwnd, foreground);
        return;
    }

    // A grace that has run out, and one whose listing is no longer a window, are both answers this debt
    // has nothing left to ask for — and a band held for it is given back here rather than left to the
    // tick's own repair, since the repair has no way of knowing where the user is (see
    // `give_the_band_back`).
    if now >= until {
        *owed = None;
        give_the_band_back(hwnd, where_the_band_goes_back(foreground, listing));
        return;
    }

    // A listing that is no longer a window is a pick nothing more can happen to: the folder was
    // closed inside the grace, and the activation that was to land is gone with it. So the debt goes
    // with the window. The walk cannot answer for it either — `pin_stands_over` reads a window that
    // is gone as one the pin does not stand over, since nothing above it ever reaches the pin — and
    // asking it would raise the pin over whatever window the user had moved to instead, which is the
    // one thing `Always On Top` off must never do (see `PinFrontOwed`).
    if !IsWindow(listing).as_bool() {
        *owed = None;
        give_the_band_back(hwnd, where_the_band_goes_back(foreground, listing));
        return;
    }

    // **The band goes back on the press's clock, and nowhere else.** The last moment the click that made
    // the pick can undo anything is a moment the press's grace covers, and asking the foreground whether
    // the click has landed is a question Windows answers a moment too early: the activation sets the
    // foreground before the window it names has been raised, so a release made on that answer could be
    // undone by the raise it was waiting for — Explorer over the pin, which is the flash this exists to
    // keep from being possible at all (see `PIN_BAND_HELD`).
    //
    // What is left at that moment is exactly the front the pick was owed. The user is in the listing the
    // click activated, and `HWND_NOTOPMOST` puts the pin at the front of the ordinary band, which is over
    // that listing: the band the press asked for and the activation that followed it settle into one
    // window over the file the user picked, with nothing between them for the user to see. Anywhere else
    // — the user moved to another program inside the grace — the pin goes behind the window they are in,
    // which is what the setting asks of a pin left over somebody else's (see `give_the_band_back`).
    //
    // The answer goes with the band, for a pin that is a window of its own: the front it owed *was* the
    // band, and a chase after it would be the pin put over a listing by a raise taken and given straight
    // back — the very shape of the fault. The stand-in kinds have no band to hold and keep the chase they
    // always had (see `PIN_RAISE_GRACE`).
    // **The answer goes with the band, where there is a band to go with.** A band the desktop answered as
    // made and did not make (see `raise_pinned_preview`) is no band this release can give back: the
    // placement below is a no-op on a window that never left the ordinary band, and dropping the answer
    // here would be a pin left behind the listing it followed — the very fault, on the one road the take
    // is not trusted. What is left there is the stand-in kinds' road, the chase, which is bounded by the
    // answer's own clock and reaches this branch again a tick later.
    if pin_is_its_own_window() && now >= band_until {
        let band_was_the_desktops = window_holds_the_topmost_band(hwnd);
        give_the_band_back(hwnd, where_the_band_goes_back(foreground, listing));
        if band_was_the_desktops || pin_stands_over(hwnd, listing) {
            *owed = None;
            return;
        }
    }

    if !pin_stands_over(hwnd, listing) {
        raise_pinned_preview(hwnd, PIN_BAND_HELD.load(Ordering::SeqCst));
    }
}

/// Where the band a pick's press held goes back to: the pin's own front when the user is in the listing the
/// click was made in, and behind the window they moved to when they are anywhere else.
///
/// **The two are one answer, not two roads.** The front of the band is where `HWND_NOTOPMOST` leaves a window
/// anyway, and it is the answer for the listing because the pin is meant to stand over the file the user picked
/// while they go on working in that listing — so releasing the band onto the user's own listing is not the pin
/// going behind what it followed, it is the pin staying exactly where it was. Anywhere else the front of the
/// band is the pin over a program the user chose, which is the one thing the setting forbids (see
/// `give_the_band_back` and `lower_pinned_preview`).
fn where_the_band_goes_back(foreground: HWND, listing: HWND) -> HWND {
    if foreground == listing {
        HWND::default()
    } else {
        foreground
    }
}

/// Give back the topmost band a pick's press held across the listing's activation, and leave the pin where
/// the user is.
///
/// A window in the topmost band is over every program the user works in, which with `Always On Top` off is
/// the one thing a pin must never be — so the band is only ever held for the length of a pick's click, and
/// every road out of the grace that held it gives it back here rather than leaving it to whichever road
/// happens to place the window next (see `raise_pinned_preview` and `keep_pin_at_the_front`). A window that
/// is not holding the band is left alone: this is no raise, and a pin standing where the user's own clicks
/// put it is not one to move.
///
/// **The band goes back where the user is, not where the pin was raised.** `HWND_NOTOPMOST` alone puts the
/// window at the front of the ordinary band — over the program the user just moved to, which is the fault the
/// hold was a fix for, one program along. Where the window the user is in is somebody else's, the pin goes
/// below it, which is what the user's own click in that window would have done; where there is no such window
/// to go behind — the activation road, the pin itself, the desktop — the front of the band is where it belongs
/// (see `lower_pinned_preview` and `foreground_is_the_desktop`).
///
/// **A window of this app's own is no exception**, though `pid_is_the_previews` is the question the tick asks
/// about the foreground: the band is given back all the same. What that gate protects there is a pin put
/// behind its own film for as long as the film is in front, which is a band the pin places again the moment it
/// is drawn (see `band_place`) — while a band left held here is over every program the user works in until
/// something else takes it, and nothing else does.
///
/// **The band is given back whatever the settings say.** `Move To Top` is about a raise — whether a pick is
/// owed the front of the band at all — and a band already held is not that question: it is a window over every
/// program the user works in, and the setting being turned off under it is one more reason to take it away.
/// Where the pin is placed after the band goes is the setting's, since that is a raise (see
/// `lower_pinned_preview`).
///
/// `behind` is the window the pin is to go below, or no window where the front of the band is the answer. It
/// is handed in rather than read here for the same reason the grace is handed the window the user is in: a test
/// process cannot make a window of its own the foreground, and the question is only ever which window that is.
unsafe fn give_the_band_back(hwnd: HWND, behind: HWND) {
    // Said first, whatever the window's band turns out to be: a hold that is over is over, and a
    // placement after it asks the setting again (see `PIN_BAND_HELD`).
    PIN_BAND_HELD.store(false, Ordering::SeqCst);

    if !pin_is_its_own_window() || pin_always_on_top() || !window_holds_the_topmost_band(hwnd) {
        return;
    }

    // The band goes first, whatever else this road does: the one thing it must never leave is a window in
    // the topmost band with no debt left to ask about it (see `PinFrontOwed`).
    let _ = SetWindowPos(
        hwnd,
        HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
    );

    if behind.is_invalid() || behind == hwnd || foreground_is_the_desktop(behind) {
        return;
    }

    lower_pinned_preview(hwnd, behind);
}

/// Whether the pin stands over `window` in the z-order right now: whether walking up from that
/// window reaches the pin. The question a pick's grace is answered with, and the same one the
/// tests ask of two windows they made (see `keep_pin_at_the_front`).
///
/// It is asked rather than remembered because the answer is not this app's to keep: another window
/// being activated is that window being put in front, and where the pin stands is read off the
/// desktop rather than worked out from what this app last asked for. A pin that is put away, and one
/// whose row is off, are owed nothing and answer yes.
pub(super) unsafe fn pin_stands_over(hwnd: HWND, window: HWND) -> bool {
    if pin_is_collapsed() || !pin_update_move_to_top() {
        return true;
    }

    if window.is_invalid() || window == hwnd {
        return true;
    }

    let mut above = window;
    while !above.is_invalid() {
        if above == hwnd {
            return true;
        }

        above = GetWindow(above, GW_HWNDPREV).unwrap_or_default();
    }

    false
}

/// The front-most window a pin belongs behind: the first visible window in the z-order that is
/// neither this app's own, nor topmost, nor the desktop.
///
/// What a press in a topmost window — the taskbar, a tooltip, another app's band window — has to
/// go behind. A pin put under such a window is still over every ordinary one, and the ordinary one
/// the user is looking at is the listing it was taken behind.
unsafe fn front_ordinary_window(except: HWND) -> HWND {
    let mut window = GetTopWindow(HWND(ptr::null_mut())).unwrap_or_default();
    let mut seen = 0;
    while !window.is_invalid() && seen < 200 {
        if window != except
            && !foreground_is_ours(window)
            && IsWindowVisible(window).as_bool()
            && !window_holds_the_topmost_band(window)
            && !foreground_is_the_desktop(window)
        {
            return window;
        }

        window = GetWindow(window, GW_HWNDNEXT).unwrap_or_default();
        seen += 1;
    }

    HWND(ptr::null_mut())
}

/// Whether the window in front belongs to this app: the pin, the menu it opens, the page the engine
/// draws, the bubble, the taskbar button that stands for the pin. A pin that owes the front owes it
/// to the listing, and the front being the app's own is the app answering itself — the press that
/// opened a menu, the click that took the pin's caption — which is not the user leaving the listing.
pub(super) unsafe fn foreground_is_ours(window: HWND) -> bool {
    let mut pid = 0u32;
    GetWindowThreadProcessId(window, Some(&mut pid));
    pid_is_the_previews(pid)
}

/// Whether a process is this app or the player it runs the film with.
///
/// **A player's window is the preview's, though its process is not.** The film is drawn by ffplay or
/// mpv — a window of somebody else's standing in the pin's media band — and what it is not is a
/// program the user is working in. Read as one, the window is a foreground the pin owes nothing to:
/// the tick put the pin behind its own film the moment that window came up in front, which is the
/// pin going out of the band with the picture left over it, and a press that picked a file went the
/// same way whenever the new film's window came up before the raise could land (see
/// `foreground_is_ours`, which is the two questions this answers for the watch and for the window a
/// demotion goes behind).
pub(super) fn pid_is_the_previews(pid: u32) -> bool {
    pid != 0 && (pid == unsafe { GetCurrentProcessId() } || pid == VIDEO_PID.load(Ordering::SeqCst))
}

/// Whether the window in front is the desktop rather than a program: the shell's `Progman`, which is
/// what a press on the wallpaper leaves in front, or the `WorkerW` an animated wallpaper is drawn
/// in.
///
/// A press on the desktop names no file, so the pin gives the front back (see
/// `lower_pinned_preview`) — but not by going *below* the desktop: the desktop is the bottom of the
/// desktop's own list, and a pin put under it is a pin behind the wallpaper. The band is given back
/// and the place the pin stands in kept, which is what an ordinary window no other window was
/// pressed over keeps.
pub(super) unsafe fn foreground_is_the_desktop(window: HWND) -> bool {
    let mut name = [0u16; 32];
    let len = GetClassNameW(window, &mut name).max(0) as usize;
    let class = String::from_utf16_lossy(&name[..len.min(name.len())]);
    class == "Progman" || class == "WorkerW"
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
        // band — and it is put back where it stands rather than asked for again: a page this
        // engine holds is a window that moves with the box rather than a document to navigate to
        // (see `webview_preview::raise`). The ask matters as much as the box does, because the pin
        // may have just been raised over the window that was covering the page as well, and a page
        // left behind would be a caption standing over the listing it was following.
        //
        // What the ask puts the page in is the band the pin now holds, and in front of the pin
        // rather than behind it: a browser's surface behind the band's hole is a resize that goes
        // black and stutters (see `Host::put_window`).
        Some(MediaType::EngineSvg) | Some(MediaType::EngineFont) => {
            if let Some((path, _)) = pinned_media_owner() {
                webview_preview::raise(
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
///
/// `paused` is the loop's own record of where a player of this app's is being held, and is the
/// only half of the play button's answer this side writes down: whether the engine is holding the
/// sound is the engine's own, and is asked here (see `a_sound_that_is_held`).
pub(super) fn pinned_audio_chrome(paused: Option<f64>) -> Option<CardChrome> {
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

    // Whether a player is running behind the card, which is what the play button's glyph says:
    // the engine's own answer where the engine is the one holding the sound, and the hold a key
    // wrote down for a player of this app's, which was ended to hold it (see
    // `a_sound_that_is_held`). It is asked behind the kind the pin is showing, because this is
    // asked on every tick with a pin of any kind up and a card that is not on screen has no
    // button to draw one either.
    let playing = if pin_shows_an_audio_card(pin) {
        let player = audio_track::playable(&pin.path)
            .map_or(Player::Ffmpeg, |track| playing_player(&pin.path, &track));

        !a_sound_that_is_held(paused, player, video_player::is_paused())
    } else {
        true
    };

    Some(CardChrome {
        playing,
        volume: pin.volume.level,
        always_on_top: pin_always_on_top(),
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

/// The box a pin of a sound is given, which is the box its hover was:
/// the box of the card as it is drawn, at the share its frame was
/// painted at, so a change to Audio Scaling that arrived while the
/// hover was up does not resize the window pinning it.
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
pub(super) fn pinned_audio_card_box(
    rect: ScreenRegion,
    path: &Path,
    dpi: u32,
    scale: PreviewScale,
) -> ScreenRegion {
    let chrome = Some(CardChrome {
        playing: false,
        volume: current_audio_volume(),
        always_on_top: pin_always_on_top(),
        hovered: None,
        pressed: None,
        window_buttons: false,
    });
    let Some(card) = audio_card(path, None, None, 0, chrome) else {
        return rect;
    };

    let room = audio_box_room(work_area_at(rect.0, rect.1), scale, dpi);
    let Some((width, height)) = audio_preview::measure(
        &card,
        room.width.max(1),
        room.height.max(1),
        dpi,
        audio_options_of(scale),
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

/// The options a pinned sound's card is drawn with: the ones its card
/// was taken up with — the options the share it remembers names — kept
/// beside the dpi and the box it was taken up at (see
/// `PinnedPreview::audio_scale`), so the card's own arithmetic — the press,
/// the seek, the volume popup's geometry, the window buttons' band — is
/// answered against the layout the card is drawn in rather than the
/// configuration's, and a change to Audio Scaling reaches the next take-up.
/// A pin that remembers none — one that shows no sound's card — is answered
/// with the configuration's.
pub(super) fn pinned_audio_options(pin: &PinnedPreview) -> AudioPreviewOptions {
    pin.audio_scale
        .map(audio_options_of)
        .unwrap_or_else(current_audio_options)
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
        pinned_audio_options(pin),
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
    y < audio_preview::window_button_band(pin.dpi, pinned_audio_options(pin))
}

/// Put a window of somebody else's — the player's — where a pinned window's media band is.
///
/// **A film mpv draws travels with the band without ever changing the order.** Its window is
/// topmost by itself (`--ontop`), and the pin stands over it; moving both with the order left
/// alone keeps it that way — a raise of the film first puts it over the pin's own chrome for a
/// frame, which is the blink, and a raise of the pin after it is a second reorder per pointer
/// move, which is the drag's slowness. So the film is moved without a raise and the pin is not
/// re-raised here at all (see `player_order`, whose hold-offs stand for mpv's film as they do
/// for FFmpeg's).
///
/// **A band whose player is parked is left alone, which is what this is for rather than an
/// afterthought of it.** The park is a hide on the drag's first pointer message, and a raise here
/// is a show as well as a place (see `ensure_video_window_topmost`): a resize drag puts the player
/// here on every pointer move, and a maximized or restored window, a relayout and a take-up each put
/// it here once more, so the film would be back on screen for as long as the drag lasted. What the
/// park holds is the window's *absence*, and there is nothing about being out of date to say the
/// absence can be ended (see `pin_player_is_parked`).
pub(super) fn ensure_pinned_sibling_box(content: ScreenRegion) {
    // **A film drawn into the pin's own host travels with the band rather than being placed at it.**
    // The host is a child of the pin's window, so a move is the compositor's and never this app's;
    // what a resize changes is the band *inside* that window, which is the host's own box — and the
    // player's geometry is never touched, because mpv's child fills whatever box the host is given
    // (see `place_the_pin_film_host`). Placed while a park stands as much as after one: the band a
    // hidden host comes back at is the band it was given (see `settle_pinned_park`).
    if pin_film_host::the_pinned_player_was_begun_in_the_host() {
        unsafe { pin_film_host::place_the_pin_film_host(content) };
        return;
    }

    if pin_player_is_parked() || pin_is_collapsed() {
        return;
    }

    if video_process_image() == MPV_IMAGE_NAME {
        let hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
        if hwnd_val == 0 {
            return;
        }
        let hwnd = HWND(hwnd_val as *mut _);
        unsafe {
            if !IsWindow(hwnd).as_bool() {
                return;
            }
            // Moved without a raise: `SWP_NOZORDER` leaves the order alone, so the pin stays
            // over the film it was standing over and no frame puts the film over the chrome.
            //
            // **And a pin that does not hold the front puts the film *under it* while it moves
            // it.** The order still travels with the box — the pair keeps its place in the desktop
            // — but the film is put directly below the pin, which is where the chrome is over it
            // and where an mpv window adopted from a hover stops being a topmost one (see
            // `put_player_under_the_pin`). A pin that holds the front keeps the old arrangement:
            // its film holds its own band up and the pin stands over it.
            if pin_holds_the_front() {
                // A film already in the band is moved without a raise: `SWP_NOZORDER` leaves the
                // order alone, and raising the film first is the band over the pin's chrome for a
                // frame (see `place_player_window`). A film the band is being asked for — the pin
                // owes the front again after a press gave it back, and the film was put under it
                // then — is put in it below the pin itself, which is the same order the two keep
                // once the film is in the band, and is one call rather than a raise of each.
                let already = window_holds_the_topmost_band(hwnd);
                let pin = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
                let insert_after = if already || pin.is_invalid() {
                    HWND_TOPMOST
                } else {
                    pin
                };
                let _ = SetWindowPos(
                    hwnd,
                    insert_after,
                    content.0,
                    content.1,
                    (content.2 - content.0).max(1),
                    (content.3 - content.1).max(1),
                    if already {
                        SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW
                    } else {
                        SWP_NOACTIVATE | SWP_SHOWWINDOW
                    },
                );
            } else {
                let _ = SetWindowPos(
                    hwnd,
                    player_insert_after(),
                    content.0,
                    content.1,
                    (content.2 - content.0).max(1),
                    (content.3 - content.1).max(1),
                    SWP_NOACTIVATE | SWP_SHOWWINDOW,
                );
            }
        }
        return;
    }

    let _ = ensure_video_window_topmost(
        content.0,
        content.1,
        (content.2 - content.0).max(1),
        (content.3 - content.1).max(1),
    );
}

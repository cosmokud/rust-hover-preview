//! What a pin *is*: the window's own state, its transport bar, its tooltip, the program a
//! file is handed to, and the keyboard it claims.

use super::*;

/// Whether the pin that is up is collapsed into the bubble that stands for it: a state in which
/// nothing of anybody else's belongs on top of the bubble — a player's window included.
///
/// Asked of the pin rather than of a flag beside it, because the two had to be written together
/// and nothing said so: a collapse set `pin.collapsed` under the pin’s lock and `PIN_COLLAPSED`
/// beside it, and a reader of the flag believed it rather than looking — and the question this
/// answers is about the window the user is looking at, so a flag that could be set without a
/// pin was a flag that could be wrong about one.
pub(super) fn pin_is_collapsed() -> bool {
    pin_state().is_some_and(|state| state.pin().is_some_and(|pin| pin.collapsed))
}

/// Whether the pinned window is the window the user is in, and so whether a key belongs to it.
///
/// It is answered of Windows rather than remembered, because Windows is what decides it: the flag
/// the app kept about itself was a guess that could be wrong in the one direction that cannot be
/// recovered from — a `SetFocus` refused by the foreground lock left the guess standing, and a
/// guess that stands keeps every later press from asking again, so the pin went deaf to the
/// arrows for the rest of its life with no way back. Asking costs one call, on a press, and
/// cannot be wrong.
pub(crate) fn pin_is_focused() -> bool {
    if !pinned() {
        return false;
    }

    // Safety: asks the keyboard for the window that holds it, and names this thread's own
    // window. Neither is a dereference and neither can fail.
    unsafe { GetFocus() == HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _) }
}

/// Give the pin the keyboard, from a press that has landed on it.
///
/// A pin is a window of this app's that stands over a listing, and a window the user presses is a
/// window the user is in: asking for the foreground is what puts the caret on a folder being named
/// into the pin instead, and the key that would have finished the name is answered here instead.
/// It is asked for on the press rather than on the pin coming up, because a pin comes up wherever
/// the pointer happens to be — a pin that took the focus as it appeared would take it out of
/// whatever the user was typing, under a preview that has not been touched at all.
///
/// The two calls that ask are two *asks*, and what is written below is written on the answer
/// rather than on the asking. Windows refuses them both under the foreground lock — the
/// foreground goes to whichever process had the last input, and a process that did not have it is
/// not given it for asking — and that refusal is ordinary rather than a fault: it is what a press
/// on the pin looks like while something else owns the input, and the hand simply presses again.
/// A claim written on the asking is a pin that holds a keyboard it was never given and a note of
/// a window the user is no longer in, and the handover that note exists for is the foreground
/// asked back from a window the user left (see `pin_release_focus`).
pub(super) unsafe fn pin_take_focus(hwnd: HWND) {
    if !pinned() {
        return;
    }

    // Where the keyboard is being taken from, remembered before it is taken: this is the only
    // moment the window in front is still the one the user was in. A window that is hidden
    // while it is still the one holding the focus leaves Windows to pick whatever it likes to
    // activate next, and for a `WS_EX_TOOLWINDOW` popup that is not reliably the Explorer
    // window that was in front a moment ago — the arrangement that leaves a desktop on which
    // nothing answers the keyboard. A popup has no owner to ask either: a window with no parent
    // has no `GW_OWNER` at all, so the window the caret was in is the one that was in front a
    // moment before the press, and the only moment it can be read is the moment the pin takes
    // over from it (see `pin_window::take_keyboard`).
    let behind = GetForegroundWindow().0 as isize;

    // The window is only focusable while a pin is up, so the style is asked to change
    // before the focus is asked for: `SetFocus` on a `WS_EX_NOACTIVATE` window is refused
    // (see `pin_set_focusable`).
    pin_set_focusable(hwnd, true);

    let _ = SetForegroundWindow(hwnd);
    let _ = SetFocus(hwnd);

    // Whether the keyboard arrived is asked of Windows, by the same question `pin_is_focused`
    // asks, because that is the one answer here that cannot be wrong: the caret is either in
    // this window or it is not, and where it is says what the pin holds rather than what the
    // calls above hoped for. A press that was refused therefore leaves the pin holding no
    // keyboard at all — the state a pin the hand has never pressed is in — and that is a state
    // which asks again rather than one that is stuck: the keyboard is not in the pin, so
    // `pin_is_focused` is false on the next press and the press is answered afresh (see
    // `WM_LBUTTONDOWN`). What such a press leaves alone is a claim an earlier one earned: the
    // keyboard in it did arrive then, and Windows drops that claim itself the moment the pin is
    // activated away from (see the `WM_ACTIVATE` above).
    take_keyboard(behind, pin_is_focused());
}

/// Hand the keyboard back, from the pin having lost it: the user has clicked into another window,
/// or the pin is on its way down.
///
/// Nothing is done to the window now in front — it has the focus already, and asking for it again
/// is the kind of thing that brings a window up over the user. What is dropped is the claim: the
/// pin is not the window the user is in, so the keys the caption walks its folder with are the keys
/// of whatever is in front of it, which is the ordinary arrangement everywhere else in Windows.
///
/// The claim goes with the window it came from, so there is nothing left to hand over to: a pin
/// holding a note of a window that took the keyboard a moment ago is a pin that will ask for the
/// foreground back from a window the user is no longer in.
pub(super) fn pin_release_focus() {
    release_keyboard();
}

/// Whether a pinned window can take the focus at all, which is what `WS_EX_NOACTIVATE` takes away
/// and a hover preview must never have: a preview is put up under a pointer that is often inside
/// a file name being renamed, and a window that took the focus out from under that name would
/// swallow the rest of it.
///
/// A pin is not that: the hand pressed it, and the user is in it. So the style is put on when a
/// pin comes up and taken off when it goes, and the window itself is created the way a hover has
/// to be created.
pub(super) unsafe fn pin_set_focusable(hwnd: HWND, focusable: bool) {
    let style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    if style == 0 {
        // A style of zero is a failure, not a window with no style: this window always carries
        // `WS_EX_LAYERED`, `WS_EX_TOOLWINDOW` and `WS_EX_TOPMOST`, so a genuine zero cannot
        // occur. Writing a style computed from it would wipe the three rather than fail, and the
        // window would stop being layered — a preview that draws nothing, with nothing to report
        // it.
        return;
    }

    let wanted = if focusable {
        style & !(WS_EX_NOACTIVATE.0 as isize)
    } else {
        style | WS_EX_NOACTIVATE.0 as isize
    };

    if wanted == style {
        return;
    }

    // Both the old and the new style come back, and the new one is the one that counts only if it
    // is not the error sentinel: a refused write leaves the window exactly as it was, which is a
    // window that cannot be focused rather than one that can.
    if SetWindowLongPtrW(hwnd, GWL_EXSTYLE, wanted) == 0 {
        return;
    }

    // The style is not a thing a window is told about live: without being asked to re-read it,
    // a window keeps refusing the focus it was created with.
    let _ = SetWindowPos(
        hwnd,
        None,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
    );
}

/// The height of the caption a pinned preview of a kind is given at a display's scale, which is
/// none at all for a sound.
///
/// It is asked of the kind rather than carried as a number because the question is about what the
/// window is *for*, and for a sound the card is all of it: a card carries its own name at its own
/// top and its own controls on its own row, so a title bar above it is a second title with a
/// close button on it and nothing underneath that the hand can take hold of but the card itself.
/// It is settled once at the take-up for the same reason the transport bar's own answer is (see
/// `PinnedPreview::caption`).
pub(super) fn pinned_caption_height(dpi: u32, kind: Option<MediaType>) -> i32 {
    match kind {
        Some(MediaType::Audio) => 0,
        _ => logical_px(dpi, PIN_CAPTION_PIXELS).max(1),
    }
}

/// The height of the transport bar a pinned preview of a kind that plays is given, and none
/// at all for a kind that has nothing to play.
pub(super) fn pinned_transport_height(dpi: u32, transport: bool) -> i32 {
    if transport {
        logical_px(dpi, PIN_TRANSPORT_PIXELS).max(1)
    } else {
        0
    }
}

/// The room a pinned window leaves its media on the display it is on: the display's work area
/// with the caption taken off the top and the transport bar off the bottom — or the work area
/// itself for a kind whose chrome is drawn over the media, which has no bands to leave room for
/// (see `pin_overlay_chrome`).
pub(super) fn pinned_room(
    bounds: ScreenBounds,
    dpi: u32,
    transport: bool,
    overlay: bool,
    caption: i32,
) -> ScreenBounds {
    if overlay {
        return bounds;
    }

    ScreenBounds {
        left: bounds.left,
        top: bounds.top + caption,
        right: bounds.right,
        bottom: bounds.bottom - pinned_transport_height(dpi, transport),
    }
}

/// The box a pinned window occupies: the media's own box with the caption above it and the
/// transport bar below it — or the media's box itself for a kind whose chrome is drawn over the
/// media, since the strips the caption and the bar are drawn in are inside it.
pub(super) fn pinned_window_box_of(
    content: ScreenRegion,
    dpi: u32,
    transport: bool,
    overlay: bool,
    caption: i32,
) -> ScreenRegion {
    if overlay {
        return content;
    }

    (
        content.0,
        content.1 - caption,
        content.2,
        content.3 + pinned_transport_height(dpi, transport),
    )
}

/// The rows of a pinned window the media is drawn in, and how many of them there are: the whole
/// window for a kind whose chrome is drawn over the media, and the rows between the two bands
/// otherwise. What the chrome is drawn *over* is what this answers — a caption over a picture, or
/// a caption sharing the window's first rows with nothing else. A kind with no caption at all has
/// no top band, so the media begins at the window's own first row (see `pinned_caption_height`).
pub(super) fn pinned_band_rows(
    height: i32,
    caption: i32,
    transport: i32,
    overlay: bool,
) -> (i32, i32) {
    if overlay {
        (0, height.max(1))
    } else {
        (caption, (height - caption - transport).max(1))
    }
}

/// The box the media of a pinned preview takes when its window is given a room: the largest box
/// of the media's own shape that fits it — the rule every other preview is placed by (see
/// `scale_dimensions`).
///
/// What a pin is never given is the letterbox a box of a fixed shape would leave: the window is
/// fitted to the media, so every band of it is filled by the thing it is a band of. At
/// fit-to-screen the media is scaled up to the room as well as down to it, which is what makes a
/// maximize a maximize — a small picture is drawn as large as the display can show it rather than
/// sitting in the middle of the screen at its own size — and at a percentage it is the size the
/// app would have hovered it at, which is what a display change keeps (see `toggle_pin_maximized`
/// and `replace_pinned_window`).
pub(super) fn pinned_media_box(
    orig_dims: (u32, u32),
    room: ScreenBounds,
    scale: PreviewScale,
) -> (i32, i32) {
    let (room_width, room_height) = room.room();
    let (width, height) = scale_dimensions(
        orig_dims.0.max(1),
        orig_dims.1.max(1),
        room_width,
        room_height,
        scale,
    );

    (width as i32, height as i32)
}

/// What a kind of preview is composited over, which is a question a hover and a pinned window
/// both ask: a texture has a backdrop of its own, the drawing of a design document another,
/// and everything else the picture's (see the `Background` submenu).
///
/// It is the same exhaustive match the hover's own answer is composed with, so a media type
/// this app grows and a kind it grows cannot each need remembering here: the three that are not
/// a picture's are named by the one router table that says which of the tray's six backdrops
/// stands behind what.
pub(super) fn preview_background(kind: MediaType) -> TransparentBackground {
    if kind.is_loading() {
        TransparentBackground::Transparent
    } else {
        current_background(match kind {
            MediaType::Dds => crate::formats::routing::Backdrop::Dds,
            MediaType::Design => crate::formats::routing::Backdrop::Design,
            MediaType::Vector => crate::formats::routing::Backdrop::Vector,
            _ => crate::formats::routing::Backdrop::Image,
        })
    }
}

/// The pinned preview: the preview that stopped being a hover.
///
/// The media is the same media — pinning is not a reload, so nothing about the picture moves
/// when the key is pressed. What the pin is, is the room the media is given beside what the
/// window adds to it: a caption above it, a transport bar below it, and a state of its own
/// that says who takes it down and what the pointer is doing on it.
///
/// It is written by the preview loop, which holds the window and the media, and read by the
/// window procedure and every repaint — which is why it is a field of the pin's state rather
/// than a local of the loop (see `pin_window::PinState`). A pin is up or it is not, and the
/// value that says so is behind the pin's own lock; `pinned()` is the same answer published
/// for the threads that ask it without wanting a lock.
pub(super) struct PinnedPreview {
    /// The file that is pinned — the hover the pin came from, kept by name so that a box that
    /// changes can be laid out again without asking the Explorer hook anything.
    pub(super) path: PathBuf,
    /// The box the media occupies on screen. Everything the window draws besides the media is
    /// chrome over or around this box, which is why the box is what the pin remembers and what a
    /// restore puts back (see `PinnedPreview::window_box`).
    pub(super) content: ScreenRegion,
    /// The longest side a swap may give this pin's media, in either direction, or nothing while the
    /// pin has no bound yet: the longest side of the box the pin went up with, or of the one a
    /// manual resize last left it at, which are the two things that write it (see `apply_pin_drag`
    /// and `pin_bound_after`).
    ///
    /// It exists because a swap fits one shape inside a box, and a box that was itself the last
    /// swap's answer can only ever lose a side of it: a pin that followed a listing past files of
    /// different shapes would walk its way down to nothing with every file picked. What a swap
    /// writes is `content`; this stays the size the user gave the pin (see `pin_update_plan`).
    ///
    /// One number rather than the box, because the files a pin is shown have shapes of their own:
    /// a pin taken up as a tall portrait would otherwise hand every widescreen file that followed
    /// it a box no wider than the portrait was. A side is what the user's own hand asked for when
    /// it dragged an edge, so a side is what every shape gets to use.
    pub(super) bound: Option<i32>,
    /// The box a restore down puts back, and `None` while the pin is not maximized: the box the
    /// window had before it was maximized, kept in step with the hand while it is one — and given up
    /// altogether the moment a hand moves the window, since a maximize the user has taken hold of
    /// is one there is nothing left to undo (see `pin_restore_box`).
    pub(super) restore: Option<ScreenRegion>,
    /// The scale of the display the pin was put up on: what the caption's measurements are
    /// multiplied by, and what the media is laid out at again when the box changes.
    pub(super) dpi: u32,
    /// The share the sound's card this pin shows was taken up at, when
    /// the pin is showing a sound: the size the card's frame was
    /// painted at at the take-up, kept beside the dpi and the box the
    /// card was taken up at.
    ///
    /// The take-up and the frame's first painting capture the share at
    /// the same moment, so this and the frame's own remembered copy
    /// (`MediaData::audio_scale`, which the take-up reads) are one
    /// value, which is why a hover that becomes a pin keeps the
    /// hover's size, and a walk to the next sound — whose media was
    /// freshly loaded — takes up the new share.
    ///
    /// The card's own arithmetic — the press, the seek, the volume
    /// popup's geometry, the window buttons' band — is answered from
    /// the options this share names rather than from the
    /// configuration, so a press is answered against the layout the
    /// card is drawn in, and a change to Audio Scaling reaches the
    /// next take-up only (see `pinned_audio_options`). Nothing at all
    /// for a pin that shows no sound's card.
    pub(super) audio_scale: Option<PreviewScale>,
    /// Whether this kind carries a transport bar, decided when the pin was taken up: it is the
    /// same answer for as long as the pin lasts, and the window's own height is measured from it.
    pub(super) transport_bar: bool,
    /// Whether that bar's controls do anything, which is a question about the engine playing the
    /// file rather than about the kind: FFmpeg's player reports no position, takes no pause, and
    /// is taken to another second only by being ended and begun again — so a pinned video it
    /// plays carries a bar that is a read-out, with no button and nothing to drag (see
    /// `pin_chrome::TransportState`). Every control on the bar is a question the media engine
    /// answers, and the two are told apart by the kind of media a pin was taken up on.
    pub(super) transport_live: bool,
    /// What the edges and the caption do for this kind, decided when the pin was taken up for the
    /// same reason the line above is (see `PinFrame`).
    pub(super) frame: PinFrame,
    /// Whether this kind's chrome is drawn over its media rather than in bands around it, decided
    /// when the pin was taken up for the same reason the two lines above are: it is a question about
    /// the kind, and the window's own box is measured from the answer (see `pin_overlay_chrome`).
    pub(super) overlay: bool,
    /// Whether the pointer can ask for this pin's chrome to come and go. It is true of the kinds
    /// whose chrome is drawn *over* their media, where a strip is in the way of the thing the
    /// window is for.
    ///
    /// The arrival window is the same one every other kind gets: a pin brought up by a key has no
    /// pointer anywhere near it, and a caption that never appeared would be a window whose close
    /// button has not been drawn yet (see `PinChrome::on_arrival`).
    pub(super) hides_chrome: bool,
    /// The room a pinned window keeps above its media for a caption, settled once at the take-up so
    /// that every question about where a point is on this window reads one number — and nothing at
    /// all for a sound, which is given no caption (see `pinned_caption_height`).
    pub(super) caption: i32,
    /// How much of that chrome is showing. A title bar painted over a picture is a strip of the
    /// picture nobody can see, so it is there when the hand is near it and gone a moment after the
    /// pin comes up otherwise — and a pin whose chrome is *not* drawn over its media has nothing to
    /// show or hide, so it is always at the one level (see `PinChrome`).
    pub(super) chrome: PinChrome,
    /// Whether the window is collapsed into the round bubble the minimize button leaves.
    pub(super) collapsed: bool,
    /// What the bubble has parked while the window is collapsed, and what puts it back when the
    /// pin is put up again: nothing while nothing was — a gate switched off, a video the bar was
    /// used to pause before the pin was collapsed, a kind with no player behind it (see
    /// `BubblePause`).
    pub(super) bubble_pause: Option<BubblePause>,
    /// The caption button the pointer is over and the one it has pressed: what the caption is
    /// painted from, and what a release acts on.
    pub(super) hovered: Option<pin_chrome::CaptionButton>,
    pub(super) pressed: Option<pin_chrome::CaptionButton>,
    /// What the caption's buttons are saying out loud, and which of them the pointer has been
    /// resting on since enough time for a name to have been read (see `PinTooltip`).
    pub(super) tooltip: PinTooltip,
    /// A drag or a resize in progress.
    pub(super) dragging: Option<PinDrag>,
    /// Whether the player behind this pin has been put away for the length of a drag.
    ///
    /// A video pin's picture is not drawn by this app: the band in the middle of the window is
    /// transparent, and FFmpeg's own window shows through it. That is what makes the picture sharp
    /// and free, and it is also what makes the window expensive to move — every pointer move puts
    /// the player's window to a new size and the compositor has to re-blit a 2560x1440 surface,
    /// which on a 144 Hz screen is a stutter the hand feels rather than sees (see `park_pinned_player`).
    ///
    /// While this is set the player's window is hidden and the band is painted a flat
    /// colour instead of being left for it to show through, so a drag is a rectangle of the pin's own
    /// background following the pointer. Blank is the point: there is nothing for the picture to be
    /// rescaled into, so there is nothing to rescale.
    ///
    /// **It carries nothing about the film**, and that is the whole of how the pair avoids pausing
    /// twice. Whether the film is held, and whether it should come back, are the transport's own
    /// facts (`PinTransport::drag_held`), written by the tick that holds it and reconciled by every
    /// path that ends the player it was made against — so a flag here that remembered whether the
    /// film was playing would be a second record of a thing that is recorded once (see
    /// `park_pinned_player`).
    pub(super) parked: bool,
    /// Where the playback of a video is, which is what the transport bar is drawn from and what
    /// a seek or a pause is measured against (see `PinTransport`).
    pub(super) transport: PinTransport,
    /// The level this pin plays at, which belongs to this window rather than to the setting it was
    /// read from (see `PinVolume`).
    pub(super) volume: PinVolume,
    /// The card's own control the pointer is over, and the one it has pressed: what the card is
    /// painted from, and what a release acts on. Nothing while the window is not showing a sound's
    /// card, and nothing at all while the card carries no controls (see `pinned_audio_chrome`).
    pub(super) audio_hovered: Option<CardControl>,
    pub(super) audio_pressed: Option<CardControl>,
    /// Whether the card's two window buttons are showing, which is what the card
    /// is painted from: a hand near the window's top border or near the buttons
    /// themselves, asked on every move and every tick (see `pinned_mouse_move`
    /// and `pin_audio_hover_refresh`).
    pub(super) audio_window_buttons: bool,
    /// The card's own menu: whether its panel is up, and which of its two
    /// pages it is showing (see `PinMenu`).
    pub(super) menu: PinMenu,
}

#[cfg(test)]
impl PinnedPreview {
    /// A pin with nothing asked of it and nothing answered about it: the state a test needs to
    /// stand a pin up in, without a box to place or a media to show.
    ///
    /// It is the same pin `overlay_pin` builds, at a box that fits on screen, so a test that
    /// wants a particular box or a particular chrome says so rather than starting here.
    pub(crate) fn for_test() -> PinnedPreview {
        Self {
            path: PathBuf::from("picture.png"),
            bound: Some(100),
            content: (0, 0, 100, 100),
            restore: None,
            dpi: 96,
            audio_scale: None,
            transport_bar: false,
            transport_live: false,
            frame: PinFrame::Shaped,
            overlay: true,
            hides_chrome: true,
            caption: pinned_caption_height(96, None),
            chrome: PinChrome::always(),
            collapsed: false,
            bubble_pause: None,
            hovered: None,
            pressed: None,
            tooltip: PinTooltip::default(),
            dragging: None,
            parked: false,
            transport: PinTransport::default(),
            volume: PinVolume::default(),
            audio_hovered: None,
            audio_pressed: None,
            audio_window_buttons: false,
            menu: PinMenu::default(),
        }
    }
}

/// What a pin's collapse into its bubble parked, and what it takes to put it back.
///
/// Two engines play a moving file and the bubble parks them differently: the media engine
/// Windows has is told to pause and holds where it stands, while FFmpeg's player takes no pause
/// at all and is therefore ended, with the second of the file it had reached kept here for the
/// player that takes its place when the pin comes up again (see `start_audio_player` and
/// `restart_pinned_player`). Which engine it is is the media's own business rather than the
/// setting's, so what is parked is asked of the engine that is playing.
#[derive(Clone, Copy)]
pub(super) enum BubblePause {
    /// The media engine was paused where it stood — a video or a sound it was playing.
    Engine,
    /// FFmpeg's player was ended at this second of the file — a video or a sound it was playing.
    Player(f64),
}

/// Where a pinned video's playback is.
///
/// Two engines play a video and they answer these questions differently — the media engine
/// reports where it is, and FFmpeg's player reports nothing at all — so what is kept here is
/// what both can be asked: the length of the file (read by the probe that measured it), where
/// this app's own clock says a player it started has got to, and whether that player has been
/// stopped where it stood (see `pin_playhead`).
#[derive(Clone, Copy, Default)]
pub(super) struct PinTransport {
    pub(super) duration: Option<f64>,
    /// When a player of this app's was started, and the second of the file it was started at.
    pub(super) started: Option<(Instant, f64)>,
    /// Where a player of this app's was stopped — a pause — and nothing while it is running.
    ///
    /// What is behind this field changed when a video FFmpeg plays became a video this app can
    /// hold: a pause used to be that player *ended*, with the second it had reached kept here
    /// for the player that took its place, and it is now a key sent to a player that keeps
    /// running (see `ffplay_key_pause`). So `paused_at` no longer implies a dead process — the
    /// process behind a held file is *alive and holding still* — and the second written here is
    /// read out of this app's own clock over the start below, which is the only position a
    /// player that reports nothing can be measured by.
    pub(super) paused_at: Option<f64>,
    /// Whether the hold in `paused_at` has reached the player yet.
    ///
    /// It is here for the one case where a hold is written down before it can be acted on: a
    /// relaunch of a held file begins a player that has no window for a moment, and a key cannot
    /// be posted to a window that does not exist (see `ffplay_key_pause`). The hold is written
    /// anyway — the file *is* meant to be held — and this flag is what says the player has not
    /// been told yet, so the loop tells it the moment there is a window to tell it through (see
    /// `settle_pending_hold`). It is a field rather than a flag beside the pin because it is the
    /// same kind of thing as the rest of this struct: something this app has asserted about a
    /// player, which a press of the play button has to be able to take back — and the press that
    /// takes it back is the answer it changes (see `toggle_pinned_playback`).
    pub(super) pending_hold: bool,
    /// Whether the hold in `paused_at` is a gesture's rather than a press of the pause button's.
    ///
    /// It is a field of this struct rather than a flag of its own beside the pin because a claim
    /// about a player belongs to the player it was made against: every path that begins, holds,
    /// lets go of, loses or replaces a player writes *here*, and a pin taken down or a file swapped
    /// takes this with it. A flag beside the pin outlives all of them, and a claim that outlives
    /// its film is a claim the next film to be dragged is answered with — and it is answered with a
    /// pause, because the key a drag posts is a toggle (see `video_drag_hold_apply`).
    ///
    /// So it is reconciled by those same writes rather than asked of anything, and the one that
    /// must *not* reconcile it is a relaunch carrying the hold: the film is still held for the
    /// gesture that is still in flight (see `PinTransport::begun`).
    pub(super) drag_held: bool,
    /// Where the pointer is dragging the bar, while it is: what the playhead is drawn at rather
    /// than where the file really is, since a drag that is still going is not a seek yet.
    pub(super) seeking: Option<f64>,
    /// Which part of the bar the pointer is over, and which it has pressed.
    pub(super) hovered: Option<pin_chrome::TransportPart>,
    pub(super) pressed: Option<pin_chrome::TransportPart>,
    /// The subtitle track this pin is showing, as the index FFmpeg's `-sst s:` specifier
    /// numbers them by — that is, from zero, counting only the file's subtitle streams — and
    /// nothing for a file with none.
    ///
    /// It is kept here rather than read back off the player, because there is nothing to read
    /// it back off: the player was started with this number and reports nothing at all about
    /// what it did with it. So this is the whole of what a track choice *is* as far as a
    /// relaunch is concerned — a seek and a resize both begin a player again, and a choice
    /// this app had not written down would silently go back to whatever the player chose for
    /// itself (see `next_subtitle`).
    pub(super) subtitle: Option<usize>,
}

impl PinTransport {
    /// A player of this app's was begun at `from`, where `up` says whether one is behind it and
    /// `holding` whether it is meant to be paused.
    ///
    /// The second is written down either way because a player that has not come up is still the
    /// one the bar is drawn against — the wait for a window is the loop's business, and a bar
    /// that forgot where it was while waiting would spring back to the beginning the moment the
    /// window appeared.
    ///
    /// `holding` is what carries a hold *through* a relaunch, and without it every relaunch of a
    /// held file would start playing: a seek, a resize settling, a change of track. Those are all
    /// this one function, so this is where a hold has to survive one, and it survives it as an
    /// assertion rather than as an action — the player that has just begun has no window to be
    /// given a key through, so the hold is written down as owing and the loop delivers it (see
    /// `pending_hold`). The second written is the relaunch's own second rather than the one the
    /// hold began at, because that is the second the file is now at and the second the bar has to
    /// agree with.
    pub(super) fn begun(&mut self, from: f64, up: bool, holding: bool) {
        self.started = up.then_some((Instant::now(), from));
        self.paused_at = holding.then_some(from);
        self.pending_hold = holding && up;
        self.seeking = None;
        // A claim of a gesture's goes with the player it was made against, so it survives a
        // relaunch exactly where the hold does: a seek or a resize settling underneath a hand in
        // flight begins another player that is held for that same gesture, and dropping the claim
        // there would leave the film frozen for good the moment the hand let go of the window. A
        // relaunch that carries no hold is a film playing on, and no gesture is holding a film that
        // is playing (see `drag_held`).
        self.drag_held = self.drag_held && holding;
    }

    /// The player was told to hold where it is, at `at`.
    ///
    /// The start is *kept*, and that is the whole of what a hold is for a video FFmpeg plays:
    /// the player is still running, it is the one holding the second, and the clock underneath
    /// `started` is what the position is read from the moment the file is let go again — so a
    /// resume is a key sent to the player that is already there rather than a second player
    /// begun from scratch (see `ffplay_key_pause`).
    ///
    /// Nothing is left owing, because the player has been told: this is what both the press of the
    /// pause button and the loop's delivery of a hold owed by a relaunch come through.
    pub(super) fn held(&mut self, at: f64) {
        self.paused_at = Some(at);
        self.pending_hold = false;
        self.seeking = None;
    }

    /// The player was let go of at `at`, and is playing on from there.
    ///
    /// The start is *rebased* rather than kept, and that is the whole of the arithmetic here. The
    /// position of a file FFmpeg plays is this app's own clock over a moment it began a player
    /// (see `pin_playhead`), and a clock keeps running while the file is held — so leaving the
    /// start where it was would make the playhead jump forward by exactly the length of the hold
    /// the moment the file was let go, which is a bar whose own pause makes it lose its place.
    ///
    /// Rebasing onto the second the file was held at is exact rather than a fudge: the player was
    /// asked to hold and did, so it did not advance while it was held, and the second written down
    /// at that moment is the second it is at when the key arrives. What is measured from there is
    /// this app's clock over a stretch the player is actually playing, which is the best a player
    /// that reports nothing can be measured by (see `B4` in the handoff, on the drift that remains).
    ///
    /// A hold that was still owing is taken back here rather than delivered afterwards, which is
    /// the only answer a press can give: the file is playing on from here, so a pause that had not
    /// reached the player yet must never reach it.
    pub(super) fn released(&mut self, at: f64) {
        self.started = Some((Instant::now(), at));
        self.paused_at = None;
        self.pending_hold = false;
        self.seeking = None;
        // A film playing on is held by nothing, so whatever put the hold there is taken back: the
        // gesture that was holding it has either just let go of it — which is the whole of what
        // this write is — or a press of the pause button has, and a claim left standing here is a
        // claim the *next* film to be dragged is answered with (see `drag_held`).
        self.drag_held = false;
    }

    /// The player behind this transport is gone, and had got to `at`.
    ///
    /// What is dropped is the claim that one is running, and what is kept is the second it had
    /// reached — because a file nothing is playing is a file that can still be *started*, and a
    /// bar that forgot where the film had got to would offer to start it again from the
    /// beginning. A hold is kept as a hold rather than dropped for the same reason: the second
    /// is the only record of where the film stopped, and nothing else holds it.
    ///
    /// Nothing is left owing either, and this is the bound on a hold a relaunch could not deliver:
    /// there is no player to deliver it to, so the flag that said it was owed goes with the claim
    /// that a player was there to owe it to.
    ///
    /// This is the reconciliation, and it is deliberately written *through* the state rather
    /// than consulted beside it: the process is the only thing here that can be *observed*, and
    /// every other field is something this app asserted, so a disagreement is settled by
    /// overwriting the assertion (see `settle_pinned_transport`).
    pub(super) fn player_gone(&mut self, at: f64) {
        self.started = None;
        self.paused_at = Some(at);
        self.pending_hold = false;
        self.seeking = None;
        // The hold is kept and the claim over it is dropped, and the asymmetry is the point: the
        // second is the only record of where the film stopped, so nothing else holds it, but the
        // claim is about a player and there is no player left to hold. So a drag that ends after
        // this finds a film that nothing is holding rather than one it believes it is.
        self.drag_held = false;
    }

    /// The file was taken to `to`, by a seek taken while it was held.
    ///
    /// A player of this app's goes where it is taken and starts from there whatever it was doing,
    /// so the hold moves with it — and a hold left behind is a bar drawn at the second the pause
    /// began at while the film is somewhere else entirely. What is owed by the hold is untouched:
    /// the file is meant to be held at the new second just as much as at the old one.
    pub(super) fn sought(&mut self, to: f64) {
        if self.paused_at.is_some() {
            self.paused_at = Some(to);
        }
        self.seeking = None;
    }
}

/// Whether a transport bar behind a player that is `up` should be drawing a pause glyph.
///
/// Three claims have to agree before this app says a file is playing, and the process is the one
/// that is *observed* while the other two are things this app asserted and can therefore be
/// wrong about: a player that has been ended leaves `started` behind, and a player that has died
/// leaves both claims standing until something notices. So liveness is asked here rather than
/// believed, and a bar can never say a file is playing once nothing is playing it.
///
/// The engine's own answers are not asked this way — it reports its own position and its own
/// liveness, so there is nothing here to reconcile (see `pin_is_playing`).
pub(super) fn transport_playing(transport: &PinTransport, up: bool) -> bool {
    up && transport.started.is_some() && transport.paused_at.is_none()
}

/// The subtitle track a press of the next-track key lands on, and why it is a number rather
/// than a name.
///
/// FFmpeg's specifier numbers subtitle streams from zero *among themselves* — `s:0` is the
/// file's first subtitle stream whatever the video and audio streams around it are numbered —
/// which is exactly the numbering a cycling key wants, and exactly why this cannot be an index
/// into the file's whole stream list: that would be a different number for the same track
/// depending on how many audio tracks the container happens to carry.
///
/// Nothing is chosen for a file with no subtitle streams at all, rather than choosing its
/// zeroth, because `s:0` against a file with none is refused by the player and a refused
/// stream specifier is a player that exits instead of a player that plays the file without
/// subtitles. A file with one stream has a key that does nothing, which is the same answer a
/// key against a sound's card gives (see `pin_toggle_target`).
pub(super) fn next_subtitle(current: Option<usize>, streams: usize) -> Option<usize> {
    (streams > 0).then(|| current.map_or(0, |index| (index + 1) % streams))
}

/// Whether a kind is one the transport bar is drawn for.
pub(super) fn pin_transport_kind(kind: Option<MediaType>) -> bool {
    matches!(kind, Some(MediaType::Video) | Some(MediaType::NativeVideo))
}

/// Whether a kind is one whose transport bar carries controls that do something.
///
/// It is the same question as whether the bar is drawn at all, and it is asked as a separate
/// name because the two stopped being different at different times and for different reasons.
/// The bar was drawn for both kinds of video from the start, but only the engine's did anything:
/// FFmpeg's player takes no pause, reports no position, and can only be taken to another second by
/// being ended and begun again, so a bar for one was drawn without a button and with nothing to
/// drag — a read-out, because a button that does nothing is a promise the app cannot keep.
///
/// Two things changed that, and neither of them is "the player got better". The player takes a
/// pause as a key posted to its own window (see `ffplay_key_pause`), and the bar's button and
/// track both reach it now. So a bar that cannot be told anything would have to be some *third*
/// kind's — and there is none, because a sound's card carries its own controls on its own row and
/// everything else has no playback behind it at all.
pub(super) fn pin_transport_live(kind: Option<MediaType>) -> bool {
    pin_transport_kind(kind)
}

/// Whether a pinned window's chrome comes and goes with the pointer, which is true of the kinds
/// whose chrome is drawn *over* their media and of nothing else.
///
/// A strip in the way of the thing the window is for is what has to be asked for rather than always
/// there, and nothing else is: a kind whose chrome is in bands around its media keeps those bands
/// whatever the pointer is doing, and a sound has no chrome at all (see `pinned_caption_height`).
pub(super) fn pin_hides_chrome(kind: Option<MediaType>) -> bool {
    pin_overlay_chrome(kind)
}

/// The name a pinned window's caption button is saying out loud, and the button it belongs to.
///
/// A caption is drawn from glyphs, and a glyph is only a picture of an action. That is enough
/// for the window's own three — a hand has been reaching for a close button on every window it
/// has ever had — and not enough for the two that hand the file to another program, where the
/// one icon cannot say whether it is reaching for the program already chosen or asking the user
/// to choose. So a button says its name while the pointer rests on it, and the name is the only
/// part of this that is read from the machine rather than written here: which program opens a
/// PNG is an answer about the user's own associations, not a fact this app has.
///
/// The name is looked up once, when the pin is taken up, rather than per repaint: it is a
/// question about the registry, and a repaint happens sixty times a second. A file the machine
/// has nothing filed against has no name, which is a button whose default is the Shell's
/// fallback and no more.
#[derive(Clone, Default)]
pub(super) struct PinTooltip {
    /// The program the Shell would open the pinned file with, as it names it - empty where the
    /// machine has no answer to give.
    pub(super) default_app: String,
    /// The button the pointer is resting on, and when it got there. A name is written after a
    /// moment rather than at once, which is the delay every tooltip has and the only reason a
    /// pointer crossing a caption does not leave a trail of words.
    pub(super) button: Option<pin_chrome::CaptionButton>,
    pub(super) since: Option<Instant>,
    /// The button whose name is on the caption right now: `button` once the wait is over, and
    /// nothing before it. Kept so that asking again can tell a change worth a repaint from the
    /// same answer twice.
    pub(super) shown: Option<pin_chrome::CaptionButton>,
}

/// How long a pointer has to rest on a caption button before the button says what it is. Long
/// enough that crossing a caption does not leave words behind it, short enough that a hand
/// arriving and stopping has been told something.
pub(super) const PIN_TOOLTIP_DELAY: Duration = Duration::from_millis(500);

impl PinTooltip {
    /// Ask what the caption should be saying, answering whether that is a change — which is the
    /// whole of what a name costs, since a caption is either painting one or not.
    ///
    /// This is the one question about the caption that is asked on the loop's tick rather than
    /// on the pointer's, because it is a question about time: a name appears half a second
    /// after a pointer arrives and not one message before, and the tick is the only clock here
    /// that is running anyway. Asking it on a move instead would mean a name that appears only
    /// when the pointer next twitches.
    pub(super) fn refresh(
        &mut self,
        button: Option<pin_chrome::CaptionButton>,
        now: Instant,
    ) -> bool {
        if self.button != button {
            self.button = button;
            self.since = button.map(|_| now);
        }

        // What should be said now: the button the pointer is on, once it has been still long
        // enough, and only if that button has anything to say.
        let showing = self
            .button
            .filter(|button| self.text_for(*button).is_some())
            .filter(|_| {
                self.since
                    .is_some_and(|since| now.duration_since(since) >= PIN_TOOLTIP_DELAY)
            });

        // Compared before it is written, because a pointer leaving a button that was naming
        // itself is as much a change as one arriving at one that was not: the name that was on
        // the caption has to come off it, and a tick that saw both answers as "nothing" would
        // leave it there for ever.
        if showing == self.shown {
            return false;
        }

        self.shown = showing;
        true
    }

    /// What the button under the pointer would say, or nothing for a button that has no name
    /// to give — which is every button but the two that hand the file away. A button with a
    /// glyph that already says what it is does not need to say it again in words, and the
    /// default's own name is worth nothing where the machine has filed the format under nothing.
    pub(super) fn text_for(&self, button: pin_chrome::CaptionButton) -> Option<String> {
        match button {
            pin_chrome::CaptionButton::OpenWith if !self.default_app.is_empty() => {
                Some(format!("Open With {}", self.default_app))
            }
            pin_chrome::CaptionButton::OpenWithList => Some("Open With...".to_string()),
            _ => None,
        }
    }
}

/// The names the Shell has given for the formats this run has asked about, keyed by the file's
/// own format — its extension, or the whole of a name that begins with a dot (see
/// `shell_format`).
///
/// Which program opens a file is a fact about the machine and the format rather than about the
/// file, and asking it is Shell work with no bound on how long it takes. A pinned window's walk
/// asks for the name on every file it lands on — the name is read when the window is shown
/// another file, and a window is shown one per press of the caption's buttons — so the question
/// is asked once per format and answered out of this map after that, and the only step that
/// pays for it is the first of its kind in the run (see `default_app_name`).
pub(super) static APP_NAMES: Lazy<Mutex<HashMap<String, String>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// How many formats the table holds. A folder walk asks about the handful of formats it holds
/// rather than one per file, so this is generous for what the app does and small enough that a
/// session sweeping across a whole drive does not grow the table by a format per window — the
/// same bound `pin_navigation::FOLDER_LIMIT` puts on the folder lists beside it.
pub(super) const APP_NAMES_LIMIT: usize = 64;

/// The program the Shell would open this file with, as it names it, or nothing where the
/// machine has no association for the format at all.
///
/// The question is the machine's rather than the file's — two files of one format open in one
/// program — so it is asked once per format this run and answered out of `APP_NAMES` after
/// that. That is what a walk of a pinned window needs: asking the Shell about one file per
/// press is work nobody asked for, and it is work that can take as long as the Shell takes
/// (see `ask_default_app_name`).
pub(super) fn default_app_name(path: &Path) -> String {
    let format = shell_format(path);

    if let Ok(names) = APP_NAMES.lock() {
        if let Some(name) = names.get(&format) {
            return name.clone();
        }
    }

    // Safety: the argument is this app's own file, and the call allocates and frees its own
    // buffers (see `ask_default_app_name`).
    let name = unsafe { ask_default_app_name(path) };

    if let Ok(mut names) = APP_NAMES.lock() {
        if names.len() >= APP_NAMES_LIMIT {
            // Whichever format the table offers up, rather than the one least recently asked
            // about. The table is a cache of answers this run already has, so what is given
            // up is a name asked again later, not one still owed to a window on screen.
            if let Some(dropped) = names.keys().next().cloned() {
                names.remove(&dropped);
            }
        }
        names.insert(format, name.clone());
    }

    name
}

/// What a file's own name says its format is, which is what the Shell's answer is filed under.
///
/// It is the extension, spelled the way the filesystem reads one — which is to say a leading dot
/// is not an extension but the start of a name, so `.gitignore` is the format and nothing
/// follows it. A file with neither is keyed by its whole name rather than being filed under no
/// format at all: two extensionless files are not one format, and a shared empty slot between
/// them would put one file's association on another file's button.
pub(super) fn shell_format(path: &Path) -> String {
    let name = path
        .file_name()
        .map(|name| name.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    match path.extension() {
        Some(extension) => format!(".{}", extension.to_string_lossy().to_lowercase()),
        None => name,
    }
}

/// The association itself, asked of the Shell the once per format: see `default_app_name`.
///
/// `ASSOCSTR_FRIENDLYAPPNAME` is the association's own display name rather than the class it is
/// stored under, so what comes back is what the user would read in Explorer's "Open with" —
/// which is the whole point of naming the default on a button that opens it.
///
/// It is Shell work and it is not bounded, which is the whole of why it is asked once per
/// format rather than once per file, and never with the pin's own lock held (see the take-up in
/// `run_preview_window`). A machine with nothing filed against the format has no name to print,
/// which is a button that says nothing rather than a fault to report.
pub(super) unsafe fn ask_default_app_name(path: &Path) -> String {
    ask_association(path, ASSOCSTR_FRIENDLYAPPNAME).unwrap_or_default()
}

/// One string out of a file's association, as the machine has it filed — which is the same
/// call whatever string is asked for, and the same three ways of getting nothing back.
///
/// `what` is the question rather than an argument to be checked here: `ASSOCSTR`'s strings
/// are the Shell's own vocabulary for "which part of this association do you want", from
/// the executable to the icon index to a verb and its command line, and asking one of them
/// is asking the same registry the same way. Two of them are asked here rather than one —
/// the name a button is written with, and the handler a dialog's answer is read out of (see
/// `ask_default_handler`) — and the two are not a list to be kept in step, because each is
/// the only one of its kind and neither is a special case of the other.
///
/// Nothing when there is nothing: a file nothing is filed against, a format the machine has
/// no program for, and a path the Shell will not take are all one answer, and `None` is
/// what says that rather than an empty string pretending to be a name. A caller that wants
/// a name to print takes the empty string instead, and a caller comparing two readings
/// wants to know that one of them is nothing, which a pair of empty strings would not say.
///
/// # Safety
///
/// `AssocQueryStringW` is a plain `extern "system"` call whose every pointer is this
/// function's own: the two buffers are allocated here, sized by the first call from the
/// count it wrote, and both are dropped before the function returns. The wide string handed
/// in is the caller's file, which outlives the call.
pub(super) unsafe fn ask_association(path: &Path, what: ASSOCSTR) -> Option<String> {
    let wide: Vec<u16> = std::ffi::OsStr::new(&plain_path(path))
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();
    let file = PCWSTR(wide.as_ptr());

    let mut length = 0u32;
    let asked = AssocQueryStringW(
        ASSOCF_NONE,
        what,
        file,
        PCWSTR::null(),
        PWSTR::null(),
        &mut length,
    );
    if asked.is_err() || length == 0 {
        return None;
    }

    let mut buffer = vec![0u16; length as usize];
    let answered = AssocQueryStringW(
        ASSOCF_NONE,
        what,
        file,
        PCWSTR::null(),
        PWSTR(buffer.as_mut_ptr()),
        &mut length,
    );
    if answered.is_err() {
        return None;
    }

    buffer.truncate(
        buffer
            .iter()
            .position(|unit| *unit == 0)
            .unwrap_or(buffer.len()),
    );
    (!buffer.is_empty()).then(|| String::from_utf16_lossy(&buffer))
}

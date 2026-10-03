//! The box of a pinned window: how far a resize may take it, what a drag of one edge does to
//! the other three, and where a maximized pin is put back to.

use super::*;

/// How far a press on a pinned window may wander before it is a drag rather than a click.
pub(super) const PIN_DRAG_SLOP_PIXELS: f32 = 4.0;
/// How much of a pinned window stays on a display. A window dragged past an edge leaves
/// this much of itself behind, so its caption stays reachable and can drag it back — the
/// question Windows answers with a maximized window's own rules and this app has to answer
/// itself, because a captionless window can be dragged anywhere at all.
pub(super) const PIN_KEEP_ON_SCREEN_PIXELS: f32 = 64.0;
/// How wide a band along a pinned window's edge begins a resize: the eight pixels Windows gives a
/// resizable window's frame at 100% — a four-pixel frame with a four-pixel pad beside it.
pub(super) const PIN_RESIZE_BORDER_PIXELS: f32 = 8.0;
/// How much wider than the band a corner's own band is. A corner is the smallest target of the
/// eight and the one a hand aims at, and the top-right one shares its place with the caption's
/// buttons, so it is given the extra room Windows gives the corners of a window of its own.
pub(super) const PIN_RESIZE_CORNER_EXTRA_PIXELS: f32 = 4.0;
/// How small a resize can take the media a pin is showing: the window around it is never given a
/// smaller body than this, whichever edge is being dragged, so a window cannot be shrunk to
/// something with no room left to grab.
pub(super) const PIN_MIN_MEDIA_PIXELS: f32 = 48.0;
/// How often a pinned window with a transport bar is painted again while its file plays: the
/// playhead is a thing that moves on its own, and a quarter of a second is what a sound's card is
/// repainted at for the same reason (see `AUDIO_CARD_REPAINT`).
pub(super) const PIN_TRANSPORT_REPAINT_MS: u64 = 250;

/// What a pinned window's edges and caption do, which is a question about the kind of thing
/// inside it.
///
/// The rule under all three is the one thing a preview is never given: a bar. A picture, a
/// video, a rendered page — anything whose pixels are the file's own shape — is resized by
/// scaling the whole box, so its box can only ever be a box of its own shape and the media fills
/// it exactly. A page that is *laid out* to the box it is given, like a document, needs no such
/// rule: any box it is handed is a box it draws text into, so its edges move one at a time. And a
/// sound's card is neither: it is its own size, and a window around it would only be a window
/// with room in it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinFrame {
    /// Shaped: every edge and corner scales both sides, keeping the media's shape.
    Shaped,
    /// Laid out to any box: an edge resizes that edge alone.
    Free,
    /// Nothing to frame: no resize and no maximize.
    None,
}

pub(super) fn pin_frame(kind: Option<MediaType>) -> PinFrame {
    match kind {
        // A sound's card is its own size, and there is no box for it to grow into: it is moved by
        // its caption and its band, and it is not framed. A video is nothing like that, whoever is
        // playing it — its pixels are the file's own shape, so a box of another shape is a picture
        // stretched into a box it was not made for, which is the rule the other kinds are given
        // and the reason a video's edges scale both sides rather than one at a time.
        //
        // It used to be framed by nothing at all, on the reasoning that a window of somebody
        // else's is not this app's to resize. What that got right is that the *stretch* is
        // somebody else's window being asked to do something (see `relayout_pinned_media`); what
        // it got wrong was the conclusion, because the player's window can be moved and told its
        // size, and a relaunch begins one at whatever size it is asked for — so a video FFmpeg
        // plays is resized by its edges and maximized by its caption like any other picture, with
        // a relaunch at the end of the drag to render it sharply at the size it settled on.
        Some(MediaType::Audio) => PinFrame::None,
        Some(MediaType::Video) | Some(MediaType::NativeVideo) => PinFrame::Shaped,
        Some(MediaType::Text) | Some(MediaType::Archive) => PinFrame::Free,
        _ => PinFrame::Shaped,
    }
}

/// The kind of media on screen, which is what the transport's own questions are answered by.
pub(super) fn current_media_type() -> Option<MediaType> {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.media_type))
}

/// Put a pinned window back on the display it is on now, after the desktop was rearranged or a
/// display's scale changed: the box is kept where the display still has room for it and pulled
/// back into the work area where it does not, and the media is laid out again for the display
/// it ended up on.
pub(super) fn replace_pinned_window() -> Option<PreviewMessage> {
    let mut pinned = pin_state()?;
    let pin = pinned.pin_mut()?;

    pin.dpi = dpi_at(pin.content.0, pin.content.1);
    let bounds = work_area_at(pin.content.0, pin.content.1);
    let room = pinned_room(bounds, pin.dpi, pin.transport_bar, pin.overlay, pin.caption);
    let shape = (
        (pin.content.2 - pin.content.0).max(1) as u32,
        (pin.content.3 - pin.content.1).max(1) as u32,
    );
    let (width, height) = pinned_media_box(shape, room, PreviewScale::Percent(100));
    let content = clamp_pinned_box(
        (
            pin.content.0,
            pin.content.1,
            pin.content.0 + width,
            pin.content.1 + height,
        ),
        pin.dpi,
        &DESKTOPS,
    );

    pin.content = content;
    pin.restore = None;
    Some(PreviewMessage::PinBox(content))
}

/// Maximize a pinned window, or restore one that is maximized.
///
/// Maximizing is Fit to Screen: the media is given the largest box of its own shape the display's
/// work area has room for — scaled *up* to it as well as down, so a small picture fills the
/// screen rather than sitting at its own size in the middle of it — and the box is put in the
/// middle of that area, with the caption above it and the transport bar below it where the kind
/// has one. Nothing about it is clipped and nothing is stretched: what is left over when the
/// display's shape and the media's disagree is a margin at the sides or above and below, which is
/// what a maximized window of any shape has always done. For a kind that is laid out to its box
/// rather than scaled, the largest box is the room itself and the page is drawn into it.
///
/// Restoring down puts back the box the window had before it was maximized, place and size alike:
/// the plain undo the button is while nothing has been done to the window since. A window the hand
/// has *moved* — carried to another place or pulled to a size — is no longer a maximized window to
/// restore at all: that maximize is given up as the drag is made (see `pin_restore_box`), so the
/// caption draws a maximize where the restore glyph was, and this maximizes the box the hand has
/// left behind.
pub(super) fn toggle_pin_maximized(request: &mut Option<PreviewMessage>) {
    // What the pin has is read out under its lock and the lock is let go of before anything is
    // measured, which is the rule every other reader of a pin keeps: a measure is a file read,
    // and this is the thread that pumps the pinned window's own messages, so a read held under
    // the lock is a window that stops answering for as long as it takes (see `pin_swap_space`).
    let Some(PinMaximizeInputs {
        path,
        dpi,
        frame,
        overlay,
        transport_bar,
        caption,
        content,
        restore,
    }) = pin_maximize_inputs()
    else {
        return;
    };

    let bounds = work_area_at(content.0, content.1);
    let room = pinned_room(bounds, dpi, transport_bar, overlay, caption);
    let shape = media_dimensions(&path, bounds, dpi).filter(|shape| !box_is_the_wait(*shape));

    // The maximize as the pin stood when it was read, kept for the write below: the two are
    // the same question asked twice, and a walk that has landed elsewhere in between is a pin
    // whose restore has been taken or written since, which is what a changed value says.
    let maximize = restore;
    let (content, restore) = pin_maximize_decided(PinMaximize {
        frame,
        content,
        restore,
        shape,
        room,
    });

    let content = clamp_pinned_box(content, dpi, &DESKTOPS);

    // The check and the write are one lock and one step, because a walk taken between the read
    // above and this write has changed what the pin is showing, and a box decided from the file
    // it was on is not the box that file's window wants. The tick that answers the walk is
    // already queued, so the right answer is to leave it to that rather than to lay out over the
    // top of it — which is a decision not to write, so it is made beside the write and not
    // before it.
    let written = pin_state()
        .and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            if pin.path != path || pin.restore != maximize {
                return None;
            }

            pin.restore = restore;
            pin.content = content;
            Some(())
        })
        .is_some();

    if written {
        *request = Some(PreviewMessage::PinBox(content));
    }
}

/// What a maximize is decided from, once the file's own shape is known: the kind the pin is
/// showing, the box it is standing in, the box the maximize put aside if it has one, and the
/// room both are laid out in.
pub(super) struct PinMaximize {
    pub(super) frame: PinFrame,
    pub(super) content: ScreenRegion,
    pub(super) restore: Option<ScreenRegion>,
    pub(super) shape: Option<(u32, u32)>,
    pub(super) room: ScreenBounds,
}

/// The box the window is put in, and the box to remember for a restore: the whole of what a
/// press of the maximize button decides.
///
/// Told apart from the reading and the writing around it, because a toggle is a state machine
/// and this is its only transition: a window with a restore put aside is put back to it and
/// gives it up, and one without is maximized to the room and remembers the box it had. Which of
/// those it is must be read from the restore that came *in*, never from the one going out —
/// the two are never the same value, and a button that checks its answer against itself is a
/// button that does nothing.
pub(super) fn pin_maximize_decided(asked: PinMaximize) -> (ScreenRegion, Option<ScreenRegion>) {
    let PinMaximize {
        frame,
        content,
        restore,
        shape,
        room,
    } = asked;

    match restore {
        // The box kept for exactly this: the one the window had before the maximize, which is still
        // here only because no hand has moved the window since (see `pin_restore_box`).
        //
        // Put back as it was, unless the file on screen is not the one it was kept for — in
        // which case it is the box *this* file wants, at the size the user chose, rather than
        // the size and shape of a file that was here before the walk moved on. A window drawn
        // into a box of another file's shape is a stretched one, and a walk through shapes is
        // the surest way to get there.
        //
        // No shape to fit it to — a file whose measure is still running, or one with no shape
        // of its own — and the box is the only answer there is, and the user's own.
        Some(previous) => (
            shape.map_or(previous, |shape| pin_restored_box(shape, previous, room)),
            None,
        ),
        None => {
            let (width, height) = match frame {
                // A kind laid out to whatever box it is given has no shape of its own to fit,
                // so the room is what maximizing means for it.
                PinFrame::Free => (
                    (room.right - room.left).max(1),
                    (room.bottom - room.top).max(1),
                ),
                // A shaped kind is the file's own shape fitted to the room — read from the
                // file, and not from the box the window happens to be standing in, which is
                // already a fitted box and so a shape of nothing but its own container.
                _ => {
                    let shape = shape.unwrap_or_else(|| {
                        // Nothing to fit — a file whose measure is still out, or one with no
                        // shape of its own. The box on screen is the only shape there is.
                        (
                            (content.2 - content.0).max(1) as u32,
                            (content.3 - content.1).max(1) as u32,
                        )
                    });
                    pinned_media_box(shape, room, PreviewScale::FitToScreen)
                }
            };

            (centred_box((width, height), room), Some(content))
        }
    }
}

/// What a pin has to decide a maximize with, read out of it in one look: see `PinSwapSpace` for
/// why it is read rather than asked for piece by piece.
pub(super) struct PinMaximizeInputs {
    pub(super) path: PathBuf,
    pub(super) dpi: u32,
    pub(super) frame: PinFrame,
    pub(super) overlay: bool,
    pub(super) transport_bar: bool,
    /// The room the caption takes above the media, which the display's room has to lose beside the
    /// media itself (see `pinned_caption_height`).
    pub(super) caption: i32,
    pub(super) content: ScreenRegion,
    pub(super) restore: Option<ScreenRegion>,
}

pub(super) fn pin_maximize_inputs() -> Option<PinMaximizeInputs> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    Some(PinMaximizeInputs {
        path: pin.path.clone(),
        dpi: pin.dpi,
        frame: pin.frame,
        overlay: pin.overlay,
        transport_bar: pin.transport_bar,
        caption: pin.caption,
        content: pin.content,
        restore: pin.restore,
    })
}

/// The box a restore down puts back for a file of a given shape: the size the user had the
/// window at, in the middle of the room.
///
/// The size is the user's and is kept exactly — a restore undoes a maximize, it does not resize.
/// The *shape* is the file's, and is fitted into that size rather than assumed to match it,
/// because the box put aside was a box for whatever file was on screen when the maximize
/// happened, and a window that has since walked along the pin to a file of another shape would
/// otherwise be drawn into it and stretched. Fitting a shape into the chosen size is the same
/// rule every preview is fitted by, asked of a box rather than of a display (see
/// `pinned_media_box`).
///
/// A file taller than the size is given the full height and a width of its own rather than
/// being squashed, and one wider than it likewise: the size is the ceiling, not the shape.
pub(super) fn pin_restored_box(
    shape: (u32, u32),
    previous: ScreenRegion,
    room: ScreenBounds,
) -> ScreenRegion {
    let chosen = ScreenBounds {
        left: 0,
        top: 0,
        right: (previous.2 - previous.0).max(1),
        bottom: (previous.3 - previous.1).max(1),
    };

    centred_box(
        pinned_media_box(shape, chosen, PreviewScale::FitToScreen),
        room,
    )
}

/// The box a restore down owes once the hand has had a maximized window: the box the maximize put
/// aside, brought in step with the drag that has just been applied — or nothing at all, where the
/// hand has ended the maximize rather than moved it.
///
/// A drag that came out on the box it began with is no drag at all — a press nothing moved, or a
/// box already held against the room's own edge — and leaves the box the maximize put aside alone.
///
/// A hand that moved the box ends the maximize, whichever way it moved it. A window *carried* to
/// another place is no longer the one the button put across the room, and one *pulled to a size* has
/// had the size the maximize gave it replaced by one the hand asked for; neither is a maximize there
/// is anything left to undo. The box comes back as nothing, which is what leaves the caption drawing
/// a maximize where the restore glyph was and the button maximizing again (see
/// `toggle_pin_maximized`).
pub(super) fn pin_restore_box(
    restore: Option<ScreenRegion>,
    dragged_from: ScreenRegion,
    dragged_to: ScreenRegion,
) -> Option<ScreenRegion> {
    let restore = restore?;

    if dragged_from == dragged_to {
        return Some(restore);
    }

    None
}

/// A box of a given size in the middle of a room: what maximize puts the media in.
///
/// A size the room cannot hold goes against the room's own top-left corner rather than half past it:
/// there is no middle to be in when the box is bigger than the place it is being put in. Centring a
/// size on a *point* rather than on a room is the other half of the same question, and is what a
/// swap asks for when another file's shape is fitted where the pin already is (see `centred_at`).
pub(super) fn centred_box(size: (i32, i32), room: ScreenBounds) -> ScreenRegion {
    let (width, height) = (size.0.max(1), size.1.max(1));
    let left = room.left + ((room.right - room.left) - width).max(0) / 2;
    let top = room.top + ((room.bottom - room.top) - height).max(0) / 2;

    (left, top, left + width, top + height)
}

/// A box of a given size centred on a point: where a swap puts another file's shape, around the
/// middle of the box the pin occupies now — and, for the room a swap is fitted into before that, the
/// middle of the room (see `pin_swap_room` and `pin_update_box`).
///
/// What the point is near is not this function's business: a box left hanging off the edge of a
/// display by a middle near one is held against that display by the keep-on-screen rule the caller
/// applies, exactly as a dragged box is (see `clamp_pinned_box`).
pub(super) fn centred_at(size: (i32, i32), centre: (i32, i32)) -> ScreenRegion {
    let (width, height) = (size.0.max(1), size.1.max(1));
    let left = centre.0 - width / 2;
    let top = centre.1 - height / 2;

    (left, top, left + width, top + height)
}

/// The pin's own answers to the chrome, once a tick: its buttons, a key it was given, and whether
/// the media behind the pin is still there at all.
///
/// A step along the walk is answered with the walk itself rather than with a message,
/// because the walk is taken up where a pick is taken up (`pin_pick`) and nowhere else: a
/// `PinUpdate` message is the Explorer's half of the question, and the loop's own match for
/// one is a no-op — it is a pick that swaps a pin's file, and a button is a pick (see
/// `step_pinned_file`).
///
/// What comes back is more than the file, because a walk is not one file: a file the pin
/// cannot be shown is stepped over rather than stopped at, and what is stepped over is
/// bounded by the list the walk is made of (see `PinStep`).
///
/// A walk answered here is one the planner had already read the folder for. The walk a
/// caption button asks for is asked of the planner instead, and comes back later as its own
/// answer — which is what `wait` is for, the arc that says the pin is still asking (see
/// `step_pinned_file`).
///
/// A key is answered here rather than in the window procedure because what it means is a
/// question about the file on screen, and the player behind that file is the loop's: a Space
/// holds a video or a sound or lets it go, and which of the two players that is comes from the
/// kind in the window. Both of them are the loop's to answer — a hold of a sound this app plays
/// is a player ended and another begun rather than a state this window can set (see
/// `toggle_pinned_by_key`). The card is asked for at once rather than at the cadence it watches
/// the clock at, so that what the key did is on screen in the frame the key was pressed in.
pub(super) fn pin_command_request(
    request: &mut Option<PreviewMessage>,
    wait: &mut Option<PinWait>,
    audio_started: &mut Option<Instant>,
    audio_start_offset: &mut f64,
    audio_paused: &mut Option<f64>,
    navigating: bool,
) -> Option<PinStep> {
    let step = match take_pin_command() {
        Some(PinCommand::Close) => {
            *request = Some(end_pin_state(Reason::Closed));
            None
        }
        Some(PinCommand::Minimize) => {
            collapse_pin();
            None
        }
        Some(PinCommand::Restore) => {
            restore_pin();
            None
        }
        Some(PinCommand::Maximize) => {
            toggle_pin_maximized(request);
            None
        }
        Some(PinCommand::Previous) => step_pinned_file(-1, wait),
        Some(PinCommand::Next) => step_pinned_file(1, wait),
        Some(PinCommand::TogglePlayback) => {
            toggle_pinned_by_key(audio_started, audio_start_offset, audio_paused);
            None
        }
        Some(PinCommand::NextSubtitle) => {
            step_pinned_subtitle();
            None
        }
        None => None,
    };

    // The thing the pin is a window onto came apart: the player's process is gone, or the engine
    // took a document and never drew it. This is the half of "until it is closed, or it comes
    // apart" that is not a button (see `pin_media_is_alive`).
    if !pin_media_is_alive(navigating) {
        *request = Some(end_pin_state(Reason::MediaGone));
    }

    step
}

/// Let go of the pointer, if this window is the one holding it.
///
/// A press on a pinned window takes the pointer for the pin (`SetCapture` in `begin_pin_drag` and
/// in the three presses that are a drag under another name), and the release is what lets it go
/// again (see `pinned_release`). That makes the release the only thing standing between a press
/// and a window that eats every mouse message on the desktop, so whatever discards the state a
/// press was written into without a release coming has to let the pointer go here instead.
///
/// A pin rebuilt over another one is the case that matters, and it is an everyday one: a window
/// being resized is a drag, and a drag answers nothing until the hand lets go — but the file it is
/// a drag *of* is free to be replaced before the hand does (see `PreviewMessage::Pin`, and
/// `step_pinned_file` for the walk that asks for one). The take-up builds a whole new pin with no
/// drag in it, so the drag the pointer was taken for is gone, and nothing is left to release it.
/// A window still holding the pointer after its own drag has been taken away out from under it is a
/// window whose caption answers nothing: every click is delivered here rather than to whatever the
/// pointer was aimed at, and the cursor keeps whichever shape the last edge gave it. It ends when
/// some other window takes the pointer for itself, which is why clicking anywhere else appeared to
/// bring the window back.
pub(super) unsafe fn release_pin_capture(hwnd: HWND) {
    if GetCapture() == hwnd {
        let _ = ReleaseCapture();
    }
}

/// Carry a pinned window's drag on: the pointer has moved, and what the press began is applied to
/// the box the window had when it began.
pub(super) unsafe fn apply_pin_drag(hwnd: HWND) {
    let (drag, dpi, transport, overlay, frame, caption) = {
        let Some(pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin() else {
            return;
        };
        let Some(drag) = pin.dragging else {
            return;
        };

        (
            drag,
            pin.dpi,
            pin.transport_bar,
            pin.overlay,
            pin.frame,
            pin.caption,
        )
    };

    let Some(point) = cursor_screen_point() else {
        return;
    };
    // A drag is a function of where the pointer is, so asking twice for a pointer that has not
    // moved is asking the same question twice: a tick carrying a drag on asks this for the same
    // sixteen milliseconds the pointer messages do, and a resize answered from the tick is a
    // repaint of the whole window for a pointer that has not shifted.
    if drag.carried == point {
        return;
    }
    let (dx, dy) = (point.0 - drag.from.0, point.1 - drag.from.1);

    let window = match drag.action {
        PinDragAction::Move => (
            drag.window.0 + dx,
            drag.window.1 + dy,
            drag.window.2 + dx,
            drag.window.3 + dy,
        ),
        PinDragAction::Resize(edge) => resize_pinned_window(
            drag.window,
            edge,
            dx,
            dy,
            dpi,
            transport,
            overlay,
            frame,
            caption,
        ),
    };
    let window = clamp_pinned_box(window, dpi, &DESKTOPS);
    let content = content_box_of(window, dpi, transport, overlay, caption);

    {
        let Some(mut pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin_mut() else {
            return;
        };
        // A hand that pulled the box to a size is the box's new bound as well as its size: the
        // side the user dragged the window out to is the side the files that follow may use, and
        // it writes one where the pin had none — a size asked for by hand is the size, not the box
        // the file on screen came out at. A move is not a size and leaves it alone (see
        // `PinnedPreview::bound`).
        if matches!(drag.action, PinDragAction::Resize(_)) {
            pin.bound = Some((content.2 - content.0).max(content.3 - content.1).max(1));
        }
        // And a maximized window gives its restore up to any hand that has moved it, carried to
        // another place or pulled to a size of its own: a window the user has taken hold of is one
        // there is no maximize left to undo, so the caption's restore glyph goes back to a maximize
        // with it and the button maximizes again (see `pin_restore_box`).
        pin.restore = pin_restore_box(pin.restore, pin.content, content);
        pin.content = content;
        // And where it was carried to, so that the next ask for a pointer that has not moved is
        // recognised as the same question (see the early return above).
        if let Some(drag) = pin.dragging.as_mut() {
            drag.carried = point;
        }
    }

    let (width, height) = (window.2 - window.0, window.3 - window.1);

    // A window that was carried has nothing to paint: what a layered window is drawn from is the
    // surface it already has, and moving one moves that surface with it — so a drag is a
    // `SetWindowPos` per pointer move and nothing else, which is the whole of why a pinned window
    // follows a hand at the pace of the pointer rather than at the pace of a repaint of a
    // display's worth of pixels.
    //
    // A resize is the drag that has a picture to draw, because the box it is being dragged to is
    // a band the frame it holds is not the size of (see `compose_media_into_band`) — and that
    // picture is the reason the new box is handed to `UpdateLayeredWindow` rather than to
    // `SetWindowPos`: one call applies the frame, its place, and the window's size together, so
    // there is no moment in which the surface of the box the drag came from stands at the place
    // of the box it is being dragged to. Sizing the window first and painting after it is that
    // moment once per pointer move, and it is seen as the whole of the window's contents shaking:
    // the surface carries the old origin and the window the new one, so the difference between
    // the two is drawn as the picture jumping — which is every edge and corner but the
    // bottom-right one, where the origin does not move at all and there is nothing to jump (see
    // `render_layered_preview_at`).
    if matches!(drag.action, PinDragAction::Resize(_)) {
        render_pinned_preview_at(hwnd, window.0, window.1);

        // The box is already the one the paint applied; what is re-asserted here is only the
        // place in the z-order a carried window gets from its own `SetWindowPos`.
        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    } else {
        // Kept in its band without touching the z-order: the order (engine above preview) is
        // established where the pin comes up, and re-asserting TOPMOST per move fights the
        // engine's own placement for the top, one DWM reorder each (see `Host::place`).
        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            window.0,
            window.1,
            width,
            height,
            SWP_NOACTIVATE | SWP_NOZORDER,
        );
    }

    // Whatever stands in the media band travels with it — after the box it stands in is the
    // window's, so that a player's window is never put where the surface has not caught up.
    place_pinned_siblings();
}

/// The window a resize drag has produced: the box the drag began on, dragged by one of its edges,
/// kept inside the room of the display that box is standing on (see `pinned_room`).
#[allow(clippy::too_many_arguments)] // Each one is a distinct fact about the pin the box is of.
pub(super) fn resize_pinned_window(
    window: ScreenRegion,
    edge: PinResize,
    dx: i32,
    dy: i32,
    dpi: u32,
    transport: bool,
    overlay: bool,
    frame: PinFrame,
    caption: i32,
) -> ScreenRegion {
    let content = content_box_of(window, dpi, transport, overlay, caption);
    let bounds = work_area_at(
        content.0 + (content.2 - content.0) / 2,
        content.1 + (content.3 - content.1) / 2,
    );

    resize_pinned_content(
        PinSpace {
            content,
            room: pinned_room(bounds, dpi, transport, overlay, caption),
            overlay,
            transport,
            caption,
        },
        edge,
        dx,
        dy,
        dpi,
        frame,
    )
}

/// A media box and the room it may be dragged within: the two boxes the geometry of a resize is
/// arithmetic on. They travel together because they are the same two numbers for the drag that
/// reads a display and for the tests that fix one — see `resize_pinned_content`.
pub(super) struct PinSpace {
    pub(super) content: ScreenRegion,
    pub(super) room: ScreenBounds,
    /// Whether the chrome of the pin being dragged is drawn over its media, which is the one fact
    /// the box a drag *comes out with* needs and neither of the other two carries: a media box is
    /// the window's for such a kind and has a caption and a bar to be given room around it
    /// otherwise (see `pinned_window_box_of`).
    pub(super) overlay: bool,
    /// Whether a transport strip is taken off the bottom, and the room taken off the top for a
    /// caption: the other two band facts, which travel with the overlay flag because the box a
    /// drag comes out with is the window's and the window's is all three of them.
    pub(super) transport: bool,
    pub(super) caption: i32,
}

/// The same drag, against a room to keep the box inside rather than against the display the box is
/// standing on: the whole of the geometry, with nothing of the desk in it.
///
/// What the media does with a box of another shape is the whole of the rule (see `PinFrame`). A
/// kind whose pixels are their own shape keeps that shape whatever edge is dragged: the edge under
/// the hand is the one the box follows and the other side is worked out from the shape, so what
/// the window is given is always a box the media fills exactly — never a picture with a band of
/// nothing beside it, and never one pulled out of shape. A kind that is laid out to its box takes
/// the drag one dimension at a time, which is what a page of text wants: a wider window is a
/// longer line, not a bigger letter.
///
/// Which part of the window stays where it is, is the other half of it, and it is the half that
/// makes a window grow away from the hand rather than under it. The edge opposite the one being
/// dragged is the one that is nailed down, and for a side the other axis is left centered on the
/// line it was on — so a window pulled wider grows around its own middle rather than downwards. A
/// corner nails down the corner opposite it, and the two answers the shape allows for a corner —
/// one worked out from the width the hand asked for, one from the height — are compared, and the
/// nearer of the two is the one taken.
///
/// What a drag cannot do is take the media past the room, or below the size a hand can still take
/// hold of (see `PIN_MIN_MEDIA_PIXELS`). A window bigger than the room it is on is a window whose
/// edges cannot be reached, and the media is drawn at the size it is shown at, so there is nothing
/// a box past the room would buy.
pub(super) fn resize_pinned_content(
    space: PinSpace,
    edge: PinResize,
    dx: i32,
    dy: i32,
    dpi: u32,
    frame: PinFrame,
) -> ScreenRegion {
    let PinSpace {
        content,
        room,
        overlay,
        transport,
        caption,
    } = space;
    let start_width = (content.2 - content.0).max(1);
    let start_height = (content.3 - content.1).max(1);

    let floor = logical_px(dpi, PIN_MIN_MEDIA_PIXELS).max(8);

    // What the room leaves the box to grow into. An axis the hand is on has an edge that stays
    // where it is, so it may grow only as far as the room reaches past that edge; an axis the hand
    // is not on is only centered on the line it was on, and its limit is the room itself — a box
    // that cannot be centered where it was put is a box moved along that line, which the placement
    // below does and which is the one thing a drag is never refused for.
    let (to_the_left, to_the_right) = (content.0 - room.left, room.right - content.2);
    let (above, below) = (content.1 - room.top, room.bottom - content.3);

    let to_grow = if edge.left {
        start_width + to_the_left
    } else if edge.right {
        start_width + to_the_right
    } else {
        room.right - room.left
    };
    let to_grow_up = if edge.top {
        start_height + above
    } else if edge.bottom {
        start_height + below
    } else {
        room.bottom - room.top
    };

    let ceiling = (
        to_grow.clamp(floor, (room.right - room.left).max(floor)),
        to_grow_up.clamp(floor, (room.bottom - room.top).max(floor)),
    );

    // What the hand asked for: the box the press began on, with the distance the pointer has
    // moved added to or taken off each side it is on.
    let wanted = (
        start_width + if edge.left { -dx } else { dx },
        start_height + if edge.top { -dy } else { dy },
    );

    let (width, height) = if frame == PinFrame::Free {
        (
            wanted.0.clamp(floor, ceiling.0),
            wanted.1.clamp(floor, ceiling.1),
        )
    } else {
        let aspect = start_width as f64 / start_height as f64;

        // The box the shape allows for a width, and for a height. Both ends of each are clamped
        // before the other side is worked out, so the box that comes out is inside the room in
        // both directions and is the shape it was asked to be.
        let from_width = |width: f64| -> (f64, f64) {
            let width = width.clamp(floor as f64, ceiling.0 as f64);
            let height = (width / aspect).clamp(floor as f64, ceiling.1 as f64);
            (height * aspect, height)
        };
        let from_height = |height: f64| -> (f64, f64) {
            let height = height.clamp(floor as f64, ceiling.1 as f64);
            let width = (height * aspect).clamp(floor as f64, ceiling.0 as f64);
            (width, width / aspect)
        };

        let (width, height) = match (edge.horizontal(), edge.vertical()) {
            (true, false) => from_width(wanted.0 as f64),
            (false, true) => from_height(wanted.1 as f64),
            _ => {
                let (by_width, by_height) =
                    (from_width(wanted.0 as f64), from_height(wanted.1 as f64));
                let error = |(width, height): (f64, f64)| {
                    (width - wanted.0 as f64).powi(2) + (height - wanted.1 as f64).powi(2)
                };

                if error(by_width) <= error(by_height) {
                    by_width
                } else {
                    by_height
                }
            }
        };

        (
            (width.round() as i32).clamp(floor, ceiling.0),
            (height.round() as i32).clamp(floor, ceiling.1),
        )
    };

    // Where the box comes out. The edge opposite the one being dragged is the one that stays; an
    // axis the hand is not on is centered on the line it was on, and moved along that line by as
    // little as the room asks rather than by the whole of what the centering wanted.
    //
    // A box the drag has already carried past the room's own edge is the exception, and it is the
    // whole of what a maximized window is: it fills the room in one dimension, so there is no room
    // left on that line for the centering to sit in, and the clamp that would have moved it back
    // has nothing to move it back *to* — it collapses onto the room's edge instead. A maximized
    // window dragged down and then pulled by an edge is the case, and the window jumps to the top
    // or the left of the screen the moment the edge is touched, resizing from the place the screen
    // says rather than the one the hand left it at. A box already off the room keeps the line the
    // drag gave it: the room cannot hold it either way, so there is nothing to be gained by
    // pretending it can (and the window is still kept on the display by `clamp_pinned_box`, which
    // is the rule for a box that must stay reachable at all).
    let off_the_room_horizontally = content.0 < room.left || content.2 > room.right;
    let off_the_room_vertically = content.1 < room.top || content.3 > room.bottom;

    let left = if edge.left {
        content.2 - width
    } else if edge.right || off_the_room_horizontally {
        content.0
    } else {
        (content.0 + (start_width - width) / 2)
            .clamp(room.left, (room.right - width).max(room.left))
    };
    let top = if edge.top {
        content.3 - height
    } else if edge.bottom || off_the_room_vertically {
        content.1
    } else {
        (content.1 + (start_height - height) / 2)
            .clamp(room.top, (room.bottom - height).max(room.top))
    };

    // And the box that comes out of it is the window's, which is the media's own for a kind whose
    // chrome is drawn over it and the media's grown by the two bands otherwise — the one place the
    // difference is not readable from the numbers either side of it, since a drag works in the
    // media's box from beginning to end and what it hands back is the window's.
    pinned_window_box_of(
        (left, top, left + width, top + height),
        dpi,
        transport,
        overlay,
        caption,
    )
}

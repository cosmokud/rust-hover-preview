//! The wheel's own view of a pinned picture: how large it is drawn and where it sits.
//!
//! A pinned picture's window *is* its media — the chrome is drawn over it rather than in bands
//! around it (see `pin_overlay_chrome`) — so a zoom has two stages and one number between them.
//! The window grows about the picture's centre until it reaches the size `Fit to Screen` would give
//! that picture, and stops there; what a notch asks for beyond that is the picture growing inside a
//! window that has stopped. Both stages are the same arithmetic: the size the picture is *drawn* at
//! is taken one step further and the window is the smaller of that and the cap, so the share between
//! the two is one until the cap is reached and more than one after it — and walking back out
//! consumes the share again before the window starts shrinking.
//!
//! A picture already larger than a fit — a huge photograph taken up at `100%`, or a box a hand has
//! dragged out past it — is given a cap of its own size rather than of the fit's, because the cap is
//! there to stop a window growing and never to shrink one (see `pin_zoom_after`).

use super::*;

/// What one wheel notch multiplies a pinned picture by.
///
/// Out is this step divided rather than a step of its own, so a notch in and a notch out leave the
/// picture exactly where it was — which is the whole of what "zoom in and zoom out" has to mean if a
/// hand is to come back from either end without counting notches.
pub(super) const PIN_ZOOM_PER_NOTCH: f32 = 1.10;

/// The share of the display's room a picture may be zoomed down to, in either direction: the hard
/// limit a zoom out stops at.
///
/// It is a share of the room rather than a size in logical pixels because the limit is about the
/// screen rather than about what is on it: a picture taken to a hundredth of a display is a mark that
/// can still be found and grown again, and a limit in pixels would be a different one on every
/// monitor.
pub(super) const PIN_ZOOM_FLOOR_SHARE: f32 = 0.01;

/// The largest a picture may be *drawn* at, which is not a limit on a zoom in but on the arithmetic:
/// a step is a power of a count the wheel hook accumulates, and a single wheel message can carry a
/// delta of thirty thousand notches, so a device — or a driver — sending one would ask for a size no
/// number holds.
///
/// What it is set to keep is the *placement*: the box a picture lands in is a whole-pixel rectangle
/// on a screen, and a drawn width at a billion is a rectangle its half still fits in an `i32` with
/// room to spare. Nothing can be looking at a picture four hundred million pixels wide, so the
/// ceiling is past every size any eye or any engine will ever be shown.
pub(super) const PIN_ZOOM_CEILING: f32 = 1.0e9;

/// How a pinned picture is drawn inside its window, or `None` for a kind the wheel does not zoom.
///
/// It is a field of the pin rather than a local of the paint because it outlives both: a picture
/// zoomed into its window keeps that view across every repaint, and it is what the *hit tests* are
/// answered against as well — whether a press on the media carries the window or the picture is this.
#[derive(Clone, Copy, PartialEq, Debug)]
pub(super) struct PinZoom {
    /// The size the picture is drawn at, in device pixels: the window's own size while the picture
    /// fills it, and larger once the window has stopped growing and the picture has not.
    ///
    /// It is a size of the *picture's* rather than of the window's, and it is the number the whole
    /// of the arithmetic walks. A step is this multiplied out, which is what keeps a picture's shape
    /// exactly the shape it started with however many notches a hand turns: a window's own box is
    /// whole pixels on a screen and a picture's shape is not, so a zoom that re-derived its shape
    /// from a rounded window would walk it out of true over a long scroll.
    pub(super) drawn: (f32, f32),
    /// Where the picture's centre sits inside the window, in device pixels, measured from the
    /// window's own centre. Nothing until a hand has carried it, and nothing at all while the picture
    /// is the window's own size — a picture that size has no room to be moved into.
    pub(super) pan: (f32, f32),
}

impl PinZoom {
    /// The picture filling its window: where every pin of a kind the wheel zooms starts, and where a
    /// box that is laid out again puts it back (see `reset_pin_zoom`).
    pub(super) fn filling(content: ScreenRegion) -> Self {
        let (width, height) = region_size(content);

        Self {
            drawn: (width as f32, height as f32),
            pan: (0.0, 0.0),
        }
    }

    /// The shape the picture is drawn at: the one thing about it a zoom never changes, and the reason
    /// the drawn size is carried in floats rather than worked out afresh from the window each notch.
    pub(super) fn shape(&self) -> f32 {
        self.drawn.1 / self.drawn.0
    }

    /// How many windows wide the picture is drawn: one while it is the window's own size, and more
    /// than one once the window has been taken as far as it may go.
    pub(super) fn share_of(&self, content: ScreenRegion) -> f32 {
        self.drawn.0 / region_size(content).0 as f32
    }

    /// The size the picture is drawn at, in whole pixels — which is what a paint and a crop are both
    /// asked for, and not what the arithmetic walks (see `PinZoom::drawn`).
    fn drawn_box(&self) -> (i32, i32) {
        drawn_pixels(self.drawn)
    }

    /// Where the picture lands inside its window: the box in the window's own coordinates, which is
    /// the whole of what a repaint and a crop both need.
    ///
    /// A picture the size of its window has nowhere to be but the whole of it, and one drawn larger
    /// is centred and then carried by the pan — so the box is the window's own wherever the picture
    /// fits, whatever the pan remembers of a larger one, and the pan only ever moves a picture with
    /// room to move in. The pan is the picture's centre's own offset from the window's, so the box
    /// steps the way the pan names: a hand carried to the right shows the left of the picture, which
    /// is a box that steps left.
    pub(super) fn picture_box(&self, window: (i32, i32)) -> (i32, i32, i32, i32) {
        let (width, height) = (window.0.max(1), window.1.max(1));
        let (drawn_width, drawn_height) = self.drawn_box();
        if drawn_width <= width && drawn_height <= height {
            return (0, 0, width, height);
        }

        let left = (width - drawn_width) / 2 + self.pan.0.round() as i32;
        let top = (height - drawn_height) / 2 + self.pan.1.round() as i32;

        (left, top, left + drawn_width, top + drawn_height)
    }
}

/// The box the window ends up with after a wheel notch, and the view the picture is left with.
pub(super) struct PinZoomPlan {
    pub(super) content: ScreenRegion,
    pub(super) zoom: PinZoom,
}

/// Everything a notch is decided from, read out of a pin in one look: the box the picture is drawn
/// in, the view the last notch left, the display it is on, and where the hand is.
///
/// The display's room is read at the middle of the *window* rather than at a corner of it, because
/// the middle is what the pair asks for: a window half over the seam of two displays is not one
/// either of them can be said to own by a corner (see `reset_pinned_size`).
pub(super) struct PinZoomInputs {
    pub(super) content: ScreenRegion,
    pub(super) zoom: PinZoom,
    pub(super) dpi: u32,
    pub(super) room: ScreenBounds,
    /// Where the pointer is on the screen, if it is anywhere: the anchor every notch after the cap is
    /// measured from, and `None` — which is a hand that has left, or a machine with no pointer to
    /// read — is a notch that keeps the picture's centre instead.
    pub(super) pointer: Option<(i32, i32)>,
}

/// Take a pinned picture one wheel's worth further in or out, or nothing where the notch asks for
/// nothing: the box the window ends up with and how the picture is drawn inside it.
///
/// `notches` is in wheel units turned into notches and is signed, up being positive — a positive
/// notch zooms in. More than one notch at once is the same gesture taken further: the step is the one
/// notch multiplied out, so a fast wheel and a slow one land on the same sizes.
///
/// The cap, the floor and the share between the picture and its window are all here and nothing of
/// them is in the window calls, because this is the whole of what a zoom decides: a caller that
/// placed a box of its own would be a second opinion about the size a picture is drawn at, which is
/// what the tests of this function exist to prevent.
pub(super) fn pin_zoom_after(
    input: &PinZoomInputs,
    notches: i32,
    displays: &dyn Displays,
) -> Option<PinZoomPlan> {
    if notches == 0 {
        return None;
    }

    let step = PIN_ZOOM_PER_NOTCH.powi(notches);
    let content = input.content;
    let (width, height) = region_size(content);
    let shape = input.zoom.shape();

    // The picture one step further out, and the floor it may not be taken under: the display's own
    // room at a hundredth of itself, in *either* direction. A picture has a shape, so the two
    // directions are asked together and it is the wider of the two shares that is the floor — the one
    // that keeps both sides of it above the limit rather than only the one the room happened to be
    // proportioned for.
    //
    // Never more than the picture already is: the floor is there to stop a zoom out and not to take
    // one back, so a box taken up already below it is left where it is rather than grown by the notch
    // that asked to shrink it.
    // shortcut: that leaves such a pin below the floor until it has been grown and shrunk again —
    // which needs a display wider than six thousand pixels before the 48-pixel floor a resize keeps
    // stops covering it; upgrade if a display that wide is seen.
    let (room_width, room_height) = input.room.room();
    let floor = (PIN_ZOOM_FLOOR_SHARE * (room_width as f32).max(room_height as f32 / shape))
        .min(input.zoom.drawn.0);

    let drawn_width = (input.zoom.drawn.0 * step).max(floor).min(PIN_ZOOM_CEILING);
    let drawn = (drawn_width, drawn_width * shape);

    // The largest the window may grow to: the picture's own width at the size `Fit to Screen` would
    // give it — the whole of what "a maximum the size of a fit" means, and read off the function the
    // maximize button reads it off so that the two are one rule.
    //
    // And a window already past that is given a cap of its own size instead, which is what a picture
    // taken up larger than a fit is left alone by: the cap's whole business is to stop a window
    // *growing*, and a rule that also shrank one would turn the first notch of a zoom into a picture
    // the hand did not ask for. It applies to growth alone, so a window already past the fit still
    // comes down when the wheel is turned the other way.
    let (fit_width, _) = pinned_media_box(
        (width.max(1) as u32, height.max(1) as u32),
        input.room,
        PreviewScale::FitToScreen,
    );
    let cap = if width as f32 >= fit_width as f32 {
        width as f32
    } else {
        fit_width.max(1) as f32
    };

    // The window is the smaller of what the picture asks for and the cap, which is the whole of the
    // two stages in one line: while the picture is under the cap the window is the picture, and once
    // it is over it the window is the cap and the share below is what the picture overhangs by.
    //
    // The height is the width through the picture's shape rather than a clamp of its own, so that the
    // two stages are decided by one number: two clamps would be two chances for the window to come
    // out of a shape no picture ever had.
    let window_width = drawn_width.min(cap).max(1.0);
    let size = (
        window_width.round().max(1.0) as i32,
        (window_width * shape).round().max(1.0) as i32,
    );

    // The picture's centre as it stands, in the screen's own coordinates, and where the window is put
    // around it: the centre is what a zoom holds still, and a window twice the size about its own
    // middle is a picture that grew in every direction by the same amount where one grown about a
    // corner would walk off the screen as it grew.
    //
    // A window that is *not* changing size keeps the box it has, whatever the pan is doing: a picture
    // carried inside a window that has stopped is the picture moving and not the window, and a window
    // re-placed about a picture's carried centre is a window that walks across the screen as the hand
    // drags the picture in it.
    let drawn_centre = pinned_centre(content, input.zoom.pan);
    let placed = if size == (width, height) {
        content
    } else {
        clamp_pinned_box(centred_at(size, drawn_centre), input.dpi, displays)
    };
    let box_centre = ((placed.0 + placed.2) / 2, (placed.1 + placed.3) / 2);

    // And where the picture sits inside it. A picture the window's own size fills it and has nowhere
    // to sit, whatever it was carrying before a zoom out brought it back to that size; one drawn
    // larger is anchored on the pointer, which is what lets a hand zoom into a corner it has pointed
    // at.
    //
    // The anchor is the point of the picture under the pointer staying under the pointer, and it is
    // measured in the window's own shares rather than in pixels — which is the whole of why the
    // overhang is used up before any of the window is: a notch that took the window out to the cap in
    // the same breath as it took the picture past it moves the picture by exactly as much of the
    // overhang as the window left room for, and never by more, so the picture cannot be pulled off a
    // window it was covering a moment ago.
    let share = input.zoom.share_of(content);
    let pan = match input.pointer.filter(|_| drawn.0 > size.0 as f32) {
        Some(point) => {
            let kept = (drawn.0 / size.0 as f32) / share;
            let (dx, dy) = (
                (point.0 - drawn_centre.0) as f32,
                (point.1 - drawn_centre.1) as f32,
            );

            (
                dx * (1.0 - kept) + (drawn_centre.0 - box_centre.0) as f32,
                dy * (1.0 - kept) + (drawn_centre.1 - box_centre.1) as f32,
            )
        }
        None => (0.0, 0.0),
    };

    // And held to as much of the picture as the window has room for, which a hand inside the window
    // can never ask for more of but a picture the hand has already carried off can: a zoom *places* a
    // picture in its window, and where it may be placed is a box that is at most the whole of the
    // overhang to one side. Carrying one further off than that is the drag's own business, so a
    // picture a hand has put out of reach is one the next notch brings back into it.
    //
    // The overhang is the whole pixels the picture and the window are drawn at and the same halving
    // `picture_box` places a carried picture by, so that the two cannot disagree by the half pixel a
    // rounded pair of floats leaves behind — which is the difference between a picture that covers its
    // window and one with a hairline of backdrop down one side of it.
    let (drawn_width, drawn_height) = drawn_pixels(drawn);
    let overhang = (
        ((drawn_width - size.0).max(0) / 2) as f32,
        ((drawn_height - size.1).max(0) / 2) as f32,
    );

    Some(PinZoomPlan {
        content: placed,
        zoom: PinZoom {
            drawn,
            pan: (
                pan.0.clamp(-overhang.0, overhang.0),
                pan.1.clamp(-overhang.1, overhang.1),
            ),
        },
    })
}

/// Take a pinned picture one wheel's worth further in or out and draw it there, from the wheel hook's
/// own count of notches.
///
/// It is the window procedure's road rather than the loop's for the reason a drag is: the notches are
/// read on a hook that every wheel message on the desktop waits for, so what they come to is written
/// and drawn here rather than handed to the thread that owns the media — and a picture is drawn from
/// the frame the window already holds, so there is nothing here the loop owns.
pub(super) unsafe fn apply_pin_zoom(hwnd: HWND, notches: i32) {
    let input = {
        let Some(pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin() else {
            return;
        };
        // A collapsed pin has no window on the screen to point at or to draw, and a drag already
        // under the hand is a box the hand is carrying: a wheel notch landing in the middle of one
        // would be two answers to the same question.
        let Some(zoom) = pin
            .zoom
            .filter(|_| !pin.collapsed && pin.dragging.is_none())
        else {
            return;
        };

        let window = pin.window_box();
        PinZoomInputs {
            content: pin.content,
            zoom,
            dpi: pin.dpi,
            room: work_area_at((window.0 + window.2) / 2, (window.1 + window.3) / 2),
            pointer: cursor_screen_point(),
        }
    };

    let Some(plan) = pin_zoom_after(&input, notches, &DESKTOPS) else {
        return;
    };

    // The reading and the write are one lock and one step, as they are for a resize: a walk taken
    // between the two has put another file in the window, and a box decided from the file that was
    // there is not the box this one wants (see `reset_pinned_size`). The tick that answers such a
    // walk is already queued, so the right answer here is to leave it to that rather than to lay out
    // over the top of it — which is a decision not to write, so it is made beside the write.
    let written = pin_state()
        .and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            if pin.content != input.content || pin.zoom != Some(input.zoom) {
                return None;
            }

            // A window the wheel has taken hold of has given up whatever maximize it was carrying,
            // for the reason a dragged one has: a maximize is the button's own arrangement of the box,
            // and this is another one (see `pin_restore_box`).
            pin.restore = pin_restore_box(pin.restore, pin.content, plan.content);
            pin.content = plan.content;
            pin.zoom = Some(plan.zoom);
            Some(())
        })
        .is_some();

    if written {
        // The frame, its place and the window's size in one call, exactly as the drag of an edge does
        // it: a window sized first and painted after is a moment in which the surface holds the old
        // box and the window the new one, which an eye reads as the picture jumping (see
        // `apply_pin_drag`).
        render_pinned_preview_at(hwnd, plan.content.0, plan.content.1);
    }
}

/// Put a pinned picture back to filling its window, for the kinds the wheel zooms and nothing at all
/// for any other.
///
/// It is what every road that lays a box out again owes: a walk to another file, a maximize or a
/// restore, the re-centre a double-click asks for, and an edge dragged to a new size are each a box
/// chosen for the whole of a window rather than for a picture inside one, and a picture left zoomed
/// into a box of another size is a picture hanging off it.
pub(super) fn reset_pin_zoom(pin: &mut PinnedPreview) {
    if pin.zoom.is_some() {
        pin.zoom = Some(PinZoom::filling(pin.content));
    }
}

/// The size a drawn picture is whole pixels at: what a paint, a crop and a pan's own limit are all
/// asked for, and not what the arithmetic walks (see `PinZoom::drawn`).
pub(super) fn drawn_pixels(drawn: (f32, f32)) -> (i32, i32) {
    (
        drawn.0.round().max(1.0) as i32,
        drawn.1.round().max(1.0) as i32,
    )
}

/// Where the picture's centre sits on the screen: a box's middle carried by the pan.
pub(super) fn pinned_centre(content: ScreenRegion, pan: (f32, f32)) -> (i32, i32) {
    (
        (content.0 + content.2) / 2 + pan.0.round() as i32,
        (content.1 + content.3) / 2 + pan.1.round() as i32,
    )
}

/// The size of a screen region, width first, and never less than one pixel in either direction: a box
/// with no room in it is a divide waiting to happen on every road this app measures one.
pub(super) fn region_size(region: ScreenRegion) -> (i32, i32) {
    ((region.2 - region.0).max(1), (region.3 - region.1).max(1))
}

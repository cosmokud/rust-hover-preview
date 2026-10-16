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

/// How fast a zoom travels to the size the wheel asked for: the rate of the spring the picture's
/// drawn size is carried by, in radians a second, doubled into the spring's own rate below.
///
/// It is a rate rather than a length of time because the thing it paces is a *travel* and not a
/// curve: a hand that turns the wheel again while the last notch is still on its way does not
/// restart anything, it only moves where the travel is going — so what a fast wheel and a slow one
/// share is this number and not a duration either of them was given.
///
/// Twenty is a step that has visibly finished in about a fifth of a second, which is long enough to
/// be seen as a movement and short enough that the picture is never behind the hand.
///
pub(super) const PIN_ZOOM_SPRING_SPEED: f32 = 20.0;

/// How near the size it was asked for a zoom has to come, and how slowly, before it is *there*: a
/// quarter of a pixel, and four pixels a second.
///
/// A spring closes on what it was asked for and never quite arrives, and what that costs is not the
/// picture — a quarter of a pixel is not a pixel — but the *window*, which is whole pixels on a
/// screen and would otherwise be given a `SetWindowPos` and a repaint on every frame for the rest of
/// the pin's life. The pair is also what ends an animation, and therefore what says when a picture
/// has stopped long enough to be worth laying out again at the size it has settled on (see
/// `pin_zoom_step`).
///
pub(super) const PIN_ZOOM_SETTLED_PIXELS: f32 = 0.25;
pub(super) const PIN_ZOOM_SETTLED_SPEED: f32 = 4.0;

/// The longest a single frame of a travel is taken to have lasted, in seconds.
///
/// A spring is arithmetic on how much time has passed, and the loop that takes its frames wakes on
/// input as well as on its own cadence: a tick that has been away for a second — a hover answered, a
/// file read, a machine that went to sleep — would otherwise advance the spring by a second's worth
/// of travel in one step, which is a picture that jumps. It is the same ceiling the animation this
/// is modelled on keeps, and it is above the frame cadence rather than at it, so an ordinary frame
/// is never shortened by it.
///
pub(super) const PIN_ZOOM_LONGEST_FRAME: f32 = 0.033;

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
    ///
    /// It is where the picture *is* and not where it is going: a notch writes `toward` and this
    /// follows it over the frames of a travel (see `pin_zoom_step`).
    pub(super) drawn: (f32, f32),
    /// Where the picture's centre sits inside the window, in device pixels, measured from the
    /// window's own centre. Nothing until a hand has carried it, and nothing at all while the picture
    /// is the window's own size — a picture that size has no room to be moved into.
    pub(super) pan: (f32, f32),
    /// The width the wheel has asked the picture to be drawn at: where `drawn` is on its way to.
    ///
    /// A width rather than a pair, because a picture's height is its width through its shape — the one
    /// thing about it a zoom never changes — so the width is the whole of what a travel carries.
    pub(super) toward: f32,
    /// How fast the drawn width is travelling toward it, in pixels a second.
    ///
    /// It is half of the spring and not a leftover: what a spring remembers between one frame and the
    /// next is where it is *and* how fast it was going, and a step taken from the size alone is a
    /// chase rather than a spring — it arrives at a rate that depends on how many frames a second the
    /// machine happens to run at.
    pub(super) velocity: f32,
    /// The screen point the travel is anchored on: the point of the picture that stays where it is
    /// while the picture grows and shrinks about it.
    ///
    /// It is read from the pointer once per notch and *kept* for every frame of the travel, which is
    /// the whole of why it is written down: a picture anchored on the pointer afresh each frame is a
    /// picture that follows a hand moving the mouse while the wheel is turning — a slide nobody asked
    /// for, since what is being turned is the wheel and not the picture.
    ///
    /// Nothing before the cap, where the anchor is the picture's own middle: a window growing with the
    /// picture has no overhang for an anchor to move the picture into, and the middle is what a
    /// picture that has not been carried is measured about (see `pin_zoom_to`).
    pub(super) anchor: Option<(i32, i32)>,
}

impl PinZoom {
    /// The picture filling its window: where every pin of a kind the wheel zooms starts, and where a
    /// box that is laid out again puts it back (see `reset_pin_zoom`).
    pub(super) fn filling(content: ScreenRegion) -> Self {
        let (width, height) = region_size(content);

        Self::still((width as f32, height as f32), (0.0, 0.0))
    }

    /// A picture standing at a size and a place with nothing on its way anywhere: what a zoom that has
    /// arrived is, and the shape a test about where a picture *lands* is written in.
    pub(super) fn still(drawn: (f32, f32), pan: (f32, f32)) -> Self {
        Self {
            drawn,
            pan,
            toward: drawn.0,
            velocity: 0.0,
            anchor: None,
        }
    }

    /// Whether this view is still on its way to the size the wheel asked for, which is what the loop
    /// keeps the frame cadence for (see `pin_zoom_in_flight`).
    pub(super) fn travelling(&self) -> bool {
        self.drawn.0 != self.toward || self.velocity != 0.0
    }

    /// The size the picture is drawn at in whole pixels, which is what a repaint crops against and what
    /// the file itself is asked for once a travel has ended (see `drawn_pixels`).
    pub(super) fn drawn_size(&self) -> (i32, i32) {
        drawn_pixels(self.drawn)
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
        let (drawn_width, drawn_height) = drawn_pixels(self.drawn);
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
///
/// What it answers is where the notches *ask* the picture to be, which is the size a travel is aimed
/// at rather than the size it is drawn at this frame: a notch is counted on a hook every wheel
/// message on the desktop waits for, so what it can do is write one number down, and the drawing
/// follows over the frames of the travel (see `pin_zoom_toward` and `pin_zoom_step`).
pub(super) fn pin_zoom_after(
    input: &PinZoomInputs,
    notches: i32,
    displays: &dyn Displays,
) -> Option<PinZoomPlan> {
    if notches == 0 {
        return None;
    }

    let mut plan = pin_zoom_to(input, pin_zoom_toward(input, notches), displays);
    // The plan is the view the picture would have if the notches landed at once, so the size it is drawn
    // at in it is a size it has *arrived* at rather than one it is still on its way to: a caller that
    // reads the plan as a state — which is what a test does, one notch at a time — is handed a settled
    // one, and a caller that wants only where the wheel has asked for reads it straight off it (see
    // `aim_pin_zoom`).
    plan.zoom.toward = plan.zoom.drawn.0;

    Some(plan)
}

/// The width a wheel's notches ask a pinned picture to be drawn at: the number a zoom accumulates
/// into, and the whole of what "a notch is a tenth" means.
///
/// It accumulates on what the wheel has *asked for* rather than on where the picture has got to, and
/// that is the whole of what makes a fast wheel and a slow one land in the same place: a notch taken
/// while the last one is still travelling would otherwise be counted from the size the picture
/// happened to be passing, and two hands turning the same number of notches would come out at two
/// different sizes (see `pin_zoom_step`).
pub(super) fn pin_zoom_toward(input: &PinZoomInputs, notches: i32) -> f32 {
    let step = PIN_ZOOM_PER_NOTCH.powi(notches);
    let floor = pin_zoom_floor(input, input.zoom.toward);

    (input.zoom.toward * step).max(floor).min(PIN_ZOOM_CEILING)
}

/// The narrowest a zoom may leave a picture: the display's own room at a hundredth of itself, in
/// *either* direction.
///
/// It is a share of the room rather than a size in logical pixels because the limit is about the
/// screen rather than about what is on it: a picture taken to a hundredth of a display is a mark that
/// can still be found and grown again, and a limit in pixels would be a different one on every
/// monitor.
///
/// A picture has a shape, so the two directions are asked together and it is the wider of the two
/// shares that is the floor — the one that keeps both sides of it above the limit rather than only
/// the one the room happened to be proportioned for.
///
/// And never more than `already`, which is the width the zoom is at: the floor is there to stop a
/// zoom out and not to take one back, so a box taken up already below it is left where it is rather
/// than grown by the notch that asked to shrink it.
// shortcut: that leaves such a pin below the floor until it has been grown and shrunk again — which
// needs a display wider than six thousand pixels before the 48-pixel floor a resize keeps stops
// covering it; upgrade if a display that wide is seen.
fn pin_zoom_floor(input: &PinZoomInputs, already: f32) -> f32 {
    let shape = input.zoom.shape();
    let (room_width, room_height) = input.room.room();

    (PIN_ZOOM_FLOOR_SHARE * (room_width as f32).max(room_height as f32 / shape)).min(already)
}

/// Where a pinned picture drawn `drawn_width` across is left: the box its window ends up with, and the
/// view the picture has inside it.
///
/// `drawn_width` is the size the picture is drawn at *now* — the number a travel walks frame by frame,
/// and the number a notch lands on the instant it is taken — and it is the whole of what the two stages
/// of a zoom are decided from.
pub(super) fn pin_zoom_to(
    input: &PinZoomInputs,
    drawn_width: f32,
    displays: &dyn Displays,
) -> PinZoomPlan {
    let content = input.content;
    let (width, height) = region_size(content);
    let shape = input.zoom.shape();

    // The picture at the size it is being drawn at, held inside the floor it may not be taken under and
    // the ceiling the arithmetic can hold. The floor is asked of the width the picture *is* rather than
    // of the width being asked for, so that a zoom out that has arrived at the floor stops there
    // instead of being taken under it by the next frame of its own travel (see `pin_zoom_floor`).
    let drawn_width = drawn_width
        .max(pin_zoom_floor(input, input.zoom.drawn.0))
        .min(PIN_ZOOM_CEILING);
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

    PinZoomPlan {
        content: placed,
        zoom: PinZoom {
            drawn,
            pan: (
                pan.0.clamp(-overhang.0, overhang.0),
                pan.1.clamp(-overhang.1, overhang.1),
            ),
            // What the travel is aimed at, how fast it is going and what it is anchored on are not
            // this function's business and are carried through it: where a picture *lands* is the whole
            // of what is decided here, and a caller that is walking a size is the caller that knows
            // what it is walking toward (see `pin_zoom_step`).
            toward: input.zoom.toward,
            velocity: input.zoom.velocity,
            anchor: input.zoom.anchor,
        },
    }
}

/// One frame of a travel: how far a picture drawn `drawn` across has got toward the size it was asked
/// for by the time `elapsed` has passed, and how fast it is going when it gets there.
///
/// It is a *critically damped spring*, and each of the three things it is not is the answer to a
/// question the other two get wrong. A curve of a fixed length — an ease of two hundred milliseconds
/// from where the picture is to where it was asked to be — cannot take a second notch in the middle of
/// its own travel: the wheel is a hand and not a script, so a picture that restarted a fixed ease from
/// wherever it had got to would crawl under a fast wheel and lurch under a slow one. A chase that
/// closed a fixed share of the distance each frame would arrive at a rate that depends on how many
/// frames a second the machine happens to run at. And a *spring* has no overshoot while it is
/// critically damped, which is the one thing a zoom may not have: a picture that bounced past the size
/// it was asked for and came back would be a picture that grew and shrank again after the hand had
/// stopped turning.
///
/// It is written in seconds rather than in frames for the same reason the chase is wrong: what the two
/// constants above say is how long a travel takes, and a machine that takes its frames twice as often
/// takes the same travel in twice as many of them.
pub(super) fn pin_zoom_spring(drawn: f32, toward: f32, velocity: f32, elapsed: f32) -> (f32, f32) {
    let dt = elapsed.clamp(0.0, PIN_ZOOM_LONGEST_FRAME);
    let omega = PIN_ZOOM_SPRING_SPEED * 2.0;

    let velocity = velocity + (omega * omega * (toward - drawn) - 2.0 * omega * velocity) * dt;
    let next = drawn + velocity * dt;

    // A spring closes on what it was asked for and never quite arrives, and a picture a quarter of a
    // pixel from its size is a picture at its size: what the snap buys is the *window*, which is whole
    // pixels on a screen and would otherwise be resized and repainted on every frame for the rest of
    // the pin's life for a difference nothing can see. It is also what says a travel is over, and so
    // what says the file is worth laying out again at the size the picture settled on (see
    // `pin_zoom_step`).
    if (toward - next).abs() < PIN_ZOOM_SETTLED_PIXELS && velocity.abs() < PIN_ZOOM_SETTLED_SPEED {
        return (toward, 0.0);
    }

    // And a frame that crosses the size it is going to *is* the frame that arrives at it: a spring
    // given a moving target — a hand turning the wheel back while the last notch is still travelling —
    // can carry past it, and a picture that grew past the size it was asked for and came back would be
    // a picture that bounced.
    if (toward - next) * (toward - drawn) < 0.0 {
        return (toward, 0.0);
    }

    (next, velocity)
}

/// Carry a travel one frame further on: where the picture has got to by the time `elapsed` has passed,
/// and the box that comes with it — or nothing at all for a picture already at the size the wheel asked
/// for.
///
pub(super) fn pin_zoom_step(
    input: &PinZoomInputs,
    elapsed: f32,
    displays: &dyn Displays,
) -> Option<PinZoomPlan> {
    if !input.zoom.travelling() {
        return None;
    }

    let (drawn, velocity) = pin_zoom_spring(
        input.zoom.drawn.0,
        input.zoom.toward,
        input.zoom.velocity,
        elapsed,
    );

    let mut plan = pin_zoom_to(input, drawn, displays);
    plan.zoom.velocity = velocity;

    Some(plan)
}

/// Everything a notch is decided from, read out of the pin in one look — or nothing at all where the
/// wheel is not this pin's to take.
///
/// The four refusals are one question — is there a picture here for the wheel to zoom — asked of the
/// four things that can say there is not: no pin, a pin whose media is a kind the wheel does not zoom,
/// a pin put away as a bubble, and a pin whose box is under a hand already. The last is a *drag* and not
/// a zoom: a wheel notch landing in the middle of one would be two answers to the same question.
fn pin_zoom_inputs() -> Option<PinZoomInputs> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    let zoom = pin
        .zoom
        .filter(|_| !pin.collapsed && pin.dragging.is_none())?;
    let window = pin.window_box();

    Some(PinZoomInputs {
        content: pin.content,
        zoom,
        dpi: pin.dpi,
        room: work_area_at((window.0 + window.2) / 2, (window.1 + window.3) / 2),
        // The point the travel is anchored on, which is the one the wheel last read and not where the
        // pointer is now: a picture anchored afresh on every frame is a picture that follows a hand
        // moving the mouse while the wheel is turning (see `PinZoom::anchor`).
        pointer: zoom.anchor,
    })
}

/// Whether a pinned picture's travel is still going, which is what the loop keeps the frame cadence
/// for: a spring takes its frames at the frame cadence or it is not a spring (see
/// `wait_before_the_next_tick`).
///
/// What it asks is `pin_zoom_inputs` and not the pin alone, so that a travel the loop could not take
/// a frame of is not a travel it keeps the cadence for: a pin with a drag under its hand is a pin the
/// wheel has given over to the hand (see `PinZoom::toward`), and a loop waking sixty times a second
/// for one is a loop waking for nothing.
pub(super) fn pin_zoom_in_flight() -> bool {
    pin_zoom_inputs().is_some_and(|input| input.zoom.travelling())
}

/// The box a pinned picture is drawn in when it is *larger than its window* — or nothing where there
/// is no picture the wheel zooms, and nothing where the picture is the size of its window.
///
/// It is what the loop asks for when the wheel has stopped: a picture drawn larger than its window is
/// a frame the window is holding scaled up, and the file itself is what it is owed. A picture that
/// fills its window is the whole of that question already — there is nothing a decode of another size
/// could add to it — so the box is nothing for one, and the loop asks nothing of the file for every pin
/// of every picture that has never been zoomed (see `relayout_pinned_media`).
pub(super) fn pin_zoom_oversize_box() -> Option<ScreenRegion> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    let (drawn_width, drawn_height) = pin.zoom?.drawn_size();
    let (width, height) = region_size(pin.content);

    (drawn_width > width || drawn_height > height).then_some((0, 0, drawn_width, drawn_height))
}

/// Aim a pinned picture at the size a wheel's notches ask for, from the hook's own count.
///
/// The picture is left exactly where it is: what is written down is where it is *going*, and the frames
/// that take it there are the loop's (see `step_pin_zoom`). That is the whole of why a notch is cheap
/// enough to be counted on a hook every wheel message on the desktop waits for — it is one number
/// written under one lock — and it is what lets a hand turning the wheel twice as fast come out twice as
/// far in, since a notch taken mid-travel moves the size the last one was already going to rather than
/// the size the picture happens to be passing (see `pin_zoom_toward`).
pub(super) fn aim_pin_zoom(notches: i32) {
    let Some(input) = pin_zoom_inputs() else {
        return;
    };
    // What the notches ask for, as if the picture landed there at once: the plan is that landed view, so
    // the size it is drawn at in it is the size the travel below is aimed at — and the box and the pan it
    // carries are where the travel is going rather than where it is (see `pin_zoom_step`).
    let Some(plan) = pin_zoom_after(&input, notches, &DESKTOPS) else {
        return;
    };
    let toward = plan.zoom.toward;
    let (content, zoom) = (input.content, input.zoom);

    // The reading and the write are one lock and one step, as they are for a resize: a walk taken
    // between the two has put another file in the window, and a size decided from the file that was
    // there is not the size this one wants (see `reset_pinned_size`). The tick that answers such a walk
    // is already queued, so the right answer here is to leave it to that rather than to write over the
    // top of it — which is a decision not to write, so it is made beside the write.
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        if pin.content != content || pin.zoom != Some(zoom) {
            return None;
        }

        let view = pin.zoom.as_mut()?;
        view.toward = toward;
        // The point this travel is anchored on: read from the pointer once, here, and kept for every
        // frame of the travel rather than re-read by each of them.
        view.anchor = cursor_screen_point();
        Some(())
    });
}

/// Carry a pinned picture's travel one frame further on and draw it where it has got to, answering
/// whether this is the frame it *arrived* at the size the wheel asked for.
///
/// It is the loop's road and not the window procedure's, because the clock a travel is measured by is
/// the loop's: a spring is written in seconds, and the loop is the one thing here that knows how long
/// its own last tick was (see `PIN_ZOOM_LONGEST_FRAME`).
///
/// The frame, its place and the window's size go to `UpdateLayeredWindow` in one call, exactly as the
/// drag of an edge does it: a window sized first and painted after is a moment in which the surface
/// holds the old box and the window the new one, which an eye reads as the picture jumping (see
/// `apply_pin_drag`).
///
/// The answer is the moment worth knowing about, and the caller is the one that knows it: a picture
/// that has just stopped is a picture whose file is worth laying out again at the size it settled on,
/// and asking for that on every frame of a travel would be a decode per frame (see `pin_zoom_drawn_box`).
pub(super) unsafe fn step_pin_zoom(hwnd: HWND, elapsed: f32) -> bool {
    let Some(input) = pin_zoom_inputs() else {
        return false;
    };

    let Some(plan) = pin_zoom_step(&input, elapsed, &DESKTOPS) else {
        return false;
    };
    let arrived = !plan.zoom.travelling();

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
        render_pinned_preview_at(hwnd, plan.content.0, plan.content.1);
    }

    written && arrived
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
fn drawn_pixels(drawn: (f32, f32)) -> (i32, i32) {
    (
        drawn.0.round().max(1.0) as i32,
        drawn.1.round().max(1.0) as i32,
    )
}

/// Where the picture's centre sits on the screen: a box's middle carried by the pan.
fn pinned_centre(content: ScreenRegion, pan: (f32, f32)) -> (i32, i32) {
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

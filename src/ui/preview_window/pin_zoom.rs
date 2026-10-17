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
//!
//! An image a hand has carried off its window — the backdrop showing around it — is brought home by
//! the zoom out that takes the picture back toward the fit, and swings a little past its place on the
//! way: the *image* inside the window is what bounces there, and the window's own box is the wheel's
//! and nothing else (see `pin_zoom_step`).

use super::*;

/// What one wheel notch multiplies a pinned picture by.
///
/// Out is this step divided rather than a step of its own, so a notch in and a notch out leave the
/// picture exactly where it was — which is the whole of what "zoom in and zoom out" has to mean if a
/// hand is to come back from either end without counting notches.
///
/// It is the step of a *whole* notch, and a wheel may report less than one: the count a notch is
/// turned into is fractional, and what the part of a notch is worth is this step taken by a fractional
/// exponent (see `notch_step`).
pub(super) const PIN_ZOOM_PER_NOTCH: f32 = 1.10;

/// How near the size it was asked for a zoom has to come before it is *there*: a quarter of a pixel.
///
/// A chase closes on what it was asked for and never quite arrives, and what that costs is not the
/// picture — a quarter of a pixel is not a pixel — but the *window*, which is whole pixels on a screen
/// and would otherwise be given a `SetWindowPos` and a repaint on every frame for the rest of the pin's
/// life. It is also what ends a travel, and therefore what says when a picture has stopped long enough
/// to be worth laying out again at the size it has settled on (see `pin_zoom_release` and
/// `pin_zoom_step`).
///
pub(super) const PIN_ZOOM_SETTLED_PIXELS: f32 = 0.25;

/// How long a wheel's ask takes to be carried out: nine tenths of what is left of it within this, and
/// the rest of it behind that.
///
/// It is a share of the *remainder* rather than of the ask, which is the whole of what makes the curve
/// the same on a machine that takes its frames twice as often: a share of the ask is a curve in frames,
/// and two frames of half the length are one frame of the whole (see `released_share`).
///
/// A tenth of a second is the rate the viewer this is modelled on releases its own wheel by, and what
/// it is worth to an eye is the same thing a notch is: long enough to be seen as a movement, and short
/// enough that the picture is never far behind the hand — measured on a spin of twelve notches in a
/// third of a second, the picture trails the wheel by a seventh of itself and the glide that follows
/// the last notch is a third of a second long.
///
pub(super) const PIN_ZOOM_RELEASE_SECONDS: f32 = 0.1;

/// The share of the ask that is still left to carry out at `PIN_ZOOM_RELEASE_SECONDS`: a tenth, which
/// is what makes the release nine tenths of the ask in a tenth of a second and ninety-nine hundredths
/// of it in twice that.
const PIN_ZOOM_RELEASE_LEFT: f32 = 0.1;

/// How long the image inside a window takes to swing back to its place once a zoom out has brought the
/// picture within a notch of the fit: a tenth of a second or so.
///
/// It is a *landing* and not a slower zoom: the image is brought home by the zoom itself, over the last
/// of the picture's overhang, and this is the bounce laid over the end of it — long enough to be read as
/// a movement and short enough that the image is never left behind the wheel (see `pin_zoom_step`).
pub(super) const PIN_ZOOM_BOUNCE_SECONDS: f32 = 0.12;

/// How far past its place the image swings at the top of that bounce, as a share of the offset it is
/// being brought home from: a sixth of it.
///
/// It is what an eye reads as a bounce rather than as a stop, and the whole of what it swings is the
/// *image's* offset inside a window whose box never moves: the picture is never taken far off its
/// window and the frame the fit asks for is reached exactly all the same (see `bounce_swing`).
const PIN_ZOOM_BOUNCE_PAST: f32 = 0.15;

/// The least a bounce is worth taking: the pixels the image would be swung by.
///
/// A picture sitting in the middle of its window has no offset to give up and one a hair off the middle
/// has a hair's swing, and a swing under a pixel is not a movement — so neither is worth the frames it
/// would cost the loop, and the zoom out of them is the motion it always was (see `pin_zoom_bounce`).
const PIN_ZOOM_BOUNCE_LEAST_PIXELS: f32 = 1.0;

/// The longest a single frame of a travel is taken to have lasted, in seconds.
///
/// A release is arithmetic on how much time has passed, and the loop that takes its frames wakes on
/// input as well as on its own cadence: a tick that has been away for a second — a hover answered, a
/// file read, a machine that went to sleep — would otherwise carry a second's worth of the wheel's ask
/// out in one step, which is a picture that jumps. It is above the frame cadence rather than at it, so
/// an ordinary frame is never shortened by it.
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
    /// How far into its bounce the image inside the window is, in seconds: nothing except while a zoom
    /// out is bringing a panned image back to its place (see `pin_zoom_step`).
    ///
    /// It is the *image's* own motion and never the window's: what a bounce swings is where the picture
    /// sits inside its window, and the frame the window is drawn in is the wheel's own and stays exactly
    /// where the wheel put it (see `swung_pan`).
    pub(super) bounce: f32,
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
            anchor: None,
            bounce: 0.0,
        }
    }

    /// Whether this view is still on its way to the size the wheel asked for, or swinging the image
    /// inside it back to its place — either of which is what the loop keeps the frame cadence for (see
    /// `pin_zoom_in_flight`).
    pub(super) fn travelling(&self) -> bool {
        self.drawn.0 != self.toward || self.bounce > 0.0
    }

    /// The pan the picture is *drawn* with: the offset the wheel and the hand left it with, swung a
    /// little past its place while the bounce of a return is still running (see `bounce_swing`).
    ///
    /// It is where the *image* is inside the window and not where the window is: the box the window takes
    /// is the wheel's own and never this one, so an image carried off its window comes home without the
    /// window moving with it (see `pin_zoom_to`).
    fn swung_pan(&self) -> (f32, f32) {
        let swing = bounce_swing(self.bounce);

        (self.pan.0 * swing, self.pan.1 * swing)
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

    /// Where the picture *really* lands inside its window: the place in the floats the zoom's own
    /// arithmetic walks, which is what a sampler is owed and what a box of whole pixels is a rounding of.
    ///
    /// **The fraction is the whole of what the wheel's anchoring is worth.** The release walks the size by
    /// a fraction of a pixel a frame and the place with it, and a frame mapped through a whole-pixel box is
    /// a picture placed on the pixel grid: every time the place crosses a half pixel the whole picture steps
    /// a pixel — which a hand that is holding still reads as the picture jolting *backwards* at the end of a
    /// zoom, because the jolt is the same size whatever the frame's own travel was. Measured on six notches
    /// taken at the cap, the point of the picture under the wheel stepped 1.00 pixels back on a frame whose
    /// travel was 0.33 forwards, and the place handed to the sampler is where that stops (see
    /// `BandTarget::exact` and `resample_into_band`).
    ///
    /// The box the picture is written *into* is `picture_box`, and the two are the same place rounded: a
    /// caller wants the exact one or the whole-pixel one and never a third answer.
    pub(super) fn picture_place(&self, window: (i32, i32)) -> (f32, f32, f32, f32) {
        let (width, height) = (window.0.max(1), window.1.max(1));
        let (drawn_width, drawn_height) = drawn_pixels(self.drawn);
        if drawn_width <= width && drawn_height <= height {
            return (0.0, 0.0, width as f32, height as f32);
        }

        let (pan_x, pan_y) = self.swung_pan();

        (
            (width as f32 - self.drawn.0) / 2.0 + pan_x,
            (height as f32 - self.drawn.1) / 2.0 + pan_y,
            self.drawn.0,
            self.drawn.1,
        )
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
    ///
    /// It is the box the picture's pixels are written *into* — which is what a fill has to know, and what a
    /// frame drawn at its own size is copied through — and where they are *drawn* is the place beside it, in
    /// the floats this box is a rounding of (see `picture_place`). The half pixel of place a copy gives up is
    /// what a copy is worth here: a frame the size of its own box sampled rather than copied is the sampler's
    /// several milliseconds a frame (see `copy_frame_into_band`).
    pub(super) fn picture_box(&self, window: (i32, i32)) -> (i32, i32, i32, i32) {
        let (width, height) = (window.0.max(1), window.1.max(1));
        let (drawn_width, drawn_height) = drawn_pixels(self.drawn);
        if drawn_width <= width && drawn_height <= height {
            return (0, 0, width, height);
        }

        let (pan_x, pan_y) = self.swung_pan();
        let left = (width - drawn_width) / 2 + pan_x.round() as i32;
        let top = (height - drawn_height) / 2 + pan_y.round() as i32;

        (left, top, left + drawn_width, top + drawn_height)
    }
}

/// How much of a panned picture's offset is still its own, by how far the picture overhangs its window:
/// the whole of it while the overhang is worth more than a notch, and none of it at the fit itself.
///
/// **This is what brings an image home.** A hand may carry a picture clean off its window — the backdrop
/// showing around it — and the wheel will follow it rather than undoing it; what is not free is the *end*
/// of that: the moment the picture is the window's own size again there is no room to sit off the middle
/// in, and an offset given up there is given up at once — a picture's worth of a jump. So it is given up
/// over the last of the overhang instead, one notch's worth of it, and the share is a smooth step, so that
/// the offset neither starts nor stops being given up with a jerk (see `pin_zoom_to`).
///
/// The share is asked of the overhang and not of a clock, which is the whole of why the image is *already*
/// home by the time the picture reaches the fit: whatever the wheel does — however fast it is turned,
/// wherever it is stopped — the fit is the one size at which the offset is nothing.
pub(super) fn home_share(overhang: f32, window: f32) -> f32 {
    let last = window * (PIN_ZOOM_PER_NOTCH - 1.0);
    if last <= 0.0 || overhang >= last {
        return 1.0;
    }

    let left = (overhang / last).clamp(0.0, 1.0);

    left * left * (3.0 - 2.0 * left)
}

/// The swing a bounce has the image at by the time `at` of it has passed, as a factor on the offset the
/// image is being brought home from: nothing at either end of the bounce, and a sixth of that offset on
/// the *other* side of its place at the top of it (see `PIN_ZOOM_BOUNCE_PAST`).
///
/// Either end being the whole of it is what makes the bounce a bounce rather than a move: it is laid over
/// a return that the zoom itself is carrying out, so an image whose bounce has run its course is exactly
/// where that return has brought it — nothing to undo and nothing left over (see `swung_pan`).
fn bounce_swing(at: f32) -> f32 {
    if at <= 0.0 || at >= PIN_ZOOM_BOUNCE_SECONDS {
        return 1.0;
    }

    let turn = std::f32::consts::PI * at / PIN_ZOOM_BOUNCE_SECONDS;

    1.0 - (1.0 + PIN_ZOOM_BOUNCE_PAST) * turn.sin()
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
/// `notches` is the wheel's own turn in notches and is signed, up being positive — a positive notch
/// zooms in — and fractional, because a wheel may report a fraction of a notch at a time and a count
/// rounded to whole ones is a count that throws most of a high-resolution wheel's travel away (see
/// `notch_step`). More than one notch at once is the same gesture taken further: the step is the one
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
    notches: f32,
    displays: &dyn Displays,
) -> Option<PinZoomPlan> {
    if notches == 0.0 {
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
pub(super) fn pin_zoom_toward(input: &PinZoomInputs, notches: f32) -> f32 {
    let step = notch_step(notches);
    let floor = pin_zoom_floor(input, input.zoom.toward);

    (input.zoom.toward * step).max(floor).min(PIN_ZOOM_CEILING)
}

/// What a wheel's own turn is worth as a factor: the step of a notch taken as many times as the wheel
/// was turned.
///
/// A whole notch is the exact power — which is what makes a notch in and a notch out leave the picture
/// where it was — and the part of a notch a high-resolution wheel sends is the same step taken by a
/// fractional exponent, because a wheel that reports less than a notch per message is a wheel whose
/// every unit is worth the same share of a notch. A count rounded to whole notches would throw two of
/// every three of a 40-unit wheel's messages away and turn the third into a whole notch, which is a
/// zoom that moves in lurches of a tenth of a picture rather than in the thirds it was asked for.
fn notch_step(notches: f32) -> f32 {
    let whole = notches.trunc();
    let part = notches - whole;

    // A whole notch is the exact power, and a turn with no part to it is the whole of what a wheel of
    // detents ever sends: the fractional step is the same factor and is exactly one where the part is
    // nothing, so the two are one expression rather than a branch either side of it.
    PIN_ZOOM_PER_NOTCH.powi(whole as i32) * PIN_ZOOM_PER_NOTCH.powf(part)
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
/// `drawn_width` is the width the picture is drawn at in the answer — the number a travel walks frame
/// by frame, and the number the wheel's notches come to when they are taken as landing at once — and it
/// is the whole of what the two stages of a zoom are decided from.
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

    // Where the window is put: about its own middle, and never about the picture's carried
    // centre inside it. A picture carried inside a window that has stopped is the picture moving
    // and not the window, so a window shrinking about a carried centre is a window that walks
    // across the screen as the hand zooms out of a pan — the window follows the window itself,
    // whatever the pan is doing, and the pan only ever moves the picture within it.
    let window_centre = ((content.0 + content.2) / 2, (content.1 + content.3) / 2);
    let placed = if size == (width, height) {
        content
    } else {
        clamp_pinned_box(centred_at(size, window_centre), input.dpi, displays)
    };
    let box_centre = (
        (placed.0 + placed.2) as f32 / 2.0,
        (placed.1 + placed.3) as f32 / 2.0,
    );

    // And where the picture sits inside it. A picture the window's own size — or one a hand has
    // shrunk under the window's floor, which is smaller than it — fills it and has nowhere to sit,
    // whatever it was carrying before a zoom out brought it back to that size; one drawn larger is
    // anchored on the pointer, which is what lets a hand zoom into a corner it has pointed at. It is
    // the question `picture_box` asks, in the same whole pixels, because the two are one rule: what a
    // pan may move is exactly what a box has room for (see `PinZoom::picture_box`).
    //
    // The anchor is the point of the picture under the pointer staying under the pointer: its own
    // offset from the picture's middle is carried out by however much the picture grew, and the
    // picture's middle is then put wherever the window's own middle leaves it. Nothing bounds that
    // placement, and nothing has to while the pointer is over the *picture*: an offset inside the
    // picture is at most half of it, so the middle moves by at most half of what the picture grew by
    // — and that is at most what the picture has outgrown the window by, since the window grows to
    // meet the picture and never past it. What is left over is the pointer over the backdrop, which
    // is a hand that has already carried the picture off its window: the drag is free to place it
    // there, and the wheel now follows it rather than undoing it.
    //
    // The bound this replaces was there for the one notch that takes the window out to the cap in the
    // same breath as the picture past it, where the whole of the pointer's offset would have pulled
    // the picture off the far side of a window it was covering. That notch grows the window by a tenth
    // at the very most, since the window grows with the picture a frame at a time (see `pin_zoom_step`),
    // and the rule above covers it by the same arithmetic as every other notch.
    // The same centre in the floats the zoom is walked in, which is what the pan below is measured from
    // and not the whole-pixel one the window's place is chosen about: a centre rounded to the pixel grid
    // is half a pixel of the last frame's rounding carried into the anchor of every frame that follows,
    // and it compounds — measured on six notches taken at the cap, the point of the picture under the
    // wheel drifted 3.49 pixels off it and this centre holds it to 0.75, which the place beside it then
    // takes to nothing (see `picture_place`).
    let (exact_x, exact_y) = (
        (content.0 + content.2) as f32 / 2.0 + input.zoom.pan.0,
        (content.1 + content.3) as f32 / 2.0 + input.zoom.pan.1,
    );
    let (drawn_width, drawn_height) = drawn_pixels(drawn);
    let pan = match input
        .pointer
        .filter(|_| drawn_width > size.0 || drawn_height > size.1)
    {
        Some(point) => {
            let grown = drawn.0 / input.zoom.drawn.0;
            let (dx, dy) = (exact_x - point.0 as f32, exact_y - point.1 as f32);

            (
                point.0 as f32 - box_centre.0 + dx * grown,
                point.1 as f32 - box_centre.1 + dy * grown,
            )
        }
        None => (0.0, 0.0),
    };

    // And what a picture coming *back* to the fit gives up of it, which is the whole of what brings an
    // image home: over the last notch's worth of overhang the offset is handed back a share at a time, so
    // that the size the hand is aiming at — the one where a picture its window's own size has nowhere to
    // sit off its middle — is reached with nothing left to give up at once. A zoom *in* gives up nothing:
    // a hand turning the wheel the other way is opening room for the offset, not taking it away.
    let pan = if drawn.0 < input.zoom.drawn.0 {
        let share = home_share(drawn.0 - size.0 as f32, size.0 as f32);

        (pan.0 * share, pan.1 * share)
    } else {
        pan
    };

    PinZoomPlan {
        content: placed,
        zoom: PinZoom {
            drawn,
            pan,
            // What the travel is aimed at and what it is anchored on are not this function's
            // business and are carried through it: where a picture *lands* is the whole of what is
            // decided here, and a caller that is walking a size is the caller that knows what it is
            // walking toward (see `pin_zoom_step`).
            toward: input.zoom.toward,
            anchor: input.zoom.anchor,
            bounce: input.zoom.bounce,
        },
    }
}

/// How far a travel has got by the time `elapsed` has passed: the width the picture is drawn at after a
/// frame, on its way to the width the wheel asked for and never past it.
///
/// A notch is a step, and a step landed whole is a jump: what a hand turning a wheel is looking at is a
/// *movement*, and a wheel that has stopped would be doing nothing at all. So the ask is carried out
/// over the frames that follow it, and the part of it that is left when the hand stops is what the
/// picture is still doing — a fast wheel builds a longer glide than a slow one, because every notch
/// leaves its own remainder behind it.
///
/// It is a share of what is *left* rather than of the ask, which is the whole of what makes the curve
/// the same on a machine that takes its frames twice as often: two frames of half the length are one
/// frame of the whole (see `released_share`). And it is a chase written in seconds rather than a spring:
/// what a spring remembers between one frame and the next is where the picture is *and* how fast it was
/// going, and a step taken from the size alone arrives at the rate the seconds say rather than at the
/// rate the frames happen to come.
///
/// And the last quarter of a pixel of it is not a release but an arrival: a chase closes on what it was
/// asked for and never quite arrives, and a release left dribbling a hundredth of a pixel a frame would
/// keep the loop at the frame cadence for the rest of the pin's life (see `PIN_ZOOM_SETTLED_PIXELS`).
pub(super) fn pin_zoom_release(drawn: f32, toward: f32, elapsed: f32) -> f32 {
    let left = toward - drawn;
    if left.abs() < PIN_ZOOM_SETTLED_PIXELS {
        return toward;
    }

    drawn + left * released_share(elapsed)
}

/// The share of what is left of a wheel's ask that one frame carries out: nine tenths of it within
/// `PIN_ZOOM_RELEASE_SECONDS`, and the rest of it behind that.
///
/// It is the *curve in seconds* and not the shape of the release, which is the whole of what a share of
/// the remainder buys: a share of the ask itself would be a curve in frames, and a machine that takes
/// its frames twice as often would take the same notch in half the time.
fn released_share(elapsed: f32) -> f32 {
    let dt = elapsed.clamp(0.0, PIN_ZOOM_LONGEST_FRAME);

    1.0 - PIN_ZOOM_RELEASE_LEFT.powf(dt / PIN_ZOOM_RELEASE_SECONDS)
}

/// Carry a travel one frame further on: where the picture has got to by the time `elapsed` has passed,
/// and the box that comes with it — or nothing at all for a picture already at the size the wheel asked
/// for.
///
/// The frame is the wheel's ask *released* and not the ask itself, which is the whole of what a notch is
/// worth to an eye: a step landed whole is a jump, and a hand that has stopped turning is a picture that
/// goes on moving for as long as the rest of the ask takes to carry out (see `pin_zoom_release`).
///
/// And the *image* inside the window is brought home as the picture comes back to the fit, with a bounce
/// laid over the end of it — the window's own box is nothing to do with either, and is the wheel's own to
/// the pixel (see `home_share` and `bounce_swing`).
pub(super) fn pin_zoom_step(
    input: &PinZoomInputs,
    elapsed: f32,
    displays: &dyn Displays,
) -> Option<PinZoomPlan> {
    if !input.zoom.travelling() {
        return None;
    }

    let drawn = pin_zoom_release(input.zoom.drawn.0, input.zoom.toward, elapsed);
    let bounce = pin_zoom_bounce(input, drawn, elapsed);
    let mut plan = pin_zoom_to(input, drawn, displays);

    plan.zoom.bounce = bounce;

    Some(plan)
}

/// How far into its bounce a picture's image is after this frame: the clock the image's swing is measured
/// by, started on the one frame a zoom out first brings the picture within a notch of the fit.
///
/// The trigger is that *edge* and not the state, so a picture that spends many frames inside the last
/// notch's worth of overhang — a slow release, a wheel turned on — swings once rather than on every frame
/// of it; and it is asked of the size the release is passing rather than of the size it was aimed at, so
/// the swing is laid over the last frames of the travel that brought it there.
fn pin_zoom_bounce(input: &PinZoomInputs, drawn: f32, elapsed: f32) -> f32 {
    if input.zoom.bounce > 0.0 {
        let at = input.zoom.bounce + elapsed;

        return if at < PIN_ZOOM_BOUNCE_SECONDS {
            at
        } else {
            0.0
        };
    }

    let window = region_size(input.content).0 as f32;
    let last = window * (PIN_ZOOM_PER_NOTCH - 1.0);
    let crossing =
        drawn < input.zoom.drawn.0 && input.zoom.drawn.0 - window > last && drawn - window <= last;

    // And a swing worth taking, which is a swing an eye could see: a picture sitting in the middle of its
    // window has no offset to give up, and one a hair off the middle has a hair's swing — neither is a
    // movement, and neither is worth the loop's frame cadence (see `PIN_ZOOM_SETTLED_PIXELS` for what a
    // hair is worth, and `PIN_ZOOM_BOUNCE_PAST` for how much of an offset a swing moves).
    let swung = input.zoom.pan.0.abs().max(input.zoom.pan.1.abs()) * (1.0 + PIN_ZOOM_BOUNCE_PAST);

    if crossing && swung > PIN_ZOOM_BOUNCE_LEAST_PIXELS {
        elapsed
    } else {
        0.0
    }
}

/// The view of a pinned picture the wheel is the owner of, or nothing at all where it is not this pin's
/// to take.
///
/// Three of the four refusals are here and the fourth is the pin's existence itself: a pin whose media
/// is a kind the wheel does not zoom has no view, a pin put away as a bubble has no window on the screen
/// to point at or to draw, and a pin whose box is under a hand already is a *drag* rather than a zoom — a
/// wheel notch landing in the middle of one would be two answers to the same question.
pub(super) fn pin_zoom_view(pin: &PinnedPreview) -> Option<PinZoom> {
    pin.zoom.filter(|_| !pin.collapsed && pin.dragging.is_none())
}

/// Everything a notch is decided from, read out of the pin in one look — or nothing at all where the
/// wheel is not this pin's to take (see `pin_zoom_view`).
fn pin_zoom_inputs() -> Option<PinZoomInputs> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    let zoom = pin_zoom_view(pin)?;
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
/// for: a travel takes its frames at the frame cadence or it is a jump (see
/// `wait_before_the_next_tick`).
///
/// What it asks is the same question `pin_zoom_inputs` asks and nothing else, so that a travel the loop
/// could not take a frame of is not a travel it keeps the cadence for: a pin with a drag under its hand
/// is a pin the wheel has given over to the hand, and a loop waking sixty times a second for one is a
/// loop waking for nothing. The room the display has is deliberately not read for it, since a travel
/// being in flight is not a question about the desk.
pub(super) fn pin_zoom_in_flight() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().and_then(pin_zoom_view))
        .is_some_and(|zoom| zoom.travelling())
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
pub(super) fn aim_pin_zoom(notches: f32) {
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

/// Carry a pinned picture's travel one frame further on and draw it where it has got to.
///
/// It is the loop's road and not the window procedure's, because the clock a travel is measured by is
/// the loop's: a release is written in seconds, and the loop is the one thing here that knows how long
/// its own last tick was (see `PIN_ZOOM_LONGEST_FRAME`).
///
/// The frame, its place and the window's size go to `UpdateLayeredWindow` in one call, exactly as the
/// drag of an edge does it: a window sized first and painted after is a moment in which the surface
/// holds the old box and the window the new one, which an eye reads as the picture jumping (see
/// `apply_pin_drag`).
///
/// Nothing is answered about whether the travel has ended: what a caller does with a size a picture has
/// settled at — laying the file out again for it — is asked of the size itself and not of the moment it
/// stopped, so that a travel ended by a hand taking the picture is owed the same frame as one that ran
/// its course (see `pin_zoom_oversize_box`).
pub(super) unsafe fn step_pin_zoom(hwnd: HWND, elapsed: f32) {
    let Some(input) = pin_zoom_inputs() else {
        return;
    };

    let Some(plan) = pin_zoom_step(&input, elapsed, &DESKTOPS) else {
        return;
    };

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

/// The size of a screen region, width first, and never less than one pixel in either direction: a box
/// with no room in it is a divide waiting to happen on every road this app measures one.
pub(super) fn region_size(region: ScreenRegion) -> (i32, i32) {
    ((region.2 - region.0).max(1), (region.3 - region.1).max(1))
}

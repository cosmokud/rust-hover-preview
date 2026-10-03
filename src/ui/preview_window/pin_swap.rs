//! Deciding what a pin should show next and what has to be waited for: the plan an update turns
//! into, the room it takes, and the swap held until the new frame is in hand.

use super::*;

/// What a pin that is up needs to be shown another file: the box the new file's media is given,
/// the scale of the display it is laid out at, and the level the pin plays at.
///
/// It is read off the pin in one look and carried by value, because what follows the plan — a
/// file read, a decode, a player started — is work the pin's own lock must not be held across:
/// that lock is what the window procedure takes to answer a drag, a button, or a tick's repaint.
pub(super) struct PinUpdate {
    pub(super) content: ScreenRegion,
    pub(super) dpi: u32,
    pub(super) volume: u32,
}

/// What a pin that is up is to do with the file it was offered.
pub(super) enum PinPlan {
    /// Show it: the box its media takes, the display it is laid out for, and the level the pin
    /// plays at are what the swap is made with (see `PinUpdate`).
    Show(PinUpdate),
    /// Show it in a moment: the file has no box of its own yet, because what was measured for it is
    /// the wait for a read or a probe that is in flight. The pin keeps the file it is showing, and
    /// the offer is made again — by name, from the answer itself — when the box lands (see
    /// `PinBox::Waiting`).
    Awaiting,
}

/// What a swap has of the file it is for: the box its media takes, or the wait for one.
///
/// The two are not the same answer and may not be used as one: the wait is a placeholder the
/// measurers hand back while the read that would know runs, and a placeholder has no shape — a swap
/// laid out at it is a square window with the file's picture stretched into it (see
/// `box_is_the_wait`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinBox {
    /// The box the new file's media takes on screen.
    Measured(ScreenRegion),
    /// No box yet, and one is on its way.
    Waiting,
}

/// What a pin that is up has to lay the file that replaces the one it is showing out with: the box
/// the pin occupies now, the bound it may not grow past, the two facts about the kind on screen
/// that turn a display's work area into the pin's own room, and whether the window is maximized.
///
/// Each one is an answer about how large the new media may be, and each is read off the pin in one
/// look and carried by value for the reason `PinUpdate` is: the lock that answered is what the
/// window procedure takes to answer a drag, and a swap must not hold it across a file read (see
/// `pin_swap_room` and `pin_update_content`).
#[derive(Clone, Copy)]
pub(super) struct PinSwapSpace {
    /// The media box the pin has now: what the new file is centred on, so a window the hand has
    /// moved keeps its place while it changes size.
    pub(super) current: ScreenRegion,
    /// The longest side the new file's media may take, in either direction: a swap's ceiling
    /// rather than the size it is given, and nothing while the pin has no bound yet (see
    /// `PinnedPreview::bound`).
    pub(super) bound: Option<i32>,
    /// Whether the kind on screen carries a transport bar.
    pub(super) transport_bar: bool,
    /// Whether its chrome is drawn over its media (see `pin_overlay_chrome`).
    pub(super) overlay: bool,
    /// The room its caption takes above the media, which the display's room has to lose as well as
    /// the media itself (see `pinned_caption_height`).
    pub(super) caption: i32,
    /// Whether the window is maximized, where the display's room is the box and the bound is left
    /// for the restore that follows.
    pub(super) maximized: bool,
    /// The room the display has, which is where a maximized window's file is laid out — the
    /// whole of what keeps a walk through shapes a window of one size (see `pin_update_content`).
    pub(super) room: ScreenBounds,
}

/// What a pin has to lay another file out with, read off it in one look: see `PinSwapSpace`.
///
/// It is read rather than asked for piece by piece so that the lock is let go of before
/// anything is measured or asked for an engine — the same rule `PinUpdate` keeps.
pub(super) fn pin_swap_space(pin: &PinnedPreview) -> PinSwapSpace {
    let bounds = work_area_at(pin.content.0, pin.content.1);
    PinSwapSpace {
        current: pin.content,
        bound: pin.bound,
        transport_bar: pin.transport_bar,
        overlay: pin.overlay,
        caption: pin.caption,
        maximized: pin.restore.is_some(),
        room: pinned_room(bounds, pin.dpi, pin.transport_bar, pin.overlay, pin.caption),
    }
}

/// The plan for showing the pinned window another file, if there is one to be had.
///
/// Nothing comes of a pin that is collapsed into its bubble — there is no media on screen to
/// replace, and a swap taken up under one would be a window that came back from its bubble
/// showing a file nobody picked in it. Nothing comes of the file the pin is already showing, which
/// is what a second ask for one — a click and the focus that moved with it, a key pressed twice on
/// one row — is: the work a swap costs is a decode, and it is not work to do twice for one file.
/// Nothing comes of a file with no shape yet either, or of one this app has no preview for: a pin
/// that is up can only keep the file it is showing.
///
/// What the new file's box is measured against is the pin's own bound rather than the box the file
/// on screen came out at, which is the whole of what keeps a pin that follows a listing from walking
/// its way down it file by file — and the room the display has where the pin has no bound yet, which
/// is every pin taken up on a file drawn to its own box that has not been shown a shape since (see
/// `PinnedPreview::bound` and `pin_swap_room`).
pub(super) fn pin_update_plan(path: &PathBuf) -> Option<PinPlan> {
    let (space, collapsed, showing) = {
        let pinned = pin_state()?;
        let pin = pinned.pin()?;
        (pin_swap_space(pin), pin.collapsed, pin.path.clone())
    };

    if collapsed || showing == *path {
        return None;
    }

    // The scale the media is laid out at is asked of the box the way the take-up that follows asks
    // it, rather than read off the pin: what the take-up computes is the scale it draws the new
    // kind's chrome at, and a media loaded at another one would be a picture and a caption that
    // disagree about how large a pixel is.
    let dpi = dpi_at(space.current.0, space.current.1);
    let bounds = work_area_at(space.current.0, space.current.1);

    // The level the file replacing the one on screen is played at is asked for by the kind of that
    // file rather than read off the pin, for the reason `pinned_level` gives: the level the pin is
    // holding is the kind it is showing, and a film is not played at the sound's.
    let volume = pinned_level(drawn_as_audio(path));

    Some(match pin_update_content(space, path, bounds, dpi)? {
        PinBox::Measured(content) => PinPlan::Show(PinUpdate {
            content,
            dpi,
            volume,
        }),
        PinBox::Waiting => PinPlan::Awaiting,
    })
}

/// The room an engine that has to draw a pin's new file is asked for: the room of the display the
/// pin is on, which is the room the ask is sized by wherever a preview is picked up — the hover
/// asks its own display for the same file the same way (see `PendingLoad::room`).
///
/// It is the display's rather than the pin's own box, and the two are different questions: what an
/// engine is told is how large the page or the picture it writes may be, and it is the pin's box
/// that answer is then fitted into (see `pin_swap_room`). Sizing the ask by the pin's box instead
/// would have a pinned document's page written at a size the same file's hover writes again — one
/// page per document and version is kept, whichever side asked for it (`document_cache`) — and a
/// page drawn for a small window is one that could never be shown any larger, which a pin that is
/// maximized afterwards asks for. What the engine is asked for is what a hover asks for, so the
/// two share what comes back.
pub(super) fn pin_engine_room() -> Option<(u32, u32)> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    if pin.collapsed {
        return None;
    }

    Some(work_area_at(pin.content.0, pin.content.1).room())
}

/// Ask whichever engine owes a pinned window's new file its page, a picture or a listing — and
/// answer what is now being waited on.
///
/// It is the ask the hover loader makes for the same file (see `request_engine_render`), made
/// here because a swap does not go through the loader: the pin keeps the file it is showing
/// until the answer lands, and the landing is taken up where it arrives — a message for Office,
/// the image converter and the listing engine, and the page itself for the two engines that
/// write one into the app's own folder (see `PinPlan::Awaiting` and `engine_page_answer`).
pub(super) fn request_pin_engine_render(path: &Path, generation: u64) -> Option<(PathBuf, u64)> {
    request_engine_render(path, generation, pin_engine_room()?)
}

/// Take an engine's answer for a pinned window, where the pin is what was waiting on it: whether
/// this answer was the pin's, and — where it was — that the file is picked up again when there is
/// something to pick up.
///
/// A page, a picture or a listing a pin is owed is taken here rather than by the hover machinery
/// the same answer feeds: what takes it up is the swap, which lays the file out for the pin's own
/// box, where a hover replayed for it would be a second preview on screen beside a window that is
/// not one. A file the engine turns down (`ok` false) is the answer too, and what it leaves is the
/// pin showing the file it has — the wait is over either way, and a refusal is remembered, so a
/// second pick of the file costs no launch (see `page_is_on_the_way` and `refused`).
pub(super) fn take_pin_engine_answer(
    path: &Path,
    ok: bool,
    awaiting: &mut Option<PathBuf>,
    pick: &mut Option<PathBuf>,
) -> bool {
    if !pinned() || awaiting.as_deref() != Some(path) {
        return false;
    }

    *awaiting = None;

    if ok {
        *pick = Some(path.to_path_buf());
    }

    true
}

/// The media box a pin is given for another file.
///
/// What a pin promises across a swap is where it stands: a window that jumped to a fresh placement
/// — beside the pointer that picked the file, or into the best room the display has — would be a
/// window taken out from under the hand that is reading it, and the point of the setting is for the
/// pin to follow the listing rather than for it to be placed again on every file. So the new media
/// goes in the middle of the box the pin has now, and how large it may be is answered by three
/// things, innermost first: the pin's bound is the ceiling, and it is a side rather than a box, so
/// every shape may use the whole of it; inside it the file is laid out at the scale a hover of the
/// same file would take; and a maximized window stays maximized, since a swap is not the gesture
/// that takes it out of the state the user put it in.
///
/// Four kinds are exceptions, and they are exceptions the same way. A page of text, a listing out
/// of an archive, a page an engine draws and a sound's card hold no shape of their own: what sizes
/// them is the room they are laid out in at the share their kind names — which is the answer a
/// hover of the same file gives, and so the one `Scaling` names — and the box comes out centred on
/// the box the pin has now rather than fitted into it. The bound is a size some other file came out
/// at, and none of these four is drawn to that or fitted into it; the window keeps its place on the
/// display and changes size about the middle of it, which is the only half of the promise a swap
/// makes about where it stands that is kept.
///
/// Which is also why a swap to one neither takes a bound nor gives one: a box laid out in the room
/// is a size the file was laid out at and no ceiling for the files after it.
///
/// And a file whose box is still being read is not laid out at all: what the measure answered is
/// the wait for a box rather than one, and a window made of it would be a square with the file's
/// picture in it at the wrong shape. The same answer is what a file gets whose media an engine
/// still owes it: the pin keeps the file it is showing until the answer lands (see `PinBox` and
/// `request_pin_engine_render`).
pub(super) fn pin_update_content(
    space: PinSwapSpace,
    path: &PathBuf,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<PinBox> {
    // A sound is the one kind drawn to its box whose media is not in hand until a probe has
    // answered: the card is built from what the machine has for the file, and that verdict is a
    // read of it — a source reader, or an `ffprobe` run — which is felt, so it is taken where a
    // hover takes it, off this thread (see `audio_box`). The measure is also what starts the
    // probe, which is the whole of why it is asked before the box is settled: what the card says,
    // and whether there is a card at all, is the answer this is for.
    if drawn_as_audio(path) {
        return match media_dimensions(path, bounds, dpi) {
            // Nothing here plays the file, so there is no card to swap in: the pin keeps what it
            // is showing. The verdict is remembered, so this is not a wait that comes back.
            None => None,
            // The probe is in flight, which is the wait `pin_swap_awaits` names.
            Some(shape) if pin_swap_awaits(path, shape) => Some(PinBox::Waiting),
            // The card is there to be laid out, and a card is its own size: the box it is given
            // is the one its own measure came out at, in the middle of the box the pin has now —
            // never the box another file left, which no card is drawn to or fitted into.
            Some(shape) => {
                let centre = (
                    (space.current.0 + space.current.2) / 2,
                    (space.current.1 + space.current.3) / 2,
                );

                Some(PinBox::Measured(centred_at(
                    (shape.0 as i32, shape.1 as i32),
                    centre,
                )))
            }
        };
    }

    // A page, a picture or a listing an engine still owes the file is a wait rather than a
    // refusal: nothing of the new file can be laid out yet, and the pin keeps the file it is
    // showing while the loop asks for what is owed (see `request_pin_engine_render`). A file an
    // engine has turned down is not on the way — the question is the one that was asked before
    // the engine was, and it answers no for a file already refused — so it comes through here
    // and is measured, or is nothing, as it always was (see `page_is_on_the_way`).
    if page_is_on_the_way(path) {
        return Some(PinBox::Waiting);
    }

    let scale = effective_preview_scale(path, current_hover_scales());
    let shape = media_dimensions(path, bounds, dpi)?;

    // A box still being read is the wait for a box rather than one, whichever file it was asked
    // for: the measure this call has just taken is what starts the read that settles it, so what
    // the pin has of the new file is a wait to keep rather than a shape to lay out (see
    // `pin_swap_awaits`).
    if pin_swap_awaits(path, shape) {
        return Some(PinBox::Waiting);
    }

    // A kind that holds no shape of its own — a page of text, a listing, a page an engine draws, a
    // sound's card — is laid out in the room rather than fitted into the bound, at the share its
    // kind names: which is the answer a hover of the same file is given, and so the one `Scaling`
    // names. A bound is a size some *other* file came out at, and none of these four is drawn to
    // that or fitted into it.
    //
    // The box it comes out with is centred on the box the pin occupies rather than fitted into the
    // room it was laid out in, so the window keeps its place on the display the hand put it and
    // only its size moves — the same answer a card is given, and for the same reason: a box fitted
    // into the room is centred on the middle of the *screen* (see `pin_update_box`), which is a
    // fresh placement rather than the one a swap promises.
    //
    // The maximize is left out of it deliberately, and not because a maximum cannot be given here:
    // it is because these are already measured against this very room, so there is nothing left for
    // it to add (see `pin_keeps_its_box`).
    //
    // The room is the pin's own, which is the one the file before this one left its chrome in. All
    // four carry no transport bar and draw no chrome over their media, so what the file before it
    // was decides only whether the room a picture left is a caption shorter than this one's — and
    // which way that goes, every other swap in this file already measures the same way (see
    // `pin_swap_room`).
    if pin_keeps_its_box(path) {
        let laid_out = pin_update_box(space.room.region(), shape, scale);

        return Some(PinBox::Measured(centred_at(
            (laid_out.2 - laid_out.0, laid_out.3 - laid_out.1),
            pin_swap_centre(space),
        )));
    }

    if space.maximized {
        // The room the display has, and not the box the file before this one came out at. A
        // maximized window is the one that fills its display, and a swap that measured the
        // new file against the last one would make every step a smaller step: 16:9 into a
        // 16:9 box, then 4:3 into that, then 1:1 into that, is a window walking itself down
        // to the smallest shape in the folder while the caption goes on drawing the restore
        // glyph, because the maximize is a state about the room and never stopped being one.
        // Measured against the room every time, a walk of shapes is a window of the same
        // size showing something else — which is what every other swap in this file is for.
        return Some(PinBox::Measured(pin_update_box(
            space.room.region(),
            shape,
            scale,
        )));
    }

    Some(PinBox::Measured(pin_update_box(
        pin_swap_room(space, bounds, dpi),
        shape,
        scale,
    )))
}

/// Whether a swap for this file has to wait for a box: what the file was measured by is the
/// placeholder rather than a size, and the read or the probe that would answer it is running.
///
/// Both halves are asked because either alone is the wrong question. A file really is the size of
/// the placeholder sometimes, and a box of that size held for one is the file's own size rather
/// than a wait — while a placeholder nothing is reading behind is a size that will never change,
/// and a pick that waited on it would wait forever.
pub(super) fn pin_swap_awaits(path: &Path, shape: (u32, u32)) -> bool {
    box_is_the_wait(shape) && (measure_waiting(path) || video_probe_due(&HoverFacts::read(path)))
}

/// Whether a measured shape is the wait for a box rather than one: the placeholder every measurer
/// answers with while it runs, which is what a video nobody has probed is measured by too (see
/// `measured_off_the_tick` and `video_box`).
///
/// It is square, which is the whole of why it may not be laid out as a file's shape: fitted into
/// the room it becomes a square box, and the file's own picture is drawn into that box with its
/// aspect gone — the 16:9 video in a 1:1 window.
pub(super) fn box_is_the_wait(shape: (u32, u32)) -> bool {
    shape == (office_preview::WAITING_BOX, office_preview::WAITING_BOX)
}

/// The room a swap fits the new file into: the bound in both directions — a square of the longest
/// side the pin has been given, so every shape may use the whole of it — no larger than the room
/// the display has, in the middle of the box the pin occupies now.
///
/// The display's own room is the second limit because a bound outlives the one it was measured on:
/// a window carried to a smaller display, or one the desktop was rearranged under, is fitted into
/// the room that is there rather than into the room that was, and the take-up that follows a swap
/// is kept on the display by the same clamp every other box is (see `clamp_pinned_box`). It is
/// asked of each side apart, which is what lets a display that is wider than it is tall keep its
/// width for the room while bounding the height: the square is cut to the room rather than shrunk
/// within it.
///
/// A pin with no bound yet is given the whole of that room: no size has been asked of this window
/// to fit anything inside, and the box a page of text or a sound's card came out at is a size the
/// file was drawn to rather than one the user gave the pin (see `PinnedPreview::bound`).
pub(super) fn pin_swap_room(space: PinSwapSpace, bounds: ScreenBounds, dpi: u32) -> ScreenRegion {
    let (room_width, room_height) = pinned_room(
        bounds,
        dpi,
        space.transport_bar,
        space.overlay,
        space.caption,
    )
    .room();
    let (room_width, room_height) = (room_width.max(1) as i32, room_height.max(1) as i32);

    // The bound is the room's own ceiling, and a pin without one is left the whole of it — which
    // is what the file swapped into such a pin is scaled into, at the scale its kind names.
    let (width, height) = match space.bound {
        Some(side) => (side.max(1).min(room_width), side.max(1).min(room_height)),
        None => (room_width, room_height),
    };

    centred_at((width, height), pin_swap_centre(space))
}

/// The middle of the box the pin occupies now: where a swap puts the box it comes out with, so a
/// window the hand has moved keeps its place while it changes size.
///
/// Every file a swap lays out is centred here — one fitted into the bound as well as one laid out
/// in the room — because that is the promise a swap makes about where it stands, and it is a
/// promise about the middle rather than about a corner: a box whose top left stood still would walk
/// its way across the display as its size changed (see `pin_update_content`).
pub(super) fn pin_swap_centre(space: PinSwapSpace) -> (i32, i32) {
    (
        (space.current.0 + space.current.2) / 2,
        (space.current.1 + space.current.3) / 2,
    )
}

/// The box a pinned window's media takes for another file's shape: the largest box of that shape
/// the room can hold at the scale it is given, in the middle of the room.
///
/// It is the rule every preview is fitted into its room by, asked of a box rather than of a display
/// (see `pinned_media_box`), and what it is for is the two promises a swap keeps: nothing of the new
/// media falls outside the room, and what is left of that room is shared equally — so a file of
/// another shape is the same window showing something else rather than a window that has moved.
pub(super) fn pin_update_box(
    room: ScreenRegion,
    shape: (u32, u32),
    scale: PreviewScale,
) -> ScreenRegion {
    let bounds = ScreenBounds {
        left: room.0,
        top: room.1,
        right: room.2,
        bottom: room.3,
    };
    let (width, height) = pinned_media_box(shape, bounds, scale);
    let centre = ((room.0 + room.2) / 2, (room.1 + room.3) / 2);

    centred_at((width, height), centre)
}

/// Whether a preview of this file is drawn to the box it is given rather than scaled into it by a
/// shape of its own: a page of text, a listing this app or an engine reads out of an archive, a page
/// an engine draws, and a sound's card are all measured against the room they are drawn in — there
/// is no size in the file to take a shape from — and they are the kinds `media_dimensions` answers
/// for itself rather than through the file's own dimensions (see `media_dimensions`).
///
/// It decides two things about a swap, and they are the same thing said twice: the box the new
/// file is laid out in is the display's whole room rather than the pin's bound, because a bound is
/// a size some other file came out at and none of these is drawn to it or fitted into it; and the
/// file neither takes a bound nor gives one, since a box laid out in the room is not a ceiling for
/// the files after it (see `pin_update_content` and `pin_bound_after`).
///
/// A sound's card is here for the second of those only, as it was: the box a swap gives a card is
/// the one its own measure came out at rather than the box the pin has (see `pin_update_content`).
pub(super) fn pin_keeps_its_box(path: &Path) -> bool {
    // Asked of the content, with the entry read before the lock rather than the whole
    // configuration copied out of it: the take-up that asks this one is a take-up that must
    // not be holding a lock a window message may be waiting on, and a copy of the
    // configuration is sixteen lists to clone to read one answer (see `content_of`).
    if let crate::formats::content_type::Content::Kind(kind) = content_of(path) {
        return matches!(
            kind,
            PreviewType::Text | PreviewType::Archives | PreviewType::Peazip | PreviewType::Audio
        );
    }

    is_text_preview(path)
        || previewed_as(path, PreviewType::Archives)
        || peazip_formats::is_peazip_file(path) && PreviewType::Peazip.enabled()
        || drawn_as_audio(path)
}

/// The bound a pin has once the file just taken up is the one on screen: a side the files that
/// follow may be fitted into, or nothing while the pin has none to give.
///
/// A bound the window already has is carried through a swap untouched, and that is not a question
/// about the new file at all: what a swap is measured against is the size the window was given
/// rather than the size the file it is showing came out at (see `PinnedPreview::bound`).
///
/// A pin that has none yet keeps none while the file on screen is one drawn to its own box — a page
/// of text, a listing, a sound's card — since that box is a size the file was drawn to and no
/// ceiling for anything else (see `pin_keeps_its_box`). The first file with a shape of its own takes
/// the longest side of the box it came out at, which is where a pin taken up on a picture gets its
/// bound too: what a swap hands such a pin is the box the file was laid out at against the room the
/// display has — the scale its kind names, which is the scale a hover of it would have been given —
/// and that box is the size this window has now shown it can hold.
pub(super) fn pin_bound_after(
    carried: Option<i32>,
    keeps_its_box: bool,
    content: ScreenRegion,
) -> Option<i32> {
    if let Some(bound) = carried {
        return Some(bound);
    }

    if keeps_its_box {
        return None;
    }

    Some((content.2 - content.0).max(content.3 - content.1).max(1))
}

/// The maximize a pin has once the file just taken up is on screen: the state the window was in,
/// kept by every kind but a sound's card.
///
/// A maximize is a state about the box rather than about the file, and a card is the one kind a
/// swap does not lay the new file out in the pin's own box: a card is its own size and offers no
/// maximize to stay in (`PinFrame::None`), so the file after it is laid out from the box the card
/// stands in, exactly as it is for any other card window — and a restore carried onto one would be
/// a state about a box the card has just left, with no button anywhere to reach it and the file
/// after the card measured by a maximum that is no longer standing (see `pin_update_content`).
pub(super) fn pin_restore_after(carried: Option<ScreenRegion>, card: bool) -> Option<ScreenRegion> {
    if card {
        return None;
    }

    carried
}

/// What a swap of a pinned window's file has come to, in whichever of the three shapes it can
/// take: the file the window is shown next, a video the pin is held for until the engine has
/// drawn a frame of it, or a file there is nothing to show.
///
/// The three are an enum rather than an `Option` because the middle one is the whole of what
/// this arm is for, and an `Option` cannot say it. `None` written as "not yet" would be the same
/// answer for a file a player would not start — where the walk is asked for the next file and
/// the window keeps what it has — and for a file that is about to be shown and is not yet, and
/// those two want opposite things done to the walk: one steps off it, the other carries it onto
/// whatever the pin is shown next (see `refuse_pinned_media` and `install_pinned_media`).
pub(super) enum PinSwap {
    /// The file is ready to go into the window now: everything the install is made with, and
    /// nothing left to wait for. This is what every kind but a video the media engine plays
    /// comes to, and what a video comes to once the engine has drawn a frame of it (see
    /// `PinSwapHold`).
    Ready(PinInstallable),
    /// A video the media engine has been started for, whose first frame is not in hand: the pin
    /// keeps showing the file it was showing, at the frame it had stopped on, and the swap goes
    /// on being waited for by the loop's own tick rather than by anything in this call.
    Holding(PinSwapHold),
    /// Nothing of this file to show: a player that would not start, or a file nothing could read
    /// at all. The file and the walk it was a step of are carried rather than dropped, because
    /// both are what the refusal is answered from (see `refuse_pinned_media`).
    Refused {
        path: PathBuf,
        walk: Option<PinStep>,
    },
}

/// A file a pinned window is to be shown, and the whole of what putting it there is made of.
///
/// It is one struct rather than the four or five things a swap hands back because a swap can be
/// answered now or a tick or several later — the second is a video held for its first frame — and
/// an answer that has to be taken apart into the loop's locals and put back together a tick later
/// is a chance for the two to disagree about which file is being shown, which is exactly the
/// class of bug a walk's own `from` exists to refuse (see `PinStep`).
pub(super) struct PinInstallable {
    /// The file, which is what the take-up below is asked for by name: the pin's window is
    /// re-taken-up on this path, and it is the only place a swap's own file is named.
    pub(super) path: PathBuf,
    /// The plan the load was made with — the box the media was laid out for and the level the
    /// pin plays at — kept rather than asked of the pin again, so that the media and the box it
    /// was loaded for can never come from two different plans (see `PinUpdate`).
    pub(super) update: PinUpdate,
    /// The frame. For a video the media engine plays this is the placeholder the load landed
    /// with until the wait has handed the engine's first frame to it, and the buffer the engine
    /// draws every frame after that into (see `take_native_video_frame`).
    pub(super) media: MediaData,
    /// The clock a sound's card is drawn against, which belongs to the loop rather than to the
    /// media and is `None` for every kind that is not a sound.
    pub(super) audio: Option<SwappedAudio>,
    /// The walk this file is a step of, kept rather than spent with the load: an engine that
    /// takes the file and then never draws a frame of it says so a tick or three from here, with
    /// nothing left to step on (see `pin_step_off`).
    pub(super) walk: Option<PinStep>,
}

/// Take the file a pinned window is showing down out of the way of the one replacing it: the
/// player behind it is ended and the frame this side holds of it is let go of.
///
/// It is one call rather than the three it is made of at each road into a swap because the order
/// is the whole of it — the player is ended before the frame it was playing into is dropped, so
/// nothing can hand a frame to a media that is no longer on screen — and because a swap that
/// defers its own take-down has to be able to leave the standing file exactly where it is for
/// as long as it is held (see `PinSwapHold`).
///
/// A browser is not touched here: whether the engine a document is drawn in goes or is pointed
/// at another document is a question about the *new* file's kind, and is asked where that kind
/// is known.
pub(super) fn take_down_pinned_media() {
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(ref mut existing) = *current {
            existing.cancel_background_work();
            stop_video_playback(existing);
        }
        *current = None;
    }
}

/// End the player the file a pinned window is standing on has, keeping that file and the frame
/// this side holds of it.
///
/// The take-down a swap performs is a take-down of the whole media for every kind but a video the
/// media engine plays, and that one is held rather than dropped so the window does not flash the
/// backdrop while the engine's first frame is a tick away (see `PinSwapHold`). The hold is about
/// the frame alone, and the player behind it is not part of it: a sound FFmpeg is playing, or a
/// film it is drawing in a window of its own, goes the moment the swap starts, or it is heard (or
/// seen) over the file that replaced it for as long as the hold is out — and nothing else ends
/// it, the leftover sweep standing down for as long as a pin is up (see `kill_stray_video_process`).
///
/// The media engine's own session needs nothing here: whatever the standing file was, the `play`
/// this is asked for stops what was playing before it starts anything of its own (see
/// `video_player::play`), which is why the standing file is not asked whether it is one first.
pub(super) fn stop_pinned_player() {
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(ref mut existing) = *current {
            kill_player_process(existing);
        }
    }
}

/// How the wait for a swap's first frame stands: whether the video is on screen now, whether the
/// wait is over some other way, or whether the engine is still starting.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinSwapWait {
    /// The engine has handed a frame of the file over: what the pin is showing from here is the
    /// video, and the file it was held on can be taken down.
    Arrived,
    /// The swap is not going to be waited for. What the pin is shown is what it would have been
    /// shown without the hold, and what happens to that afterwards is the standing machinery's
    /// (see `install_pinned_media`).
    Abandoned,
}

/// How the wait for a swap's first frame stands. `None` is an engine that has not handed one over
/// yet, which is a wait that goes on.
///
/// A frame in hand arrives whatever the engine's own state says, and that is the one place this
/// parts company with `player_wait`, where a window with no player behind it is not a preview.
/// The reasoning is the difference between the two questions: a player process is the *only* thing
/// a FFmpeg video exists as, so a window with none behind it is a window onto nothing, whereas a
/// session that has drawn a frame has shown the file. `failing_before_a_frame` is careful to leave
/// a session that has drawn one out of its own answer for the same reason, and a film that played
/// and then met a bad sector is a film the pin was legitimately shown — taking the window down
/// over it is `pin_media_is_alive`'s question a tick later and not this one's.
///
/// What is read before that is an engine that is gone and an engine that has said outright that
/// it cannot play the file. Neither is worth going on for: a frame that is not coming is a file to
/// install the way a swap with no hold at all would have installed it, and the machinery that
/// watches a pinned video for an engine that never drew (`pin_media_failed_before_a_frame`) is
/// already standing behind that answer and runs to a bound of its own.
///
/// Elapsed time is deliberately *not* one of the answers, and there is no bound on this wait. It
/// used to be a second, on the reasoning that a pin frozen on the previous film reads as a window
/// this app has stopped answering in, and that a refinement which costs the user a second of a
/// dead window is not one. But a first frame lands in tens of milliseconds only for a file whose
/// header is already read, and this is the swap of a 3.3 GB file on a cold read, where the bound
/// was reached before the engine had finished opening it — and a give-up that fires installs the
/// very placeholder the hold exists to remove, which is the backdrop flash this was written to
/// delete. A wait that has not ended yet costs the old picture for as long as it takes, which is
/// the honest picture of a window still loading; what ends it is the engine's word, either way.
pub(super) fn pin_swap_wait(
    frame_in_hand: bool,
    playing: bool,
    failing: bool,
) -> Option<PinSwapWait> {
    if frame_in_hand {
        return Some(PinSwapWait::Arrived);
    }

    if !playing || failing {
        return Some(PinSwapWait::Abandoned);
    }

    None
}

/// A swap of a pinned window's file that is waiting for the media engine's first frame, and what
/// is to be installed when it comes.
///
/// It is the hover's own `FirstFrameWait` in the shape a *swap* needs rather than a hover's. That
/// one holds back a window that has not been put up yet, and so has to know the hover it was for
/// is still wanted — which is what its generation and its hide count are. None of that is in
/// question here: the window is up, it belongs to the user rather than to the pointer, and the
/// file the swap is for is the file the walk landed on. What is in question is only whether the
/// engine has drawn anything yet, which is why this is an ordinary tick's question answered in
/// one place rather than a wait matched against anything.
///
/// The file being installed is carried whole rather than a name to be looked up. It is the frame
/// the engine plays into, and it is the whole of what the loop's own per-tick take must not be
/// pointed at while the wait is out: the engine's frames are pulled into *this* buffer, at this
/// file's size, and the file standing on screen is left exactly as it was. That is the whole
/// difference between a pin frozen on the last frame of the previous film and a pin showing the
/// next one drawn into the previous one's box (see `take_native_video_frame`, and the gate on
/// the loop's own take).
pub(super) struct PinSwapHold {
    /// The file to be installed once the wait is over, held in hand rather than asked of the pin
    /// again: the answer a load came back with is what the install is made of, and taking it to
    /// pieces to be carried in the loop and putting it back together a tick later is a chance for
    /// the two to disagree about which file the window is being shown (see `PinInstallable`).
    pub(super) file: PinInstallable,
    /// The arc the load was painted with, carried on rather than taken down when the load answered.
    /// The pin is still waiting for a file, and the whole of what the user is shown while it does
    /// is the file it already had with the arc turning over it — so the arc stays up across the
    /// hold, and comes down with the install (see `pin_arc_set`).
    pub(super) arc: PinArc,
}

impl PinSwapHold {
    /// Take whatever frame the engine has ready into the held file's own buffer, answering
    /// whether the wait is over either way.
    ///
    /// The take is this file's, and never the standing one's: a frame pulled into the media the
    /// pin is still showing is a frame of the new film in the old file's buffer at the old file's
    /// size, which is what the loop's own per-tick take would be doing if this were not here to
    /// do it first. What it hands the decision is the one question the wait is about — has the
    /// engine drawn anything — and it is asked of this file's engine and no other, so the standing
    /// file's player is not consulted about a session that has nothing to do with it (see
    /// `video_player::failing_before_a_frame`).
    pub(super) fn settle(&mut self) -> Option<PinSwapWait> {
        let frame_in_hand = self.file.media.take_native_video_frame();
        let playing = video_player::is_playing();
        let failing = video_player::failing_before_a_frame().is_some();

        pin_swap_wait(frame_in_hand, playing, failing)
    }

    /// The file to install, for a wait that has come to either of its ends.
    ///
    /// Consuming rather than borrowing is what lets the buffer the engine has been drawing into be
    /// moved into the window rather than copied: that buffer is the frame the pin is about to be
    /// showing, and a copy of it would be a second megabyte a frame for a video that is running.
    pub(super) fn into_installable(self) -> PinInstallable {
        let PinSwapHold { file, .. } = self;
        file
    }
}

/// Give a held swap up, ending the engine it was started for and letting the frame go.
///
/// A hold is given up on rather than installed when a pick supersedes it or the pin has closed,
/// and in both of those the engine behind it is this app's own and nobody is going to take the
/// frame it has been drawing into: an engine left running would go on decoding a file into a
/// buffer nothing reads, for as long as the app runs. The take-down is the one any road out of a
/// video uses, called on the hold's own media rather than on the standing file's — the standing
/// file may well be a film FFmpeg's player is drawing, and that player is not what this session
/// is (see `stop_video_playback`).
pub(super) fn abandon_pin_swap(hold: &mut Option<PinSwapHold>) {
    if let Some(mut held) = hold.take() {
        stop_video_playback(&mut held.file.media);
    }
}

/// The swap this tick installs, out of the two waits a pinned window can be in.
///
/// A hold is settled first, because it is the older of the two: what has come to one of the hold's
/// two ends — a frame in hand, or a wait with nothing left to wait for — goes through the same
/// install a swap answered at once goes through, and which of the two it was makes no difference.
/// A load in hand and a hold outstanding cannot both exist anyway, because a pick gives the hold up
/// before it starts the load that would answer here.
///
/// A load that has answered is put to the swap here, on this thread, and does not always finish in
/// this call: a video the media engine plays is started and then held, and what is installed for it
/// is the same file on a later tick (see `PinSwapHold`).
pub(super) fn settle_pin_swap(
    hold: &mut Option<PinSwapHold>,
    load: &mut Option<PinLoad>,
) -> Option<PinSwap> {
    let mut swap: Option<PinSwap> = None;

    if let Some(mut held) = hold.take() {
        swap = if held.settle().is_some() {
            Some(PinSwap::Ready(held.into_installable()))
        } else {
            // An engine that has not drawn anything yet. The wait goes on, and the pin keeps
            // showing the file it was frozen on with the arc still turning over it — which is the
            // whole of what the user sees, and the whole of what makes the swap worth holding for.
            *hold = Some(held);
            None
        };
    }

    // Asked only where this tick has made no swap of its own, and asked before anything started
    // for a new file, because `swap_pinned_media` starts an engine and the hold's own take-down
    // ends whichever engine is running (see `abandon_pin_swap`).
    //
    // The gate is on the *answer*, not on the hold being spent. `hold.is_none()` is just as true
    // on the tick a hold has just come to an end as it is on a tick that never had one, and the
    // first of those is the one tick there is something to install: reading the load there writes
    // its answer over an install that already holds the frame the engine drew, and the file is
    // gone with it. A load that answered into such a tick waits in its slot for the next one, which
    // is the answer that is right rather than the one that is fast.
    if swap.is_none() && hold.is_none() {
        swap = take_pin_load(load).map(swap_pinned_media);
    }

    swap
}

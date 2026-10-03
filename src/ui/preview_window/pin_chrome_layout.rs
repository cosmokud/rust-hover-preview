//! The pin's chrome: the heights and near-misses its caption, transport bar and volume button
//! are drawn at, and the level they are at.

use super::*;

/// What a pinned preview adds to the box its media is in, in the units every other margin
/// this app's placement is written in and multiplied by the display's scale: the caption
/// above the media, the transport bar below it, and the round bubble the pin collapses into.
pub(super) const PIN_CAPTION_PIXELS: f32 = 30.0;
pub(super) const PIN_TRANSPORT_PIXELS: f32 = 30.0;
pub(super) const PIN_BUBBLE_PIXELS: f32 = 44.0;
/// How far from the strip it is drawn in the pointer asks for a pinned window's chrome, and how
/// long a pin shows its chrome for whether or not the pointer is near it — a moment after the pin
/// is taken up, which is when a hand is looking for the buttons.
///
/// There is no fade between those two states, and none is wanted: the chrome is drawn into the
/// window's own rows rather than composited over anything, so a half-shown one is a strip with the
/// picture missing behind it rather than a fainter strip — and what is cheap is the state rather
/// than a level, since a chrome that has been asked for is drawn whole in the paint that notices.
pub(super) const PIN_CHROME_NEAR_PIXELS: f32 = 24.0;
pub(super) const PIN_CHROME_ARRIVAL_SECONDS: f32 = 1.5;
/// The volume a pinned preview is playing at.
///
/// A pin is given the level the tray names at the moment it is taken up — which is the level the
/// preview it came from is already playing at — and everything done to it after that belongs to
/// that window alone. A hand on the pin's volume control moves this and writes nothing: the setting
/// under `Volume → Video` stays exactly where the user put it, and the next preview begins at that
/// (see `current_video_volume`). That is why the level is kept here rather than read from the
/// configuration where it is needed.
#[derive(Clone, Copy, Default)]
pub(super) struct PinVolume {
    /// Which of the tray's two level settings this level is: `Volume → Audio` where it is a sound's
    /// card's, `Volume → Video` for every other kind.
    ///
    /// The two settings stay apart, and this is what says which one the level in hand belongs to —
    /// so that a file of another kind is played at its own setting rather than at the level the pin
    /// is holding. One knob and one bar serve both, and the bar a hand moved is a hand on the file
    /// under it, not on the file that comes to it next (see `pinned_level`).
    pub(super) audio: bool,
    /// The level this pin is playing at, 0-100.
    pub(super) level: u32,
    /// The level this pin was playing a file of the *other* kind at, and which kind that was:
    /// the level one slot cannot hold across a walk from a sound onto a film and back.
    ///
    /// Only a level `Remember` says is kept is kept, and it is only asked for while it says so —
    /// with the switch off the knob belongs to the window it was turned on and to no other file
    /// whatever, so there is nothing here for a walk to come back to (see `pin_level_is_remembered`).
    pub(super) kept: Option<(bool, u32)>,
    /// The level the player that is running now was started at. FFmpeg's player is told nothing
    /// while it runs — a level is another player begun at that level — so what says whether one is
    /// owed is this against `level`, and what the bar is drawn from is always `level`.
    pub(super) playing_at: u32,
    /// Whether the popup is open.
    pub(super) open: bool,
    /// Whether the knob is being held: a drag is a hand on the level, and what a drag must not be
    /// mistaken for is the pointer having gone away from the popup (see `refresh_pin_volume`).
    pub(super) dragging: bool,
}

/// Answer about the pin that is up with a change to it, where there is one.
pub(super) fn with_pin(change: impl FnOnce(&mut PinnedPreview)) {
    if let Some(mut pinned) = pin_state() {
        if let Some(pin) = pinned.pin_mut() {
            change(pin);
        }
    }
}

/// The level the pin that is up plays a film at, and the setting where there is no pin: what a
/// player this app starts for a pinned file is given (see `restart_pinned_player`).
pub(super) fn pinned_volume_level() -> u32 {
    pinned_level(false)
}

/// The level the pin that is up plays at, for a file of the kind named: the pin's own where the
/// level it is holding belongs to that kind — a knob turned on the bar is a hand on the file the bar
/// is under, and it is kept for the next file of that kind — the level it was playing that kind at
/// before it walked off it, where `Remember` says a level moved on a knob is kept — and the tray's
/// setting for the kind otherwise.
///
/// Reading the two apart is the whole of what a pin walking off a sound and onto a film needs: the
/// level in hand is the sound's, so a film played at `Volume → Video` — 0% and silent, as the tray
/// is set more often than not — is played at `Volume → Audio` if the sound's level is carried onto
/// it, which is a film at full volume out of a tray that says it should make no noise at all. And
/// the walk back is what one slot cannot answer on its own: a sound at 100% walked onto a film and
/// back is a sound at 100% again while `Remember` says so, and one slot has been holding the film's
/// 0% by then (see `PinVolume::kept`).
pub(super) fn pinned_level(audio: bool) -> u32 {
    pin_level_for(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.volume)),
        audio,
    )
}

/// The same answer as `pinned_level`, over a level in hand rather than over the pin it is on — the
/// one place both the read above and the take-up below are asked it, so the level a swap is made
/// with and the level a player is begun at cannot be two answers to one question.
pub(super) fn pin_level_for(holding: Option<PinVolume>, audio: bool) -> u32 {
    let tray = || {
        if audio {
            current_audio_volume()
        } else {
            current_video_volume()
        }
    };

    let Some(holding) = holding else {
        return tray();
    };

    if holding.audio == audio {
        return holding.level;
    }

    match holding.kept {
        Some((kind, level)) if kind == audio && pin_level_is_remembered(audio) => level,
        _ => tray(),
    }
}

/// What a pin is up at after it has been taken up on a file of the kind named: `pin_level_for`'s
/// answer, of that kind, with the level it was holding kept for the kind it walked off.
///
/// A file of the kind the pin is already showing carries its level whole rather than coming through
/// here, so what this is for is the walk from one kind to the other: the level in hand is that other
/// kind's and is answered by that other kind's setting, and the level it is replaced with is kept so
/// that walking back is answered by the pin rather than by whatever the setting says by then.
pub(super) fn pin_volume_taken_up(holding: Option<PinVolume>, audio: bool) -> PinVolume {
    let level = pin_level_for(holding, audio);
    // Kept where `Remember` says the level in hand is the level the next file of its kind is played
    // at, and nowhere else: with it off the knob belongs to the window it was turned on, so a level
    // that leaves this window's kind leaves it (see `pin_level_is_remembered`).
    let kept = holding.and_then(|holding| {
        pin_level_is_remembered(holding.audio).then_some((holding.audio, holding.level))
    });

    PinVolume {
        audio,
        level,
        playing_at: level,
        kept,
        ..Default::default()
    }
}

/// Whether a level moved on a pin's own knob is the level the next file of that kind is played at:
/// the `Remember` row under the level it was moved against, which is what says whether a level
/// outlives a walk onto a file of another kind (see `PinVolume::kept` and `remember_pin_volume`).
pub(super) fn pin_level_is_remembered(audio: bool) -> bool {
    CONFIG
        .lock()
        .map(|config| {
            if audio {
                config.remember_audio_volume
            } else {
                config.remember_video_volume
            }
        })
        .unwrap_or(false)
}

/// Whether the pin that is up has its volume popup open, which is what the tick's re-assertion of
/// the player's window is held off by: the popup is drawn over the media, and the media of a video
/// FFmpeg plays is that window.
pub(super) fn pin_volume_open() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.volume.open && !pin.collapsed))
        .unwrap_or(false)
}

/// Put the pin's volume popup away, answering whether it was open — which is whether the window
/// owes a repaint for it.
pub(super) fn close_pin_volume() -> bool {
    let mut closed = false;
    if let Some(mut pinned) = pin_state() {
        if let Some(pin) = pinned.pin_mut() {
            closed = pin.volume.open;
            pin.volume.open = false;
            pin.volume.dragging = false;
        }
    }

    closed
}

/// Whether the player on screen is playing, for a caller that has no transport in hand.
///
/// It is the pin's own answer where there is a pin, and `true` where there is not — which is the
/// honest answer rather than a convenient one. A player with no pin in front of it is a hover, and
/// a hover is begun at zero and given FFmpeg's own loop, so it is never the player a rewind is owed
/// to; and nothing is holding it, because there is no transport that could be holding it (see
/// `pin_is_playing` and `video_loop_tick`).
pub(super) fn pin_is_playing_current() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin_is_playing(&pin.transport)))
        .unwrap_or(true)
}

/// Whether the pin's keyboard should be asked for again: it was claimed, and Windows says the pin
/// does not have it.
///
/// This is the whole of what makes an arrow a walk of the pin's own folder rather than a key of
/// whatever is in front of it. The claim is written when a press hands the pin the keyboard and
/// dropped when the pin is activated away from it (`pin_take_focus`, `pin_release_focus`), so the
/// two can disagree for as long as it takes the tick that notices — and something *can* activate the
/// pin away: the player is launched again on every navigation, every seek and every resize, and
/// FFmpeg's window is created without `WS_EX_NOACTIVATE` (measured) and only styled after this app
/// finds it on a monitor thread's next pass. A window that is briefly activatable, in front of the
/// pin, and coming and going on every arrow press is a window an arrow press can be swallowed by.
///
/// The two facts are the whole question and neither is asked for. The claim is this app's own
/// record and needs no window to read. `GetFocus` is Windows' own answer about where the caret is,
/// and it is the only answer that cannot be wrong about a window that does not exist in any test.
pub(super) fn pin_keyboard_wanted_back(claimed: bool, has_focus: bool) -> bool {
    claimed && !has_focus
}

/// Ask the keyboard back into a pin that still holds a claim and no longer has it, answering
/// whether it was asked for.
///
/// It is `pin_take_focus` with the press left out, and everything in that function's long note
/// applies to it unchanged — including the rule that the claim is written on the *answer* rather
/// than on the asking, so a refusal by the foreground lock leaves the pin holding nothing and
/// asking again on the next tick rather than stuck believing it holds a keyboard it was never given.
pub(super) unsafe fn pin_ask_keyboard_back(hwnd: HWND) -> bool {
    if !pin_keyboard_wanted_back(pin_window::pin_claims_the_keyboard(), pin_is_focused()) {
        return false;
    }

    pin_take_focus(hwnd);
    true
}

/// Which strips of a pinned window's chrome are showing, and what is asking for them.
///
/// Only the kinds whose chrome is drawn over their media have anything to show or hide: for those
/// the caption is a strip over the top of the picture and the transport bar a strip over the
/// bottom, and each is in the way of the thing the window is for. The two are asked about *apart*,
/// because they are two different things a hand reaches for: a pointer at the bottom of a video has
/// not asked for its title bar, and a bar drawn across the picture for a hand that was nowhere near
/// it is the thing this whole arrangement exists to avoid. What asks for a strip is the pointer
/// being near that strip and nothing else — a press or a drag on the media is a hand on the
/// picture, which is not a hand asking for a title bar — and coming up is one of the moments a hand
/// is looking for a close button, so a pin shows the whole of its chrome for a moment whether or
/// not the pointer is near it (see `PIN_CHROME_ARRIVAL_SECONDS`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) struct PinChrome {
    /// Whether the pointer is asking for the caption, and for the transport bar. There is no
    /// half-shown state: a strip is drawn into the window's own rows, so one painted at a share
    /// of itself would be a strip with the picture gone behind it. Each is drawn whole or not at
    /// all (see the note over `PIN_CHROME_NEAR_PIXELS`).
    pub(super) caption: bool,
    pub(super) bar: bool,
    /// When the arrival window closes, while one is open. While it is open both strips are shown.
    pub(super) until: Option<Instant>,
}

impl PinChrome {
    /// The chrome a pinned window comes up with: both strips, and staying that way for a moment
    /// whether or not the pointer is near them.
    pub(super) fn on_arrival(now: Instant) -> Self {
        Self {
            caption: true,
            bar: true,
            until: Some(now + Duration::from_secs_f32(PIN_CHROME_ARRIVAL_SECONDS)),
        }
    }

    /// The chrome of a kind that draws it in bands around its media: always there, never hidden,
    /// and not something the pointer can ask for or away.
    pub(super) fn always() -> Self {
        Self {
            caption: true,
            bar: true,
            until: None,
        }
    }
}

/// Ask which strips of a pinned window's chrome the pointer is calling for, answering whether
/// either answer changed — which is whether the window owes a repaint, since there is nothing
/// between shown and hidden for a frame to be drawn at.
///
/// What is asked is the pointer's place on the screen rather than anything the window was sent: a
/// strip that has gone is not a region the mouse can be over, so a window that waited to be told
/// the pointer had arrived would never be told (see `pin_chrome_near`).
pub(super) fn refresh_pin_chrome(
    pin: &mut PinnedPreview,
    now: Instant,
    cursor: Option<(i32, i32)>,
) -> bool {
    // The volume popup is asked about on this clock before the strips are, and asked about for
    // every kind: it is the transport bar's own, and a kind whose chrome is never hidden still has
    // a popup that belongs on screen only while it is being used (see `refresh_pin_volume`).
    let closed = refresh_pin_volume(pin, cursor);

    // A bubble has no caption and no window of its own on the screen, so it has no name to be
    // saying and no button to say it about: a name is put away with the window it belongs to.
    if pin.collapsed {
        return closed || pin.tooltip.refresh(None, now);
    }

    // A kind whose chrome the pointer can ask for has strips the pointer asks for, and this is
    // the question about them. A kind whose chrome is in bands around its media and is never
    // hidden — a page of text, an archive listing — has a caption that is always there instead,
    // and so has nothing to be asked: `PinChrome::always` already carries that answer in the pin
    // (see `pin_hides_chrome`).
    let mut shown = false;
    if pin.hides_chrome {
        // The arrival window is spent the moment it closes, and there is no bringing it back:
        // what asks for a strip after it is the pointer and nothing else.
        if pin.chrome.until.is_some_and(|until| now >= until) {
            pin.chrome.until = None;
        }

        let (caption, bar) = match pin.chrome.until.is_some() {
            true => (true, true),
            false => pin_chrome_near(pin, cursor),
        };

        // The strip the popup came out of stays showing while it is open, whatever the pointer
        // is doing: the popup is drawn in this window's own rows above that strip, and a bar
        // that went away underneath it would take the button that opened it with it. Only a kind
        // that draws one is held up by it, though — a sound's popup hangs off a button on the
        // card rather than off a strip, and `bar` for a kind with no strip would otherwise be
        // stuck true for as long as the popup is open.
        let bar = pin.transport_bar && (bar || pin.volume.open);

        shown = (pin.chrome.caption, pin.chrome.bar) != (caption, bar);
        pin.chrome.caption = caption;
        pin.chrome.bar = bar;
    }

    // A sound's card carries its own buttons, and what the pointer is on one of them is a question
    // asked on this clock as well as on the pointer's, for the same reason the strips are: a button
    // lit under a pointer that has walked away from the window is exactly what a caption refuses to
    // let happen to its own. It is asked here and not only in `pinned_mouse_move` because a pointer
    // that leaves the window sends no more moves — the lit button would simply stay lit. It is asked
    // of every kind rather than of the ones whose chrome hides, because a card is a media frame and
    // has no arrival window to be spent in: it is showing for as long as the window is up.
    let card = pin_audio_hover_refresh(pin, cursor);

    // A name belongs to a button, and a button belongs to a pointer that is still on it. The
    // pointer is read from the cursor rather than from `pin.hovered`, because that is only
    // written on a move *this* window is sent, and a pointer that has left the pin — over a
    // modal dialog, or simply off to one side — sends it nothing more. A name left up for a
    // pointer that has gone is a caption naming a button the hand is not on. This is the same
    // reading `pin_chrome_near` does, and for the same reason: a strip that is not under the
    // pointer cannot be asked for by looking at where the pointer was.
    //
    // It is asked on every kind, because a caption is a caption whichever kind it belongs to and
    // the two hand-off buttons are on all of them — the name hangs below the strip rather than
    // inside it, so it is drawn over the media band a kind without a hidden caption still has.
    //
    // It is asked *before* the two above are answered, and that is the whole of why: `||`
    // short-circuits, so a tick that was already going to repaint because a strip came or went
    // would never reach this question — and the name it was about to start saying, or stop
    // saying, would sit wrong on the caption until the pointer happened to move again. All
    // three are asked every tick; whether any of them changed is the answer.
    let spoken = cursor
        .filter(|(x, y)| {
            let window = pin.window_box();
            let band = pin.caption;
            *x >= window.0 && *x < window.2 && *y >= window.1 && *y < window.1 + band
        })
        .and(pin.hovered);
    let spoken = pin.tooltip.refresh(spoken, now);

    spoken || shown || closed || card
}

/// The card's own control the pointer is on, asked of the cursor's place on the screen rather than
/// of anything the window was sent, and answered whether it changed — which is whether the card
/// owes itself a repaint.
///
/// A pointer resting on the card is asked about every control at once rather than one at a time,
/// because the card is a media frame and not a strip of chrome: there is no band of it to come and
/// go, so what a hand is on is settled on this clock and nowhere else.
pub(super) fn pin_audio_hover_refresh(pin: &mut PinnedPreview, cursor: Option<(i32, i32)>) -> bool {
    let window = pin.window_box();
    let hovered = cursor.and_then(|(x, y)| {
        (x >= window.0 && x < window.2 && y >= window.1 && y < window.3)
            .then(|| pin_audio_control_at(pin, x - window.0, y - window.1))
            .flatten()
    });

    let changed = pin.audio_hovered != hovered;
    pin.audio_hovered = hovered;

    // The card is a media frame and not chrome, so repainting the window alone would redraw the
    // *old* card with the old button still lit under it: what shows the wash is the card's own
    // paint, and this is the one place in the pin that has to know that.
    if changed {
        AUDIO_CARD_DIRTY.store(true, Ordering::Release);
    }

    changed
}

/// Whether the volume popup is still wanted, answering whether it has been put away — which is
/// whether the window owes a repaint for it.
///
/// It belongs to the button that opened it and to nothing else, so it is kept while the pointer is
/// anywhere in that control — the panel itself, or the strip the button sits in — and put away once
/// the pointer is away from the whole of it. A knob being held keeps it whatever the pointer is
/// doing: a drag is the pointer's, and a drag that has left the panel is a level being taken to an
/// end rather than a popup being dismissed.
pub(super) fn refresh_pin_volume(pin: &mut PinnedPreview, cursor: Option<(i32, i32)>) -> bool {
    if !pin.volume.open || pin.volume.dragging {
        return false;
    }

    // Asked of the pin this function is handed rather than of whichever pin is installed, because
    // this is the tick's reader and it must be talking about the window under the hand: a panel
    // answered against another window is a panel that closes while the pointer is still on it (see
    // `pinned_volume_panel`).
    let Some(popup) = pinned_volume_panel(pin) else {
        return false;
    };

    let margin = logical_px(pin.dpi, PIN_CHROME_NEAR_PIXELS).max(1);
    let window = pin.window_box();
    let on_the_popup = cursor.is_some_and(|(x, y)| {
        let (x, y) = (x - window.0, y - window.1);
        x >= popup.panel.left - margin
            && x < popup.panel.right + margin
            && y >= popup.panel.top - margin
            && y < popup.panel.bottom + margin
    });

    let (near_caption, near_bar) = pin_chrome_near(pin, cursor);
    if on_the_popup || near_caption || near_bar {
        return false;
    }

    pin.volume.open = false;
    true
}

/// Which of the strips a pinned window's chrome is drawn in the pointer is near — near enough that
/// the strip comes out.
///
/// They are answered one at a time rather than as one question, and that is the whole of what this
/// function is for: the caption is over the picture's first rows and the bar over its last, and a
/// hand at one end of a window is not a hand that has asked for what is at the other end of it. What
/// they have in common is the room beside them — a strip is asked for from either side of the
/// window's own edge, since a pointer that has not quite arrived is on its way there — so a pointer
/// out beyond the window's sides is near neither.
pub(super) fn pin_chrome_near(pin: &PinnedPreview, cursor: Option<(i32, i32)>) -> (bool, bool) {
    let Some((x, y)) = cursor else {
        return (false, false);
    };

    let window = pin.window_box();
    let margin = logical_px(pin.dpi, PIN_CHROME_NEAR_PIXELS).max(1);
    if x < window.0 - margin || x > window.2 + margin {
        return (false, false);
    }

    let caption = pin.caption;
    let transport = pinned_transport_height(pin.dpi, pin.transport_bar);

    // A caption drawn *over* its media comes out when the pointer is near it and nowhere else,
    // because it is in the way of the picture the window is for. A caption in a band of its own
    // above the media is in the way of nothing and is never asked for, and a kind with no caption
    // at all is not asked about either: `caption` is nothing and the first band of the window is
    // the media's own first row (see `pinned_caption_height`).
    let at_the_top = y >= window.1 - margin && y <= window.1 + caption + margin;
    let at_the_bottom =
        transport > 0 && y >= window.3 - transport - margin && y <= window.3 + margin;

    (at_the_top, at_the_bottom)
}

/// Whether a pinned window's chrome — the caption and the transport bar — is drawn *over* the
/// media rather than in bands above and below it.
///
/// Two things have to be true of a kind for that. The band has to be this app's own pixels: the
/// window FFmpeg's player has, and the page the browser draws an SVG or a font on, are windows of
/// somebody else's standing *in* the band and asserted over this one, so a caption this app drew
/// over one of those would be a caption underneath it. And the frame has to be the media's own
/// shape, which is what makes the window box and the media box the same box: the strips the chrome
/// is drawn in are then inside the picture rather than beside it, and there is no band of empty
/// window where a bar used to be.
///
/// A page that is laid out to whatever box it is given — a text preview, an archive listing — and
/// a sound's card are neither, and neither is a video below a minimized one: the first two keep
/// their bars the way they have always had them, and the player's window keeps the band it stands
/// in.
pub(super) fn pin_overlay_chrome(kind: Option<MediaType>) -> bool {
    match kind {
        Some(kind) => {
            !kind.is_engine()
                && kind != MediaType::Video
                && pin_frame(Some(kind)) == PinFrame::Shaped
        }
        None => false,
    }
}

/// The second of a pinned sound a press on its card's bar asked the file to be taken to, which
/// the preview loop answers on its next tick.
///
/// The clock a card is drawn from is the loop's, and so is the player behind a sound this app
/// plays, and a seek in one of those is a player ended and another begun. The engine Windows has
/// is the exception — it is told where to go rather than replaced — and the seek is left for the
/// loop all the same, so that a press has one door and the card is asked for again in the same
/// tick the file was taken (see `settle_pinned_audio_seek`).
pub(super) static PIN_AUDIO_SEEK: Lazy<Mutex<Option<f64>>> = Lazy::new(|| Mutex::new(None));

/// That a play/pause pressed on a pinned sound's card has been asked for, which the loop
/// answers on its next tick for the same reason a seek is left for it (see
/// `settle_pinned_audio_seek`): the clock a card is drawn from and the player behind a sound this
/// app plays are both the preview loop's, and the window procedure cannot touch either.
pub(super) static PIN_AUDIO_TOGGLE: AtomicBool = AtomicBool::new(false);

/// Ask for a pinned sound to be held or set going again, from a play/pause button on its card.
pub(super) fn ask_pin_audio_toggle() {
    PIN_AUDIO_TOGGLE.store(true, Ordering::Release);
}

/// Hold a pinned sound where it stands, or set it going again, on the loop's tick — the whole of
/// what a click on the card's own play button asks for (see `toggle_pinned_audio`).
///
/// It is asked here rather than acted on in the window procedure because the three things it acts
/// on — the clock, the player and whether either is running — are the loop's own, and a pin's
/// button that reached for them directly would be reaching under a lock this thread holds.
pub(super) fn settle_pinned_audio_toggle(
    started: &mut Option<Instant>,
    offset: &mut f64,
    paused: &mut Option<f64>,
) {
    if !PIN_AUDIO_TOGGLE.swap(false, Ordering::AcqRel) {
        return;
    }

    toggle_pinned_audio(started, offset, paused);
}

/// Where a pinned sound's clock stands after a seek has been answered, and whether the sound is
/// left being held: the player's start, the second it was started at, and the second it is held at.
///
/// A sound that was playing is taken to the second and goes on playing it, so what is left is a
/// clock begun there and no hold. A sound a key was holding is not begun again by a seek — it is
/// held at the second instead, and the player that takes the hold's place is begun there when the
/// key comes — so what is left is the hold and no clock. A player that did not come up at all is
/// neither: a card with no clock rather than a hold over a sound that is not playing.
///
/// The one answer this must never give is a hold standing over a sound a player is playing. That
/// is a card frozen at a single second over a sound going on, and — because a hold is what a later
/// press reads as a pause rather than a sound to take somewhere — a seek after it does nothing at
/// all until a key is pressed and lifted again.
pub(super) fn pinned_audio_after_seek(
    was_playing: bool,
    player_came_up: bool,
    seconds: f64,
) -> (Option<Instant>, f64, Option<f64>) {
    if !was_playing {
        return (None, seconds, Some(seconds));
    }

    (player_came_up.then(Instant::now), seconds, None)
}

/// Take a pinned sound to the second a press on its card's bar asked for, in the way the player
/// playing it answers to.
///
/// The engine is told where to go and gets there on its own clock, which is the clock the card is
/// drawn from, so nothing is written down for it. A player of this app's is ended and another
/// begun at that second, with the clock the card is measured against set to it — the same bargain
/// a resume from a hold makes, and the same one the bubble's park makes (see
/// `restart_pinned_audio`).
///
/// A seek made while a sound is held is a hold that has moved: the second it is held at is the
/// one the press named, so a key pressed afterwards lets it go from where the hand put it rather
/// than from where it was left, and the press does not begin a sound the user is not listening to
/// (see `pinned_audio_after_seek`).
pub(super) fn settle_pinned_audio_seek(
    started: &mut Option<Instant>,
    offset: &mut f64,
    paused: &mut Option<f64>,
) {
    let seconds = {
        let Ok(mut request) = PIN_AUDIO_SEEK.lock() else {
            return;
        };
        let Some(seconds) = request.take() else {
            return;
        };
        seconds
    };

    let Some((path, _)) = pinned_media_owner() else {
        return;
    };

    // Which player is playing the file is the answer of the player itself, not the probe's (see
    // `playing_player`).
    let Some(track) = audio_track::playable(&path) else {
        return;
    };

    match playing_player(&path, &track) {
        Player::Native => video_player::seek(seconds),
        Player::Ffmpeg => {
            let was_playing = paused.is_none();
            let player = was_playing
                .then(|| restart_pinned_audio(&path, seconds))
                .flatten();

            (*started, *offset, *paused) =
                pinned_audio_after_seek(was_playing, player.is_some(), seconds);
        }
    }

    // The card is at the second the hand named, and is asked for again at once (see
    // `AUDIO_CARD_DIRTY`).
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

/// Move the pin's level to `level`, giving it to a player that can be told one while it runs.
///
/// The media engine takes a level while it plays, which is what makes a knob dragged on the pin
/// heard as it moves — and the engine plays sounds as well as videos, so a knob dragged on a
/// sound's card is heard as it moves for exactly the same reason. FFmpeg's player takes one only by
/// being started at it, so nothing is given there and nothing is marked as settled: a level marked
/// as given to a player that was never told it is a level `settle_pin_volume` has nothing left to
/// do about, which is what a sound normalized on FFmpeg used to get (see
/// `audio_preview_level_is_owed_to_the_engine`). Where the level is remembered, the configuration
/// is written as the knob moves rather than at the end of the drag, because the next sound's player
/// is started from it and a drag is not over until the hand lets go (see
/// `remember_pin_volume`).
///
/// A level moved on a knob is recorded as the kind of the file under it as well, which is what makes
/// it the pin's own: the pin keeps the level a hand moved on a bar for the next file *of that kind*
/// and not for the next file whatever it is (see `pinned_level`). A walk off a sound and onto a film
/// is answered by the tray, and a knob turned on the film's bar afterwards makes the level the
/// film's, again.
pub(super) fn set_pin_volume(level: u32) {
    let level = level.min(100);
    let audio = matches!(current_media_type(), Some(MediaType::Audio));
    with_pin(|pin| {
        pin.volume.level = level;
        pin.volume.audio = audio;
    });

    let immediate = match current_media_type() {
        Some(MediaType::NativeVideo) => true,
        Some(MediaType::Audio) => audio_preview_level_is_owed_to_the_engine(),
        _ => false,
    };
    if immediate {
        video_player::set_volume(level);
        with_pin(|pin| pin.volume.playing_at = level);
    }

    remember_pin_volume(level);

    // The speaker on a sound's card says what the level is, and the card is a media frame rather
    // than chrome: repainting the window alone would redraw the *old* card with the old speaker on
    // it. A video's level is drawn in the strip, which the repaint does reach (see
    // `pin_audio_hover_refresh`).
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

/// Write a pin's level into the configuration where the setting beside it says the level is
/// remembered, and only then: this is the configuration a *later* player is started from, so a
/// knob that is not remembered leaves the next sound where this one is.
///
/// The file is written once and when the knob is let go of, not once per pixel of a drag (see
/// `settle_pin_volume`).
pub(super) fn remember_pin_volume(level: u32) {
    if let Ok(mut config) = CONFIG.lock() {
        match current_media_type() {
            Some(MediaType::Audio) if config.remember_audio_volume => {
                config.audio_volume = level;
            }
            Some(MediaType::NativeVideo) | Some(MediaType::Video)
                if config.remember_video_volume =>
            {
                config.video_volume = level;
            }
            _ => {}
        }
    }
}

/// Write the file a pin's level is remembered in once the knob is let go of, so that a drag is one
/// write rather than the hundred the knob moved through.
///
/// One write, and it is the write the `Remember` row beside the level promises: where the row is on
/// the file is brought up to the level in hand, and where it is off nothing is written at all (see
/// `put_remembered_pin_volume`). It is asked for whether or not a player was owed a level: a knob
/// turned on a paused pin settles nothing to a player but is still a level the next sound is wanted
/// at (see `remember_pin_volume`).
pub(super) fn save_remembered_pin_volume(level: u32) {
    let Ok(mut config) = CONFIG.lock() else {
        return;
    };
    if put_remembered_pin_volume(&mut config, current_media_type(), level) {
        config.save();
    }
}

/// Put the level a pin's knob was let go at where `Remember` says a level moved on that knob is
/// kept, answering whether the file is owed a write because of it.
///
/// Nothing here compares the level against what the configuration already holds, and that is the
/// whole of what it used to get wrong: by the time the knob is let go the configuration holds
/// exactly this level, because every step of the drag put it there and so did the click that moved
/// nothing before the release (`remember_pin_volume`). A comparison of the two is a comparison of a
/// thing with itself, which is why the file was never written at all.
///
/// What it answers is asked of a configuration in hand rather than of the running one, so the
/// question can be asked of a level without a file being written for the asking.
pub(super) fn put_remembered_pin_volume(
    config: &mut AppConfig,
    kind: Option<MediaType>,
    level: u32,
) -> bool {
    match kind {
        Some(MediaType::Audio) if config.remember_audio_volume => {
            config.audio_volume = level;
            true
        }
        Some(MediaType::NativeVideo) | Some(MediaType::Video) if config.remember_video_volume => {
            config.video_volume = level;
            true
        }
        _ => false,
    }
}

/// What a level let go of on a pinned FFmpeg video is settled by.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinLevelSettling {
    /// A player is behind the pin and takes a level only by being begun at one, so it is replaced:
    /// begun at the pin's own level, and — where the file is held — begun holding it, because the
    /// replacement is the player that will be asked to hold (see `restart_pinned_player`).
    Relaunch { holding: bool },
    /// No player is behind the pin at all. Nothing is begun: the level is written down as the one
    /// playing, and the player that takes it is the one a later press begins at it.
    Recorded,
}

/// Whether a level moved on a pinned FFmpeg video is owed to a player replaced or written down, and
/// whether that player is begun held.
///
/// The question is whether a *player* is behind the pin rather than whether the film is playing,
/// which is what makes a held file owe the level at all: a held file is still a file playing at a
/// level, and the knob moved while it was held would otherwise reach a player that never learns of
/// it — the film would go on at the level it was at when the hand stopped it while the bar claimed
/// the level it was set to (see `settle_pin_volume`).
///
/// The hold travels across the replacement rather than being dropped by it, so a level turned while
/// the file was held does not hand the film back to the sound: the level and the hold are the same
/// relaunch, and a player begun playing where the file was held is a file the bar says is held and
/// the desk says is not.
///
/// A player this app has lost the handle to is still a player while it is running, so a process
/// behind a bar that has stopped claiming playback is still owed the level; only a pin with neither
/// a claim nor a process is owed nothing.
///
/// Both answers are taken before the decision rather than one short-circuiting the other, so a knob
/// let go of on a playing film reconciles the player behind it exactly as one let go of on a held
/// film does — and the lock that costs is one the relaunch is about to take anyway (see
/// `restart_pinned_player`).
pub(super) fn pin_level_settled_by(playing: bool, running: bool, held: bool) -> PinLevelSettling {
    if playing || running {
        PinLevelSettling::Relaunch { holding: held }
    } else {
        PinLevelSettling::Recorded
    }
}

/// Give the pin's level to the player it is owed to, where giving it costs a player replaced.
///
/// FFmpeg's player is told nothing once it is running, so the level it is to play at is another
/// player begun at it — the same bargain a seek makes, and it is taken to the second the hand let
/// go of the knob at rather than back to the beginning (see `pin_playhead`).
///
/// A pin that is *held* owes it too, which it did not before a hold could be a running player, and
/// what is asked in that arm is which of two things the level is settled by (see
/// `pin_level_settled_by`). What is genuinely owed nothing is a pin whose player has gone, where the
/// file is held because there is nothing holding it: the player that begins when the file is let go
/// of is begun at the level this reads.
///
/// A sound is the one kind whose settling is the loop's rather than this one's: the clock behind
/// its card and the player a level is owed to are both the loop's, so the door is a flag rather
/// than a call (see `settle_pinned_audio_volume`).
///
/// A level that is remembered is written to the file here, before anything is settled: this is
/// where the knob is let go of, and it is the one moment a drag is over (see
/// `save_remembered_pin_volume`).
pub(super) fn settle_pin_volume() {
    let Some((path, content, transport, volume)) = pinned_playback_state() else {
        return;
    };
    save_remembered_pin_volume(volume.level);
    if volume.playing_at == volume.level {
        return;
    }

    match current_media_type() {
        Some(MediaType::NativeVideo) => {
            video_player::set_volume(volume.level);
            with_pin(|pin| pin.volume.playing_at = pin.volume.level);
        }
        Some(MediaType::Video) => match pin_level_settled_by(
            pin_is_playing(&transport),
            is_video_process_running(),
            transport.paused_at.is_some(),
        ) {
            PinLevelSettling::Relaunch { holding } => {
                let playhead = pin_playhead(&transport).unwrap_or(0.0);
                restart_pinned_player(&path, content, playhead, holding);
            }
            PinLevelSettling::Recorded => {
                with_pin(|pin| pin.volume.playing_at = pin.volume.level);
            }
        },
        Some(MediaType::Audio) => {
            let playing = audio_preview_level_is_owed_to_the_engine();
            if playing {
                // The live path already gave it to the engine, so this arm is normally a no-op —
                // and it is kept for the case where the knob was turned before a session was
                // there to take it (see `set_pin_volume`).
                video_player::set_volume(volume.level);
                with_pin(|pin| pin.volume.playing_at = pin.volume.level);
            } else {
                // Otherwise the player that takes a level only by being started at one is owed it,
                // and the clock behind a sound's card is the loop's: so the flag, and the loop
                // starts the player at the pin's level on its next tick (see
                // `settle_pinned_audio_volume`).
                PIN_AUDIO_VOLUME.store(true, Ordering::Release);
            }
        }
        _ => {}
    }
}

/// That a pin's level has been settled for a sound, which the loop answers on its next tick
/// because the clock behind a sound's card and the player that takes a level are both its own
/// (see `settle_pin_volume`).
pub(super) static PIN_AUDIO_VOLUME: AtomicBool = AtomicBool::new(false);

/// Whether the sound on screen is played by the media engine, which is the one player of the two
/// that can be told a level while it is running. It is asked of the player and not of the probe,
/// for the same reason the clock is (see `playing_player`).
pub(super) fn audio_preview_level_is_owed_to_the_engine() -> bool {
    let Some((path, _)) = pinned_media_owner() else {
        return false;
    };
    let Some(track) = audio_track::playable(&path) else {
        return false;
    };

    playing_player(&path, &track) == Player::Native
}

/// Restart the player behind a pinned sound at the pin's level, from the second the hand let go of
/// the knob at.
///
/// The player is the loop's and the clock is the loop's, which is why this is the loop's tick and
/// not the release: a knob dragged across a card is a hundred releases a second, and a player
/// begun for each of them is a sound that never settles. What is settled at the end of the drag is
/// one player at one level from one second, which is the whole of what the transport bar's own
/// knob does for a video (see `settle_pin_volume`).
pub(super) fn settle_pinned_audio_volume(
    started: &mut Option<Instant>,
    offset: &mut f64,
    paused: &mut Option<f64>,
) {
    if !PIN_AUDIO_VOLUME.swap(false, Ordering::AcqRel) {
        return;
    }

    let Some((path, _)) = pinned_media_owner() else {
        return;
    };

    // Only the player of this app's is owed anything, and only while it is the one playing: the
    // engine was given the level as the knob moved, and there is nothing left to settle.
    let Some(track) = audio_track::playable(&path) else {
        return;
    };
    if playing_player(&path, &track) == Player::Native {
        return;
    }

    // A sound that was playing goes on playing from the second the knob was let go at, and a
    // sound that was held stays held there — the same two answers a seek gives, because a level
    // moved is a player begun and a player begun at a second is a seek (see
    // `pinned_audio_after_seek`).
    let seconds = match paused {
        Some(at) => *at,
        None => audio_clock(&path, *started, *offset, None).0.unwrap_or(0.0),
    };
    let was_playing = paused.is_none();
    let player = was_playing
        .then(|| restart_pinned_audio(&path, seconds))
        .flatten();

    // A level of nothing is a player begun at nothing rather than no player at all, so this is
    // `pinned_audio_after_seek` written as it is for every other level: a clock begun at the second
    // the knob was let go at, and a sound going on in silence (see `start_audio_playback_at`).
    (*started, *offset, *paused) = pinned_audio_after_seek(was_playing, player.is_some(), seconds);

    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

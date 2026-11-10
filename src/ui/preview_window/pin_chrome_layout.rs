//! The pin's chrome: the heights and near-misses its caption, transport bar and volume button
//! are drawn at, and the level they are at.

use super::pin_trace::note_pin_trace;
use super::playback_trace::note_playback_trace;
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
/// There is no fade between those two states, and none is wanted: the chrome is carried in
/// windows of its own rather than composited over anything, so a half-shown one is a strip with
/// the picture missing behind it rather than a fainter strip — and what is cheap is the state
/// rather than a level, since a chrome that has been asked for is drawn whole in the paint that
/// notices.
pub(super) const PIN_CHROME_NEAR_PIXELS: f32 = 24.0;
/// The arrival window tests still exercise; no real install arms one — every pin comes up hidden
/// and only the pointer asks a strip back (see `PinChrome::hidden`).
#[cfg(test)]
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
    /// The level this pin is playing at, 0-100: the *applied* level, which the popup is drawn at
    /// and the engine is told. A wheel notch moves `target` and this glides there (see
    /// `ramp_pin_volume`).
    pub(super) level: u32,
    /// The level a wheel notch asked for, in the wheel's own fine units (1/24 percent, 120 to a
    /// notch), while the applied level is still on its way there, and nothing once it has landed.
    ///
    /// Fine units rather than whole percent because that is where a notch is exact: a 5% notch is
    /// 120 units and the applied level settles on it precisely.
    pub(super) target: Option<u32>,
    /// Where the glide has got to, in `target`'s units; only meaningful while `target` is set.
    pub(super) applied: u32,
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
    /// The popup's box as it stood when the drag began, held so the panel and the groove the hand is
    /// measured against are the ones that were drawn under it: a pin whose chrome or card re-lays
    /// out mid-drag would otherwise move the groove out from under the pointer and the thumb would
    /// jump (see `pinned_volume_panel`).
    pub(super) held: Option<pin_chrome::VolumePopup>,
    /// Whether the level in hand is still owed to the engine, which one write per preview tick
    /// coalesces a drag or wheel burst to the latest level (see `flush_pin_volume_level`).
    pub(super) pending: bool,
    /// Ticks left before a level is quiet enough to settle, FFmpeg's player taking one only by
    /// being begun again (see `settle_pin_volume_after_quiet`).
    pub(super) quiet: u32,
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

    let taken = PinVolume {
        audio,
        level,
        playing_at: level,
        kept,
        ..Default::default()
    };
    note_playback_trace!(
        "volume takeup audio {} level {} kept {}",
        u8::from(audio),
        taken.level,
        u8::from(taken.kept.is_some())
    );
    taken
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

/// A level one of the tray's `Volume` rows names was moved, and the pin up on that kind is
/// told it now rather than at the next file it is shown.
///
/// A pin's level is its own, which is the whole of what a knob is for and why the setting is
/// not read back where a level is needed: the two are separate on purpose, and a level moved
/// on a knob never belongs to the tray (see `PinVolume`). What this is for is the one road the
/// other way — a row moved in the tray is a hand on the level whatever is already up should
/// be heard at, and a pin left at the level it was taken up at goes on being heard at that
/// until the window is closed and opened again, which reads as a tray that has stopped
/// working.
///
/// It is the knob's own road rather than a second one beside it: the level is moved exactly as
/// a drag moves it and settled exactly as a release settles it, so every route answers with
/// what it always has — mpv and the media engine are told the level while they run, FFmpeg's
/// player is told it at its mixer session, and a *held* film is left holding its
/// level for the player that begins when it is let go of (see `set_pin_volume` and
/// `settle_pin_volume`).
///
/// The other half's row is about the other kind of file, and a pin answers only the half it is
/// playing: a sound's level moved is not a hand on a film, which is the same answer a film
/// stepped onto from a sound is given (see `pinned_level`). Where no pin of that kind is up, the row
/// is a hover's player's instead, which is FFmpeg's and is told at the mixer (see
/// `ask_the_ffplay_its_level`).
pub(super) fn tray_volume_moved(audio: bool) {
    let held = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.volume.dragging));
    let row_level = if audio {
        current_audio_volume()
    } else {
        current_video_volume()
    };
    note_playback_trace!(
        "volume tray audio {} level {} dragging {}",
        u8::from(audio),
        row_level,
        u8::from(held == Some(true))
    );
    // A hand holding the knob owns the level: a publisher of any kind — the config watcher, the
    // tray's own reload, a wheel notch — must not move the thumb under the pointer.
    if held == Some(true) {
        note_playback_trace!(
            "volume tray audio {} level {} dragging 1 left to the hand",
            u8::from(audio),
            row_level
        );
        return;
    }

    let kind = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.volume.audio));
    if kind == Some(audio) {
        set_pin_volume(row_level);
        settle_pin_volume();
        return;
    }

    // **A hover follows the row too**, and it is not a pin's level to settle: the player is FFmpeg's
    // and its level is its mixer session's, asked for at once (see `ask_the_ffplay_its_level`).
    let on_screen = current_media_type();
    let row = if audio {
        Some(MediaType::Audio)
    } else {
        Some(MediaType::Video)
    };
    if !pinned() && on_screen == row {
        ask_the_ffplay_its_level(row_level);
    }
}

/// Whether the pin that is up has its volume popup open, which is what the tick's re-assertion of
/// the player's window is held off by: the popup is drawn over the media, and the media of a video
/// FFmpeg plays is that window.
pub(super) fn pin_volume_open() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.volume.open && !pin.collapsed))
        .unwrap_or(false)
}

/// The level the pin that is up is at, which is what a wheel notch over its volume control
/// is taken from — the popup and the button are the same level (see `set_pin_volume`).
pub(super) fn pin_volume_level() -> Option<u32> {
    pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.volume.level))
}

/// Whether the pin that is up is showing any strip of its chrome — a pin that is up
/// rather than collapsed into its bubble, with its caption or its transport bar drawn.
/// This is the popup's hold-off, generalized: a strip is painted into the pin's own
/// rows, and the rows of a video FFmpeg plays are that window, which is the one on top
/// for as long as no strip is showing. Both raises of it are held off while any strip
/// is up (see `apply_noactivate_to_hwnd`, `ensure_video_window_topmost`), and the tick
/// answers a strip coming up with a raise of the pin's own window in return.
pub(super) fn pin_chrome_up() -> bool {
    pin_state()
        .and_then(|pinned| {
            pinned
                .pin()
                .map(|pin| !pin.collapsed && (pin.chrome.caption || pin.chrome.bar))
        })
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
            pin.volume.held = None;
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
    // **A pin waiting on its hosted film's frame is holding that film, whatever its transport
    // says.** The player is begun paused on the frame the hover froze on (see
    // `pin_film_host::settle_pin_handover`), and the transport's clock is the only thing that says
    // otherwise — so the loop is told the truth here, and a film frozen a tenth of a second from its
    // end is not rewound and set going underneath the wait.
    if pin_film_host::pin_handover_armed() {
        return false;
    }

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

    pin_take_focus(hwnd, &Win32PinWindow);
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
    /// half-shown state: a strip stands in a window of its own, so one painted at a share of
    /// itself would be a strip with the picture gone behind it. Each is drawn whole or not at
    /// all (see the note over `PIN_CHROME_NEAR_PIXELS`).
    pub(super) caption: bool,
    pub(super) bar: bool,
    /// When the arrival window closes, while one is open. While it is open both strips are shown.
    pub(super) until: Option<Instant>,
}

impl PinChrome {
    /// The chrome a pinned window comes up with in the tests that still exercise an arrival
    /// window: both strips, and staying that way for a moment whether or not the pointer is near
    /// them. Nothing in the app arms one: every install comes up hidden.
    #[cfg(test)]
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

    /// The chrome of a kind that hides it after a step: neither strip showing, and no arrival
    /// window armed — what the swap's cover shows under, until the pointer asks for a strip back
    /// (see `take_up_pinned_window`).
    pub(super) fn hidden() -> Self {
        Self {
            caption: false,
            bar: false,
            until: None,
        }
    }
}

/// Whether the chrome is held down by something other than the pointer: a hosted film's handover.
/// The cheap term only, so a hand that moves over the pin can ask it; the player read is
/// `refresh_pin_chrome`'s alone (see `film_window_without_a_picture` there).
pub(super) fn pin_chrome_is_held_down() -> bool {
    pin_film_host::pin_handover_armed()
}

/// Ask the chrome question now, outside the loop's tick, answering whether it changed. The caller
/// decides what that is owed — the strips are synced where they are (see `refresh_pin_chrome`).
pub(super) fn refresh_the_pin_chrome_now() -> bool {
    let Some(mut pinned) = pin_state() else {
        return false;
    };
    let Some(pin) = pinned.pin_mut() else {
        return false;
    };
    refresh_pin_chrome(pin, Instant::now(), cursor_screen_point())
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

    // The box a wheel notch over the volume control belongs to is published on this clock as well,
    // because the tick is where the panel's own state is asked of the pin rather than of a lock the
    // hook cannot take (see `publish_pin_volume_region`).
    publish_pin_volume_region(Some(pin));

    // The card's menu is asked about on this clock as well: it is a panel over the card, and
    // where the pointer is on it is a question the same tick answers — the row under the hand
    // to wash, the Seek row's hover timer, and whether a flyout that is up is still wanted (see
    // `refresh_pin_menu`). It is asked of every kind and answers nothing where no menu is up.
    let menu = refresh_pin_menu(pin, now, cursor);

    // A bubble has no caption and no window of its own on the screen, so it has no name to be
    // saying and no button to say it about: a name is put away with the window it belongs to.
    if pin.collapsed {
        return closed || menu || pin.tooltip.refresh(None, now);
    }

    // A kind whose chrome the pointer can ask for has strips the pointer asks for, and this is
    // the question about them. A kind whose chrome is in bands around its media and is never
    // hidden — a page of text, an archive listing — has a caption that is always there instead,
    // and so has nothing to be asked: `PinChrome::always` already carries that answer in the pin
    // (see `pin_hides_chrome`).
    let mut shown = false;
    if pin.hides_chrome {
        // A pin waiting on its hosted film's first frame shows no chrome meanwhile: the strips
        // belong to the pin-mode frame, and that frame is not on screen yet. The handover shows
        // no chrome either — it is the pin's first frame, not an arrival, and only the pointer
        // asks a strip back (see `pin_film_host::settle_pin_handover`).
        //
        // A file step holds the strips down while it is on its way: the strips go down with the
        // step, not with the ask, and the file that replaces it asks them back (see
        // `pin_chrome_is_held_down`).
        //
        // A hosted film whose player has a window but no frame yet is a band that would be black
        // behind an *arrival's* chrome: the strips of an arrival window wait for the picture the
        // player itself says it drew. What they do not wait for is the pointer's chrome, which is
        // the strip a hand is already resting on: a file step's incoming player has its window under
        // the host for a tick or more before its frame is proved, and a step is not an arrival —
        // taking the hand's strip down there is the blink a step used to have (see
        // `take_up_pinned_window`).
        let pid = VIDEO_PID.load(Ordering::SeqCst);
        let film_window_without_a_picture = pin_film_host::the_hosted_player_is_drawing(pid)
            && !pin_film_host::the_hosted_frame_is_up(pid);
        let (caption, bar) = if pin_chrome_is_held_down() {
            (false, false)
        } else {
            // The arrival window is spent the moment it closes, and there is no bringing it back:
            // what asks for a strip after it is the pointer and nothing else.
            if pin.chrome.until.is_some_and(|until| now >= until) {
                pin.chrome.until = None;
            }

            match pin.chrome.until.is_some() {
                true if film_window_without_a_picture => (false, false),
                true => (true, true),
                false => {
                    let (caption, bar) = pin_chrome_near(pin, cursor);
                    // Near is only geometric: a pointer over a window standing over the pin is near
                    // it without hovering it, so the strips stay hidden there (see
                    // `pin_chrome_occluded`).
                    let (caption, bar) = if (caption || bar) && pin_chrome_occluded(pin, cursor) {
                        (false, false)
                    } else {
                        (caption, bar)
                    };
                    // **A strip that is already on the screen stays while the film it belongs to has
                    // no frame yet**: the band behind it is black either way, and taking it out from
                    // under a hand resting on it is the blink a step's wait used to be. What the gate
                    // above keeps is a strip that is *down* down (see `PinChrome::hidden`).
                    if film_window_without_a_picture {
                        (caption && pin.chrome.caption, bar && pin.chrome.bar)
                    } else {
                        (caption, bar)
                    }
                }
            }
        };

        // The strip the popup came out of stays showing while it is open, whatever the pointer
        // is doing: the popup is drawn in this window's own rows above that strip, and a bar
        // that went away underneath it would take the button that opened it with it. Only a kind
        // that draws one is held up by it, though — a sound's popup hangs off a button on the
        // card rather than off a strip, and `bar` for a kind with no strip would otherwise be
        // stuck true for as long as the popup is open.
        // A strip the popup came out of stays showing while it is open, whichever of the bar's two
        // own controls opened it (see `TransportPart::Volume` and `TransportPart::Fps`).
        let bar = pin.transport_bar && (bar || pin.volume.open || pin.fps.open);

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

    // And the playlist band under a card, which is a band of rows the same way: the pointer that
    // has left the pin sends it no more moves, and a row left marqueed under one that is somewhere
    // else is a name travelling across a row nobody is looking at (see
    // `pin_playlist::pin_playlist_hover_refresh`, which asks the window the pointer is in for the
    // row rather than the box the pin is laid out in).
    let playlist =
        pin_playlist::pin_playlist_hover_refresh(pin, pin_lifecycle::pinned_window(), cursor);

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

    // And the find bar's card over a pinned page, which is a card over a media band the same way a
    // sound's is: the pointer that has left the pin sends it no more moves, and a control left lit
    // under a pointer that is somewhere else is exactly what a caption refuses to let happen to its
    // own. It is asked of the pin in hand rather than through `text_find::hover`, which asks for this
    // lock again (see `text_find::hover_over`).
    let window = pin.window_box();
    let find = cursor
        .filter(|(x, y)| *x >= window.0 && *x < window.2 && *y >= window.1 && *y < window.3)
        .map(|(x, y)| text_find::hover_over(pin, x - window.0, y - window.1))
        .unwrap_or_else(|| text_find::hover_over(pin, -1, -1));

    spoken || shown || closed || card || playlist || menu || find
}

/// The card's own control the pointer is on, asked of the cursor's place on the screen rather than
/// of anything the window was sent, and answered whether it changed — which is whether the card
/// owes itself a repaint.
///
/// A pointer resting on the card is asked about every control at once rather than one at a time,
/// because the card is a media frame and not a strip of chrome: there is no band of it to come and
/// go, so what a hand is on is settled on this clock and nowhere else. What the hand must also be
/// near is the card's own command bar, so that a hand on the artwork asks for nothing (see
/// `pin_audio_control_revealed`).
pub(super) fn pin_audio_hover_refresh(pin: &mut PinnedPreview, cursor: Option<(i32, i32)>) -> bool {
    let window = pin.window_box();
    let hovered = cursor.and_then(|(x, y)| {
        (x >= window.0 && x < window.2 && y >= window.1 && y < window.3)
            .then(|| pin_audio_control_at(pin, x - window.0, y - window.1))
            .flatten()
            .filter(|control| pin_audio_control_revealed(pin, y - window.1, *control))
    });

    // The two window buttons come and go with the hand as well, and are
    // asked here as well as on every move because the tick is what paints
    // the card again: a hand in the top band — the band as tall as the
    // buttons' own hit boxes — and across the columns the buttons stand in
    // is a hand near the buttons themselves, and a hand anywhere else is a
    // hand asking for the name's corner back (see
    // `pin_audio_window_buttons_at`).
    let shown = cursor.is_some_and(|(x, y)| {
        x >= window.0
            && x < window.2
            && y >= window.1
            && y < window.3
            && pin_audio_window_buttons_at(pin, x - window.0, y - window.1)
    });

    let changed = pin.audio_hovered != hovered || pin.audio_window_buttons != shown;
    pin.audio_hovered = hovered;
    pin.audio_window_buttons = shown;

    // And the facts line, which is the one band of the card that is not a control: a pointer resting
    // on it is what engages its scroll, and one that has left it winds the line home (see
    // `pin_audio_facts_at` and `audio_preview::NameScroll::set_engaged`).
    let on_the_line = cursor.is_some_and(|(x, y)| {
        x >= window.0
            && x < window.2
            && y >= window.1
            && y < window.3
            && pin_audio_facts_at(pin, x - window.0, y - window.1)
    });
    let changed = changed || pin.audio_facts_hovered != on_the_line;
    pin.audio_facts_hovered = on_the_line;

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
    let Some(button) = pinned_volume_button(pin) else {
        return false;
    };

    let margin = logical_px(pin.dpi, PIN_CHROME_NEAR_PIXELS).max(1);
    let window = pin.window_box();
    let on_the_popup = cursor.is_some_and(|(x, y)| {
        let (x, y) = (x - window.0, y - window.1);
        // The button itself counts as the control it is, and not only what its margin
        // reaches: a wheel notch there opens the panel, and one that closed it on the next
        // tick would flash a panel the user is still turning.
        (x >= button.left && x < button.right && y >= button.top && y < button.bottom)
            || (x >= popup.panel.left - margin
                && x < popup.panel.right + margin
                && y >= popup.panel.top - margin
                && y < popup.panel.bottom + margin)
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
///
/// Near is geometric only: whether the pointer is really over the pin rather than over a window
/// standing over it is the hover question's own, asked where the answer is used (see
/// `pin_chrome_occluded`).
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

/// Whether the pointer the chrome's hover question is about is over another window rather than
/// the pin's own: a hand past the pin's edge counts as near, but one over a window standing over
/// the pin is a hand somewhere else — Notepad over the pin is not a hand asking for the pin's
/// caption (see `refresh_pin_chrome`).
///
/// **A window of the film the pin is showing is the pin, not a window over it.** A player begun for
/// the pin is born at the band and shows itself there before the host has it (see
/// `pin_film_thumbnail::the_source_launch_origin`), so the hand resting on a strip during a step is
/// over a window of a player this app began for a film — the incoming one, or the outgoing one it is
/// kept on screen until the incoming one draws — and a film the pin is showing is not a program
/// standing over it. The film's host answers the same way for a step whose player is already drawing
/// in it.
///
/// **The player is asked of the app's own record rather than of the pid the host happens to hold**,
/// and that is what makes the answer the same at every tick of a handover: the host names whichever
/// player the step last recorded, while the window under the hand is whichever of the two is in
/// front, and a rule keyed on that one pid answers the same window as somebody else's on the ticks
/// where the two disagree — the intermittent strip (see `engine_processes::holds_a_player`).
///
/// **A window of this app's own is the pin, not a window over it**, whether it is a window of a film
/// the pin is showing or one of the app's own: with `Pin Mode → Always On Top` off a pin on the screen
/// wears the taskbar stand-in, which is a window of this app's own and can stand over the pin's
/// caption — the intermittent strip the user's trace named, by a window that was neither the film's
/// host nor the player. The question of whose window it is is the one the rest of the app asks (see
/// `pin_lifecycle::pid_is_the_previews`), asked here of the same pids.
///
/// Asked of the window under the point rather than of the order, because the pin's children —
/// the strips — are disabled and `WindowFromPoint` skips them, so the pin's own window comes back
/// wherever the pointer is genuinely over the pin's chrome; the film host standing over the band is
/// read as the pin's own as well, because the band is the pin's picture (see
/// `a_window_of_the_pins_own`). A pin with no window of its own has nothing standing over it to ask
/// about, so there the geometric question stands as it did.
///
/// **The margin is no place a window can stand over the pin**: `pin_chrome_near` reaches
/// `PIN_CHROME_NEAR_PIXELS` past the pin's own edge, and the desktop there answers whatever
/// the hand left — Explorer, the desktop — which is not a cover of this pin but the world
/// outside it. So only a point inside `pin.window_box()` is asked what stands under it, and
/// the hand one pixel above the caption gets the caption it is reaching for.
fn pin_chrome_occluded(pin: &PinnedPreview, cursor: Option<(i32, i32)>) -> bool {
    let Some((x, y)) = cursor else {
        return false;
    };
    let window = pin.window_box();
    if x < window.0 || x >= window.2 || y < window.1 || y >= window.3 {
        return false;
    }
    let Some(hwnd) = pin_real_window::real_pin_window() else {
        return false;
    };
    let under = window_at_point(x, y);
    let occluded = under != hwnd.0 as isize && !a_window_of_the_pins_own(under);
    if occluded {
        note_pin_trace!(
            "chrome occluded by {:?} over the pin {:?}: the strips the pointer asked for go down",
            HWND(under as *mut _),
            hwnd,
        );
    }
    occluded
}

/// Whether a window under the pointer belongs to the pin rather than to somebody else: the pin's own
/// window and its children, the film's host — a window of this app's standing at the band the pin is in
/// — a window of any player this app began for a film, and a window of this app's own process at all
/// (see `pin_chrome_occluded`). No window at all under the point is nobody's, which is the same answer
/// as the pin's: nothing stands over a pin in open space.
///
/// The player is asked for by its process, which is the same question the hover's own probe asks of
/// a film's window — and of the app's record of the players it began rather than of the host, because
/// the host names one player and the handover has two (see `pin_chrome_occluded`). A window of any
/// other process is somebody else's however close it stands.
fn a_window_of_the_pins_own(under: isize) -> bool {
    // **Nothing at all under the point is nothing standing over the pin.** `WindowFromPoint` answers
    // no window where no window is there — a hand past every window the desktop has — and a rule that
    // read that as somebody else's would take a strip away from a pointer that is near the pin in
    // open space.
    if under == 0 {
        return true;
    }

    // **The pin's own children are the pin.** The chrome strips are children of its window, and a
    // strip that `WindowFromPoint` does answer (a platform or a setting that lets it) is the chrome
    // the pointer is asking for rather than a window over it.
    if pin_real_window::real_pin_window()
        .is_some_and(|pin| unsafe { GetAncestor(HWND(under as *mut _), GA_ROOT) == pin })
    {
        return true;
    }

    if pin_film_host::film_host().is_some_and(|host| host.0 as isize == under) {
        return true;
    }
    let mut owner = 0;
    unsafe { GetWindowThreadProcessId(HWND(under as *mut _), Some(&mut owner)) };
    pin_lifecycle::pid_is_the_previews(owner)
        || engine_processes::holds_a_player(owner)
        || pin_film_host::hosted_player_pid() == Some(owner)
}

/// The window under a screen point, which a test stands in rather than reading off the desktop:
/// a headless run has no pointer, and the desktop under it has no pin in it (see
/// `pin_chrome_occluded`).
fn window_at_point(x: i32, y: i32) -> isize {
    #[cfg(test)]
    if let Ok(held) = WINDOW_AT_POINT_STAND_IN.lock() {
        if let Some(answer) = *held {
            return answer;
        }
    }

    use windows::Win32::UI::WindowsAndMessaging::WindowFromPoint;
    unsafe { WindowFromPoint(POINT { x, y }) }.0 as isize
}

/// Stand an answer for `window_at_point` in place of the desktop's own, or give the desktop's
/// own answer back with `None` — put back the same way.
#[cfg(test)]
pub(super) fn stand_window_at_point_in(answer: Option<isize>) {
    if let Ok(mut held) = WINDOW_AT_POINT_STAND_IN.lock() {
        *held = answer;
    }
}

/// The answer `window_at_point` gives while a test stands one in (see `stand_window_at_point_in`).
#[cfg(test)]
static WINDOW_AT_POINT_STAND_IN: Mutex<Option<isize>> = Mutex::new(None);

/// Whether a pinned window's chrome — the caption and the transport bar — is drawn *over* the
/// media rather than in bands above and below it.
///
/// Two things have to be true of a kind for that. The band has to be one this app can
/// paint on: the page the browser draws an SVG or a font on is a window of somebody else's
/// standing *in* the band while that page runs, so a caption this app drew over one of those would
/// be a caption underneath it. The window FFmpeg's player has is the other window of somebody else's, and it
/// is drawn over all the same: the pin's own window is the one on top for as long as any strip
/// of its chrome is showing, and the player's own raise is held off for that while (see
/// `pin_chrome_up`), so what this app paints in the band is what is seen. And the frame has to
/// be the media's own shape, which is what makes the window box and the media box the same box:
/// the strips the chrome is drawn in are then inside the picture rather than beside it, and
/// there is no band of empty window where a bar used to be.
///
/// A page that is laid out to whatever box it is given — a text preview, an archive listing — and
/// a sound's card are neither: the first two keep their bars the way they have always had them,
/// the card is its own size with no band to fill, and a pin collapsed into its bubble has no
/// chrome of either kind.
///
/// **A page the browser draws is the two answers in one kind, and `page_runs` is the whole of
/// which**: a document or a specimen is a picture the pin draws through the Shell's thumbnail of a
/// source window, so its band is this app's to paint a strip over; a page that runs is a window of
/// the user's standing in front of the pin, and a caption painted over one of those would be a
/// caption underneath it (see `webview_preview`'s `put_window`). `page_runs` is read and unused by
/// every other kind.
pub(super) fn pin_overlay_chrome(kind: Option<MediaType>, page_runs: bool) -> bool {
    match kind {
        Some(MediaType::EngineSvg) | Some(MediaType::EngineFont) => !page_runs,
        Some(kind) => !kind.is_engine() && pin_frame(Some(kind)) == PinFrame::Shaped,
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

/// Tell the pin's running video player a level, and move the guard's own record of it where the
/// player was told.
///
/// **The guard stands for the level this app last asked a player for** (see
/// `audio_guard::the_level_moved`), and every road a hand reaches a running player by arrives here:
/// a step of the knob's drag, the wheel and the press behind it (`pin_input`), the release's
/// re-assertion, and a row moved in the tray — which is the knob's road and not a second one (see
/// `tray_volume_moved`). A level of nothing is the whole of what this closes: a pin dragged to 0
/// leaves a player with its track and `volume 0`, which a guard holding the launch's level would
/// take for a film something else had silenced and take back up ten seconds later.
///
/// Only a write that reached a player moves it, which is `PinVolume::playing_at`'s rule as well: a
/// level no player was told is still owed to the next one, and the player behind the bar is still at
/// the level the guard already holds.
fn set_the_pin_players_level(level: u32) -> bool {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let told = mpv_ipc::set_volume(pid, level);
    if told {
        audio_guard::the_level_moved(pid, level);
    }
    told
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
/// How much of the full range one wheel notch over a pin's volume control moves the level:
/// five percent, which is the smallest step a hand turning a wheel can aim at.
pub(super) const VOLUME_WHEEL_STEP: u32 = 5;

/// How many ticks a level must be quiet for before FFmpeg's player is begun again at it — the
/// ~250 ms a hand needs to stop turning, at the loop's ~60 ticks a second (see
/// `settle_pin_volume_after_quiet`).
pub(super) const PIN_VOLUME_SETTLE_TICKS: u32 = 15;

/// A wheel's own units per percent of the level: one 120-unit notch is still exactly
/// `VOLUME_WHEEL_STEP` percent (`WHEEL_DELTA / VOLUME_WHEEL_STEP`).
pub(super) const WHEEL_LEVEL_UNITS_PER_PERCENT: i64 = WHEEL_DELTA as i64 / VOLUME_WHEEL_STEP as i64;

/// One wheel message's travel added to a requested level held in wheel units, so that a smooth
/// wheel or a precision touchpad moves the level continuously rather than only once a whole
/// notch has accumulated. The units are exact — no travel is dropped or rounded per message.
pub(super) fn volume_after_wheel_delta(level_units: i64, delta: i32) -> i64 {
    (level_units + i64::from(delta)).clamp(0, 100 * WHEEL_LEVEL_UNITS_PER_PERCENT)
}

/// The integer level a wheel-unit request is drawn and played at, rounded to the nearest percent.
pub(super) fn volume_from_wheel_units(level_units: i64) -> u32 {
    let per = WHEEL_LEVEL_UNITS_PER_PERCENT;
    ((level_units + per / 2) / per).clamp(0, 100) as u32
}

/// Which wheel-unit level a notch in hand is added to: the carry it kept, but only where the carry
/// still belongs to the level the pin is at.
///
/// A carry is sub-notch travel on top of a level, so the level it was built from is
/// `volume_from_wheel_units(carry)` within a percent. A drag, a tray row or a key moved the level
/// since, and applying the notch to that carry would move the thumb back to where the older level
/// was — the jump a wheel turned straight after a drag used to make.
///
/// A glide in flight is the pin's own level still on its way to where the wheel asked, so its carry
/// is where the wheel is going and is kept whole rather than re-based on a level it has already
/// passed through (see `ramp_pin_volume`).
pub(super) fn wheel_base_units(carry: Option<i64>, level: u32, glide_pending: bool) -> i64 {
    let fresh = i64::from(level) * WHEEL_LEVEL_UNITS_PER_PERCENT;
    match carry {
        Some(units) if glide_pending || volume_from_wheel_units(units) == level => units,
        _ => fresh,
    }
}

pub(super) fn set_pin_volume(level: u32) {
    let level = level.min(100);
    let audio = matches!(current_media_type(), Some(MediaType::Audio));
    note_playback_trace!("volume set {} audio {}", level, u8::from(audio));
    let immediate = pin_level_is_immediate();
    with_pin(|pin| {
        pin.volume.level = level;
        pin.volume.audio = audio;
        // A drag, a tray row or a key is a level in hand: it is heard at once, and no glide left over
        // from a wheel notch before it is carried on with (see `ramp_pin_volume`).
        pin.volume.target = None;
        pin.volume.applied = level * WHEEL_LEVEL_UNITS_PER_PERCENT as u32;
        // A level a running player can be told is coalesced: the knob and the bar move now, the
        // engine hears the latest level once on the next tick (see `flush_pin_volume_level`).
        pin.volume.pending |= immediate;
        pin.volume.quiet = PIN_VOLUME_SETTLE_TICKS;
    });

    remember_pin_volume(level);

    // The speaker on a sound's card says what the level is, and the card is a media frame rather
    // than chrome: repainting the window alone would redraw the *old* card with the old speaker on
    // it. A video's level is drawn in the strip, which the repaint does reach (see
    // `pin_audio_hover_refresh`).
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

/// Whether the player behind the pin takes a level while it runs, which is what decides whether a
/// moved level is owed the engine on the next tick (see `flush_pin_volume_level`).
fn pin_level_is_immediate() -> bool {
    match current_media_type() {
        Some(MediaType::NativeVideo) => true,
        Some(MediaType::Audio) => {
            audio_preview_level_is_owed_to_the_engine() || ffplay_holds_the_level()
        }
        Some(MediaType::Video) => mpv_takes_a_level_as_it_moves() || ffplay_holds_the_level(),
        _ => false,
    }
}

/// How much of the gap between the applied level and a wheel's target one tick closes: about a
/// third, which is a glide of a handful of 16 ms ticks per notch (see `ramp_pin_volume`).
pub(super) const PIN_VOLUME_RAMP_SHARE: f32 = 0.35;

/// Ask the pin's level to go to `units`, a wheel's request in its own fine units (1/24 percent,
/// 120 to a notch): the popup is drawn at once from where the level is and the applied level glides
/// there on the next ticks (see `ramp_pin_volume`).
pub(super) fn set_pin_volume_target(units: i64) {
    let target = units.clamp(0, 100 * WHEEL_LEVEL_UNITS_PER_PERCENT) as u32;
    let audio = matches!(current_media_type(), Some(MediaType::Audio));
    note_playback_trace!("volume target {}", target);
    with_pin(|pin| {
        pin.volume.audio = audio;
        pin.volume.target = Some(target);
        pin.volume.applied = pin.volume.level * WHEEL_LEVEL_UNITS_PER_PERCENT as u32;
        // Every notch is a hand still turning, so the FFmpeg settle waits for the last of them (see
        // `settle_pin_volume_after_quiet`).
        pin.volume.quiet = PIN_VOLUME_SETTLE_TICKS;
    });

    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

/// Step the applied level one tick toward a wheel's target, never past it and never backwards.
///
/// The point of the glide is that a notch is heard as a short ramp rather than as a step of five
/// percent: the engine is told the applied level once a tick (see `flush_pin_volume_level`) and the
/// popup is drawn at it, so the thumb moves continuously. The last tick writes the exact target, so
/// a whole notch settles on precisely `VOLUME_WHEEL_STEP` percent rather than near it.
pub(super) fn ramp_pin_volume() {
    let per = WHEEL_LEVEL_UNITS_PER_PERCENT;
    let immediate = pin_level_is_immediate();
    let level = {
        let Some(mut pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin_mut() else {
            return;
        };
        let volume = &mut pin.volume;
        let Some(target) = volume.target else {
            return;
        };

        let gap = i64::from(target) - i64::from(volume.applied);
        if gap == 0 {
            volume.target = None;
            return;
        }

        // A share of the gap, and at least one fine unit while any gap is left: a small gap would
        // otherwise round to no step at all and the level would stick short of its target.
        let share = (gap as f32 * PIN_VOLUME_RAMP_SHARE).round() as i64;
        let step = share.abs().max(1) * gap.signum();
        // Never past the target: the share is under the gap, and the floor above only fires on a gap
        // it cannot overshoot.
        let applied = (i64::from(volume.applied) + step).clamp(0, 100 * per) as u32;
        let before = volume.applied;
        volume.applied = applied;
        let landed = applied == target;
        volume.level = volume_from_wheel_units(if landed {
            i64::from(target)
        } else {
            i64::from(applied)
        });
        if landed {
            volume.target = None;
        }
        // Told to a player that takes a level while it runs, once for every tick of the glide (see
        // `flush_pin_volume_level`).
        volume.pending |= immediate;
        (volume.level, applied != before, volume.target)
    };
    let (level, moved, target) = level;
    if moved {
        note_playback_trace!(
            "volume ramp applied {} level {} target {}",
            level,
            level,
            target.map_or_else(|| "none".to_string(), |units| units.to_string())
        );
    }

    remember_pin_volume(level);
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

/// Whether a level is still on its way: a wheel's target not yet reached, a level the engine is
/// still owed, a wheel notch waiting to be read, or FFmpeg's settle window running.
///
/// It is what keeps the loop at the frame cadence while a wheel is turning the volume — the glide
/// and the flush are only as smooth as the ticks they land on — and what lets it drop back to the
/// pin's own slower wait once the level has settled (see `wait_before_the_next_tick`).
pub(super) fn pin_volume_is_moving() -> bool {
    if crate::app::seam::pin_volume_delta_pending() {
        return true;
    }

    pin_state()
        .and_then(|pinned| {
            pinned.pin().map(|pin| {
                pin.volume.target.is_some() || pin.volume.pending || pin.volume.quiet > 0
            })
        })
        .unwrap_or(false)
}

/// Whether a pinned preview owes its window one repaint for its chrome alone, spending the flag a
/// wheel notch left: `AUDIO_CARD_DIRTY` is the pin's chrome as much as a sound's card, but only a
/// sound's arm is where a card is refreshed, so without this drain a film's thumb stepped on
/// whatever incidental repaint came next (see the preview loop's tick).
///
/// One flag per tick however many notches landed in it, and a sound's card is left to its own arm:
/// its paint is the card rather than the chrome, which is the one thing a bare repaint cannot
/// redraw.
pub(super) fn pin_chrome_owes_a_repaint() -> bool {
    pinned()
        && current_media_kind() != Some(PreviewType::Audio)
        && AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel)
}

/// Whether a wheel's target is still waiting for the glide, which is what says a knob that moved
/// the level between ticks was a hand other than the wheel's (see the wheel's own carry in the
/// preview loop).
pub(super) fn pin_volume_target_pending() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.volume.target.is_some()))
        .unwrap_or(false)
}

/// Tell the engine the latest level a knob or wheel asked for, at most once per preview tick.
///
/// A drag and a smooth wheel ask for a level on every message; opening, writing and closing mpv's
/// pipe on each of those is what made the knob stutter. The level stays in `PinVolume::level` — so
/// the popup is drawn at it at once — and only the engine command waits here for the tick.
pub(super) fn flush_pin_volume_level() {
    let pending = {
        let Some(mut pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin_mut() else {
            return;
        };
        if !pin.volume.pending {
            return;
        }
        pin.volume.pending = false;
        pin.volume.level
    };

    let told = match current_media_type() {
        Some(MediaType::NativeVideo) => video_player::set_volume(pending),
        Some(MediaType::Audio) => {
            if audio_preview_level_is_owed_to_the_engine() {
                video_player::set_volume(pending)
            } else if ffplay_holds_the_level() {
                // One-slot last-write-wins ask to the `mixer-level` thread, cheap enough at 60 Hz.
                ask_the_ffplay_its_level(pending);
                true
            } else {
                false
            }
        }
        Some(MediaType::Video) => {
            if mpv_takes_a_level_as_it_moves() {
                set_the_pin_players_level(pending)
            } else if ffplay_holds_the_level() {
                ask_the_ffplay_its_level(pending);
                true
            } else {
                false
            }
        }
        _ => false,
    };
    // The claim of what the player is playing at moves only where the player was actually told:
    // writing it for a command that never reached one would leave the bar claiming a level the
    // film never took, and the player that begins next owed nothing (see `settle_pin_volume`).
    if told {
        with_pin(|pin| {
            if pin.volume.level == pending {
                pin.volume.playing_at = pending;
            }
        });
    }
}

/// Whether the film on screen is mpv's, the one player here that takes a level while it runs.
///
/// It is asked of the route rather than of a player's liveness: a level turned with no player up is
/// written down for the player that begins next either way, and it is the route that says whether
/// that player can be told a level live at all (see `set_pin_volume` and `mpv_ipc`).
fn mpv_takes_a_level_as_it_moves() -> bool {
    pinned_media_owner().is_some_and(|(path, _)| video_route(&path) == VideoRoute::Mpv)
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
    /// A player is behind the pin and takes the level at its mixer session, so nothing is begun,
    /// ended or restarted: the level is asked for and the claim moves (see
    /// `ask_the_ffplay_its_level`).
    Mixed,
    /// A player is behind the pin and it is *holding* the file, so it is left exactly where it is
    /// and the level is owed to the player that begins when the file is let go of.
    ///
    /// Nothing is written for it here, and what is not written is the point: the claim of what the
    /// running player was begun at stays standing, which is what says a level is owed and what
    /// makes that player refuse a pause key until it has been replaced (see `video_key_pause`).
    OwedToTheNextPlayer,
    /// No player is behind the pin at all. Nothing is begun: the level is written down as the one
    /// playing, and the player that takes it is the one a later press begins at it.
    Recorded,
}

/// Whether a level moved on a pinned FFmpeg video is told to the running player, is left owed to the
/// player that begins next, or is simply written down.
///
/// The question is whether a *player* is behind the pin rather than whether the film is playing,
/// which is what makes a held file a case of its own: a held file is still a file playing at a
/// level, and the knob moved while it was held would otherwise reach a player that never learns of
/// it — the film would go on at the level it was at when the hand stopped it while the bar claimed
/// the level it was set to (see `settle_pin_volume`).
///
/// **A held file's level is still left to the next player**, because a tune at the mixer of a player
/// that is holding a frame is not a thing the user asked for until the film is let go of; and the
/// claim of what the running player was begun at is what keeps it owed.
///
/// A player this app has lost the handle to is still a player while it is running, so a process
/// behind a bar that has stopped claiming playback is still owed the level; only a pin with neither
/// a claim nor a process is owed nothing, and a held pin with neither is owed nothing for the same
/// reason — there is no player holding it to owe the level to.
pub(super) fn pin_level_settled_by(playing: bool, running: bool, held: bool) -> PinLevelSettling {
    if held && running {
        PinLevelSettling::OwedToTheNextPlayer
    } else if playing || running {
        PinLevelSettling::Mixed
    } else {
        PinLevelSettling::Recorded
    }
}

/// Give the pin's level to the player it is owed to, where a player is running to take it.
///
/// **FFmpeg's player takes the level where it runs**: the write is the mixer session's, and nothing
/// is begun, ended or restarted for it (see `ask_the_ffplay_its_level`). The claim then moves, which
/// is what says the player owes nothing more (see `PinVolume::playing_at`).
///
/// A pin that is *held* owes the level to the player that begins when the file is let go of, and
/// what is asked in that arm is which of three things the level is settled by (see
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
    // A release settles what it was dragging, so any wheel's quiet countdown is over too, and the
    // latest level is told before anything asks whether the player already has it.
    with_pin(|pin| pin.volume.quiet = 0);
    flush_pin_volume_level();

    let Some((path, _content, transport, volume)) = pinned_playback_state() else {
        return;
    };
    note_playback_trace!(
        "volume settle level {} playing_at {}",
        volume.level,
        volume.playing_at
    );
    save_remembered_pin_volume(volume.level);
    if volume.playing_at == volume.level {
        return;
    }

    match current_media_type() {
        Some(MediaType::NativeVideo) => {
            // The engine takes a level while it plays, and the claim moves only where it did: a
            // session that has already gone heard nothing, and the level stays owed to the next
            // player rather than being recorded as given (see `video_player::set_volume`).
            if video_player::set_volume(volume.level) {
                with_pin(|pin| pin.volume.playing_at = pin.volume.level);
            }
        }
        Some(MediaType::Video)
            if video_route(&path) == VideoRoute::Mpv && set_the_pin_players_level(volume.level) =>
        {
            // mpv was told the level as the knob moved (see `set_pin_volume`), so there is no player
            // to replace here: the command is re-asserted and the claim moves only where it reached
            // the player. One that did not falls to the arm below, which is FFmpeg's player told at
            // its mixer session (see `pin_level_settled_by`).
            with_pin(|pin| pin.volume.playing_at = pin.volume.level);
        }
        Some(MediaType::Video) => match pin_level_settled_by(
            pin_is_playing(&transport),
            is_video_process_running(),
            transport.paused_at.is_some(),
        ) {
            PinLevelSettling::Mixed => {
                ask_the_ffplay_its_level(volume.level);
                with_pin(|pin| pin.volume.playing_at = pin.volume.level);
            }
            PinLevelSettling::OwedToTheNextPlayer => {
                // Nothing is done, and nothing is written: the claim of what the running player
                // was begun at stays standing, which is the whole of what says the level is owed
                // (see `PinVolume::playing_at` and `video_key_pause`).
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
                if video_player::set_volume(volume.level) {
                    with_pin(|pin| pin.volume.playing_at = pin.volume.level);
                }
            } else {
                // Otherwise the player at the mixer is owed it, and the clock behind a sound's card
                // is the loop's: so the flag, and the loop asks for the level on its next tick (see
                // `settle_pinned_audio_volume`).
                PIN_AUDIO_VOLUME.store(true, Ordering::Release);
            }
        }
        _ => {}
    }
}

/// Settle a level once the wheel or drag has left it alone for `PIN_VOLUME_SETTLE_TICKS`: a knob is a
/// burst of notches, and one settle when the hand stops is what keeps the claims in step with the
/// player rather than a write per notch.
pub(super) fn settle_pin_volume_after_quiet() {
    let quiet = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.volume.quiet))
        .unwrap_or(0);
    if quiet == 0 {
        return;
    }
    if quiet > 1 {
        with_pin(|pin| pin.volume.quiet = quiet - 1);
        return;
    }
    // The last tick is held while a wheel's level is still gliding: settling FFmpeg's player on a
    // level the glide has not reached yet would leave it playing at one level and the bar claiming
    // another, with nothing left to settle the rest (see `ramp_pin_volume`).
    if pin_volume_target_pending() {
        return;
    }
    with_pin(|pin| pin.volume.quiet = 0);
    settle_pin_volume();
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

/// Give a pinned sound's player the level the pin was moved to, where that player is FFmpeg's.
///
/// **The write is the mixer session's**, so the player is neither restarted nor seeked and the clock
/// behind the card is left exactly where it was: the level a sound plays at is not a fact of when it
/// was begun any more (see `audio_session::ask_the_mixer`). A sound that is held, or one whose player
/// has gone, is left alone — the player that begins next is begun at the pin's level.
pub(super) fn settle_pinned_audio_volume() {
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

    if let Some(level) = pin_volume_level() {
        note_playback_trace!("volume audio settle {}", level);
        ask_the_ffplay_its_level(level);
    }

    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

//! A hand on a pinned window: the move that carries it, the press that begins a drag, and the
//! caption, transport bar and volume button a press can land on.

use super::*;

/// The point the pointer is at, in screen coordinates: what a drag is measured in, since a
/// window that is being resized moves its own origin out from under a client coordinate.
pub(super) fn cursor_screen_point() -> Option<(i32, i32)> {
    let mut point = POINT::default();
    unsafe { GetCursorPos(&mut point) }.ok()?;
    Some((point.x, point.y))
}

/// The media box a window box implies: the window less the caption above it and the transport bar
/// below it — or the window itself, for a kind whose chrome is drawn over its media and whose two
/// boxes are therefore one box (see `pin_overlay_chrome`).
pub(super) fn content_box_of(
    window: ScreenRegion,
    dpi: u32,
    transport: bool,
    overlay: bool,
    caption: i32,
) -> ScreenRegion {
    if overlay {
        return window;
    }

    (
        window.0,
        window.1 + caption,
        window.2,
        window.3 - pinned_transport_height(dpi, transport),
    )
}

/// What the pointer is doing on a pinned window, in the window's own coordinates: which caption
/// button it is over — a caption lights up as a pointer crosses it, the way a Windows one does —
/// and, while a press is being held, the drag or the resize it began.
pub(super) unsafe fn pinned_mouse_move(hwnd: HWND, x: i32, y: i32) {
    let Some(caption) = pinned_caption_geometry() else {
        return;
    };

    let dragging = {
        let Some(pinned) = pin_state() else {
            return;
        };
        pinned.pin().and_then(|pin| pin.dragging)
    };

    if dragging.is_some() {
        apply_pin_drag(hwnd);
        return;
    }

    // A drag of the volume knob, which is a level being aimed rather than a window being moved
    // (see `pinned_volume_drag`).
    if pinned_volume_drag(hwnd, y) {
        return;
    }

    // A drag along the transport bar, which is a seek being aimed rather than a window being
    // moved (see `pinned_transport_drag`).
    if pinned_transport_drag(hwnd, x) {
        return;
    }

    let hovered = (caption.wanted && y < caption.height)
        .then(|| {
            pin_chrome::button_at(
                x,
                y,
                caption.width,
                caption.height,
                caption.dpi,
                caption.frame != PinFrame::None,
            )
        })
        .flatten();
    let changed = {
        let Some(mut pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin_mut() else {
            return;
        };

        // A hand in one of the strips is a hand asking for that strip, and it is asked here as well
        // as on the loop's tick because a press can arrive in the same handful of messages as the
        // move that brought the pointer there: the loop would answer the question a tick too late
        // for a click that is already on its way (see `pin_chrome_near`). One strip at a time, the
        // way the tick asks it: it is the strip the pointer is in, not the window it is over.
        let (_, height) = pin.window_size();
        let bar_top = height - pinned_transport_height(pin.dpi, pin.transport_bar);
        if y < caption.height {
            pin.chrome.caption = true;
        }
        if pin.transport_bar && y >= bar_top {
            pin.chrome.bar = true;
        }

        let changed = pin.hovered != hovered;
        pin.hovered = hovered;
        changed
    };

    // The same question of the transport bar, which is the other strip a pointer lights up.
    let transport_hovered = pinned_transport_geometry().and_then(|bar| {
        (bar.wanted && y >= bar.top)
            .then(|| {
                pin_chrome::transport_part_at(
                    x,
                    y - bar.top,
                    bar.width,
                    bar.height,
                    bar.dpi,
                    bar.live,
                )
            })
            .flatten()
    });
    let transport_changed = pin_state()
        .and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            let changed = pin.transport.hovered != transport_hovered;
            pin.transport.hovered = transport_hovered;
            Some(changed)
        })
        .unwrap_or(false);

    // And the sound's card's own row of buttons, which is the other strip a pointer lights up —
    // and the only one a sound has (see `pin_transport_kind`). Asked here as well as on the loop's
    // tick so that a press arriving in the same handful of messages as the move that brought the
    // pointer over a button finds it already lit.
    let audio_changed = pin_state()
        .and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            let hovered = pin_audio_control_at(pin, x, y);
            let changed = pin.audio_hovered != hovered;
            pin.audio_hovered = hovered;

            // The card's own paint is what shows the wash, and a card is a media frame rather than
            // chrome: repainting the window alone would redraw the *old* card with the old button
            // still lit under it (see `pin_audio_hover_refresh`, which is the same question asked
            // on the tick).
            if changed {
                AUDIO_CARD_DIRTY.store(true, Ordering::Release);
            }
            Some(changed)
        })
        .unwrap_or(false);

    if changed || transport_changed || audio_changed {
        render_layered_preview(hwnd);
    }
}

/// What the pointer's questions about a pinned window's caption need: how tall the strip is, the
/// scale it is drawn at, the window's own width, and what this pin's edges and buttons do.
pub(super) struct PinnedCaption {
    /// The strip itself, which is where the buttons are, what `button_at` is asked with, and what a
    /// point is inside or outside of. It is nothing at all for a sound, which is given no caption,
    /// so every question that starts `y < height` is asked and answered false — which is the whole
    /// of what keeps a sound's card a card and not a title bar with a picture under it.
    pub(super) height: i32,
    pub(super) dpi: u32,
    pub(super) width: i32,
    pub(super) frame: PinFrame,
    /// Whether this strip is showing: the buttons are answers to a hand that has come for the
    /// caption, and a caption that is not drawn is not something a press may act on — what is under
    /// it then is the picture, which is a handle for moving the window (see `PinChrome`).
    pub(super) wanted: bool,
}

pub(super) fn pinned_caption_geometry() -> Option<PinnedCaption> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    (!pin.collapsed).then(|| {
        let (width, _) = pin.window_size();
        PinnedCaption {
            height: pin.caption,
            dpi: pin.dpi,
            width,
            frame: pin.frame,
            wanted: pin.chrome.caption,
        }
    })
}

/// Where a pinned window's transport bar is, as the pointer's questions about it need it: the
/// window's own width, the row the bar begins at, how tall it is, the scale it is drawn at, and
/// the transport's own state for the parts a press acts on.
pub(super) struct PinnedTransportBar {
    pub(super) width: i32,
    pub(super) top: i32,
    pub(super) height: i32,
    pub(super) dpi: u32,
    /// Whether the player behind the bar can be told anything, which is whether its parts are the
    /// pointer's to press at all (see `PinnedPreview::transport_live`).
    pub(super) live: bool,
    /// Whether the bar is showing, which is the same question the caption's buttons are answered
    /// by: a bar that is not drawn is not a bar a press may act on (see `PinnedCaption::wanted`).
    pub(super) wanted: bool,
}

pub(super) fn pinned_transport_geometry() -> Option<PinnedTransportBar> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    if pin.collapsed || !pin.transport_bar {
        return None;
    }

    let (width, height) = pin.window_size();
    let band = pinned_transport_height(pin.dpi, true);

    Some(PinnedTransportBar {
        width,
        top: height - band,
        height: band,
        dpi: pin.dpi,
        live: pin.transport_live,
        wanted: pin.chrome.bar,
    })
}

/// A press on a pinned window's transport bar, answering whether it was the bar's to act on: the
/// button pauses and resumes, and the bar is taken hold of where it was pressed.
///
/// A drag that is still going is not a seek — with FFmpeg it would be a player restarted for
/// every pixel of the drag — so what a press on the bar does is put the playhead under the hand
/// and what the release does is take the file there (see `seek_pinned_playback`).
pub(super) unsafe fn pinned_transport_press(hwnd: HWND, x: i32, y: i32) -> bool {
    let Some(bar) = pinned_transport_geometry() else {
        return false;
    };
    if y < bar.top || !bar.wanted {
        return false;
    }

    let Some(part) =
        pin_chrome::transport_part_at(x, y - bar.top, bar.width, bar.height, bar.dpi, bar.live)
    else {
        return false;
    };

    match part {
        pin_chrome::TransportPart::Play => {
            update_pin_transport(|transport| transport.pressed = Some(part));
        }
        pin_chrome::TransportPart::Seek => {
            let share = pin_chrome::transport_share_at(x, bar.width, bar.height, bar.dpi, bar.live);
            let aim = pin_state().and_then(|pinned| {
                pinned
                    .pin()
                    .and_then(|pin| pin_seconds_at(&pin.transport, share))
            });

            // **A press with no second to seek to claims nothing at all, and that is the whole of
            // what this arm refuses.** There is no aim to move and no film to hold, so nothing
            // below has anything to arm — but the capture at the end of this function is taken by
            // the press whatever the arm decided, and every release arm answers out of what the
            // press armed. A press that armed nothing therefore left the window holding the
            // pointer for the rest of the process, and every mouse message on the desktop arrived
            // here instead of at whatever it was aimed at: a pin whose title bar, edges and bar had
            // gone dead and whose cursor kept the last edge's shape, until the app was closed.
            //
            // Which is what a pin is between a file step and the probe that answers for the file in
            // it: the bar is drawn before the length is known, so the band is seekable and the file
            // has no second in it yet. Answering the press rather than refusing it is what made
            // that window unusable (see `release_pin_capture`, which is where the pointer a press
            // arms nothing for is given back).
            if !seek_press_arms(aim) {
                return false;
            }
            update_pin_transport(|transport| {
                transport.seeking = aim;
            });
            // Parked at seek-start, so the relaunch the release makes runs behind a cover: the
            // old player is retired at relaunch with the hole transparent and the replacement up
            // later, which is the desktop flash. Only a player of this app's is parked — the
            // engine's draws into this app's own surface and has no window to put away, and a
            // cover over one is a frozen frame nothing ends. Only an aimed second arms the
            // gesture at all: a press with no second to seek to parks nothing and holds nothing
            // (see `seek_press_arms`). The film is held with the same gesture hold as a drag —
            // audio with the picture, for the whole gesture — and the release carries it onto
            // the relaunch, which the swap ends without a key (see `seek_press_hold` and
            // `settle_seek_hold_after_swap`).
            if current_media_type() == Some(MediaType::Video) {
                let at = window_origin(hwnd)
                    .map(|origin| (origin.0, origin.1))
                    .or_else(|| pinned_content().map(|content| (content.0, content.1)))
                    .unwrap_or((0, 0));
                // The kill road freezes the clock before the park reads the
                // frame, and kills only after the cover stands: silence by
                // construction, and no blind toggle anywhere on the gesture.
                // A press the kill road refuses — no player to kill, or one
                // already dead — keeps the legacy hold.
                park_pinned_player_for_seek(hwnd, at);
                if gesture_press_freeze() {
                    kill_pinned_player_async();
                } else {
                    seek_press_hold();
                }
                // A newer gesture supersedes whatever relaunch is still in flight behind the
                // cover: the bump kills it before it can publish. After the hold rather than
                // before it — the key is posted to the live player, and there is no window
                // to post one through once it has gone.
                bump_pinned_generation();
            }
        }
        // The volume button is held rather than acted on where it is pressed, like every button a
        // window has: what a click does is open the popup or put it away, and that is a release
        // (see `pinned_transport_release`).
        pin_chrome::TransportPart::Volume => {
            update_pin_transport(|transport| transport.pressed = Some(part));
        }
    }

    let _ = SetCapture(hwnd);
    render_layered_preview(hwnd);
    true
}

/// Where a drag along the bar has taken the playhead: the share of the bar under the hand turned
/// into a second of the file, for a file whose length is known.
///
/// The length is asked for the way the bar *draws* one rather than read out of what was written
/// down when the pin was taken up: a container that does not say how long it is, and a length the
/// probe could not read, are both filled in from the engine's own answer — so a bar that draws a
/// length and a bar that can be dragged to one are the same bar rather than two questions that
/// can disagree.
pub(super) fn pin_seconds_at(transport: &PinTransport, share: f64) -> Option<f64> {
    pin_duration(transport).map(|duration| (duration * share).clamp(0.0, duration))
}

/// Carry a drag along the transport bar: the playhead follows the hand, and the file is taken
/// there only when the pointer lets go.
pub(super) unsafe fn pinned_transport_drag(hwnd: HWND, x: i32) -> bool {
    let Some(bar) = pinned_transport_geometry() else {
        return false;
    };

    let dragging =
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.transport.seeking.is_some()));

    if dragging != Some(true) {
        return false;
    }

    let share = pin_chrome::transport_share_at(x, bar.width, bar.height, bar.dpi, bar.live);
    update_pin_transport(|transport| transport.seeking = pin_seconds_at(transport, share));
    render_layered_preview(hwnd);
    true
}

/// What a release on the transport bar does: a click on the button pauses or resumes, and a drag
/// that has let go of the bar takes the file to where the hand stopped.
pub(super) unsafe fn pinned_transport_release(
    hwnd: HWND,
    x: i32,
    y: i32,
    window: &dyn PinWindow,
) -> bool {
    let (part, seeking, transport) = {
        // An arm that declines does not touch the pointer, and this arm declining is the ordinary
        // case: a release on the caption is a hand on a title bar, and the bar has nothing to say
        // about it. Letting go of the pointer here would re-enter `pin_capture_lost` before the
        // caption's own arm had read what the press armed, and that road drops the pressed button
        // out from under it — so every button on a pinned window would be a button that does
        // nothing. The capture a press takes is given back by the arm that owns it, and by the
        // procedure's own last resort where no arm owns it (see `release_pin_capture`).
        let Some(mut pinned) = pin_state() else {
            return false;
        };
        let Some(pin) = pinned.pin_mut() else {
            return false;
        };

        let part = pin.transport.pressed.take();
        let seeking = pin.transport.seeking;
        (part, seeking, pin.transport)
    };

    if part.is_none() && seeking.is_none() {
        return false;
    }

    // The kill road the press may have armed, read before the pointer is let
    // go of: the release below re-enters `pin_capture_lost` synchronously
    // through `WM_CAPTURECHANGED`. Through the window seam rather than a bare
    // `ReleaseCapture` so a test can stand in that re-entrancy; on the machine
    // the two are the same release, the press having taken the capture.
    let kill_road = gesture_snapshot_active();
    release_the_pointer(window, hwnd.0 as isize);

    if let Some(part) = part {
        // A button is clicked where the pointer is still on it, which is the rule every caption
        // button of this app's follows.
        let still_on_it = pinned_transport_geometry()
            .map(|bar| {
                y >= bar.top
                    && pin_chrome::transport_part_at(
                        x,
                        y - bar.top,
                        bar.width,
                        bar.height,
                        bar.dpi,
                        bar.live,
                    ) == Some(part)
            })
            .unwrap_or(false);

        if still_on_it {
            match part {
                pin_chrome::TransportPart::Play => {
                    if let Some((path, content, _, _)) = pinned_playback_state() {
                        toggle_pinned_playback(&path, content, transport);
                    }
                }
                pin_chrome::TransportPart::Volume => toggle_pin_volume(hwnd),
                pin_chrome::TransportPart::Seek => {}
            }
        }

        update_pin_transport(|transport| transport.pressed = None);
        render_layered_preview(hwnd);
        return true;
    }

    // A drag of the bar: the file is taken to the second the hand stopped at, which is the one
    // moment an FFmpeg player is ended and begun again (see `seek_pinned_playback`). The relaunch
    // runs behind the press's cover carrying the seek's hold, and the swap bound is re-armed onto
    // it: the scrub's ticks spent the stamp the press took, and without this the first tick after
    // the release swaps onto whatever window merely exists (see `rearm_seek_cover_for_relaunch`).
    if let Some(seconds) = seeking {
        // The kill road owns the gesture's one relaunch, at the aimed second
        // rather than the press's: the take inside `relaunch_gesture_end_at`
        // makes it the only one across the capture-lost re-entered above and
        // this release, whichever runs first.
        if kill_road {
            if let Some((_, content, _, _)) = pinned_playback_state() {
                relaunch_gesture_end_at(content, Some(seconds));
                rearm_seek_cover_for_relaunch();
            }
        } else if let Some((path, content, _, _)) = pinned_playback_state() {
            seek_pinned_playback(&path, content, seconds);
            rearm_seek_cover_for_relaunch();
            // A relaunch that never came up leaves no swap to disarm the
            // dead interval: the press's snapshot goes with it instead.
            disarm_gesture_if_no_relaunch();
        }
    }

    render_layered_preview(hwnd);
    true
}

/// The popup a pin's volume button opens, as the pointer's questions about it need it: where it
/// is, or nothing when it is not up.
///
/// It is one question answered for two kinds, and it is asked of the pin rather than recomputed at
/// each of the four places that want the panel — which is what keeps the panel a press is answered
/// against the panel that is drawn. A kind with a transport bar hangs it off that bar's own volume
/// button, and a sound hangs the very same panel off the button on its card instead (see
/// `pinned_audio_volume_popup` and `pin_chrome::volume_popup_from_button`).
pub(super) fn pinned_volume_geometry() -> Option<pin_chrome::VolumePopup> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    pin.volume.open.then(|| pinned_volume_panel(pin)).flatten()
}

/// Where *this* pin's volume button opens its panel from, whichever of the two buttons it is: the
/// transport strip's own where the pin has a strip, and a sound's card for a pin whose controls are
/// the card's — one panel, hung from a button either way (see `pin_chrome::
/// volume_popup_from_button`).
///
/// It takes the pin rather than asking for the one that is installed, because two of its four
/// callers are handed a pin they are already holding and one of them is a tick: a popup answered
/// against a window other than the one under the hand is a panel that closes when it should not
/// (see `refresh_pin_volume`).
pub(super) fn pinned_volume_panel(pin: &PinnedPreview) -> Option<pin_chrome::VolumePopup> {
    if pin.collapsed {
        return None;
    }

    let strip = if pin.transport_bar {
        let (width, height) = pin.window_size();
        let band = pinned_transport_height(pin.dpi, true);
        Some(pin_chrome::volume_popup_layout(
            width,
            (height - band).max(0),
            band,
            pin.dpi,
        ))
    } else {
        None
    };

    strip.or_else(|| pinned_audio_volume_popup(pin))
}

/// Where a pinned sound's card's volume button opens its panel from, in the window's own
/// coordinates: the card fills the pin's media box and is drawn in the band between the chrome's
/// (see `pinned_band_rows`), so the box the card lays its row out in is the box the window draws
/// it at, moved down by the band's own row.
///
/// Nothing about the two windows standing in each other's place applies to it — the panel floats
/// over this app's own card rather than over a player's window, so there is no window to put in
/// front and no tick to hold off (see `pin_volume_open`).
pub(super) fn pinned_audio_volume_popup(pin: &PinnedPreview) -> Option<pin_chrome::VolumePopup> {
    if !pin_shows_an_audio_card(pin) {
        return None;
    }

    let (width, height) = pin.window_size();
    let (top, _) = pinned_band_rows(
        height,
        pin.caption,
        pinned_transport_height(pin.dpi, pin.transport_bar),
        pin.overlay,
    );

    // The button's own box, from the card's own arithmetic rather than from anything kept beside
    // it: a popup hung off a button that has moved is a popup that is not where its button is
    // (see `audio_preview::control_box`).
    let button = audio_preview::control_box(
        CardControl::Volume,
        (pin.content.2 - pin.content.0).max(1) as u32,
        pin.dpi,
        current_audio_options(),
        true,
    )?;

    Some(pin_chrome::volume_popup_from_button(
        RECT {
            top: button.top + top,
            ..button
        },
        width,
        pin.dpi,
    ))
}

/// Whether the popup's knob is being held.
pub(super) fn pin_volume_dragging() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.volume.dragging))
        .unwrap_or(false)
}

/// A press on a pinned window's volume popup, answering whether it was the popup's: the knob is
/// taken hold of where the hand landed, which is the level it is drawn at from there.
///
/// The panel is the target rather than the groove inside it: a level is aimed at with the whole of
/// what is drawn, and a hand a few pixels off a six-pixel groove is a hand that meant to move it.
/// It is asked before anything else the band could be — what is under the popup is the picture it
/// floats over, and a press there is not a hand on that picture.
pub(super) unsafe fn pinned_volume_press(hwnd: HWND, x: i32, y: i32) -> bool {
    let Some(popup) = pinned_volume_geometry() else {
        return false;
    };

    if x < popup.panel.left
        || x >= popup.panel.right
        || y < popup.panel.top
        || y >= popup.panel.bottom
    {
        return false;
    }

    let _ = SetCapture(hwnd);
    with_pin(|pin| pin.volume.dragging = true);
    // The kill road, once at the first step iff a player lives: the press
    // snapshots the playhead and freezes the clock, the park captures the
    // frame and stands the cover, and only then is the player killed. Rapid
    // steps find the snapshot standing and rewrite the owed level only — one
    // relaunch at the settle, at the latest level.
    if current_media_type() == Some(MediaType::Video) && gesture_press_freeze() {
        let at = window_origin(hwnd)
            .map(|origin| (origin.0, origin.1))
            .or_else(|| pinned_content().map(|content| (content.0, content.1)))
            .unwrap_or((0, 0));
        // The record the cover's own end will wait on, written by the road that raised it and by
        // nobody else: a press that arrives over a file step's cover extends that record and
        // leaves its arm alone (see `park_pinned_player_for_gesture`).
        park_pinned_player_for_gesture(hwnd, at, false);
        kill_pinned_player_async();
        bump_pinned_generation();
    }
    set_pin_volume((pin_chrome::volume_share_at(y, popup.track) * 100.0).round() as u32);
    render_layered_preview(hwnd);
    true
}

/// Carry a drag of the volume knob: the knob follows the hand, and the level of a player that can
/// be told one follows it too (see `set_pin_volume`).
pub(super) unsafe fn pinned_volume_drag(hwnd: HWND, y: i32) -> bool {
    if !pin_volume_dragging() {
        return false;
    }
    let Some(popup) = pinned_volume_geometry() else {
        return false;
    };

    set_pin_volume((pin_chrome::volume_share_at(y, popup.track) * 100.0).round() as u32);
    render_layered_preview(hwnd);
    true
}

/// A release on the volume popup: the knob is let go, and the player that takes a level only by
/// being started at one is settled with it (see `settle_pin_volume`).
pub(super) unsafe fn pinned_volume_release(hwnd: HWND, window: &dyn PinWindow) -> bool {
    let dragging = {
        // Declining is free of side effects, for the same reason it is in the transport arm above,
        // and this arm is asked of every release on a pinned window before any other.
        let Some(mut pinned) = pin_state() else {
            return false;
        };
        let Some(pin) = pinned.pin_mut() else {
            return false;
        };

        let dragging = pin.volume.dragging;
        pin.volume.dragging = false;
        dragging
    };

    if !dragging {
        // The knob was taken hold of by a press and let go of again by whatever stood between that
        // press and this release — the swap's own reconcile above all, which drops the drag state
        // so a dead aim cannot answer (see `reconcile_swap_take_up`). Nothing here is armed, so
        // there is no capture of this arm's to give back.
        return false;
    }

    // The kill road the press may have armed, read before the pointer is let
    // go of: the release below re-enters `pin_capture_lost` synchronously
    // through `WM_CAPTURECHANGED`, and that road relaunches — so the check
    // after it would find the take already spent and settle the level a
    // second time behind the first relaunch. Through the window seam rather
    // than a bare `ReleaseCapture` so a test can stand in that re-entrancy;
    // on the machine the two are the same release, the press having taken the
    // capture.
    let kill_road = gesture_snapshot_active();
    release_the_pointer(window, hwnd.0 as isize);
    // A gesture that killed settles in the one relaunch its press armed: at
    // the snapshot second, the box the hand left behind, the latest level —
    // carrying was-held or the gesture's hold for the swap to settle without
    // a key. The take inside `relaunch_gesture_kill_at` makes it the only one
    // across the capture-lost re-entered above and this release, whichever
    // runs first. The remembered level is still written once, where the knob
    // was let go of.
    if kill_road {
        if let Some((_, content, _, volume)) = pinned_playback_state() {
            save_remembered_pin_volume(volume.level);
            relaunch_gesture_kill_at(content);
        } else {
            take_gesture_snapshot();
        }
        render_layered_preview(hwnd);
        return true;
    }
    settle_pin_volume();
    render_layered_preview(hwnd);
    true
}

/// Open the volume popup, or put it away: what a click on the volume button does.
///
/// Opening it puts the pin's window above the player's, because what the panel floats over is the
/// media — and the media of a video FFmpeg plays is a window of somebody else's standing in that
/// band, asserted over this one. The tick that keeps that window in front of Explorer is held off
/// while the popup is open (`pin_volume_open`), which is the whole of what the two windows owe each
/// other; nothing has to be done when it closes, since the pin's window is transparent wherever it
/// is not painting and the player is in front of it again on the tick.
///
/// For a sound the panel floats over this app's own card instead, so raising the window is
/// harmless rather than necessary — which is why this function asks nothing about the kind (see
/// `pinned_audio_volume_popup`).
pub(super) unsafe fn toggle_pin_volume(hwnd: HWND) {
    let opened = {
        let Some(mut pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin_mut() else {
            return;
        };

        pin.volume.open = !pin.volume.open;
        if !pin.volume.open {
            pin.volume.dragging = false;
        }

        pin.volume.open
    };

    if opened {
        raise_pinned_window(hwnd);
    }

    // A sound's card is a media frame and its volume button is one of the card's own pixels: the
    // card has to be painted again for a popup opening or closing to be seen at all, and
    // repainting the window alone would redraw the old card (see `set_pin_volume`).
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
    render_layered_preview(hwnd);
}

/// Put the pin's window above the window standing in its media band, without taking the focus and
/// without moving anything: the order the tick keeps for the player, asked the other way round for
/// as long as a volume popup is open.
pub(super) unsafe fn raise_pinned_window(hwnd: HWND) {
    if let Some((window, width, height)) = pinned_window_box() {
        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            window.0,
            window.1,
            width,
            height,
            SWP_NOACTIVATE,
        );
    }
}

/// A press on a pinned window, answering whether it was the pin's to act on.
///
/// Four things a press can be, in the order a hand finds them: an edge of the window, which begins
/// a resize, one of the caption's buttons, the rest of the caption — which is a title bar, and
/// beginning a move is what a title bar does — and the media itself, which moves the window under
/// a hand that drags it. What is inside the media is left to the media: a text preview's scrollbar
/// and its selection are the pointer's own, and a press on either is not a move.
///
/// The frame is asked about before the caption, which is the order Windows itself has and the
/// order the two are drawn in: the band along the top of a window is over the strip the caption is
/// painted in, so a caption asked about first swallows the top edge and the two corners at the top
/// of the window — three of the eight places a window can be resized from, and the three a hand
/// reaches for on a window it wants wider. What the caption keeps is everything outside the band,
/// which is where its buttons are: a corner's share of a button is a resize, and the rest of it is
/// the button, exactly as it is on a window whose frame the system draws.
pub(super) unsafe fn pinned_press(hwnd: HWND, x: i32, y: i32) -> bool {
    let Some(caption) = pinned_caption_geometry() else {
        return false;
    };
    let framed = caption.frame != PinFrame::None;

    // Anything but the volume button itself puts its popup away: the panel floats over the media,
    // and a hand that has come for the picture, the title bar or an edge is not a hand on the
    // level. The button is left out because a click on it is what puts the popup away — closing it
    // here would have the release open it straight back up (see `pinned_transport_release`).
    if !pin_point_is_volume_button(x, y) && close_pin_volume() {
        render_layered_preview(hwnd);
    }

    if framed {
        let edge = {
            let Some(pinned) = pin_state() else {
                return false;
            };
            pinned.pin().and_then(|pin| pin.resize_edge(x, y))
        };
        if let Some(edge) = edge {
            begin_pin_drag(hwnd, &Win32PinWindow, PinDragAction::Resize(edge), true);
            return true;
        }
    }

    if y < caption.height {
        // The buttons are drawn over the picture only while the chrome is there to be used: with
        // it gone, the strip across the top of a pinned picture is the picture, and a press on it
        // is the handle every other part of the media is (see `PinChrome`). A kind with no caption
        // has no strip at all, so nothing is answered from here and every press below it is the
        // card's own (see `pinned_caption_height`).
        if caption.wanted {
            let button =
                pin_chrome::button_at(x, y, caption.width, caption.height, caption.dpi, framed);
            if let Some(button) = button {
                if let Some(mut pinned) = pin_state() {
                    if let Some(pin) = pinned.pin_mut() {
                        pin.pressed = Some(button);
                    }
                }
                let _ = SetCapture(hwnd);
                render_layered_preview(hwnd);
                return true;
            }
        }

        begin_pin_drag(hwnd, &Win32PinWindow, PinDragAction::Move, true);
        return true;
    }

    // The transport bar, which is the strip along the bottom of a kind that plays.
    if pinned_transport_press(hwnd, x, y) {
        return true;
    }

    // And then the four controls a sound's card carries, which are drawn on the card rather than in
    // a strip of their own, and so are asked of before the media becomes a handle for moving the
    // window: a button is a button, and a hand on the rest of the card is still a hand carrying the
    // window (see `pinned_audio_control_press`).
    if pinned_audio_control_press(hwnd, x, y) {
        return true;
    }

    if pinned_content_is_the_pins(x, y) {
        begin_pin_drag(hwnd, &Win32PinWindow, PinDragAction::Move, true);
        return true;
    }

    false
}

/// A press on one of the controls a pinned sound's card carries: the walk either side of the
/// play/pause button, the bar itself, and the button that opens the volume popup — the four
/// controls drawn on the card's own row rather than in a strip of the window, and the only way of
/// moving a pinned sound within its own file or between files.
///
/// The controls are asked of the card's own layout rather than of anything this app draws here,
/// because the marks being pressed are ones the card's painter drew (see
/// `audio_preview::control_at`). The seek is the odd one out: it is taken where it landed, with
/// nothing captured and no drag following, since a bar on a card is a place to press rather than a
/// thing to be carried — the whole of the window is already the thing to be carried, and a press on
/// the rest of the card does that instead. The other three are held and clicked on the release, the
/// rule every caption button of this app's follows (see `pinned_audio_control_release`).
///
/// Nothing is *played* from here: the clock a card is drawn from and the player behind a sound this
/// app plays are the preview loop's, so the play/pause is left for it and the seek with it (see
/// `settle_pinned_audio_toggle` and `settle_pinned_audio_seek`).
pub(super) unsafe fn pinned_audio_control_press(hwnd: HWND, x: i32, y: i32) -> bool {
    // Asked of the pin's own three facts rather than of the media's kind, so that this procedure —
    // which the media's own lock can be held while, since the loop's audio block repaints the card
    // under it — never has to take that lock to know what the window is showing (see
    // `pin_shows_an_audio_card`).
    let shows = pin_state()
        .and_then(|pinned| pinned.pin().map(pin_shows_an_audio_card))
        .unwrap_or(false);
    if !shows {
        return false;
    }

    // The card fills the pin's media box, because a card is its own size and is not framed into one
    // (see `PinFrame`), so the box is the width the row is laid out across. The point is taken in
    // the media's own coordinates, which is the frame the card is painted in (see `media_point`).
    let Some(content) = pinned_content() else {
        return false;
    };
    let Some((path, dpi)) = pinned_media_owner() else {
        return false;
    };
    let width = (content.2 - content.0).max(1) as u32;

    // A file that does not say how long it is is drawn with a block crossing its track rather than
    // a played part of it, and there is no second of such a file for a press to mean — so the *seek*
    // is refused and nothing else is. That is a change from what this used to do, and it is the
    // difference between a card whose whole row is dead and a card whose one dead control is the
    // seek: the buttons either side of it are the caption's own walk and the level this window is
    // playing at, and neither of them is a question about the length of a file.
    let duration = pinned_audio_duration(&path).filter(|length| *length > 0.0);

    let (media_x, media_y) = media_point(x, y);
    let Some(control) =
        audio_preview::control_at(media_x, media_y, width, dpi, current_audio_options(), true)
    else {
        return false;
    };

    match control {
        CardControl::Seek => {
            let Some(share) = audio_preview::bar_share_at(
                media_x,
                media_y,
                width,
                dpi,
                current_audio_options(),
                true,
            ) else {
                return false;
            };
            let Some(duration) = duration else {
                return false;
            };

            if let Ok(mut request) = PIN_AUDIO_SEEK.lock() {
                *request = Some((duration * share).clamp(0.0, duration));
            }
            AUDIO_CARD_DIRTY.store(true, Ordering::Release);
        }
        _ => {
            with_pin(|pin| pin.audio_pressed = Some(control));
            let _ = SetCapture(hwnd);
            render_layered_preview(hwnd);
        }
    }

    true
}

/// What a release on a pinned sound's card's own controls does: a button is clicked where the
/// pointer is still on it, the same rule every caption button of this app's follows — and the seek
/// does nothing at all, because a seek was already taken where it was pressed.
///
/// The two ends of the walk are the caption's own walk rather than this card's: `PinCommand::
/// Previous` and `PinCommand::Next` are what the caption's arrows ask for, so a step from a button
/// on the card and a step from a button on the caption are one walk and not two.
pub(super) unsafe fn pinned_audio_control_release(hwnd: HWND, x: i32, y: i32) -> bool {
    // No button of the card was held, so this arm has nothing to answer and nothing of the
    // pointer's to give back: the capture a press took belongs to the arm that armed it, and a
    // release here armed nothing (see `pinned_release`).
    let Some(pressed) = pin_state().and_then(|mut pinned| pinned.pin_mut()?.audio_pressed.take())
    else {
        return false;
    };

    let _ = ReleaseCapture();

    let still_on_it = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            Some(pin_audio_control_at(pin, x, y) == Some(pressed))
        })
        .unwrap_or(false);

    if still_on_it {
        match pressed {
            // The player's own state — whether it is going — is the loop's, so this is a door
            // rather than a call (see `settle_pinned_audio_toggle`).
            CardControl::Play => ask_pin_audio_toggle(),
            CardControl::Previous => ask_pin(PinCommand::Previous),
            CardControl::Next => ask_pin(PinCommand::Next),
            // The panel is over this app's own card rather than over a player's window, so raising
            // the pin's window for it is harmless rather than necessary (see `toggle_pin_volume`).
            CardControl::Volume => toggle_pin_volume(hwnd),
            // The two window buttons in the top margin act on the pin rather than on the sound,
            // and are the one way out of a pinned sound that does not ask for the keyboard: the
            // minimize shrinks the pin into the bubble it leaves, and the close ends the pin, the
            // player and the window together. Both are the pin's own commands, which the loop
            // answers the way it answers a caption button (see `pin_command_request`).
            CardControl::Minimize => ask_pin(PinCommand::Minimize),
            CardControl::Close => ask_pin(PinCommand::Close),
            CardControl::Seek => {}
        }
    }

    render_layered_preview(hwnd);
    true
}

/// A release on a pinned window, answering whether it was the pin's to act on: the button a press
/// landed on is clicked if the pointer is still on it, and a drag — which is over wherever the
/// pointer left it — asks for the media to be laid out again at the box the window ended up with.
///
/// **An arm that declines does not touch the pointer, and only the arm that owns a press may give
/// it back.** `ReleaseCapture` answers this window procedure before the call returns, so a release
/// raised by an arm that has nothing to say would run `pin_capture_lost` before the caption's own
/// arm had read what the press armed — and that road drops the pressed button, the bar's aim and
/// the knob out from under the hand. Every button on a pinned window is a button that is pressed
/// and then released while some arm before it has nothing to answer, which is why an arm's own
/// refusal has to cost nothing at all. A press no arm owns is given back once, after every arm has
/// declined, by the procedure's last resort (see `release_pin_capture`).
pub(super) unsafe fn pinned_release(hwnd: HWND, x: i32, y: i32, window: &dyn PinWindow) -> bool {
    // A drag of the volume knob first, which is a hand on the level rather than on anything else:
    // it is the one press on a pinned window that is let go of somewhere other than where it began
    // (see `pinned_volume_press`).
    if pinned_volume_release(hwnd, window) {
        return true;
    }

    // The transport bar next: a bar a press has taken hold of is the bar's pointer until it lets
    // go, whatever else is under it.
    if pinned_transport_release(hwnd, x, y, window) {
        return true;
    }

    // And then the four buttons a sound's card carries, which are the same rule and the same
    // order: a press that has taken hold of one is that button's until it lets go.
    if pinned_audio_control_release(hwnd, x, y) {
        return true;
    }

    // Only the button is taken here. The drag is not, because letting go of one is the same work
    // whichever end it is asked from, and `finish_pin_drag` is that work: a press that arrived as
    // a message ends here, and one read off the hook's published button state is ended by the
    // tick instead (see `settle_pinned_engine_drag`). A window's own press does not press a
    // caption button and drag the window at once, so the two are not in competition.
    // No pin at all is the tail's own answer rather than this gate's: `finish_pin_drag` below is
    // reached either way, and it lets the pointer go of a capture it cannot find a drag for.
    let (pressed, caption_height, dpi, width, framed) = {
        let Some(mut pinned) = pin_state() else {
            return false;
        };
        let Some(pin) = pinned.pin_mut() else {
            return false;
        };

        let pressed = pin.pressed.take();
        let (width, _) = pin.window_size();
        (
            pressed,
            pin.caption,
            pin.dpi,
            width,
            pin.frame != PinFrame::None,
        )
    };

    if let Some(button) = pressed {
        let _ = ReleaseCapture();

        let still_on_it = y < caption_height
            && pin_chrome::button_at(x, y, width, caption_height, dpi, framed) == Some(button);

        if still_on_it {
            match button {
                pin_chrome::CaptionButton::Minimize => ask_pin(PinCommand::Minimize),
                pin_chrome::CaptionButton::Maximize => ask_pin(PinCommand::Maximize),
                pin_chrome::CaptionButton::Close => ask_pin(PinCommand::Close),
                pin_chrome::CaptionButton::Previous => ask_pin(PinCommand::Previous),
                pin_chrome::CaptionButton::Next => ask_pin(PinCommand::Next),
                // The file is in hand here and the Shell takes it as it stands, so this one
                // is asked for where it was pressed rather than written down for the loop
                // to pick up: nothing about opening a file is the loop's business.
                //
                // Except the pin's end, which is the loop's own: a preview window left on
                // screen is a window over the program the button was pressed to start, and
                // nothing the program does afterwards is any of this app's business. So both
                // buttons that open something hand the file over and then ask for the end,
                // with no question asked of when the opened thing is finished with (see
                // `show_open_with_dialog`).
                //
                // The end is asked for whatever came back, and that is a decision rather than
                // an oversight, because the arrangement it replaces is the one that read the
                // Shell's answer and kept the pin up on an error: a format with nothing
                // filed against it left a preview standing over a file the user was still
                // reading, and the list button's end had to be *detected* — a process
                // parented to the rundll32 running the dialog, and the file's own handler
                // read either side of it — which is a question about a window this app does
                // not own, on a thread it does not run, and can come back empty for a
                // cancelled dialog as readily as for a launch. So three things are given up
                // to have the one that matters: a cancelled list leaves no pin, a hand-off
                // that started nothing still takes the pin down, and nothing is left watching
                // a process. Each costs a hover to undo. What is not given up is a topmost
                // window sitting over the program the button exists to start.
                pin_chrome::CaptionButton::OpenWith => {
                    if let Some(path) = pinned_path() {
                        open_path_with_default_app(&path);
                        request_pin_end();
                    }
                }
                // The Shell's own list, the same way round: the file is in hand, the dialog
                // is the system's, and the pin comes down as the dialog goes up.
                pin_chrome::CaptionButton::OpenWithList => {
                    if let Some(path) = pinned_path() {
                        show_open_with_dialog(&path);
                        request_pin_end();
                    }
                }
            }
        }

        render_layered_preview(hwnd);
        return true;
    }

    finish_pin_drag(hwnd, window)
}

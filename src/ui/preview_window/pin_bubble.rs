//! The pin's bubble: the smaller window that stands beside it for a collapsed pin, the mark
//! drawn in it, and the drag that carries it.

use super::*;

/// The class the round bubble a collapsed pin leaves is created from.
pub(super) const PIN_BUBBLE_CLASS: PCWSTR = w!("RustHoverPreviewPinBubble");

/// The bubble's window, or zero while no pin is collapsed.
pub(super) static PIN_BUBBLE_HWND: AtomicIsize = AtomicIsize::new(0);

/// A drag of the bubble in progress, as the two things the press measured: where in the bubble the
/// hand took hold of it — the pointer's offset from the window's own corner — and where the pointer
/// itself was on the screen. The offset is what the bubble is placed by, so that it is carried from
/// the point it was grabbed at rather than having its middle put under the hand; the press point is
/// what a press that turns out to be a click is measured against (see `drag_pin_bubble`).
pub(super) type BubbleDrag = ((i32, i32), (i32, i32));

pub(super) static PIN_BUBBLE_DRAG: Lazy<Mutex<Option<BubbleDrag>>> = Lazy::new(|| Mutex::new(None));

/// How long the bubble's own wait runs before the drag is looked at again, while the button that
/// took hold of it is down.
///
/// The wait itself ends on the pointer's next message — a mouse either moves or it does not, and
/// the one message that is not a move is the release — so this is only the ceiling on a drag that
/// has lost its capture without one being sent, which is what keeps the wait from being a hang.
pub(super) const PIN_BUBBLE_DRAG_WAIT_MS: u32 = 100;

/// How long a bubble drag may hold its loop before the drag is given back.
///
/// The same shape as `PIN_DRAG_WAIT_MS` and for the same reason: the wait above is a ceiling on a
/// drag whose capture has gone, not a bound on how long a hand may be moving. A drag that has
/// lost its capture without a message saying so ended at the first wait that returned; a drag
/// with the button genuinely held down is bounded by this instead, and past it the drag is
/// looked at again by whatever asked for it rather than by a loop that never gives it back.
pub(super) const PIN_BUBBLE_DRAG_CARRY_MS: u64 = 2_000;

/// Whether the drag in progress has moved at all: what tells a click on the bubble — which puts
/// the window back up — from a hand that was carrying it somewhere.
pub(super) static PIN_BUBBLE_MOVED: AtomicBool = AtomicBool::new(false);

/// Collapse a pinned window into the round bubble it leaves: the window and everything standing
/// in it come off the screen, and a small circle takes their place. A collapsed pin is still a
/// pin — previews are still held back until it is restored or closed, and what it is playing is
/// held where it is by the two `Pin Mode → Pause Preview` switches (see `PIN_COLLAPSED` and
/// `settle_bubble_playback`).
pub(super) fn collapse_pin() {
    let anchor = {
        let Some(mut pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin_mut() else {
            return;
        };
        if pin.collapsed {
            return;
        }

        // The bubble takes the place of the button the hand went for, which is
        // looked up in the layout of what the pin is showing — a sound's
        // minimize is one of the two window buttons its card carries, and that
        // question is asked while the pin still shows its card, before the
        // collapse below takes it off it (see `pinned_minimize_box`).
        let anchor = pinned_minimize_box(pin);

        pin.collapsed = true;

        // A bubble has no bar and no level to be read off: the popup goes with the window it was
        // drawn over, and a pin put back up is put back up without it.
        pin.volume.open = false;
        pin.volume.dragging = false;

        anchor
    };

    // A window that is being taken off the screen is not a window the user is in, and the bubble
    // that takes its place cannot be — it is a circle standing in for a window, and it is made the
    // way every preview is. The keyboard goes back now rather than whenever Windows notices the
    // window has gone, so that a pin collapsed with the caret on it does not leave the caret on a
    // window that is not there.
    //
    // Asked for outside the lock above because giving a focus up is a message to this app's own
    // window procedure, and a window procedure that arrives back here while the lock is held would
    // be a second thread waiting on it — this one, from its own loop.
    give_the_keyboard_back(&Win32PinWindow);

    unsafe {
        hide_pinned_windows();
        show_pin_bubble(anchor);
    }
}

/// The box of a pinned window's minimize button, in screen coordinates: what each collapse puts its
/// bubble on.
///
/// The button is *found* rather than the window's corner assumed, because the two are not the same
/// place and never were: a caption puts its buttons against its right edge in Windows' order —
/// minimize, maximize, close — so a window's corner is the *close* button's place, and a pin with
/// nothing to maximize carries two buttons where another carries three, which moves the minimize
/// without moving the corner (see `pin_chrome::button_boxes`). What the buttons are laid out in is
/// the caption strip across the top of the window (see `pinned_caption_height`).
///
/// A sound has no caption, and its minimize is one of the two window buttons its card carries in
/// its top margin instead: the box a press on it is answered against, which reaches to the window's
/// own top edge and so puts the bubble in the top-right corner where every other pin's bubble
/// stands (see `audio_preview::control_box`).
pub(super) fn pinned_minimize_box(pin: &PinnedPreview) -> ScreenRegion {
    let window = pin.window_box();

    if pin_shows_an_audio_card(pin) {
        let (width, _) = pin.window_size();
        if let Some(box_) = audio_preview::control_box(
            CardControl::Minimize,
            width as u32,
            pin.dpi,
            pinned_audio_options(pin),
            true,
        ) {
            return (
                window.0 + box_.left,
                window.1 + box_.top,
                window.0 + box_.right,
                window.1 + box_.bottom,
            );
        }
    }

    let caption = pin.caption;

    let minimize = pin_chrome::button_boxes(
        window.2 - window.0,
        caption,
        pin.dpi,
        pin.frame != PinFrame::None,
    )
    .into_iter()
    .find(|button| button.kind == pin_chrome::CaptionButton::Minimize);

    // Every caption carries a minimize (see `pin_chrome::button_boxes`); a window whose caption
    // somehow did not is a window whose own box is the place a bubble goes.
    let Some(button) = minimize else {
        return window;
    };

    (
        window.0 + button.rect.left,
        window.1 + button.rect.top,
        window.0 + button.rect.right,
        window.1 + button.rect.bottom,
    )
}

/// Put a collapsed pin back up: the bubble goes, the window comes back where it was, whatever
/// stands in its media band — the player's window, the browser's — is put back with it, and what
/// the collapse parked is started again by the tick this runs in (see `settle_bubble_playback`).
pub(super) fn restore_pin() {
    {
        let Some(mut pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin_mut() else {
            return;
        };
        if !pin.collapsed {
            return;
        }

        pin.collapsed = false;

        // The window goes back up where a preview of it would have been put at the pointer the
        // bubble was clicked with, rather than where it stood when it was collapsed: the bubble is
        // the bit of the pin the hand is on, and what the hand gets back is a window placed beside
        // it the way this app places everything else (see `placed_pin_box`).
        if let Some(window) = placed_pin_box(pin, &DESKTOPS) {
            pin.content =
                content_box_of(window, pin.dpi, pin.transport_bar, pin.overlay, pin.caption);
        }
    }

    hide_pin_bubble();

    unsafe {
        let hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
        if !hwnd.is_invalid() {
            show_pinned_window(hwnd);
        }
    }

    // A document or a specimen the browser was drawing went down with the window it stood in,
    // so it is asked for again the way a hover asks for it — and a player's window, which
    // nothing of this app's owns, is simply put back where the media band is.
    if let Some((path, _)) = pinned_media_owner() {
        let kind = CURRENT_MEDIA
            .lock()
            .ok()
            .and_then(|media| media.as_ref().map(|media| media.media_type));
        if matches!(
            kind,
            Some(MediaType::EngineSvg) | Some(MediaType::EngineFont)
        ) {
            if let Some(content) = pinned_content() {
                webview_preview::show(
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
        } else {
            place_pinned_siblings();
        }
    }
}

/// What the bubble a collapsed pin left has parked, if anything (see `BubblePause`).
pub(super) fn pin_bubble_pause() -> Option<BubblePause> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    pin.bubble_pause
}

/// Write an answer about what the bubble has parked back into the pin, if there is still one.
pub(super) fn update_pin_bubble_pause(park: Option<BubblePause>) {
    with_pin(|pin| pin.bubble_pause = park);
}

/// Keep a pin's playback in step with its bubble: hold what is playing when the pin is collapsed
/// — each kind behind the switch of its own, a video behind `Pin Mode → Pause Preview → Video`
/// and a sound behind `... → Audio` — and put it back the way it was when the pin comes up again.
///
/// It runs on the tick beside the pin's own commands, and the sound's clock is half of why: a
/// sound FFmpeg plays is timed by this app over the moment its player was started and the second
/// of the file it was started at, both of which are the loop's own (see `audio_clock`), so a park
/// made anywhere else would be a sound put back at the beginning of its file. The other half is
/// that the media, the player and the engine's session are this thread's, so this is the thread
/// the whole of it can be asked of (see `video_player`).
///
/// The switches are read every tick rather than once at the collapse, so one thrown while the pin
/// is a bubble is answered on the next tick: a film left running because its switch was off is
/// held the moment the switch is turned on. The other way round is deliberately not symmetric — a
/// park is not given back until the pin is up again, because a player started beside a bubble
/// would be a picture on screen next to the one thing a collapse leaves there.
pub(super) fn settle_bubble_playback(
    audio_started: &mut Option<Instant>,
    audio_start_offset: &mut f64,
) {
    let (collapsed, parked) = {
        let Some(pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin() else {
            return;
        };

        (pin.collapsed, pin.bubble_pause)
    };

    if collapsed {
        if parked.is_some() {
            return;
        }

        if let Some(park) = bubble_playback_to_park(audio_started, audio_start_offset) {
            update_pin_bubble_pause(Some(park));
        }
    } else if let Some(park) = parked {
        put_back_bubble_playback(park, audio_started, audio_start_offset);
        update_pin_bubble_pause(None);
    }
}

/// What the bubble owes the media on screen, if anything: what is playing, behind the switch that
/// kind answers to, and what it takes to put it back (see `BubblePause`).
pub(super) fn bubble_playback_to_park(
    audio_started: &mut Option<Instant>,
    audio_start_offset: &mut f64,
) -> Option<BubblePause> {
    let kind = current_media_type()?;

    match kind {
        MediaType::NativeVideo | MediaType::Video if pin_pause_video() => {
            let (path, content, transport, _) = pinned_playback_state()?;

            // What is not playing is not parked: a video the bar was used to pause before the
            // pin was collapsed is one a restore has nothing to put back, and the second it is
            // held at is the bar's business rather than the bubble's.
            if !pin_is_playing(&transport) {
                return None;
            }

            let playhead = pin_playhead(&transport).unwrap_or(0.0);
            toggle_pinned_playback(&path, content, transport);

            Some(match kind {
                MediaType::NativeVideo => BubblePause::Engine,
                _ => BubblePause::Player(playhead),
            })
        }
        MediaType::Audio if pin_pause_audio() => {
            // A sound FFmpeg plays is a process, and a process is parked by being ended: where it
            // had got to is read before it goes, because the clock that measures it goes with it,
            // and the player that takes its place is begun at that second (see
            // `put_back_bubble_playback`). A sound the engine Windows has is held where it
            // stands, the way a video is.
            let playing = CURRENT_MEDIA
                .lock()
                .ok()
                .and_then(|media| media.as_ref().map(|media| media.video_process.is_some()))
                .unwrap_or(false);

            if playing {
                let (path, _) = pinned_media_owner()?;
                // Nothing to do about a hold here: a sound a key has paused has no player to
                // park, and a restore that parks nothing is a restore that changes nothing
                // (see `toggle_pinned_audio`).
                let played = audio_clock(&path, *audio_started, *audio_start_offset, None)
                    .0
                    .unwrap_or(0.0);

                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                    if let Some(media) = current.as_mut() {
                        kill_player_process(media);
                    }
                }
                // The clock the card is drawn from is taken down with the player: seconds that
                // went on moving under a bubble would be a card claiming a sound is playing, and
                // the park is what put this one where it is (see `audio_clock`).
                *audio_started = None;

                Some(BubblePause::Player(played))
            } else if video_player::is_playing() {
                video_player::set_paused(true);

                Some(BubblePause::Engine)
            } else {
                None
            }
        }
        _ => None,
    }
}

/// Put back what the bubble parked, in the way the engine that parked it answers to.
///
/// A video goes back through the transport, which is where a pause of this app's is kept: the
/// media engine is told to go on and the second the bar was held at is cleared, and a player
/// FFmpeg's is begun again at that second (see `toggle_pinned_playback`). A sound the engine
/// Windows has is simply told to go on; one FFmpeg plays is begun again at the second the park
/// wrote down, with the clock the card is drawn from set to it, so the card goes on from where
/// the sound was rather than from the beginning of the file or from the moment the pin returned.
pub(super) fn put_back_bubble_playback(
    park: BubblePause,
    audio_started: &mut Option<Instant>,
    audio_start_offset: &mut f64,
) {
    match (current_media_type(), park) {
        (Some(MediaType::Video) | Some(MediaType::NativeVideo), _) => {
            if let Some((path, content, transport, _)) = pinned_playback_state() {
                if !pin_is_playing(&transport) {
                    toggle_pinned_playback(&path, content, transport);
                }
            }
        }
        (Some(MediaType::Audio), BubblePause::Engine) => video_player::set_paused(false),
        (Some(MediaType::Audio), BubblePause::Player(from)) => {
            if let Some((path, _)) = pinned_media_owner() {
                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                    if let Some(media) = current.as_mut() {
                        if start_audio_playback(&path, media, from) && media.video_process.is_some()
                        {
                            *audio_started = Some(Instant::now());
                            *audio_start_offset = from;
                        }
                    }
                }
            }
        }
        _ => {}
    }
}

/// Where a collapsed pin's window goes back up when its bubble is clicked: the box a preview of the
/// pin would have been given at the bubble, which is the placement the tray's `Placement → Position`
/// setting asks for and is made by the same algorithm every hover's preview is placed by
/// (`compute_mouse_layout`).
///
/// The bubble's own place is what is handed over as the pointer's — the hand that clicked it was
/// there — so what comes out is where a preview of this pin would have gone at that point.
///
/// What is placed is the pin's *window* — the media box with the caption above it and the transport
/// bar below — rather than the media alone, so the display's room has to hold the whole of it.
///
/// A position mode decides *where* a window goes, so the box it keeps is the box it has: the
/// placement is asked for the size the window already is, the media is not measured again the way a
/// hover measures it, and a window the room has no space for keeps its own size against the room's
/// edge rather than being shrunk into a corner of it.
///
/// A pin collapsed while it was maximized is the one box that is not placed at all: it is a window
/// the user has already settled, and placing it again would be the app moving a window the user put
/// there. It goes back exactly as it went down (see `pin_restore_box`).
///
/// A hover's placement steps around the name the file is listed under (see `avoiding_text`); a
/// bubble is on no name, so the pointer's own standoff is the whole of the clearance.
///
/// The display arrives as an argument, and that is the second of the two functions in this file
/// that could not be tested because it worked the display out for itself. Where a box goes back
/// up is a decision about the machine it is going up on, and the one that matters is which
/// display that is: a bubble left at the right-hand edge of a display the second monitor begins
/// beside must come back up on the second, or it comes back up under the edge of the first and
/// off the bottom of the desktop.
pub(super) fn placed_pin_box(pin: &PinnedPreview, displays: &dyn Displays) -> Option<ScreenRegion> {
    if pin.restore.is_some() {
        return None;
    }

    let window = pin.window_box();
    let width = (window.2 - window.0).max(1);
    let height = (window.3 - window.1).max(1);

    let (anchor_x, anchor_y) = pin_bubble_centre().or_else(cursor_screen_point)?;
    let display = displays::display_at(displays, anchor_x, anchor_y);
    let bounds = display.work_area;
    let dpi = display.dpi;

    let layout = compute_mouse_layout(
        anchor_x,
        anchor_y,
        HoverPlacement {
            orig_dims: (width as u32, height as u32),
            avoid: None,
            follow_cursor: CONFIG
                .lock()
                .map(|config| config.follow_cursor)
                .unwrap_or(true),
            preview_scale: PreviewScale::Percent(100),
            at_the_pointer_corner: false,
        },
        bounds,
        dpi,
    )?;

    let x = layout
        .pos_x
        .clamp(bounds.left, (bounds.right - width).max(bounds.left));
    let y = layout
        .pos_y
        .clamp(bounds.top, (bounds.bottom - height).max(bounds.top));

    Some((x, y, x + width, y + height))
}

/// The middle of the round bubble a collapsed pin left, while one is up: the place a restore is
/// placed at, since what clicked the bubble was a hand on the bubble (see `placed_pin_box`).
pub(super) fn pin_bubble_centre() -> Option<(i32, i32)> {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return None;
    }

    let (left, top, width, height) = window_origin(HWND(hwnd as *mut _))?;
    Some((left + width / 2, top + height / 2))
}

/// Take a pinned window and everything of somebody else's that stands in it off the screen:
/// what a collapse leaves behind, and what a pin that ends leaves to the ordinary take-down.
pub(super) unsafe fn hide_pinned_windows() {
    let hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
    if !hwnd.is_invalid() {
        let _ = ShowWindow(hwnd, SW_HIDE);
    }

    let video = VIDEO_HWND.load(Ordering::SeqCst);
    if video != 0 {
        let _ = ShowWindow(HWND(video as *mut _), SW_HIDE);
    }

    if webview_preview::is_showing() {
        webview_preview::hide();
    }
}

/// Take the round bubble down, if one is up.
pub(super) fn hide_pin_bubble() {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return;
    }

    unsafe {
        let _ = ShowWindow(HWND(hwnd as *mut _), SW_HIDE);
    }
}

/// Put the round bubble up, created the first time a pin is collapsed.
///
/// It is put *on the box it is given*, centered on it: that box is the minimize button's own (see
/// `pinned_minimize_box`), so the circle a collapse leaves stands where the button that collapsed
/// the window stood, every time — a collapse is the same gesture twice and lands in the same place
/// twice, and where the hand last left a bubble is not a place the next collapse would be expected
/// to go. It is painted before it is shown, for the reason every other layered window of this app's
/// is: what one shows between two paints is the surface it already has.
pub(super) unsafe fn show_pin_bubble(anchor: ScreenRegion) {
    let dpi = dpi_at(anchor.0, anchor.1);
    let side = logical_px(dpi, PIN_BUBBLE_PIXELS).max(16);

    let centre_x = (anchor.0 + anchor.2) / 2;
    let centre_y = (anchor.1 + anchor.3) / 2;
    let (x, y) = (centre_x - side / 2, centre_y - side / 2);
    let clamped = clamp_pinned_box((x, y, x + side, y + side), dpi, &DESKTOPS);
    let (x, y) = (clamped.0, clamped.1);

    let Some(hwnd) = pin_bubble_window() else {
        return;
    };

    if !IsWindowVisible(hwnd).as_bool() {
        let _ = MoveWindow(hwnd, x, y, side, side, false);
    }

    render_pin_bubble(hwnd);
    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        x,
        y,
        side,
        side,
        SWP_NOACTIVATE | SWP_SHOWWINDOW,
    );
    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
}

/// The bubble's window, created once for the run: a layered popup that never takes focus and is
/// never in the taskbar, exactly as the preview window itself is.
pub(super) unsafe fn pin_bubble_window() -> Option<HWND> {
    let existing = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if existing != 0 {
        return Some(HWND(existing as *mut _));
    }

    let hinstance = GetModuleHandleW(None).ok()?;
    let hwnd = CreateWindowExW(
        WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE,
        PIN_BUBBLE_CLASS,
        w!("Pinned preview"),
        WS_POPUP,
        0,
        0,
        1,
        1,
        None,
        None,
        hinstance,
        None,
    )
    .ok()?;

    PIN_BUBBLE_HWND.store(hwnd.0 as isize, Ordering::SeqCst);
    Some(hwnd)
}

/// Paint the bubble: the round window a collapsed pin leaves, with the picture the pin was
/// showing inside it where its kind has one (see `pin_chrome::paint_bubble`).
pub(super) unsafe fn render_pin_bubble(hwnd: HWND) {
    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return;
    }

    let width = (rect.right - rect.left).max(1) as u32;
    let height = (rect.bottom - rect.top).max(1) as u32;
    let Some(bits) = ensure_layered_surface(hwnd.0 as isize, width, height) else {
        return;
    };
    let out = std::slice::from_raw_parts_mut(bits, width as usize * height as usize * 4);

    let Some(palette) = pin_chrome::ChromePalette::current() else {
        return;
    };

    let (thumbnail, mark) = bubble_art(width);
    let art = thumbnail.as_deref().map(|pixels| (pixels, width, height));

    pin_chrome::paint_bubble(out, width, height, &palette, art, mark);

    let dst_point = POINT {
        x: rect.left,
        y: rect.top,
    };
    let size = SIZE {
        cx: width as i32,
        cy: height as i32,
    };
    let src_point = POINT { x: 0, y: 0 };
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };

    if let Some(mem_dc) = layered_surface_dc(hwnd.0 as isize) {
        let _ = UpdateLayeredWindow(
            hwnd,
            None,
            Some(&dst_point),
            Some(&size),
            mem_dc,
            Some(&src_point),
            COLORREF(0),
            Some(&blend),
            ULW_ALPHA,
        );
    }
}

/// What the bubble is drawn with: the picture the pin holds, where its kind has one, and the
/// mark its kind is otherwise.
pub(super) fn bubble_art(side: u32) -> (Option<Vec<u8>>, pin_chrome::BubbleMark) {
    let Ok(media) = CURRENT_MEDIA.lock() else {
        return (None, pin_chrome::BubbleMark::Picture);
    };
    let Some(media) = media.as_ref() else {
        return (None, pin_chrome::BubbleMark::Picture);
    };

    let mark = media.media_type.bubble_mark();
    let thumbnail = media
        .media_type
        .has_bubble_picture()
        .then(|| bubble_thumbnail(media, side))
        .flatten();

    (thumbnail, mark)
}

/// The picture inside the bubble: the frame the pin is holding, shrunk to cover a square and
/// cropped to the middle of it.
///
/// It is taken from the frame already in memory rather than from the file: a bubble is drawn at
/// the moment a window collapses, and a decode for it would be a second read of a file whose
/// picture this app is holding at that very moment.
pub(super) fn bubble_thumbnail(media: &MediaData, side: u32) -> Option<Vec<u8>> {
    let (width, height) = (media.current_width(), media.current_height());
    if width == 0 || height == 0 || side == 0 {
        return None;
    }

    let frame = media.current_pixels();
    let expected = width as usize * height as usize * 4;
    if frame.len() < expected {
        return None;
    }

    // A frame is BGRA and the image crate works in RGBA, so the two channels are exchanged on
    // the way in and on the way out.
    let rgba: Vec<u8> = frame[..expected]
        .as_chunks::<4>()
        .0
        .iter()
        .flat_map(|pixel| [pixel[2], pixel[1], pixel[0], pixel[3]])
        .collect();
    let picture = image::RgbaImage::from_raw(width, height, rgba)?;

    // Scaled to *cover* the circle rather than to fit inside it, and cropped to the middle: what
    // a bubble shows of a picture is the middle of it, at the bubble's own size.
    let scale = (side as f32 / width as f32).max(side as f32 / height as f32);
    let scaled_width = ((width as f32 * scale).round() as u32).max(side);
    let scaled_height = ((height as f32 * scale).round() as u32).max(side);
    let scaled = image::imageops::resize(
        &picture,
        scaled_width,
        scaled_height,
        image::imageops::FilterType::Triangle,
    );
    let cropped = image::imageops::crop_imm(
        &scaled,
        (scaled_width - side) / 2,
        (scaled_height - side) / 2,
        side,
        side,
    )
    .to_image();

    let mut out = Vec::with_capacity(side as usize * side as usize * 4);
    for pixel in cropped.pixels() {
        out.extend_from_slice(&[pixel[2], pixel[1], pixel[0], pixel[3]]);
    }

    Some(out)
}

/// The bubble's own window procedure: a click puts the pinned window back, a drag carries the
/// bubble anywhere on the desktop, and a right click takes the pin down — the three things a
/// collapsed pin can be asked for, on the one piece of it that is still on screen.
pub(super) unsafe extern "system" fn pin_bubble_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_LBUTTONDOWN => {
            let cursor = cursor_screen_point();
            let origin = window_origin(hwnd).map(|rect| (rect.0, rect.1));

            // Two things are measured once, here, and held for the whole drag: where in the bubble
            // the hand took hold of it — the pointer's offset from the window's own corner, which is
            // what every move from here on is placed by — and where the pointer itself was, which is
            // what a press that turns out to be a click is measured against (see `drag_pin_bubble`).
            if let (Ok(mut drag), Some(cursor), Some(origin)) =
                (PIN_BUBBLE_DRAG.lock(), cursor, origin)
            {
                *drag = Some(((cursor.0 - origin.0, cursor.1 - origin.1), cursor));
            }
            PIN_BUBBLE_MOVED.store(false, Ordering::Release);
            let _ = SetCapture(hwnd);
            // And the drag is carried here rather than left to the preview loop's next tick: what
            // has hold of the bubble is the pointer, and the pointer does not wait for a tick of
            // anything (see `carry_bubble_drag`).
            carry_bubble_drag(hwnd);
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            // Where the pointer is *now*, read from the cursor rather than from the message: the
            // coordinates a mouse message carries are measured in the window this drag moves.
            drag_pin_bubble(hwnd);
            LRESULT(0)
        }
        WM_LBUTTONUP => {
            let dragging = PIN_BUBBLE_DRAG
                .lock()
                .map(|mut drag| drag.take().is_some())
                .unwrap_or(false);

            if dragging {
                let _ = ReleaseCapture();
                if !PIN_BUBBLE_MOVED.load(Ordering::Acquire) {
                    // A press that did not move is a click, and a click on the bubble is what
                    // puts the window back up.
                    ask_pin(PinCommand::Restore);
                }
            }
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_CAPTURECHANGED => {
            // A capture taken away is a drag that is over wherever the bubble stood: a window
            // that is not holding the pointer is not being carried by it, and the state is what
            // the drag's own wait is ended by (see `carry_bubble_drag`).
            drop_bubble_drag();
            LRESULT(0)
        }
        WM_RBUTTONUP => {
            ask_pin(PinCommand::Close);
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_LBUTTONDBLCLK => {
            ask_pin(PinCommand::Restore);
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_PAINT => {
            let mut ps = PAINTSTRUCT::default();
            let _ = BeginPaint(hwnd, &mut ps);
            render_pin_bubble(hwnd);
            let _ = EndPaint(hwnd, &ps);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

/// Carry the bubble with the pointer: the bubble is put where the pointer is, less the offset it was
/// taken hold of by, and that is the whole of the calculation.
///
/// The drag this replaces *moved* the window by the pointer's messages instead — a delta from the
/// coordinates of one message to the next — and that arithmetic is what was wrong with it, because
/// the frame those coordinates are measured in is the window the drag is moving. A message posted
/// before the window moved reports a delta that this app has already applied; one posted after it
/// reports a delta it never applied; and which of the two a given message is depends on nothing but
/// the timing between the pointer's messages and this app's own `SetWindowPos`. A window moved by
/// its own messages therefore spends the drag correcting its own movement, which is what a hand on
/// it reads as a bubble stuttering, overshooting and shaking rather than sitting under the cursor.
///
/// The cursor is read fresh on every message and the window is put at it less the offset the press
/// took hold of the bubble by. Nothing of the last message, and nothing of the window's own
/// coordinates, reaches the box, so a message that arrives late, twice, or not at all cannot put
/// the window anywhere but under the hand that is holding it.
///
/// That offset is the one thing a *placing* drag cannot do without: a bubble whose middle was put
/// under the pointer would leap on the first move, and a window that jumps the moment it is picked
/// up reads as having been dropped rather than grabbed. It is measured once, at the press.
///
/// A press that has not yet moved past the slop places nothing: a click on the bubble is a click,
/// however much the mouse shivers while it is being made.
pub(super) unsafe fn drag_pin_bubble(hwnd: HWND) {
    let Ok(drag) = PIN_BUBBLE_DRAG.lock() else {
        return;
    };
    let Some((grab, press)) = *drag else {
        return;
    };
    drop(drag);

    let Some((x, y)) = cursor_screen_point() else {
        return;
    };

    let slop = (logical_px(96, PIN_DRAG_SLOP_PIXELS)).max(2);
    if (x - press.0).abs() + (y - press.1).abs() > slop {
        PIN_BUBBLE_MOVED.store(true, Ordering::Release);
    }
    if !PIN_BUBBLE_MOVED.load(Ordering::Acquire) {
        return;
    }

    let Some((_, _, width, height)) = window_origin(hwnd) else {
        return;
    };
    let dpi = dpi_at(x, y);
    let target = clamp_pinned_box(
        (
            x - grab.0,
            y - grab.1,
            x - grab.0 + width,
            y - grab.1 + height,
        ),
        dpi,
        &DESKTOPS,
    );

    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        target.0,
        target.1,
        0,
        0,
        SWP_NOSIZE | SWP_NOACTIVATE,
    );
}

/// Whether a drag of the bubble is in progress: what the wait below is waiting to be over, and
/// what a release is read as a click or as the end of a carry against.
pub(super) fn bubble_is_being_dragged() -> bool {
    PIN_BUBBLE_DRAG
        .lock()
        .map(|drag| drag.is_some())
        .unwrap_or(false)
}

/// Forget a drag of the bubble, if one is in progress.
pub(super) fn drop_bubble_drag() {
    if let Ok(mut drag) = PIN_BUBBLE_DRAG.lock() {
        *drag = None;
    }
}

/// Carry the bubble for as long as the button that took hold of it is held.
///
/// This is the one drag that cannot be left to the preview loop, because it is the one drag of a
/// window with nothing in it: the loop's own wait is a sixteenth of a second with something on
/// screen and longer with nothing, and every one of those waits is time the pointer has moved
/// through and the bubble has not — a small window with the pointer on it is read against the
/// pointer itself, so what a tick of lag costs is the bubble trailing the hand rather than
/// arriving with it. So while the button is down the thread waits on its own message queue rather
/// than on the loop's clock, and a move is carried where it lands (`drag_pin_bubble`) instead of
/// being drained in a batch a tick later.
///
/// What is owed the loop while this runs is the note that the loop is still alive: a bubble being
/// dragged down a desktop for a few seconds is a preview thread at work, not one that has stopped
/// answering, and the hook that ends the engines of a stopped loop reads that note (see
/// `preview_stall_ms`).
///
/// It is over when the drag is: a release is answered by the handler that has always answered it,
/// which is dispatched from here like any message, and a capture taken away says so with the
/// message that takes it away. The wait is bounded for the one case neither leaves a message for:
/// a drag whose state has been emptied for another reason still leaves it within the bound.
pub(super) unsafe fn carry_bubble_drag(hwnd: HWND) {
    // The same bound the other carry takes, and the same refusal to re-enter a drain: this loop
    // and `carry_pin_drag_with_the_hand` were two hand-written loops with two different bounds,
    // and a bubble drag answered from inside a pump would dispatch the same message twice — a
    // click that lands in the wrong place (see `Carry`).
    let mut carry = Carry::begin(PIN_BUBBLE_DRAG_CARRY_MS);

    loop {
        if carry.is_exhausted() {
            return;
        }
        note_preview_alive();

        // Everything the queue holds, answered where it is: a move carries the bubble, and the
        // release ends the drag in the handler it always has.
        if !carry.pump_window_messages() {
            return;
        }

        if !RUNNING.load(Ordering::SeqCst) || !bubble_is_being_dragged() || GetCapture() != hwnd {
            return;
        }

        // Nothing left to answer: what is waited for is the pointer's next message rather than
        // the loop's next tick — and the wait ends on it, so the bubble moves with the hand
        // rather than after it.
        let _ = MsgWaitForMultipleObjectsEx(
            None,
            PIN_BUBBLE_DRAG_WAIT_MS,
            QS_ALLINPUT,
            MWMO_INPUTAVAILABLE,
        );
    }
}

/// The box of a window, as a screen rectangle and its size.
pub(super) fn window_origin(hwnd: HWND) -> Option<(i32, i32, i32, i32)> {
    let mut rect = RECT::default();
    unsafe {
        GetWindowRect(hwnd, &mut rect).ok()?;
    }

    Some((
        rect.left,
        rect.top,
        rect.right - rect.left,
        rect.bottom - rect.top,
    ))
}

/// A box a drag has left a pinned window at, which the preview loop lays the media out for on
/// its next tick.
///
/// The window procedure owns a drag — the pointer is its to follow, and the window has to keep
/// up with it — and the loop owns the media, so one asks and the other does: what a window
/// dragged by an edge owes is the media laid out at the box it ended up with, and that is the
/// same work a maximized window owes (see `PreviewMessage::PinBox`).
pub(super) static PIN_BOX_REQUEST: Lazy<Mutex<Option<ScreenRegion>>> =
    Lazy::new(|| Mutex::new(None));

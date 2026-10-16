//! Compositing a media into a window: the layered surfaces drawn at a position, the bands a
//! pinned window is divided into, and the compose, stretch, resample and blend each is filled
//! by.

use super::*;

pub(super) unsafe fn render_layered_preview(hwnd: HWND) {
    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return;
    }

    // The window is two things: the hover's own frame, and — while a preview is pinned — a
    // window with a caption and a bar of its own. Which of the two it is showing is not a
    // question about its size, so it is asked here rather than worked out by each painter.
    if pinned() {
        render_pinned_preview_at(hwnd, rect.left, rect.top);
    } else {
        render_layered_preview_at(hwnd, rect.left, rect.top);
    }
}

/// Paint the frame the window is holding at a given place on screen, sizing the
/// window to the frame.
///
/// The window is the frame's size, which is what makes drawing a frame into the box
/// the layout planned a rule every loader follows rather than a detail of one of them
/// (see `load_media`): the box is what was fitted to the display, so a frame that grew
/// past it takes the preview past the display's edge with it.
///
/// `UpdateLayeredWindow` applies the place, the size and the surface in one call,
/// which is what lets a window that is already on screen take a frame of another
/// size — the spinner's box first, the page's after it — without a moment of the
/// frame it is holding stretched into the new box: between one call and the next, a
/// layered window shows the surface it already has, at whatever size the window
/// has.
pub(super) unsafe fn render_layered_preview_at(hwnd: HWND, x: i32, y: i32) {
    let Some((width, height)) = (|| {
        let media_guard = CURRENT_MEDIA.lock().ok()?;
        let media = media_guard.as_ref()?;

        // Three kinds are not this window's to draw: a video is played by the player's own
        // window, and a document and a font specimen are drawn by the engine's.
        if matches!(
            media.media_type,
            MediaType::Video | MediaType::EngineSvg | MediaType::EngineFont
        ) {
            return None;
        }

        let width = media.current_width();
        let height = media.current_height();
        let expected_size = width as usize * height as usize * 4;
        if width == 0 || height == 0 || media.current_pixels().len() < expected_size {
            return None;
        }

        // The spinner is nothing but an arc, and what is behind it is the desktop:
        // a backdrop of the configured kind would put back the square its frame is
        // transparent to avoid. Everything else this window draws is composited over
        // the backdrop of its kind — a picture's, the one the tray keeps for a texture, or
        // the one it keeps for the picture a design document is previewed from, each a
        // setting of its own for the reason `dds_image` gives. A document is composited by
        // the engine, over the backdrop of its own, and none of them reaches here.
        let background = preview_background(media.media_type);
        let bits = ensure_layered_surface(hwnd.0 as isize, width, height)?;
        let out = unsafe { std::slice::from_raw_parts_mut(bits, expected_size) };

        // The frame is composed once, and the spinner — where there is one — is drawn over
        // what came out rather than into a copy of the frame that has to be composed again:
        // what is copied is the box the arc sits in, a few thousand pixels of it, and what is
        // composed back is that same box (see `OVERLAY_SCRATCH`).
        compose_preview_pixels_into(
            media.current_pixels(),
            width,
            height,
            background,
            media.current_frame_is_opaque(),
            out,
        );

        if media.should_draw_streaming_overlay() {
            let elapsed = media
                .loading_start
                .map(|s| s.elapsed().as_secs_f32())
                .unwrap_or(0.0);
            let angle = elapsed * 2.0 * std::f32::consts::PI * 1.2;

            if let Some(area) = spinner_overlay_box(width, height) {
                OVERLAY_SCRATCH.with(|cell| {
                    let mut box_pixels = cell.borrow_mut();
                    copy_frame_box_into(media.current_pixels(), width, area, &mut box_pixels);
                    overlay_loading_spinner(&mut box_pixels, area, width, height, angle);
                    compose_preview_block_into(&box_pixels, area, background, out, width);
                });
            }
        }

        Some((width, height))
    })() else {
        return;
    };

    let Some(mem_dc) = layered_surface_dc(hwnd.0 as isize) else {
        return;
    };

    let dst_point = POINT { x, y };
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

    // The hold is deliberately not published from here. Both render paths used to publish it on
    // every frame, which is four lock round-trips and a `GetWindowRect` per painted frame, and it
    // defeated the throttle entirely: the tick's own publish (see `POINTER_HOLD_HEARTBEAT_MS`)
    // asks the same question 500 ms later and was always arriving second. What a paint changes
    // is the window's position, and the tick's key already carries the generation a move or
    // resize bumps, so the region is one heartbeat behind the paint at worst.
}

/// Whether the band of a pinned window is a place Windows hands to the pin, or a hole it hands to
/// whatever stands behind it.
///
/// A pin whose media is a film a player of its own draws is a pin whose band is not its pixels at
/// all: the band is a hole onto that player's window, and a hole is click-through, because hit
/// testing of a layered window is answered by the shape of its pixels. That is what FFmpeg's
/// player is owed, and it is what a preview this app draws for itself costs nothing. **mpv is the
/// one answer that is neither.** Its film is not parked for a gesture, so the window it draws in
/// travels with the band rather than being left behind a strip of the pin's chrome, and the film
/// is therefore the pin's own drag handle: a hand on the picture carries the window, the chrome
/// and the film as one thing. A handle cannot be a hole, so the band is given the least coverage
/// a pixel can have (see `fill_band_hit_testable`).
///
/// Which player is up is read off the recorded player image, which is the record the carry is
/// read off too — a second way to know which player is up would be a second thing to keep in step
/// with this one, and the two arms would be free to disagree about the same film (see
/// `video_process_image`). The kind is read with it because a film a player draws is the only band
/// that is a hole in the first place: every other kind's band is this app's own pixels already,
/// and a band a park has painted flat is left alone, since what stands in it is the parked film.
pub(super) fn film_band_is_hit_testable(media: MediaType, image: &str, parked: bool) -> bool {
    !parked && matches!(media, MediaType::Video) && image == MPV_IMAGE_NAME
}

/// Paint a pinned window: the caption, the media in the band below it, and — for a kind that
/// plays — the transport bar below that.
///
/// It is the same composition `render_layered_preview_at` performs, for a window that is no
/// longer the frame's size: the surface is the window's whole box, the frame is composed into
/// the band the media occupies, and the chrome is drawn over the rest. What makes a pinned
/// preview worth a painter of its own is exactly that band: a frame lands at the row its band
/// starts at rather than at the top of the surface, and the window is as large as the media
/// plus what was added above and below it.
///
/// Three kinds leave the band empty of this app's own pixels, because the thing really
/// drawn there is a window of somebody else's: the two the browser engine draws, whose
/// windows stand in the band, and a video a player of this app's plays, whose window
/// stands in the whole of it — a pin whose chrome is drawn over its media is media the
/// whole of the window down (see `pinned_band_rows`). What the band is then is nothing at
/// all — alpha zero, which is what lets the window underneath be seen through it and,
/// more to the point, be *clicked*: hit testing of a layered window is answered by the
/// shape of its pixels, so a band of transparent ones is a band the player keeps for
/// itself. The one exception is a film mpv draws, whose band is the pin's own handle and
/// is therefore painted the least a pixel can be (see `film_band_is_hit_testable`).
pub(super) unsafe fn render_pinned_preview_at(hwnd: HWND, x: i32, y: i32) {
    let Some(paint) = pinned_paint() else {
        return;
    };

    let (width, height) = (paint.width, paint.height);
    let caption_height = paint.caption_height;
    let Some(bits) = ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32) else {
        return;
    };
    let out = std::slice::from_raw_parts_mut(bits, width as usize * height as usize * 4);
    out.fill(0);

    // The media band. A frame of the band's own size is composed into it row for row; one that is
    // not — which is a window whose edge is under the hand — is sampled into it, so the band is
    // filled by the media at the size it is being dragged to rather than holding a frame of the
    // size it had (see `compose_media_into_band`). Where the band is is a question about the kind:
    // a pin whose chrome is drawn over its media is media the whole of the window down, and one
    // whose chrome has bands of its own is media between them (see `pinned_band_rows`).
    let (band_top, band_height) = pinned_band_rows(
        height,
        caption_height,
        paint.transport_height,
        paint.overlay,
    );
    let band_height = band_height.max(1) as u32;
    let band_width = width.max(1) as u32;
    let mem_dc = layered_surface_dc(hwnd.0 as isize);

    // Where in that band the media lands. A picture is the band's own box unless the wheel has
    // zoomed it into a window that stopped growing, in which case it lands on a box larger than the
    // band and is carried around inside it — and a box outside the band is left to GDI to clip, since
    // what makes it whole is the placement and not the size (see `PinZoom::picture_box`).
    //
    // What `picture_box` answers is a *rectangle* — a left, a top and the two edges the picture runs
    // to — and what a band is filled from is a left, a top and the two *extents* it covers, so the one
    // is turned into the other here. Handing an edge over as an extent is a picture drawn at the
    // window's own size plus the size it overhangs by over two, which is a picture that shrinks as it
    // is zoomed in and is carried the wrong way as it is panned: the whole of what a zoom past the cap
    // must not be.
    let frame = match paint.zoom {
        Some(zoom) => {
            let (left, top, right, bottom) =
                zoom.picture_box((band_width as i32, band_height as i32));

            (left, top, right - left, bottom - top)
        }
        None => (0, 0, band_width as i32, band_height as i32),
    };

    // A player parked for the length of a drag is painted over rather than painted through: the
    // band is transparent everywhere else so that FFmpeg's own window shows through it, and a
    // transparent band with the player hidden is a hole in the desktop shaped like a video. So the
    // band is filled with the last frame the player had on screen, scaled to the box being dragged
    // to exactly as a picture's own frame is (see `compose_parked_band`), and black under it rather
    // than the configured `Background -> Video` — that setting is allowed to be `Transparent`, and
    // a deliberately transparent band is exactly the hole this is here not to leave.
    if paint.parked {
        if let Some(mem_dc) = mem_dc {
            let held = held_video_frame();
            compose_parked_band(
                held.as_ref()
                    .map(|frame| (frame.pixels.as_slice(), (frame.width, frame.height))),
                band_width,
                BandTarget {
                    out,
                    dc: mem_dc,
                    width: band_width,
                    origin_y: band_top.max(0) as u32,
                    height: band_height,
                    frame,
                },
            );
        } else {
            fill_band_opaque(out, band_width, band_top.max(0) as u32, band_height);
        }
    }

    if let Some(mem_dc) = mem_dc {
        if let Ok(media) = CURRENT_MEDIA.lock() {
            if let Some(media) = media.as_ref() {
                if !media.media_type.is_engine() && !matches!(media.media_type, MediaType::Video) {
                    // What the band is filled with: a picture the wheel has taken past its window is
                    // drawn over the tray's own backdrop for a zoom, and every other picture over the
                    // backdrop its own kind asks for. The two are two different things to look at —
                    // what stands behind a picture nobody has touched is the picture's business, and
                    // what fills the band beside one a hand has carried off is the window's (see
                    // `pinned_band_backdrop`).
                    let zoomed = frame != (0, 0, band_width as i32, band_height as i32);
                    let backdrop = pinned_band_backdrop(media.media_type, zoomed);

                    // The rows a zoomed picture has been carried off are the backdrop's rather than
                    // nobody's, and they are filled before the picture goes on top: a layered window
                    // is hit-tested by the shape of its pixels, so a band left transparent where the
                    // picture used to be would be a band a press falls through (see
                    // `fill_band_backdrop`).
                    if zoomed {
                        fill_band_backdrop(
                            out,
                            band_width,
                            band_top.max(0) as u32,
                            band_height,
                            backdrop,
                        );
                    }

                    compose_media_into_band(
                        media.current_pixels(),
                        media.current_width(),
                        media.current_height(),
                        (band_width, band_height),
                        backdrop,
                        media.current_frame_is_opaque(),
                        BandTarget {
                            out,
                            dc: mem_dc,
                            width: band_width,
                            origin_y: band_top.max(0) as u32,
                            height: band_height,
                            frame,
                        },
                    );
                } else if film_band_is_hit_testable(
                    media.media_type,
                    video_process_image(),
                    paint.parked,
                ) {
                    // Nothing is composed into this band — the film is in a window under it — so
                    // what is written here is only ever the pin's own claim on the pixel (see
                    // `film_band_is_hit_testable`).
                    fill_band_hit_testable(out, band_width, band_top.max(0) as u32, band_height);
                }
            }
        }
    }

    // The arc of a window that is waiting for a file, over the band and under the chrome: a
    // caption drawn over it is a caption a hand can still read, which is the whole of what the
    // wait is not allowed to cost the window (see `paint_pin_spinner`).
    paint_pin_spinner(out, width.max(1) as u32, band_top, band_height as i32);

    // The chrome, for kinds that draw it over their media: each strip whole, or not painted at
    // all, which is what makes a picture whose chrome has gone a picture and nothing else — and the
    // two are asked about one at a time, so a hand at one end of a window does not bring out what
    // is at the other end of it (see `PinChrome`). The palette both strips are drawn in is read
    // once for the paint, and not at all where neither is drawn. The caption, across the window's
    // first rows.
    let palette = if paint.caption || (paint.transport_height > 0 && paint.bar) {
        pin_chrome::ChromePalette::current()
    } else {
        None
    };

    if paint.caption {
        if let Some(palette) = palette.as_ref() {
            let caption = pin_chrome::Caption {
                title: &paint.title,
                maximized: paint.maximized,
                maximizable: paint.maximizable,
                hovered: paint.hovered,
                pressed: paint.pressed,
            };

            CHROME_SURFACE.with(|cell| {
                let mut surface = cell.borrow_mut();
                let wanted = (width.max(1) as u32, caption_height.max(1) as u32);
                if surface
                    .as_ref()
                    .map(|surface| (surface.width, surface.height))
                    != Some(wanted)
                {
                    *surface = DibSurface::create(wanted.0, wanted.1);
                }
                if let Some(surface) = surface.as_ref() {
                    pin_chrome::paint_caption(surface, palette, &caption, paint.dpi);
                    copy_surface_rows_into(surface, out, width as u32, 0);
                }
            });
        }
    }

    // The transport bar, for the kinds that play, across the window's last rows.
    if paint.transport_height > 0 && paint.bar {
        if let Some(palette) = palette.as_ref() {
            let state = pin_chrome::TransportState {
                interactive: paint.transport_live,
                playing: paint.playing,
                position: paint.position,
                duration: paint.duration,
                hovered: paint.transport.hovered,
                pressed: paint.transport.pressed,
                volume: paint.volume.level,
                volume_open: paint.volume.open,
            };

            TRANSPORT_SURFACE.with(|cell| {
                let mut surface = cell.borrow_mut();
                let wanted = (width.max(1) as u32, paint.transport_height.max(1) as u32);
                if surface
                    .as_ref()
                    .map(|surface| (surface.width, surface.height))
                    != Some(wanted)
                {
                    *surface = DibSurface::create(wanted.0, wanted.1);
                }
                if let Some(surface) = surface.as_ref() {
                    pin_chrome::paint_transport(surface, palette, &state, paint.dpi);
                    copy_surface_rows_into(
                        surface,
                        out,
                        width as u32,
                        (height - paint.transport_height).max(0) as u32,
                    );
                }
            });
        }
    }

    // The volume popup, over the media above the button it came out of: it is drawn last of
    // everything because it is over everything, and the band it floats over is the media's own —
    // this app's pixels for a picture the media engine draws, this app's own card for a sound, and
    // the hole the player's window stands in for a video FFmpeg plays (see `pin_volume_open`).
    //
    // Which of the three it hangs off is the pin's own question and not this one's: a video's
    // button is on the transport strip and a sound's is on its card, and both open the same panel
    // from the button's own box (see `pinned_volume_geometry`).
    if let (true, Some(popup)) = (paint.volume.open, paint.volume_popup.as_ref()) {
        if let Some(palette) = pin_chrome::ChromePalette::current() {
            pin_chrome::paint_volume_popup(
                out,
                width,
                &palette,
                popup,
                paint.volume.level,
                paint.volume.dragging,
            );
        }
    }

    // The caption's tooltip, floated over the media below the strip — which is where it is
    // because a name like "Open With Adobe Photoshop" is wider than the buttons it describes,
    // and a tooltip drawn inside the strip would cover them for as long as it was up (see
    // `tooltip_layout`). It is drawn last of everything, like the volume popup, because it is
    // over everything.
    paint_pin_tooltip(out, width, height, &paint, caption_height);

    let dst_point = POINT { x, y };
    let size = SIZE {
        cx: width,
        cy: height,
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

    // As in `render_layered_preview_at`: the hold is the tick's to publish, because a publish
    // per painted frame is what made the 500 ms throttle suppress nothing.
}

/// Everything a repaint of a pinned window needs, taken in one look: what the caption says and
/// how large the bands are.
///
/// It is taken as a value rather than borrowed because a repaint may not hold the lock the
/// pointer's own answers are written under — the hold regions are asked for at the end of it,
/// and that question takes the same lock (see `publish_pointer_hold`).
pub(super) struct PinnedPaint {
    pub(super) width: i32,
    pub(super) height: i32,
    pub(super) caption_height: i32,
    pub(super) transport_height: i32,
    /// How the media is drawn inside its band, or nothing for a kind the wheel does not zoom: what
    /// a picture the wheel has taken inside its window is placed by, rather than the band itself (see
    /// `PinZoom`).
    ///
    /// Nothing means the media lands on the whole of its band, which is every repaint of everything
    /// but a zoomed picture.
    pub(super) zoom: Option<PinZoom>,
    /// Whether the chrome is drawn over the media rather than in bands around it, which is what
    /// says where the media's band of the window is and even whether there is one.
    pub(super) overlay: bool,
    /// Whether each strip of it is drawn this frame: each is whole or not there at all, with
    /// nothing in between to draw, and the two are asked about apart (see `PinChrome`).
    pub(super) caption: bool,
    pub(super) bar: bool,
    pub(super) dpi: u32,
    pub(super) title: String,
    pub(super) maximized: bool,
    /// Whether this pin's caption offers a maximize at all (see `PinFrame`).
    pub(super) maximizable: bool,
    pub(super) hovered: Option<pin_chrome::CaptionButton>,
    pub(super) pressed: Option<pin_chrome::CaptionButton>,
    /// What the caption's button under the pointer is saying, and which one it is — nothing,
    /// and nothing at all, while the pointer is somewhere that says nothing.
    ///
    /// It is a value rather than a reference because the paint takes the pin's lock in one look
    /// and gives it up again before anything is drawn, so a name has to be copied out of the
    /// lock rather than held across the drawing (see `PinTooltip::text_for`).
    pub(super) tooltip: Option<PinTooltipPaint>,
    pub(super) transport: PinTransport,
    /// The level this pin plays at, and whether its popup is open (see `PinVolume`).
    pub(super) volume: PinVolume,
    /// Where the volume popup goes, or nothing while it is closed. It is asked of the pin rather
    /// than recomputed at the paint, so that the panel drawn is the panel a press is answered
    /// against (see `pinned_volume_geometry`).
    pub(super) volume_popup: Option<pin_chrome::VolumePopup>,
    /// Whether the bar's controls do anything for the engine playing this file (see
    /// `PinnedPreview::transport_live`).
    pub(super) transport_live: bool,
    pub(super) playing: bool,
    pub(super) position: Option<f64>,
    pub(super) duration: Option<f64>,
    /// Whether this pin's player has been put away for the length of a drag, which is what makes
    /// the band a flat colour rather than a hole the player is supposed to show through.
    ///
    /// It is a field rather than a question asked of the pin at the paint because the paint may
    /// not hold the pin's lock, and because the answer changes without anything about the window
    /// changing — the same window, the same media, the same box, painted a different colour
    /// because a hand is on its edge (see `park_pinned_player`).
    pub(super) parked: bool,
}

/// A repaint's one tooltip, taken out of the pin as a value: which button is saying it and what
/// it is saying. Owned rather than borrowed because the pin's lock is let go of before anything
/// is drawn, and a caption drawn from a name that lived under a lock would be a caption holding
/// that lock for the length of a paint.
pub(super) struct PinTooltipPaint {
    pub(super) kind: pin_chrome::CaptionButton,
    pub(super) text: String,
}

/// What a repaint of a pinned window is composed of, read out of the pin in one look: the pin's
/// lock is let go of before the media engine is asked anything, because the playhead and the
/// length are COM calls and the Explorer hook asks `pinned_path` on a tick of its own (see
/// `pin_media_is_alive`, which lets go for the same reason).
pub(super) fn pinned_paint() -> Option<PinnedPaint> {
    let mut paint = {
        let pinned = pin_state()?;
        let pin = pinned.pin()?;
        let (width, height) = pin.window_size();

        PinnedPaint {
            width,
            height,
            caption_height: pin.caption,
            transport_height: pinned_transport_height(pin.dpi, pin.transport_bar),
            zoom: pin.zoom,
            overlay: pin.overlay,
            // A kind with no caption of its own is asked for none whatever its chrome is doing:
            // there is no row above the media for one to be drawn in, and a strip painted over the
            // first row of a card is a title bar with the controls underneath it (see
            // `pinned_caption_height`).
            //
            // And a caption too narrow to be a command bar is drawn none at all, whatever the pointer
            // is asking for: a window taken down to a hundredth of a display has no room for a strip,
            // and a strip drawn into one is a title bar with no title in it and buttons that do not
            // fit (see `caption_fits`). It is the same question the press is answered by, asked of the
            // same width, so what the caption refuses to draw is what it also refuses to answer.
            caption: pin.caption > 0
                && pin.chrome.caption
                && pin_chrome::caption_fits(width, pin.dpi, pin.frame != PinFrame::None),
            bar: pin.chrome.bar,
            dpi: pin.dpi,
            title: pin
                .path
                .file_name()
                .map(|name| name.to_string_lossy().to_string())
                .unwrap_or_default(),
            maximized: pin.restore.is_some(),
            maximizable: pin.frame != PinFrame::None,
            hovered: pin.hovered,
            pressed: pin.pressed,
            tooltip: pin.tooltip.shown.and_then(|button| {
                pin.tooltip
                    .text_for(button)
                    .map(|text| PinTooltipPaint { kind: button, text })
            }),
            position: None,
            duration: None,
            playing: false,
            transport: pin.transport,
            volume: pin.volume,
            volume_popup: None,
            transport_live: pin.transport_live,
            parked: pin.parked,
        }
    };

    // The panel, asked for after the pin's lock is let go of rather than read out of it, because
    // where it goes is a question about the geometry of a window this one only holds a lock for
    // while it is up — and because a panel placed from the pin's own numbers is the same panel
    // every press and every drag is answered against (see `pinned_volume_geometry`).
    paint.volume_popup = pinned_volume_geometry();

    // Where the bar is drawn: where the pointer has dragged it while a drag is going, and
    // where the file really is otherwise (see `PinTransport`).
    let transport = paint.transport;
    paint.position = transport.seeking.or_else(|| pin_playhead(&transport));
    paint.duration = pin_duration(&transport);
    paint.playing = pin_is_playing(&transport);

    Some(paint)
}

/// Where a frame that is not the whole of a window is composed: the surface it goes on and the
/// memory DC that surface is selected into, its own row width, the row the band the frame goes in
/// begins at, and how tall that band is.
///
/// The DC travels with the pixels because there are two ways to fill a band that is not the
/// frame's size and one of them is not this app's: a frame that is opaque is scaled by GDI,
/// straight into the surface `UpdateLayeredWindow` will read, which is the difference between a
/// window that follows a hand while an edge is dragged and one that stutters behind it (see
/// `stretch_into_band`).
pub(super) struct BandTarget<'a> {
    pub(super) out: &'a mut [u8],
    pub(super) dc: HDC,
    pub(super) width: u32,
    pub(super) origin_y: u32,
    pub(super) height: u32,
    /// Where in the band the media lands, `(left, top, width, height)` measured from the band's own
    /// top-left — the whole of the band for every caller but a zoomed pin, whose picture is drawn
    /// larger than the window it is inside and carried around by the hand (see `PinZoom`).
    ///
    /// It is what makes the band a *box* rather than a frame to be filled: the rows and columns
    /// outside it are left exactly as they were, so a caller that wants a backdrop around a picture
    /// smaller than its band fills one in first (see `fill_band_backdrop`).
    pub(super) frame: (i32, i32, i32, i32),
}

/// Put the media of a pinned window into the band its window has left for it.
///
/// A frame that is already the band's size is composed into it row for row, which is the whole of
/// the cost of an ordinary paint. One that is not is *scaled* into it, which is the one thing a
/// window being resized needs and the reason this is a question rather than a call: a box that
/// changed a moment ago holds a frame of the size it had, and what a drag has to show is that
/// frame filling the box it is being dragged to rather than standing at its old size inside it.
///
/// The scaling itself is GDI's wherever it can be, because the box is under a hand while it
/// happens: a frame whose own pixels are opaque everywhere is stretched by the window's own
/// surface DC, and the pixel-at-a-time sampler below is kept for the frames that need it — one
/// with an alpha channel to composite and a backdrop to composite it over, which is a per-pixel
/// question GDI cannot be asked (see `stretch_into_band` and `resample_into_band`).
pub(super) fn compose_media_into_band(
    bgra: &[u8],
    width: u32,
    height: u32,
    band: (u32, u32),
    background: TransparentBackground,
    opaque: bool,
    mut target: BandTarget<'_>,
) {
    if width == 0 || height == 0 || band.0 == 0 || band.1 == 0 {
        return;
    }

    // A frame that is the band's own size and lands on the whole of it is composed row for row,
    // which is the whole of the cost of an ordinary paint. A frame that lands on *part* of the
    // band is not that case however its size compares, since the rows it lands on are the ones that
    // matter and the rest of the band belongs to the caller (see `BandTarget::frame`).
    let frame = target.frame;
    let lands_on_everything = frame == (0, 0, band.0 as i32, band.1 as i32);

    if lands_on_everything && (width, height) == band {
        compose_preview_pixels_into_band(bgra, width, height, background, opaque, target);
    } else if !opaque || !stretch_into_band(bgra, (width, height), &mut target) {
        resample_into_band(bgra, (width, height), band, background, opaque, target);
    }
}

/// Fill a band of another size with the frame, scaled by GDI, answering whether it was drawn.
///
/// It is `resample_into_band`'s fast road and the reason it exists: the frame is copied into a
/// surface of its own — the media's pixels, not the window's — and `StretchBlt` scales it into
/// the band, which for a box under a hand is the difference between a picture that keeps up with
/// the pointer and one the hand outruns. What GDI is asked for is the quality a still picture is
/// scaled at rather than the fastest one it has (`HALFTONE`), since what is being watched is a
/// file, and what the drag owes the eye at the end of it is still a layout at the box the window
/// ended up with rather than the stretch it was being shown (see `relayout_pinned_media`).
///
/// Only a frame whose every pixel is opaque comes here. That is what makes the stretch the whole
/// of the picture: there is no alpha to composite and no backdrop behind it to composite it over,
/// so the band can be filled by a copy of the pixels at another size — while a frame with an
/// alpha channel is a per-pixel question (what is behind it, and where the checkerboard's squares
/// fall) that GDI cannot be asked, and is sampled the long way instead.
///
/// The alpha byte is forced opaque over the band afterwards rather than trusted to the stretch: a
/// 32-bit `BI_RGB` DIB has no alpha channel as far as GDI is concerned, and what it does with that
/// byte — copy it, interpolate it as a fourth channel, or write zero over it — is not something a
/// layered window can afford to be wrong about (see `UpdateLayeredWindow`'s `ULW_ALPHA`).
pub(super) fn stretch_into_band(
    bgra: &[u8],
    source: (u32, u32),
    target: &mut BandTarget<'_>,
) -> bool {
    let (source_width, source_height) = source;
    let row_bytes = source_width as usize * 4;
    let wanted = row_bytes * source_height as usize;

    if row_bytes == 0 || source_height == 0 || bgra.len() < wanted {
        return false;
    }

    let (out_width, band_height, origin_y, dc) =
        (target.width, target.height, target.origin_y, target.dc);
    let out_row_bytes = out_width as usize * 4;

    // What GDI is asked for is the *whole* of the picture's frame, wherever it lands, and never the
    // part of it the band has room for: the frame's height is the number the source is scaled by, so
    // a destination cut to the rows still on screen is the source stretched into those rows instead
    // — a picture squashed by however much of it has been carried off, which is the whole of what a
    // zoom past the cap must not be (see `BandTarget::frame`).
    //
    // Nothing is lost by asking for all of it: a frame bigger than its window reaches here only for
    // a kind whose chrome is drawn *over* its media, and such a pin's band is the whole of its
    // window — so the rows a picture overhangs by are rows the surface has not got, and GDI clips
    // them away exactly as it clips the columns of a picture carried off the side.
    let frame = target.frame;
    let top = origin_y as i32 + frame.1;
    if frame.2 <= 0 || frame.3 <= 0 || out_row_bytes == 0 {
        return false;
    }

    BAND_SOURCE.with(|cell| {
        let mut held = cell.borrow_mut();
        let size = (source_width.max(1), source_height.max(1));
        if held.as_ref().map(|surface| (surface.width, surface.height)) != Some(size) {
            *held = DibSurface::create(size.0, size.1);
        }

        let Some(source_surface) = held.as_ref() else {
            return false;
        };

        // The frame's own pixels, into the surface GDI scales out of. It is the frame that is
        // copied rather than the window, and the surface is kept between repaints: what a still
        // picture costs a drag is this copy and the stretch, per pointer move.
        let source_bits = unsafe {
            std::slice::from_raw_parts_mut(source_surface.bits(), row_bytes * size.1 as usize)
        };
        source_bits[..wanted].copy_from_slice(&bgra[..wanted]);

        unsafe {
            // A stretch's brush origin is the DC's, and a DC that keeps the one it had is a
            // halftone pattern placed by whatever drew on it last.
            let _ = SetStretchBltMode(dc, HALFTONE);
            let _ = SetBrushOrgEx(dc, 0, 0, None);
            let _ = StretchBlt(
                dc,
                frame.0,
                top,
                frame.2,
                frame.3,
                source_surface.dc,
                0,
                0,
                source_width as i32,
                source_height as i32,
                SRCCOPY,
            );
            // GDI batches its calls, and what reads this surface is not a GDI call: the bits are
            // handed to `UpdateLayeredWindow` by the caller, so anything still in the batch is
            // work nobody waits for.
            let _ = GdiFlush();
        }

        // A `BI_RGB` surface carries no alpha as far as GDI is concerned, and what a layered
        // window is drawn from is premultiplied coverage: an opaque frame's is 255 everywhere.
        //
        // The coverage handed back is the frame's own, and not the band's: the rows and columns a
        // picture smaller than its band has been carried off belong to whatever the caller put
        // there, and a stretch that sealed them would be sealing somebody else's backdrop (see
        // `fill_band_backdrop`).
        //
        // The columns are cut to the surface before the range is built rather than after it, so that a
        // frame carried wholly off one side is a frame with no columns in the band at all: a sum that
        // came out negative would be a range's own end read as an enormous unsigned one, which is the
        // whole of the band sealed opaque with nothing drawn on it.
        //
        // And the rows are cut to the band, for the same reason and one more: the frame is what the
        // picture is *drawn* at and not the band, so a picture carried a long way off is a frame whose
        // rows are mostly not on screen at all, and walking over the whole of one to seal it would be
        // a paint that takes as long as the zoom is deep.
        let first = frame.0.clamp(0, out_width as i32) as usize;
        let last = (frame.0 + frame.2).clamp(0, out_width as i32) as usize;
        let first_row = top.max(origin_y as i32).max(0);
        let last_row = (top + frame.3)
            .min(origin_y as i32 + band_height as i32)
            .max(first_row);
        let out = &mut *target.out;
        for row in first_row as usize..last_row as usize {
            let start = row * out_row_bytes;
            let Some(destination) = out.get_mut(start..start + out_row_bytes) else {
                break;
            };

            for pixel in destination.as_chunks_mut::<4>().0[first..last].iter_mut() {
                pixel[3] = 255;
            }
        }

        true
    })
}

/// Fill the band a parked player's window has left, answering whether a picture filled it.
///
/// **This is the whole of what a drag of a pinned film is drawn from, and both halves of it are
/// load-bearing.** The opaque fill is there whatever else happens: the band is transparent
/// everywhere a film is playing, and a transparent band with nothing behind it is a hole in the
/// desktop shaped like a video, so a picture that could not be taken has to leave black rather
/// than leave a hole. And the picture is the *last frame the player had on screen*, scaled to the
/// box the window is being dragged to by the same road a resized picture is scaled by — so a band
/// grown shows the film larger and a band shrunk shows it smaller, and a hand is choosing a size
/// against the film rather than against a rectangle (see `compose_media_into_band`).
///
/// A frame that is no longer held is not a reason to paint black twice over: the flat fill is the
/// fallback and the frame goes on top of it, so a frame of a size that does not cover the whole
/// band — or a road that declines to draw at all — leaves a band that is still opaque.
pub(super) fn compose_parked_band(
    held: Option<(&[u8], (u32, u32))>,
    width: u32,
    target: BandTarget<'_>,
) -> bool {
    fill_band_opaque(&mut *target.out, width, target.origin_y, target.height);

    let Some((pixels, (frame_width, frame_height))) = held else {
        return false;
    };

    compose_media_into_band(
        pixels,
        frame_width,
        frame_height,
        (width, target.height),
        TransparentBackground::Black,
        true,
        target,
    );
    true
}

/// `compose_preview_pixels_into` for a frame that is not the whole surface: the frame is
/// composed into the rows the media band of a pinned window occupies, which is what makes a
/// window larger than its frame possible.
///
/// The frame's own rows are what the checkerboard is placed by, so a picture keeps the same
/// squares it had at its own size.
pub(super) fn compose_preview_pixels_into_band(
    bgra: &[u8],
    width: u32,
    height: u32,
    background: TransparentBackground,
    opaque: bool,
    target: BandTarget<'_>,
) {
    let BandTarget {
        out,
        width: out_width,
        origin_y,
        height: band_height,
        ..
    } = target;

    if width == 0 || height == 0 || out_width == 0 || band_height == 0 {
        return;
    }

    let row_bytes = width as usize * 4;
    let out_row_bytes = out_width as usize * 4;
    let offset_x = out_width.saturating_sub(width) as usize / 2;
    let start_x = offset_x * 4;
    let rows = (height as usize).min(band_height as usize);

    for (row, src_row) in bgra.chunks_exact(row_bytes).take(rows).enumerate() {
        let destination_row = origin_y as usize + row;
        let start = destination_row * out_row_bytes + start_x;
        let end = start + row_bytes;
        if end > out.len() {
            break;
        }

        let destination = &mut out[start..end];
        if opaque && src_row.len() == row_bytes {
            destination.copy_from_slice(src_row);
            continue;
        }

        compose_preview_row(src_row, destination, background, 0, row as u32);
    }
}

/// Sample a frame into a band of another size: what a pinned window draws while an edge of it is
/// being dragged.
///
/// It is the scaling an image viewer does with the picture it is already holding — a destination
/// pixel is the four source pixels around the point that maps back from it, weighted by how far
/// into the square between them it falls — rather than the media laid out again for the box, which
/// is a decode and not something a hand can be made to wait for on every pixel of a drag. The
/// media is laid out properly once the drag is over, so what this owes the eye is a picture that
/// fills the window while it moves rather than a sharper one.
///
/// What goes in is the straight-alpha BGRA a frame is composed in, and what comes out is what the
/// layered surface wants: the sample composited over the backdrop at the place it lands, with the
/// squares of a checkerboard placed by where a pixel is in the *band* rather than by where it came
/// from, so a picture being resized keeps the backdrop it had instead of dragging it around.
pub(super) fn resample_into_band(
    bgra: &[u8],
    source: (u32, u32),
    band: (u32, u32),
    background: TransparentBackground,
    opaque: bool,
    target: BandTarget<'_>,
) {
    let BandTarget {
        out,
        width: out_width,
        origin_y,
        height: band_height,
        frame,
        ..
    } = target;

    if out_width == 0 || band_height == 0 || frame.2 <= 0 || frame.3 <= 0 {
        return;
    }

    let (source_width, source_height) = source;
    let row_bytes = source_width as usize * 4;
    let expected = row_bytes * source_height as usize;
    if row_bytes == 0 || bgra.len() < expected {
        return;
    }

    let out_row_bytes = out_width as usize * 4;
    // Where a destination pixel's center maps back to in the source, in 16.16 fixed point: the
    // mapping is two multiplies a row and two a pixel, and there is no division in the loop.
    //
    // The two are the *frame's* own widths rather than the band's, because the picture lands in the
    // frame and not in the band: a picture drawn larger than its window maps a frame bigger than the
    // band back to the whole of the source, and the rows and columns outside that frame are not its
    // to write (see `BandTarget::frame`).
    let step_x = ((source_width as u64) << 16) / frame.2 as u64;
    let step_y = ((source_height as u64) << 16) / frame.3 as u64;
    let rows = (band.1 as usize).min(band_height as usize);

    for y in 0..rows {
        let destination_row = origin_y as usize + y;
        let start = destination_row * out_row_bytes;
        let Some(destination) = out.get_mut(start..start + out_row_bytes) else {
            break;
        };

        // A row the picture does not land on is left exactly as the caller left it, which is the
        // backdrop a picture carried off its band has left behind (see `fill_band_backdrop`).
        let offset_y = y as i32 - frame.1;
        if offset_y < 0 || offset_y >= frame.3 {
            continue;
        }

        let sy = (offset_y as u64 * step_y + step_y / 2).saturating_sub(0x8000);
        let top_row = ((sy >> 16) as usize).min(source_height as usize - 1);
        let bottom_row = (top_row + 1).min(source_height as usize - 1);
        let wy = ((sy & 0xFFFF) >> 8) as u32;
        let top = &bgra[top_row * row_bytes..top_row * row_bytes + row_bytes];
        let bottom = &bgra[bottom_row * row_bytes..bottom_row * row_bytes + row_bytes];

        for (x, pixel) in destination.as_chunks_mut::<4>().0.iter_mut().enumerate() {
            let offset_x = x as i32 - frame.0;
            if offset_x < 0 || offset_x >= frame.2 {
                continue;
            }

            let sx = (offset_x as u64 * step_x + step_x / 2).saturating_sub(0x8000);
            let left_column = ((sx >> 16) as usize).min(source_width as usize - 1);
            let right_column = (left_column + 1).min(source_width as usize - 1);
            let wx = ((sx & 0xFFFF) >> 8) as u32;

            let [b, g, r, a] =
                sample_bilinear(top, bottom, left_column * 4, right_column * 4, wx, wy);

            if opaque {
                pixel.copy_from_slice(&[b, g, r, a]);
                continue;
            }

            match background {
                TransparentBackground::Transparent => {
                    let alpha = a as u32;
                    pixel[0] = ((b as u32 * alpha + 127) / 255) as u8;
                    pixel[1] = ((g as u32 * alpha + 127) / 255) as u8;
                    pixel[2] = ((r as u32 * alpha + 127) / 255) as u8;
                    pixel[3] = a;
                }
                TransparentBackground::Black => {
                    blend_pixel_over(&[b, g, r, a], pixel, 0, 0, 0);
                }
                TransparentBackground::White => {
                    blend_pixel_over(&[b, g, r, a], pixel, 255, 255, 255);
                }
                TransparentBackground::Checkerboard => {
                    let (square_r, square_g, square_b) = checkerboard_color(x as u32, y as u32);
                    blend_pixel_over(
                        &[b, g, r, a],
                        pixel,
                        square_b as u32,
                        square_g as u32,
                        square_r as u32,
                    );
                }
            }
        }
    }
}

/// One sample of a frame at a point between four of its pixels: the four are weighted by how far
/// the point is from each of them, in 256ths, and the sum of the four weights is 65536.
#[inline]
pub(super) fn sample_bilinear(
    top: &[u8],
    bottom: &[u8],
    left: usize,
    right: usize,
    wx: u32,
    wy: u32,
) -> [u8; 4] {
    let (left_near, right_near) = (256 - wx, wx);
    let (top_near, bottom_near) = (256 - wy, wy);
    let weights = [
        left_near * top_near,
        right_near * top_near,
        left_near * bottom_near,
        right_near * bottom_near,
    ];
    let corners = [left, right, left, right];

    let mut sample = [0u8; 4];
    for channel in 0..4 {
        let mut value = 0u32;
        for (index, weight) in weights.iter().enumerate() {
            let row = if index < 2 { top } else { bottom };
            value += row[corners[index] + channel] as u32 * weight;
        }

        sample[channel] = ((value + 32_768) >> 16) as u8;
    }

    sample
}

/// Draw the name of a pinned window's caption button over its media, if it has one to say.
///
/// It is a panel of its own hanging below the caption rather than a part of the caption, and it
/// is the one piece of chrome that is not a strip of its own: it needs a name measured through
/// GDI, a surface with a device context to measure and draw it on, and a place in the window
/// that is neither the strip nor the bar (see `pin_chrome::tooltip_layout`).
///
/// Nothing about it is asked of the window while it is up, and nothing is asked of it beyond
/// the one string the pin already holds: a name is a word or two of a program's own name, and
/// this is the whole of the drawing.
///
/// # Safety
///
/// `out` must be the window's own premultiplied surface, of at least `width * height` pixels,
/// and `width`/`height` that surface's own size. Nothing else is dereferenced.
pub(super) unsafe fn paint_pin_tooltip(
    out: &mut [u8],
    width: i32,
    height: i32,
    paint: &PinnedPaint,
    caption_height: i32,
) {
    let (Some(tooltip), Some(palette)) =
        (paint.tooltip.as_ref(), pin_chrome::ChromePalette::current())
    else {
        return;
    };

    // The buttons the caption was drawn with, so the name hangs off the button it belongs to
    // rather than off a box worked out a second time and a half a step out of step with it.
    let Some(anchor) =
        pin_chrome::button_boxes(width, caption_height, paint.dpi, paint.maximizable)
            .into_iter()
            .find(|button| button.kind == tooltip.kind)
            .map(|button| button.rect)
    else {
        return;
    };

    // A surface of the window's own size, because the name is measured and drawn through GDI
    // and GDI needs a device context — the layered surface has none of its own, and the
    // caption's is only as tall as the bar. It is a DIB section, so the text lands in memory
    // that is carried across rather than on a device, and it is kept between paints and rebuilt
    // only when the window changes size (see `TOOLTIP_SURFACE`).
    let wanted = (width.max(1) as u32, height.max(1) as u32);
    TOOLTIP_SURFACE.with(|cell| {
        let mut surface = cell.borrow_mut();
        if surface
            .as_ref()
            .map(|surface| (surface.width, surface.height))
            != Some(wanted)
        {
            *surface = DibSurface::create(wanted.0, wanted.1);
        }
        let Some(surface) = surface.as_ref() else {
            return;
        };

        // The surface is blanked before it is used rather than left as the last name left it:
        // what is carried off it afterwards is decided by the alpha byte GDI leaves behind, so a
        // run left opaque from the last paint is a run drawn twice.
        let blanked = unsafe {
            std::slice::from_raw_parts_mut(
                surface.bits(),
                surface.width as usize * surface.height as usize * 4,
            )
        };
        blanked.fill(0);

        let text_width = pin_chrome::measure_caption_text(surface, &tooltip.text, paint.dpi);
        let Some(panel) = pin_chrome::tooltip_layout(
            width,
            height,
            caption_height,
            anchor,
            text_width,
            paint.dpi,
        ) else {
            return;
        };

        pin_chrome::paint_tooltip(
            out,
            width,
            &palette,
            panel,
            pin_chrome::TooltipText {
                text: &tooltip.text,
                width: text_width,
                surface,
            },
            paint.dpi as f32 / 96.0,
        );
    });
}

/// A window that is not on screen is moved before the spinner is installed, so a
/// `WM_DPICHANGED` reset from crossing displays cannot discard it, and the spinner
/// is painted before the window is revealed, so the previous preview cannot flash
/// at the new place. It is also what moves a spinner whose box has changed size
/// while it was up: the frame is drawn at the size of the box it goes into.
pub(super) unsafe fn show_loading_spinner(hwnd: HWND, pl: &PendingLoad) {
    // A wait for a hover the pointer has left is not put up, and nothing of the
    // spinner is built for it — no media, no frame, no window — so a load the hook
    // has already dismissed cannot blink a box on screen in the meantime. The
    // guard is held for the rest of the function, which is what makes this check
    // and the window it writes one step (see `HIDDEN_EPOCH`), and the pointer is
    // asked the same question the reveal is: a wait belongs at the hand that is
    // still on the file, and not at one that has crossed to another item
    // (see `HOVER_POINTER_BOX`).
    let hidden = HIDDEN_EPOCH.lock().ok();
    if !hover_still_wanted(&hidden, pl) || !pointer_on_the_hovered_item() {
        return;
    }

    let (x, y) = pl.spinner_pos;
    let side = pl.spinner_side as i32;

    // A window that is already on screen is not moved ahead of its frame, for the
    // reason the page's install gives: what a layered window shows between one
    // paint and the next is the surface it already has, stretched into whatever
    // box the window has, so a spinner whose box changed under it would be drawn
    // as a bar until the new frame lands. `UpdateLayeredWindow` applies the place
    // and the size with the frame, which is where the move below happens instead.
    let visible = IsWindowVisible(hwnd).as_bool();
    if !visible {
        let _ = MoveWindow(hwnd, x, y, side, side, false);
    }

    let loading = create_loading_media(pl.spinner_side, pl.spinner_side);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(loading);
    }

    if visible {
        render_layered_preview_at(hwnd, x, y);
    } else {
        render_layered_preview(hwnd);
    }
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

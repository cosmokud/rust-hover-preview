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
/// windows stand in the band, and a video FFmpeg's player plays, whose window stands in
/// the whole of it — a pin whose chrome is drawn over its media is media the whole of the
/// window down (see `pinned_band_rows`). What the band is then is nothing at all — alpha
/// zero, which is what lets the window underneath be seen through it and, more to the
/// point, be *clicked*: hit testing of a layered window is answered by the shape of its
/// pixels, so a band of transparent ones is a band the player keeps for itself.
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
                    compose_media_into_band(
                        media.current_pixels(),
                        media.current_width(),
                        media.current_height(),
                        (band_width, band_height),
                        preview_background(media.media_type),
                        media.current_frame_is_opaque(),
                        BandTarget {
                            out,
                            dc: mem_dc,
                            width: band_width,
                            origin_y: band_top.max(0) as u32,
                            height: band_height,
                        },
                    );
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

    // The card's menu, over the media below the cell it came out of:
    // drawn after everything but the caption's own name, because it is
    // over everything but that (see `pinned_menu_geometry`).
    if let Some(menu) = paint.menu.as_ref() {
        if let Some(palette) = pin_chrome::ChromePalette::current() {
            // A surface of the window's own size, because the rows'
            // labels are measured and drawn through GDI and GDI needs a
            // device context of its own — the same surface the caption's
            // tooltip is written on, kept between paints for the same
            // reason and blanked before each use for the same one (see
            // `TOOLTIP_SURFACE`).
            let wanted = (width.max(1) as u32, height.max(1) as u32);
            MENU_SURFACE.with(|cell| {
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

                // The surface is blanked before it is used rather than
                // left as the last menu left it, for the reason the
                // tooltip's is: what is carried off it afterwards is
                // decided by the alpha byte GDI leaves behind.
                let blanked = unsafe {
                    std::slice::from_raw_parts_mut(
                        surface.bits(),
                        surface.width as usize * surface.height as usize * 4,
                    )
                };
                blanked.fill(0);

                pin_chrome::paint_menu_popup(
                    out,
                    width,
                    &palette,
                    &menu.popup,
                    &menu.rows,
                    surface,
                    paint.dpi as f32 / 96.0,
                );
            });
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
    /// The card's own menu, or nothing while its panel is closed: the panel
    /// hung from the cell the mark is drawn in and the rows it holds, asked
    /// of the pin for the reason the volume popup's panel is (see
    /// `pinned_menu_geometry`).
    pub(super) menu: Option<PinMenuPaint>,
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
            overlay: pin.overlay,
            // A kind with no caption of its own is asked for none whatever its chrome is doing:
            // there is no row above the media for one to be drawn in, and a strip painted over the
            // first row of a card is a title bar with the controls underneath it (see
            // `pinned_caption_height`).
            caption: pin.caption > 0 && pin.chrome.caption,
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
            menu: None,
            transport_live: pin.transport_live,
            parked: pin.parked,
        }
    };

    // The panel, asked for after the pin's lock is let go of rather than read out of it, because
    // where it goes is a question about the geometry of a window this one only holds a lock for
    // while it is up — and because a panel placed from the pin's own numbers is the same panel
    // every press and every drag is answered against (see `pinned_volume_geometry`).
    paint.volume_popup = pinned_volume_geometry();

    // The card's menu, asked for after the pin's lock is let go of for
    // the reason the volume popup's panel is: where it goes is a
    // question about the geometry of a window this one only holds a
    // lock for while it is up, and the panel drawn is the panel a press
    // is answered against (see `pinned_menu_geometry`).
    paint.menu = pinned_menu_geometry();

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

    if (width, height) == band {
        compose_preview_pixels_into_band(bgra, width, height, background, opaque, target);
    } else if !opaque || !stretch_into_band(bgra, (width, height), band, &mut target) {
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
    band: (u32, u32),
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
    let rows = (band.1 as usize).min(band_height as usize);
    if rows == 0 || out_row_bytes == 0 {
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
                0,
                origin_y as i32,
                out_width as i32,
                rows as i32,
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
        let out = &mut *target.out;
        for row in 0..rows {
            let start = (origin_y as usize + row) * out_row_bytes;
            let Some(destination) = out.get_mut(start..start + out_row_bytes) else {
                break;
            };

            for pixel in destination.as_chunks_mut::<4>().0 {
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
        ..
    } = target;

    if out_width == 0 || band_height == 0 {
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
    let step_x = ((source_width as u64) << 16) / band.0 as u64;
    let step_y = ((source_height as u64) << 16) / band.1 as u64;
    let rows = (band.1 as usize).min(band_height as usize);

    for y in 0..rows {
        let destination_row = origin_y as usize + y;
        let start = destination_row * out_row_bytes;
        let Some(destination) = out.get_mut(start..start + out_row_bytes) else {
            break;
        };

        let sy = (y as u64 * step_y + step_y / 2).saturating_sub(0x8000);
        let top_row = ((sy >> 16) as usize).min(source_height as usize - 1);
        let bottom_row = (top_row + 1).min(source_height as usize - 1);
        let wy = ((sy & 0xFFFF) >> 8) as u32;
        let top = &bgra[top_row * row_bytes..top_row * row_bytes + row_bytes];
        let bottom = &bgra[bottom_row * row_bytes..bottom_row * row_bytes + row_bytes];

        for (x, pixel) in destination.as_chunks_mut::<4>().0.iter_mut().enumerate() {
            let sx = (x as u64 * step_x + step_x / 2).saturating_sub(0x8000);
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

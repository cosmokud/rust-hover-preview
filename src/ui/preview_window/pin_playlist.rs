//! The playlist a pinned sound's card grows under it: the rows of the pin's own walk, the press
//! that picks one, the drag that sets the band's height, the notch the wheel steps it by, and the
//! name the hand sets moving.
//!
//! The band is a session thing rather than a pin's: it is opened from the card's own menu, it is
//! remembered as it was left — up or down, and how tall — so the next pin comes up with it, and
//! nothing of it survives the run. What is the pin's is only the rows: the walk it is standing on,
//! read afresh each tick and rebuilt only where that walk has moved, because building them reads
//! every file's own facts and a paint may not pay for that (see `refresh_the_rows`).
//!
//! The height is a number of pixels rather than a share of anything, remembered and clamped when it
//! is set: the window's own box is asked for on paths that a low-level hook walks — a wheel notch
//! landing on the band reads it — and a work area read there would be a desktop walk in a callback
//! that may do nothing of the kind.

use super::*;
use std::sync::atomic::{AtomicBool, AtomicI32};

/// Whether the band is up, for the session: what the card's menu row turns over and what the next
/// pin comes up with (see `pin_playlist_band`).
static OPEN: AtomicBool = AtomicBool::new(false);

/// How tall the band is, in the display's own pixels, as the hand last left it — clamped where it
/// was set rather than where it is read (see the module documentation).
static HEIGHT: AtomicI32 = AtomicI32::new(0);

/// The most the band may ever be: nine tenths of the work area of the display it was opened on,
/// kept beside the height so that a band remembered from a larger display is still clamped by the
/// one it is standing on without walking the desktop to find out.
static CEILING: AtomicI32 = AtomicI32::new(0);

/// The band's own state: the rows it holds, where it has scrolled to, which row the hand is on and
/// which it has picked, and the name a row is marooned with.
static THE_PLAYLIST: Lazy<Mutex<PinPlaylist>> = Lazy::new(|| Mutex::new(PinPlaylist::default()));

/// A name too long for its row, scrolled across it for as long as the hand is on the row: how far
/// it has moved, which way it is going and when the hold at either end runs out — the card's own
/// marquee, in a row rather than in a header (see `audio_preview::NameScroll`).
#[derive(Debug, Clone, Copy)]
struct Marquee {
    row: usize,
    offset: i32,
    /// Pixels it may move before the end of the name is in sight.
    travel: i32,
    /// Pixels one advance of the font is: the unit the speed is counted in, so a band at any scale
    /// scrolls the same characters a second.
    advance: i32,
    direction: i32,
    hold_until: Instant,
    /// When this marquee was last advanced, which is the cadence the next step is counted in.
    last: Instant,
}

#[derive(Default)]
struct PinPlaylist {
    /// The rows in the order they are drawn, and the files they are the rows of — the two are one
    /// list, so a press answered in rows is answered in files.
    rows: Vec<pin_chrome::PlaylistRow>,
    order: Vec<PathBuf>,
    /// The file the pin was showing when the rows were last reconciled, and the row it is on: kept
    /// so that a walk step moves the stripe without rebuilding every row's facts.
    path: Option<PathBuf>,
    showing: Option<usize>,
    /// Whole rows, never a row of a row: the band scrolls by rows because a row that is half shown
    /// is a row whose two lines are half readable.
    scroll: usize,
    hovered: Option<usize>,
    picked: Option<usize>,
    /// Whether the press the pointer is still holding is the band's, which is what the release that
    /// follows it is answered by.
    ///
    /// A press on a band and a press on the card are the same window's, and only the one that took
    /// the press may give the release up: an arm here that answered every release inside the band
    /// would swallow the release of a drag the card began and leave the window carrying the pointer
    /// for the rest of the run.
    pressed: bool,
    /// The height's drag in flight: the window row the pointer was at when it began, and the height
    /// it began at.
    resizing: Option<(i32, i32)>,
    marquee: Option<Marquee>,
}

impl PinPlaylist {
    /// The rows the walk has moved to, and the row the pin is on: rebuilt only where the order is
    /// not the one in hand, since every row of it is a file read.
    fn hold_the_walk(&mut self, order: Vec<PathBuf>, path: PathBuf) {
        if self.order != order {
            self.rows = order
                .iter()
                .map(|path| pin_chrome::playlist_row(path))
                .collect();
            self.order = order;
            // A list that is not the one a pick was made in is not a list a pick means anything in.
            self.scroll = 0;
            self.hovered = None;
            self.picked = None;
            self.pressed = false;
            self.marquee = None;
        }

        self.showing = self.order.iter().position(|row| *row == path);
        self.path = Some(path);
    }

    /// The marquee the row the hand is on needs, set up when the row becomes that row rather than
    /// on every frame the name moves: what it costs is a measure of the name at the scale it is
    /// drawn at.
    fn set_the_marquee(&mut self, layout: &pin_chrome::PlaylistLayout, scale: f32) {
        let row = self.hovered.or(self.picked);
        if self.marquee.map(|marquee| marquee.row) == row {
            return;
        }

        let room = layout.name_right - layout.text_left;
        let now = Instant::now();
        self.marquee = row
            .filter(|row| self.rows.get(*row).is_some())
            .and_then(|row| {
                let (travel, advance) = pin_chrome::name_travel(&self.rows[row].name, room, scale)?;

                Some(Marquee {
                    row,
                    offset: 0,
                    travel,
                    advance,
                    direction: 1,
                    hold_until: now + audio_preview::NAME_HOLD,
                    last: now,
                })
            });
    }

    /// How far the marqueed name has been moved, which is nothing at all for a band with no marquee.
    fn name_offset(&self) -> i32 {
        self.marquee.map_or(0, |marquee| marquee.offset)
    }
}

/// The band this pin carries, as the row it begins at and how tall it is, or nothing where the pin
/// is showing no sound's card or the band is down.
///
/// It is asked of the pin rather than of the state alone because a band is only ever a sound's
/// card's: every other kind's window is measured without one, and a card's is measured with one
/// exactly when this answers (see `PinnedPreview::window_box`).
pub(super) fn pin_playlist_band(pin: &PinnedPreview) -> Option<(i32, i32)> {
    let height = playlist_band_height(pin);
    if height <= 0 {
        return None;
    }

    let (_, _, _, bottom) = pinned_window_box_of(
        pin.content,
        pin.dpi,
        pin.transport_bar,
        pin.overlay,
        pin.caption,
    );

    Some((bottom, height))
}

/// The rows the band takes off the pin's media band, and none at all where the pin carries no band
/// (see `pinned_band_rows`).
pub(super) fn playlist_rows(pin: &PinnedPreview) -> i32 {
    playlist_band_height(pin).max(0)
}

/// How tall the band this pin carries is, in the display's own pixels — nothing where it is down,
/// or where the pin is showing anything but a sound's card.
pub(super) fn playlist_band_height(pin: &PinnedPreview) -> i32 {
    if !OPEN.load(Ordering::Acquire) || !pin_shows_an_audio_card(pin) {
        return 0;
    }

    HEIGHT.load(Ordering::Acquire).max(0)
}

/// Whether the band is up at all, which is the one question a tick with nothing pinned pays.
pub(super) fn the_band_is_open() -> bool {
    OPEN.load(Ordering::Acquire)
}

/// Turn the band over, from the card's own menu row, and settle the box that comes out of it: a
/// band opened for the first time in a run takes the share of the display it stands on that a band
/// opens at, and every opening after it takes the height the hand last left.
///
/// The window is not moved here. Nothing of the pin's own box moves when a band opens — its top and
/// its left are the card's — and the paint that follows the press is what gives the window its new
/// height (see `render_layered_preview`).
pub(super) fn toggle(pin: &PinnedPreview) {
    let open = !OPEN.load(Ordering::Acquire);
    OPEN.store(open, Ordering::Release);

    if !open {
        return;
    }

    let bounds = work_area_at(
        (pin.content.0 + pin.content.2) / 2,
        (pin.content.1 + pin.content.3) / 2,
    );
    let room = (bounds.bottom - bounds.top).max(1);
    let scale = pin.dpi as f32 / 96.0;
    let floor = pin_chrome::playlist_row_height(scale) * TWO_ROWS;
    let ceiling = (room as f32 * MOST_OF_THE_DISPLAY) as i32;

    let chosen = match HEIGHT.load(Ordering::Acquire) {
        remembered if remembered > 0 => remembered,
        _ => (room as f32 * DEFAULT_SHARE_OF_THE_DISPLAY) as i32,
    };

    HEIGHT.store(chosen.clamp(floor, ceiling.max(floor)), Ordering::Release);
    CEILING.store(ceiling.max(floor), Ordering::Release);
}

/// What the band's own geometry is asked of: the window row it begins at, how tall it is, the
/// window's own width, and the scale it is drawn at — or nothing where this pin carries no band.
///
/// The pin's lock is taken and let go of here rather than held across a caller's work, because
/// every caller below then takes the playlist's own lock and the two are never asked together.
fn the_band_of_the_pin() -> Option<(i32, i32, i32, f32)> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    let (top, height) = pin_playlist_band(pin)?;
    let (width, _) = pin.window_size();

    Some((top, height, width, pin.dpi as f32 / 96.0))
}

/// The tick's own work for the band, answering whether the window is owed a paint for it: the rows
/// rebuilt where the pin's walk has moved, the notch the wheel left here, and the name a row is
/// marooned with. One answer rather than three, because a tick owes one paint.
pub(super) fn the_tick_is_owed_a_repaint() -> bool {
    if !the_band_is_open() {
        // A notch counted for a band that has since gone away is a notch nothing is going to read.
        let _ = wheel_input::take_pin_playlist_delta();
        return false;
    }

    let mut owed = refresh_the_rows();

    let notches = wheel_input::take_pin_playlist_delta();
    if notches != 0 {
        owed |= step_the_notches(notches);
    }

    owed | advance_the_marquee()
}

/// The rows the band draws, rebuilt where the pin's walk is not the one they were built from.
///
/// The comparison is made against the walk's own list rather than against a copy of it: a folder of
/// a thousand files is a thousand paths, and the tick asks this sixty times a second.
///
/// **It must not be called with the pin's own lock held**, which is why the menu row that opens the
/// band asks for it after the pin has been let go of rather than from inside `toggle`: what it
/// reads first is the pin, and a pin asked twice on one thread blocks for ever.
pub(super) fn refresh_the_rows() -> bool {
    let (order, path) = {
        let Some(pinned) = pin_state() else {
            return false;
        };
        let Some(pin) = pinned.pin() else {
            return false;
        };

        // A pin that holds no walk is a pin of one file, which is a band of one row.
        let order: &[PathBuf] = match pinned.walk() {
            Some(walk) if !walk.order.is_empty() => &walk.order,
            _ => std::slice::from_ref(&pin.path),
        };
        let Ok(playlist) = THE_PLAYLIST.lock() else {
            return false;
        };
        if playlist.order == order && playlist.path.as_deref() == Some(pin.path.as_path()) {
            return false;
        }

        (order.to_vec(), pin.path.clone())
    };

    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };
    playlist.hold_the_walk(order, path);
    true
}

/// The wheel's notches, stepped as whole rows and clamped to what the band can scroll to:
/// answering whether the band moved.
fn step_the_notches(notches: i32) -> bool {
    let Some((_, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };

    let layout = pin_chrome::playlist_layout(width, height, scale, playlist.rows.len());
    let rows = (notches * pin_chrome::PLAYLIST_WHEEL_ROWS) / WHEEL_DELTA;
    let max = pin_chrome::playlist_scroll_max(playlist.rows.len(), layout.visible.max(0) as usize);
    // A notch up is the list come back toward its first row, which is the wheel's own reading of
    // the two everywhere else in the app.
    let scrolled = (playlist.scroll as i32 - rows).clamp(0, max as i32) as usize;

    let moved = scrolled != playlist.scroll;
    playlist.scroll = scrolled;
    moved
}

/// Move the marqueed name on by one tick, answering whether it moved: it rests at either end for
/// `NAME_HOLD` and ping-pongs between them, which is the card's own rule for the same name (see
/// `audio_preview::NameScroll`).
fn advance_the_marquee() -> bool {
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };
    let Some(mut marquee) = playlist.marquee else {
        return false;
    };

    let now = Instant::now();
    let cadence = now.saturating_duration_since(marquee.last);
    if now < marquee.hold_until || marquee.travel <= 0 {
        marquee.last = now;
        playlist.marquee = Some(marquee);
        return false;
    }

    let step = ((marquee.advance as u128 * cadence.as_millis()) / audio_preview::NAME_ADVANCE_MS)
        .max(1) as i32;
    let next = marquee.offset + step * marquee.direction;

    if next >= marquee.travel {
        marquee.offset = marquee.travel;
        marquee.direction = -1;
        marquee.hold_until = now + audio_preview::NAME_HOLD;
    } else if next <= 0 {
        marquee.offset = 0;
        marquee.direction = 1;
        marquee.hold_until = now + audio_preview::NAME_HOLD;
    } else {
        marquee.offset = next;
    }
    marquee.last = now;

    let moved = playlist.marquee.map(|held| held.offset) != Some(marquee.offset);
    playlist.marquee = Some(marquee);
    moved
}

/// A pointer over the band, answering whether the band changed for it: which row the hand is on is
/// the marquee's own answer, and a row a hand has come to is a name that starts moving.
pub(super) unsafe fn hover(y: i32) -> bool {
    let Some((top, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };

    let layout = pin_chrome::playlist_layout(width, height, scale, playlist.rows.len());
    let hovered =
        pin_chrome::playlist_row_at(&layout, playlist.scroll, playlist.rows.len(), y - top);

    let changed = playlist.hovered != hovered;
    playlist.hovered = hovered;
    playlist.set_the_marquee(&layout, scale);

    changed
}

/// A press on the band, answering whether the press was the band's: the height's own handle at its
/// bottom edge takes the pointer and drags the band, and a press on a row picks it — never carries
/// the window, which is what every other press on a sound's card does.
///
/// A press that is the second of a pair shows the file that row is, the way a press on a row of the
/// listing behind the pin shows one, and without the configuration gate the hook's own picks are
/// behind: the hand is already inside this app's own window and has named the file itself.
pub(super) unsafe fn press(hwnd: HWND, y: i32, double_click: bool) -> bool {
    let Some((top, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };
    let local = y - top;
    if local < 0 || local >= height {
        return false;
    }

    if pin_chrome::playlist_handle_holds(height, scale, local) {
        let _ = SetCapture(hwnd);
        if let Ok(mut playlist) = THE_PLAYLIST.lock() {
            playlist.pressed = true;
            playlist.resizing = Some((y, HEIGHT.load(Ordering::Acquire)));
        }
        return true;
    }

    let opened = {
        let Ok(mut playlist) = THE_PLAYLIST.lock() else {
            return true;
        };
        let layout = pin_chrome::playlist_layout(width, height, scale, playlist.rows.len());
        let Some(index) =
            pin_chrome::playlist_row_at(&layout, playlist.scroll, playlist.rows.len(), local)
        else {
            // The pad between two rows and the room below the last of them are the band's own: a
            // press there picks nothing and is still not a press the card answers.
            return true;
        };

        playlist.pressed = true;
        playlist.picked = Some(index);
        playlist.set_the_marquee(&layout, scale);

        double_click
            .then(|| playlist.order.get(index).cloned())
            .flatten()
    };

    if let Some(path) = opened.as_deref() {
        show_pinned_file(path);
    }
    render_layered_preview(hwnd);

    true
}

/// The pointer's own move with the height's drag in flight, answering whether the move was the
/// band's.
///
/// The window's top-left never moves for a band, so a row of the window is the measurement the
/// height is dragged by: the pointer's travel down the screen is the height's, and what a move
/// writes is where the band's bottom edge is.
pub(super) unsafe fn drag(hwnd: HWND, y: i32) -> bool {
    let Some((start, start_height)) = THE_PLAYLIST
        .lock()
        .ok()
        .and_then(|playlist| playlist.resizing)
    else {
        return false;
    };

    let scale = the_band_of_the_pin().map_or(1.0, |(_, _, _, scale)| scale);
    let floor = pin_chrome::playlist_row_height(scale) * TWO_ROWS;
    let ceiling = CEILING.load(Ordering::Acquire).max(floor);
    let height = (start_height + (y - start)).clamp(floor, ceiling);

    if height != HEIGHT.load(Ordering::Acquire) {
        HEIGHT.store(height, Ordering::Release);
        render_layered_preview(hwnd);
    }

    true
}

/// The release of a press the band took, answering whether it was the band's: a drag of the height
/// is over and the pointer given back, and a press on a row is over with nothing left to do.
///
/// The pointer is given back only where a press took it, which is the rule every arm of a release
/// on a pinned window keeps: a release that let go of a capture it never took would run
/// `pin_capture_lost` before the arm that owns one had read what its press armed.
pub(super) unsafe fn release(hwnd: HWND) -> bool {
    let (pressed, dragging) = match THE_PLAYLIST.lock() {
        Ok(mut playlist) => (playlist.pressed, playlist.resizing.take()),
        Err(_) => return false,
    };
    if !pressed {
        return false;
    }

    if dragging.is_some() {
        let _ = ReleaseCapture();
        render_layered_preview(hwnd);
    }

    true
}

/// The capture taken for a drag of the height has gone elsewhere, which is the drag ended: nothing
/// is left holding a pointer that is no longer this window's.
pub(super) fn the_capture_was_lost() {
    if let Ok(mut playlist) = THE_PLAYLIST.lock() {
        playlist.pressed = false;
        playlist.resizing = None;
    }
}

/// The band's own box on the screen, for the wheel hook to answer a notch against — asked without
/// blocking a hook that every mouse message on the desktop is waiting on (see
/// `pin_zoom_keep_alive_try`, whose rules this one follows).
pub(crate) fn pin_playlist_keep_alive_try() -> Option<(i32, i32, i32, i32)> {
    if !the_band_is_open() {
        return None;
    }

    let pinned = pin_window::pin_state_try()?;
    let pin = pinned.pin()?;
    let (top, height) = pin_playlist_band(pin)?;
    let (left, window_top, right, _) = pin.window_box();

    Some((left, window_top + top, right, window_top + top + height))
}

/// The band's pixels, painted into the surface the paint keeps for it and carried onto the
/// window's own rows.
///
/// The band is opaque — the theme's page at full alpha — and that is not a matter of taste: a
/// layered window is hit-tested by the shape of its pixels, so a band left see-through is a band a
/// press falls through to the listing behind the pin, and a row a hand is not on is a window the
/// user cannot click.
pub(super) unsafe fn paint_the_band(
    out: &mut [u8],
    surface: &DibSurface,
    stride: usize,
    width: i32,
    height: i32,
    paint: &PinnedPaint,
) {
    if paint.playlist <= 0 || width <= 0 || paint.playlist > height {
        return;
    }

    let Some(palette) = pin_chrome::ChromePalette::current() else {
        return;
    };

    let band_top = height - paint.playlist;
    let scale = paint.dpi as f32 / 96.0;

    // The playlist's own lock is held for the paint rather than a copy of its rows taken: the band
    // is drawn from the rows themselves, and a band of ten names is not worth ten allocations on
    // every frame of a marquee.
    let Ok(playlist) = THE_PLAYLIST.lock() else {
        return;
    };
    let layout = pin_chrome::playlist_layout(width, paint.playlist, scale, playlist.rows.len());
    let first = playlist.scroll.min(playlist.rows.len());
    let last = (first + layout.visible.max(0) as usize).min(playlist.rows.len());
    let band = pin_chrome::PlaylistBand {
        layout,
        rows: &playlist.rows[first..last],
        first,
        total: playlist.rows.len(),
        showing: playlist.showing,
        hovered: playlist.hovered,
        picked: playlist.picked,
        name_offset: playlist.name_offset(),
    };

    pin_chrome::paint_playlist(surface, &palette, &band, scale);
    drop(playlist);

    // What a band is carried across as is its rows and nothing else: it is painted opaque, so there
    // is nothing under it left to composite through and a copy is the whole of it.
    let source = std::slice::from_raw_parts(
        surface.bits(),
        surface.width as usize * surface.height as usize * 4,
    );
    let rows = (height - band_top).min(surface.height as i32).max(0) as usize;
    let count = width as usize * 4;

    for row in 0..rows {
        let from = row * surface.width as usize * 4;
        let to = (band_top as usize + row) * stride;
        if from + count > source.len() || to + count > out.len() {
            break;
        }
        out[to..to + count].copy_from_slice(&source[from..from + count]);
    }
}

/// The share of the display a band opens at when nothing is remembered, and the most it may ever
/// be: a band is for reading a folder's worth of names, and a band that has taken the whole screen
/// has taken the window.
const DEFAULT_SHARE_OF_THE_DISPLAY: f32 = 0.6;
const MOST_OF_THE_DISPLAY: f32 = 0.9;

/// The shortest a band may be dragged: two rows, because a band of one is a row a hand cannot
/// scroll to the next one of.
const TWO_ROWS: i32 = 2;

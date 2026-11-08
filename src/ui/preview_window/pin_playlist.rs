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
use std::sync::atomic::{AtomicBool, AtomicI32, AtomicU64};

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

/// Whether the band owes the tick a paint the tick has not given it yet: the band turned over, or a
/// row's own facts landing from the prober below. One flag rather than a count, because a tick owes
/// one paint however many things asked for it.
///
/// It is a flag of the band's own rather than an ask of the card's because the band is painted from
/// the pin's window whether or not the sound in it is playing: a card that is held is a card the
/// loop repaints on nothing, and a band opened over one would be a band nothing drew.
static DIRTY: AtomicBool = AtomicBool::new(false);

/// How many files the prober has landed a reading for, which is what tells the tick that a row's
/// facts have changed: the rows are rebuilt when this number moves and never otherwise (see
/// `start_the_prober`).
static PROBED: AtomicU64 = AtomicU64::new(0);

/// Which run of the prober the files of the current order are being read by: bumped when the order
/// moves, so that a thread reading a folder the pin has left stops rather than reading on into a
/// list nobody is looking at.
#[cfg(not(test))]
static PROBE_RUN: AtomicU64 = AtomicU64::new(0);

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

/// The height's drag in flight, as the press that began it left it (see `press` and `drag`).
#[derive(Clone, Copy)]
struct PlaylistResize {
    edge: pin_chrome::PlaylistEdge,
    /// The *screen* row the pointer was at when the drag began, which every move is measured from
    /// rather than the move before it: a drag that accumulated its own rounding would creep, and one
    /// measured in the window's own rows would chase a window the drag is itself moving.
    from: i32,
    /// The band's height when the drag began.
    height: i32,
    /// The card's box when the drag began, which a drag of the window's top edge carries with it.
    content: ScreenRegion,
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
    /// How many of the prober's readings the rows were last built from, which is the one number that
    /// says a row's facts have changed without a row being asked (see `start_the_prober`).
    probed: u64,
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
    /// The height's drag in flight: which of the window's two edges is being pulled, the screen row
    /// the pointer was at when it began, the height the band had, and the box the card stood in —
    /// which is the one thing a drag of the top edge moves and a drag of the bottom one does not.
    resizing: Option<PlaylistResize>,
    /// The scrollbar's thumb being carried, as the share of it the press took hold of: the bar's own
    /// drag rather than the band's, and the one thing a press on the track begins.
    thumb: Option<i32>,
    marquee: Option<Marquee>,
}

impl PinPlaylist {
    /// The rows the walk has moved to, and the row the pin is on: rebuilt only where the order is
    /// not the one in hand or a reading has landed since, since every row of it is a file read.
    ///
    /// Answers whether the order itself moved, which is the one thing that starts the prober: a
    /// walk step moves the pin's own file and leaves the list exactly as it was.
    fn hold_the_walk(&mut self, order: Vec<PathBuf>, path: PathBuf, probed: u64) -> bool {
        let moved = self.order != order;
        if moved {
            self.order = order;
            // A list that is not the one a pick was made in is not a list a pick means anything in.
            self.scroll = 0;
            self.hovered = None;
            self.picked = None;
            self.pressed = false;
            self.marquee = None;
        }

        // A row of a file nothing has read yet is a row of the extension and the room the file takes
        // and nothing else, so the rows are built again as the prober's readings land.
        if moved || self.probed != probed {
            self.rows = self
                .order
                .iter()
                .map(|path| pin_chrome::playlist_row(path))
                .collect();
            self.probed = probed;
        }

        self.showing = self.order.iter().position(|row| *row == path);
        self.path = Some(path);
        moved
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

    let (_, top, _, bottom) = pinned_window_box_of(
        pin.content,
        pin.dpi,
        pin.transport_bar,
        pin.overlay,
        pin.caption,
    );

    // **The row the band begins at is the window's own**, and the box above is measured on the
    // screen: the two are one subtraction. Without it every press in the band is answered against a
    // screen row no window coordinate ever reaches — a band nothing can be clicked in, a row nothing
    // can be picked on, and a region the wheel hook never finds the pointer in (see
    // `pin_playlist_keep_alive_try`, which adds the window's own place back to it).
    Some((bottom - top, height))
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
    // The window's own box moved with the flag, and it is the paint that follows that resizes it:
    // a band put up is a surface the window has to be given, and a band taken away is one it has to
    // be given back, so neither is left for whatever happens to repaint next (see `the_box_has_moved`).
    the_box_has_moved();

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

/// Ask the tick for the paint the band's new height is owed.
///
/// Nothing of the window is moved here and nothing is drawn here: the band's own box is the pin's
/// answer, and the paint that follows this is what gives the window its new height — including the
/// shrinking one, which is a surface the window has to be given back as much as it was given the
/// taller one.
pub(super) fn the_box_has_moved() {
    DIRTY.store(true, Ordering::Release);
    wake_preview_loop();
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
        // A band taken down is still owed the paint that gives the window its rows back, so the flag
        // is answered rather than thrown away: it is the one thing this side has to say about a band
        // that is no longer there, and a window that shrank for nothing else never would (see
        // `the_box_has_moved`).
        return DIRTY.swap(false, Ordering::AcqRel);
    }

    let mut owed = DIRTY.swap(false, Ordering::AcqRel) | refresh_the_rows();

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
    let probed = PROBED.load(Ordering::Acquire);
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
        if playlist.order == order
            && playlist.path.as_deref() == Some(pin.path.as_path())
            && playlist.probed == probed
        {
            return false;
        }

        (order.to_vec(), pin.path.clone())
    };

    let moved = match THE_PLAYLIST.lock() {
        Ok(mut playlist) => playlist.hold_the_walk(order, path, probed),
        Err(_) => return false,
    };
    if moved {
        start_the_prober();
    }
    true
}

/// Read the facts of every row nothing has read yet, on a thread of its own: a probe opens the file
/// through this machine's own codecs, which is not work a tick may pay for, and the rows are drawn
/// without those facts until each one lands (see `PROBED`).
///
/// The order the band was built from is the thread's whole list, and a list the pin has moved off is
/// a thread that stops at its next file rather than reading a folder nobody is looking at any more.
#[cfg(not(test))]
fn start_the_prober() {
    use crate::readers::audio_track::{self, Probed};

    let run = PROBE_RUN.fetch_add(1, Ordering::AcqRel) + 1;
    let order = match THE_PLAYLIST.lock() {
        Ok(playlist) => playlist.order.clone(),
        Err(_) => return,
    };

    let _ = std::thread::Builder::new()
        .name("pin-playlist-probe".to_string())
        .spawn(move || {
            for path in order {
                if PROBE_RUN.load(Ordering::Acquire) != run {
                    return;
                }
                if !matches!(audio_track::probed(&path), Probed::NotAsked) {
                    continue;
                }

                let probed = audio_engine::probe_audio_track_briefly(&path);
                audio_track::remember(
                    &path,
                    match probed {
                        Some(track) => Probed::Track(track),
                        None => Probed::Nothing,
                    },
                );
                PROBED.fetch_add(1, Ordering::AcqRel);
                DIRTY.store(true, Ordering::Release);
                wake_preview_loop();
            }
        });
}

/// A suite has no folder and no codecs, and a reading landed over a fact a test has written down is
/// a row no test put there: the prober is a machine's own business and is not asked for here (see
/// `start_the_prober`).
#[cfg(test)]
fn start_the_prober() {}

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
///
/// `y` is the window's own row, which the band's begins some way down: a hand above the band is a
/// hand on no row at all, and one on the card's own bar is not one either.
pub(super) unsafe fn hover(y: i32) -> bool {
    let Some((top, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };

    let local = y - top;
    let local = (0..height).contains(&local).then_some(local);

    set_the_hover((height, width, scale), local)
}

/// The row under the pointer the tick reads for itself, which is the one answer a move this window
/// is never sent cannot give: a hand that has left the window sends no more moves, and a row left
/// marqueed under a pointer that has gone is a name travelling for nobody (see
/// `pin_audio_hover_refresh`, which is the same question asked of the card's own buttons).
///
/// It is asked of the pin in hand rather than through `hover`, which takes this lock again.
pub(super) fn pin_playlist_hover_refresh(pin: &PinnedPreview, cursor: Option<(i32, i32)>) -> bool {
    let Some((top, height)) = pin_playlist_band(pin) else {
        return false;
    };
    let (width, _) = pin.window_size();
    let (left, window_top, right, _) = pin.window_box();
    let band_top = window_top + top;

    let local = cursor.and_then(|(x, y)| {
        (x >= left && x < right && y >= band_top && y < band_top + height).then(|| y - window_top)
    });

    set_the_hover((height, width, pin.dpi as f32 / 96.0), local)
}

/// The band's hover, written in the one place: which row a band-local row of the window is on, and
/// the marquee that comes with it, answering whether either changed.
///
/// It is handed the band's own geometry rather than asking the pin for it because its two callers
/// stand on either side of the pin's lock: a pointer's move arrives with the lock free, and the
/// tick that reads the cursor for itself holds it.
fn set_the_hover(band: (i32, i32, f32), local: Option<i32>) -> bool {
    let (height, width, scale) = band;
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };

    let layout = pin_chrome::playlist_layout(width, height, scale, playlist.rows.len());
    let hovered = local.and_then(|y| {
        pin_chrome::playlist_row_at(&layout, playlist.scroll, playlist.rows.len(), y)
    });

    let changed = playlist.hovered != hovered;
    playlist.hovered = hovered;
    playlist.set_the_marquee(&layout, scale);

    changed
}

/// A press on the band, answering whether the press was the band's: either of the window's two
/// edges takes the pointer and drags the band's height, and a press on a row picks it — never carries
/// the window, which is what a press anywhere else on a sound's card does.
///
/// A press that is the second of a pair shows the file that row is, the way a press on a row of the
/// listing behind the pin shows one, and without the configuration gate the hook's own picks are
/// behind: the hand is already inside this app's own window and has named the file itself.
///
/// **It is asked before anything the card owns** — its buttons, its bar, its walk — because the band
/// is the pin's own rows and none of it is the card's: a band answered by the card's arms is a band a
/// press picks nothing in and a double press swaps nothing with (see `pinned_press_on`).
pub(super) unsafe fn press(hwnd: HWND, x: i32, y: i32, double_click: bool) -> bool {
    let Some((top, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };

    // The band's two edges are the window's two edges, and a press on one is the height being
    // dragged rather than a row being picked: the card is never resized, and the band is the only
    // thing a pinned sound's window has to be resized by (see `PlaylistEdge`).
    if let Some(edge) = pin_chrome::playlist_edge_at(top, height, scale, y) {
        // Except where the card's own window buttons stand: they are four buttons in the window's
        // first rows, and a border drawn over them would be a border that eats the close button.
        if !the_edge_is_the_cards(edge, x, y) {
            let _ = SetCapture(hwnd);
            let content = pin_state()
                .and_then(|pinned| pinned.pin().map(|pin| pin.content))
                .unwrap_or_default();
            if let Ok(mut playlist) = THE_PLAYLIST.lock() {
                playlist.pressed = true;
                playlist.resizing = Some(PlaylistResize {
                    edge,
                    from: screen_y(hwnd, x, y),
                    height: HEIGHT.load(Ordering::Acquire),
                    content,
                });
            }
            return true;
        }
    }

    let local = y - top;
    if local < 0 || local >= height {
        return false;
    }

    let opened = {
        let Ok(mut playlist) = THE_PLAYLIST.lock() else {
            return true;
        };
        let layout = pin_chrome::playlist_layout(width, height, scale, playlist.rows.len());

        // **The scrollbar, before the rows it stands beside.** The bar is a mark on the window's own
        // right edge and a press on it carries the list rather than picking the row it happens to be
        // level with — and a press that is not on the thumb takes hold of the middle of it, which is
        // the one gesture a bar answers by jumping to where the hand landed (see
        // `pin_chrome::playlist_thumb_grab`).
        if pin_chrome::playlist_track_holds(&layout, x, local) {
            let grab = pin_chrome::playlist_thumb_grab(
                &layout,
                playlist.scroll,
                playlist.rows.len(),
                local,
            );
            // The pointer is taken for the bar, as it is for the band's own edge: a hand carrying the
            // thumb out of the window would otherwise stop being followed the moment it left it.
            let _ = SetCapture(hwnd);
            playlist.pressed = true;
            playlist.thumb = Some(grab);
            playlist.scroll =
                pin_chrome::playlist_scroll_at(&layout, playlist.rows.len(), local - grab);
            None
        } else {
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
        }
    };

    if let Some(path) = opened.as_deref() {
        show_pinned_file(path);
    }
    render_layered_preview(hwnd);

    true
}

/// The pointer's own move with something of the band's in flight, answering whether the move was the
/// band's: the thumb of its scrollbar being carried, or one of its two edges being pulled.
///
/// **Every travel is measured from where the drag began and on the screen**, which is the whole of
/// what keeps a pull of the window's top edge from shaking: that edge moves the window, so the window
/// coordinate of a pointer that has not moved changes with it, and a drag measured in one is a drag
/// that chases itself a pixel at a time. The screen is what both edges are measured against, and a
/// drag of the bottom one is the same arithmetic because its window does not move at all.
pub(super) unsafe fn drag(hwnd: HWND, x: i32, y: i32) -> bool {
    let (resize, thumb) = match THE_PLAYLIST.lock() {
        Ok(playlist) => (playlist.resizing, playlist.thumb),
        Err(_) => return false,
    };

    let Some((top, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };

    if let Some(grab) = thumb {
        let Ok(mut playlist) = THE_PLAYLIST.lock() else {
            return true;
        };
        let layout = pin_chrome::playlist_layout(width, height, scale, playlist.rows.len());
        let scrolled = pin_chrome::playlist_scroll_at(&layout, playlist.rows.len(), y - top - grab);

        if scrolled != playlist.scroll {
            playlist.scroll = scrolled;
            drop(playlist);
            render_layered_preview(hwnd);
        }
        return true;
    }

    let Some(resize) = resize else {
        return false;
    };

    let floor = pin_chrome::playlist_row_height(scale) * TWO_ROWS;
    let ceiling = CEILING.load(Ordering::Acquire).max(floor);
    let travel = screen_y(hwnd, x, y) - resize.from;

    let height = match resize.edge {
        pin_chrome::PlaylistEdge::Bottom => resize.height + travel,
        pin_chrome::PlaylistEdge::Top => resize.height - travel,
    }
    .clamp(floor, ceiling);

    if height == HEIGHT.load(Ordering::Acquire) {
        return true;
    }
    HEIGHT.store(height, Ordering::Release);

    // **The top edge is the bottom edge's own drag and a window moved under it.** The band grows by
    // exactly the rows the bottom edge would have given it, and the window is then carried up by them
    // so that the card goes with it and the window's own bottom stays where it is — which is the
    // illusion the hand is asking for, and why the card is never given a size of its own (see
    // `PlaylistEdge`).
    if matches!(resize.edge, pin_chrome::PlaylistEdge::Top) {
        let carried = resize.height - height;
        with_pin(|pin| {
            pin.content = (
                resize.content.0,
                resize.content.1 + carried,
                resize.content.2,
                resize.content.3 + carried,
            );
        });

        // The window is moved before it is painted, and only for the top edge: the paint is placed at
        // the window's own rectangle, and a top edge pulled up is a rectangle whose place has moved as
        // well as its size — while a bottom edge is the same place and a size the paint commits on its
        // own, which is how every other box of a pin's is committed (see
        // `render_pinned_preview_at`).
        if let Some((window, width, height)) = pinned_window_box() {
            let _ = SetWindowPos(
                hwnd,
                HWND::default(),
                window.0,
                window.1,
                width,
                height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }
    }
    render_layered_preview(hwnd);

    true
}

/// The pointer's place down the screen, which is what a drag of the window's own top edge is measured
/// by (see `drag`).
fn screen_y(hwnd: HWND, x: i32, y: i32) -> i32 {
    client_to_screen(hwnd, x, y).1
}

pub(super) fn edge_at(x: i32, y: i32) -> Option<pin_chrome::PlaylistEdge> {
    let (top, height, _, scale) = the_band_of_the_pin()?;
    let edge = pin_chrome::playlist_edge_at(top, height, scale, y)?;

    (!the_edge_is_the_cards(edge, x, y)).then_some(edge)
}

/// Whether the pointer on one of the band's edges is really on something the card owns there: the
/// four window buttons a pinned sound's card carries in the window's own first rows.
///
/// The band's top edge is the window's, and the card's buttons are in it. A border that answered for
/// them would be a border that ate the close button, so the buttons keep the rows they stand on and
/// the edge begins where they end (see `pin_audio_window_buttons_at`).
fn the_edge_is_the_cards(edge: pin_chrome::PlaylistEdge, x: i32, y: i32) -> bool {
    if !matches!(edge, pin_chrome::PlaylistEdge::Top) {
        return false;
    }

    pin_state()
        .and_then(|pinned| {
            pinned
                .pin()
                .map(|pin| pin_audio_window_buttons_at(pin, x, y))
        })
        .unwrap_or(false)
}

/// The release of a press the band took, answering whether it was the band's: a drag of the height
/// is over and the pointer given back, and a press on a row is over with nothing left to do.
///
/// The pointer is given back only where a press took it, which is the rule every arm of a release
/// on a pinned window keeps: a release that let go of a capture it never took would run
/// `pin_capture_lost` before the arm that owns one had read what its press armed.
pub(super) unsafe fn release(hwnd: HWND) -> bool {
    let (pressed, dragging, thumb) = match THE_PLAYLIST.lock() {
        Ok(mut playlist) => {
            // The press is spent here rather than left standing, and it is the whole of what keeps
            // the band's arm from swallowing the releases that follow it: a flag left set is the
            // release of every drag and every caption button the card begins afterwards, taken for
            // a press the band answered, and a window left carrying the pointer for the run.
            let pressed = std::mem::take(&mut playlist.pressed);
            let thumb = playlist.thumb.take();
            (pressed, playlist.resizing.take(), thumb)
        }
        Err(_) => return false,
    };
    if !pressed {
        return false;
    }

    // The pointer is given back for either drag the band can begin: the height's and the bar's, both
    // of which take it for as long as the hand is on them.
    if dragging.is_some() || thumb.is_some() {
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
        playlist.thumb = None;
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

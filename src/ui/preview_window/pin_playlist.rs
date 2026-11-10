//! The playlist a pinned sound's card grows under it: the rows of the pin's own walk, the press
//! that picks one, the drag that sets the band's height, the notch the wheel steps it by, and the
//! names the band keeps travelling — the row the hand is on, the one it has picked and the one the
//! pin is showing, each of which parks back at its resting place when the band lets it go.
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
/// which it has picked, and the names the band is travelling.
static THE_PLAYLIST: Lazy<Mutex<PinPlaylist>> = Lazy::new(|| Mutex::new(PinPlaylist::default()));

/// A name too long for its row, scrolled across it for as long as the band is travelling that row:
/// how far it has moved, which way it is going and when the hold at either end runs out — the card's
/// own marquee, in a row rather than in a header (see `audio_preview::NameScroll`).
#[derive(Debug, Clone, Copy)]
struct Marquee {
    /// The row whose name this is.
    row: usize,
    /// The travel itself, held as the card holds its own: a name the band has let go of winds back to
    /// its resting place rather than stopping where it stands (see `audio_preview::NameScroll`).
    scroll: audio_preview::NameScroll,
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

struct PinPlaylist {
    /// The rows in the order they are drawn, and the files they are the rows of — the two are one
    /// list, so a press answered in rows is answered in files.
    rows: Vec<pin_chrome::PlaylistRow>,
    order: Vec<PathBuf>,
    /// The positions into `rows` the search's query lets through: the drawing, the scroll, the
    /// scrollbar and a press all count in these, while `showing`, `picked`, `hovered` and the
    /// marquees stay the rows' own indices, which the filter never renumbers (see
    /// `position_of_row` and `row_at_position`).
    shown: Vec<usize>,
    /// The search field at the band's bottom: the query, its caret and whether the keys are going to
    /// it. It filters the rows by their own name, and nothing of a row's facts line or its length.
    query: text_find_field::Field,
    search_focused: bool,
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
    /// The names the band is travelling, one per row that is one of them — the row the hand is on,
    /// the row it has picked and the row the pin is showing — because two of those can be in sight at
    /// once. A name the band has let go of is kept here until it has wound home (see
    /// `set_the_marquee`).
    marquees: Vec<Marquee>,
    /// Whether the pointer is anywhere in the band, which is one of the two things that holds the
    /// field in sight and the only one that changes as a hand moves (see `retarget_the_search`).
    pointer_in_the_band: bool,
    /// Whether the field is meant to be in sight, and how much of it is: the target is the pointer in
    /// the band, the caret or a query hiding rows, and the slide is the share it has travelled toward
    /// it — `1.0` whole, `0.0` parked under the window's own edge (see `advance_the_search`).
    search_in_sight: bool,
    search_slide: f32,
    /// When the slide was last advanced, which is the cadence the next step of it is counted in.
    search_last: Instant,
}

impl Default for PinPlaylist {
    fn default() -> Self {
        Self {
            rows: Vec::new(),
            order: Vec::new(),
            shown: Vec::new(),
            query: text_find_field::Field::default(),
            search_focused: false,
            path: None,
            showing: None,
            probed: 0,
            scroll: 0,
            hovered: None,
            picked: None,
            pressed: false,
            resizing: None,
            thumb: None,
            marquees: Vec::new(),
            // **A band opens with its field parked.** The field is the pointer's and the caret's, and
            // a band drawn for a hand that has not reached for it yet has neither: it slides in when
            // the pointer comes into the band, which is the one gesture that asks for it (see
            // `advance_the_search`).
            pointer_in_the_band: false,
            search_in_sight: false,
            search_slide: 0.0,
            search_last: Instant::now(),
        }
    }
}

impl PinPlaylist {
    /// The rows the walk has moved to, and the row the pin is on: rebuilt only where the order is
    /// not the one in hand or a reading has landed since, since every row of it is a file read.
    ///
    /// Answers whether the order itself moved, which is the one thing that starts the prober: a
    /// walk step moves the pin's own file and leaves the list exactly as it was. And whether the row
    /// the pin is showing moved, which is a name the band travels for and one it lets go of (see
    /// `set_the_marquee`).
    fn hold_the_walk(&mut self, order: Vec<PathBuf>, path: PathBuf, probed: u64) -> (bool, bool) {
        let moved = self.order != order;
        if moved {
            self.order = order;
            // A list that is not the one a pick was made in is not a list a pick means anything in.
            self.scroll = 0;
            self.hovered = None;
            self.picked = None;
            self.pressed = false;
            self.marquees.clear();
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

        let showing = self.order.iter().position(|row| *row == path);
        let showing_moved = showing != self.showing;
        self.showing = showing;
        self.path = Some(path);

        // The rows are new, so the positions the query lets through are new: a filter worked out over
        // the list it replaced is a band drawn from rows nothing is holding any more. Nothing of what
        // the hand has picked or the pin is showing survives a row that is no longer drawn.
        self.refresh_the_shown(true);
        (moved, showing_moved)
    }

    /// The positions the query lets through, worked out afresh from the rows in hand, and everything
    /// that names a row given up where that row is gone from them.
    ///
    /// **The one place the filter is applied**, and the one place that says what a filter costs the
    /// scroll, by the reason it was applied rather than by counting the ways it could have been:
    ///
    /// - `keep_the_scroll` is a band whose *rows* moved — a facts line landing, or a pick that walked
    ///   the pin onto the file under the hand. **The scroll stays where the hand left it**, clamped
    ///   into a list that may be shorter than it was; a list that jumped home under a pick is a hand
    ///   that pressed a row and lost the place it was reading (see `refresh_the_rows`).
    /// - A query that moved is a list the hand has just asked a new question of, and the first
    ///   position the query leaves is the one it is reading from.
    fn refresh_the_shown(&mut self, keep_the_scroll: bool) {
        self.shown = pin_chrome::playlist_shown(&self.rows, self.query.text());

        let held = |row: Option<usize>| row.filter(|row| self.shown.contains(row));
        self.hovered = held(self.hovered);
        self.picked = held(self.picked);
        self.marquees
            .retain(|marquee| self.shown.contains(&marquee.row));

        self.scroll = if keep_the_scroll {
            self.scroll.min(self.shown.len().saturating_sub(1))
        } else {
            0
        };
    }

    /// The slide moved on by one tick of `cadence`, answering whether it moved: a whole slide takes
    /// `SEARCH_SLIDE_MS`, whatever its length, so the step is the share of it the cadence buys and a
    /// slow tick is a long step rather than a tick that travels nothing.
    ///
    /// **The cadence is handed in rather than read from the clock here**, because the clock is the
    /// preview loop's and a test's ticks are not a loop's: a slide measured by the wall is a slide no
    /// test can assert an intermediate position of (see `advance_the_search`).
    fn slide_the_search_by(&mut self, cadence: std::time::Duration) -> bool {
        let step = cadence.as_secs_f32() / (SEARCH_SLIDE_MS as f32 / 1000.0).max(f32::EPSILON);
        let target = if self.search_in_sight { 1.0 } else { 0.0 };
        let before = self.search_slide;
        let distance = target - before;

        self.search_slide = if step >= distance.abs() {
            target
        } else {
            before + distance.signum() * step
        };

        // A slide that has come to rest asks for nothing more: the tick that lands it asks for the
        // paint that draws it landed, and every tick after is a band of names repainted for a field
        // already parked.
        before != self.search_slide
    }

    /// **The row the hand picked and the row the pin is showing are rows the band draws**: the scroll
    /// is moved — and never the pick — so that both are inside the rows that fit, and only then is it
    /// clamped into the list. A pick a walk answered from off the bottom of the band is a pick the
    /// hand made on a row it could see a moment ago, and scrolling it out of sight is losing it.
    fn keep_the_selection_in_sight(&mut self, visible: usize) {
        for row in [self.showing, self.picked].into_iter().flatten() {
            if let Some(position) = self.position_of_row(row) {
                self.scroll = self.scroll_that_shows(position, visible);
            }
        }

        self.scroll = self
            .scroll
            .min(pin_chrome::playlist_scroll_max(self.shown.len(), visible));
    }

    /// The field's own edits, run in the one place so that the filter, the paints and the keys are all
    /// told the same thing by the same change: the band is dirtied where the query moved, as it is for
    /// every other move of its own (see `the_box_has_moved`).
    fn edit_the_query(&mut self, edit: impl FnOnce(&mut text_find_field::Field)) {
        let before = self.query.text().to_string();
        edit(&mut self.query);
        if self.query.text() == before {
            return;
        }

        self.refresh_the_shown(false);
        // **A query is what keeps the field in sight**, so a keystroke that leaves one behind is a
        // keystroke that holds the pill open — the target is asked again here because a hand typing
        // into a field the pointer has left is a hand the pointer's own answer knows nothing about.
        self.retarget_the_search();
        the_box_has_moved();
    }

    /// Put the query back to nothing, which is the whole list again: what `Esc` over the field asks.
    fn clear_the_query(&mut self) {
        self.edit_the_query(|query| query.clear());
    }

    /// Where a row stands among the positions the query lets through, or nothing where it is one of
    /// the rows the query hides.
    fn position_of_row(&self, row: usize) -> Option<usize> {
        self.shown.iter().position(|shown| *shown == row)
    }

    /// The row a position of the drawn list is, or nothing where the list is shorter than that.
    fn row_at_position(&self, position: usize) -> Option<usize> {
        self.shown.get(position).copied()
    }

    /// The rows the band draws, which is the whole walk whatever the query is: a paint slices the
    /// positions out of them rather than being handed a list that is not the walk's (see
    /// `paint_the_band`).
    fn the_search(&self, focused: bool) -> pin_chrome::PlaylistSearch<'_> {
        pin_chrome::PlaylistSearch {
            text: self.query.text(),
            caret: self.query.caret(),
            focused,
            slide: self.search_slide,
        }
    }

    /// **The one place the field's target is decided**, and the one place its slide is ever put
    /// somewhere without travelling: it is in sight while the pointer is anywhere in the band, while
    /// it has the caret, and while its query is not empty — the query is what is hiding rows, and a
    /// field a hand cannot see is a list it cannot explain.
    ///
    /// A field with either of the last two is put in sight **whole** rather than sliding in: it is not
    /// on its way anywhere, it is a field a hand is typing into, and a query that is filtering a list
    /// cannot be typed into a pill that is still on its way up. The pointer's own answer is the only
    /// one that animates, which is the whole of the parking the slide is for. Nothing is cleared to
    /// hide it, and no query is dropped to park it.
    fn retarget_the_search(&mut self) {
        let held = self.search_focused || !self.query.text().trim().is_empty();
        self.search_in_sight = self.pointer_in_the_band || held;
        if held {
            self.search_slide = 1.0;
        }
    }

    /// Where the row a pick is on has to be for the band to be drawing it: a pick made on a row the
    /// scroll no longer shows is a pick nothing on the screen says is the hand's.
    fn scroll_that_shows(&self, position: usize, visible: usize) -> usize {
        let visible = visible.max(1);
        let scroll = self.scroll;

        if position < scroll {
            position
        } else if position >= scroll + visible {
            position + 1 - visible
        } else {
            scroll
        }
    }

    /// The names the band's rows travel, set up when a row becomes one the band travels rather than
    /// on every frame a name moves: what it costs is a measure of the name at the scale it is drawn
    /// at.
    ///
    /// **Three rows a name travels for** — the one the hand is on, the one it has picked and the one
    /// the pin is showing — so that a name too long for its column is readable without the pointer
    /// being held over it. A row that stops being any of the three is *let go of* rather than dropped
    /// where it stands: its name winds back to its resting place first, which is the card's own rule
    /// for the line beside its name (see `audio_preview::NameScroll::set_engaged`, and
    /// `advance_the_marquees`).
    fn set_the_marquee(&mut self, layout: &pin_chrome::PlaylistLayout, scale: f32) {
        let travelling = [self.hovered, self.picked, self.showing];
        let room = layout.name_right - layout.text_left;
        let now = Instant::now();

        for marquee in &mut self.marquees {
            marquee
                .scroll
                .set_engaged(travelling.contains(&Some(marquee.row)));
        }

        for row in travelling.into_iter().flatten() {
            if self.marquees.iter().any(|marquee| marquee.row == row) {
                continue;
            }
            let Some((travel, advance)) = self
                .rows
                .get(row)
                .and_then(|row| pin_chrome::name_travel(&row.name, room, scale))
            else {
                continue;
            };

            self.marquees.push(Marquee {
                row,
                // The run's own width is its travel plus the room it is drawn in, which is the one
                // subtraction `fitted` works a travel out of (see `audio_preview::NameScroll`).
                scroll: audio_preview::NameScroll::fitted(travel + room, room, advance),
                last: now,
            });
        }
    }

    /// The names the band is travelling as the paint reads them: which row, and how far its name has
    /// been carried — nothing at all for a band with no marquee, and nothing for a row the query has
    /// hidden, since a row the band does not draw is a name it has no reason to carry.
    fn scrolled(&self) -> Vec<(usize, i32)> {
        self.marquees
            .iter()
            .filter(|marquee| self.shown.contains(&marquee.row))
            .map(|marquee| (marquee.row, marquee.scroll.offset()))
            .collect()
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

/// Whether the band is up, for the walk a press asks for: a sound's band is a list of the folder's
/// sounds, so the walk is `Category` while it is (see `pin_navigation::nav_mode_for`).
///
/// Only the anchor gates it — the band is a sound's card's and nothing else's — so a film's or a
/// page's walk is the same walk whatever this flag is left at.
pub(super) fn playlist_band_is_open() -> bool {
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

    // The band is what narrows the walk behind a sound's card to the folder's sounds, and what
    // gives the setting's own order back when it is put down, so the walk is asked again both ways
    // (see `pin_walk::ask_the_walk_again`). Asked as a message rather than walked from here: this
    // is the context menu's own press, and the pin's lock is held across it.
    send_preview(PreviewMessage::PinWalkAgain);

    if !open {
        // **A band put away is a list the next opening starts fresh.** The query a band was filtered
        // by, the caret it held and the pick its own rows had are this opening's and not the next
        // one's: a band reopened onto half a folder the hand was looking at last night is a list with
        // rows nothing on the screen explains the absence of.
        if let Ok(mut playlist) = THE_PLAYLIST.lock() {
            playlist.search_focused = false;
            playlist.query.clear();
            playlist.refresh_the_shown(false);
        }
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

    // **A band put up is a band whose field starts parked.** Whatever the band before it was left
    // holding — a caret, a query, a pointer that is long gone — is not this band's, and a field put up
    // part way in sight is a field the hand has to wait for. The pointer that is over the band now
    // brings it in on the tick after this (see `retarget_the_search`).
    if let Ok(mut playlist) = THE_PLAYLIST.lock() {
        playlist.search_focused = false;
        playlist.query.clear();
        playlist.pointer_in_the_band = false;
        playlist.search_in_sight = false;
        playlist.search_slide = 0.0;
        playlist.search_last = Instant::now();
        playlist.refresh_the_shown(false);
    }
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
        let _ = crate::app::seam::take_pin_playlist_delta();
        // A band taken down is still owed the paint that gives the window its rows back, so the flag
        // is answered rather than thrown away: it is the one thing this side has to say about a band
        // that is no longer there, and a window that shrank for nothing else never would (see
        // `the_box_has_moved`).
        return DIRTY.swap(false, Ordering::AcqRel);
    }

    let mut owed = DIRTY.swap(false, Ordering::AcqRel) | refresh_the_rows();

    let notches = crate::app::seam::take_pin_playlist_delta();
    if notches != 0 {
        owed |= step_the_notches(notches);
    }

    owed | advance_the_marquees() | advance_the_search()
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

    let (moved, showing_moved) = match THE_PLAYLIST.lock() {
        Ok(mut playlist) => playlist.hold_the_walk(order, path, probed),
        Err(_) => return false,
    };
    if moved {
        start_the_prober();
    }

    // **A walk that moved is a band that keeps the rows the hand is on where they are.** The rows a
    // filter reworks under a pick are the pick's own rows — the file the press named, its facts
    // landing, the stripe moving onto it — so the scroll the hand left is the scroll it is read at,
    // and only the two rows that mean something are brought back into sight: the picked row and the
    // one the pin is showing (see `keep_the_selection_in_sight`).
    if !moved {
        if let Some((_, height, width, scale)) = the_band_of_the_pin() {
            if let Ok(mut playlist) = THE_PLAYLIST.lock() {
                let visible =
                    pin_chrome::playlist_layout(width, height, scale, playlist.shown.len())
                        .visible
                        .max(0) as usize;
                playlist.keep_the_selection_in_sight(visible);
            }
        }
    }

    // A pin that has walked to another file is a row whose name travels now, and the row it has
    // walked off is a name let go of — set up here rather than left to the next move of a hand that
    // may never come, since a highlighted row's name is one the band travels without one (see
    // `set_the_marquee`).
    if showing_moved {
        if let Some((_, height, width, scale)) = the_band_of_the_pin() {
            if let Ok(mut playlist) = THE_PLAYLIST.lock() {
                let layout =
                    pin_chrome::playlist_layout(width, height, scale, playlist.shown.len());
                playlist.set_the_marquee(&layout, scale);
            }
        }
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
    use crate::readers::audio_track;

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

                // One ask per file per sweep, and a file nothing could be read from is left
                // unanswered for the sweep after this one rather than retried here: a folder of
                // files no probe can read must not become a thread reading them round. A row of a
                // folder read in an earlier run is answered by the disk, with nothing started at
                // all (see `audio_track::ask`).
                if audio_track::ask(&path, || audio_engine::probe_audio_track_briefly(&path))
                    .is_none()
                {
                    continue;
                }

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

    let layout = pin_chrome::playlist_layout(width, height, scale, playlist.shown.len());
    let rows = (notches * pin_chrome::PLAYLIST_WHEEL_ROWS) / WHEEL_DELTA;
    let max = pin_chrome::playlist_scroll_max(playlist.shown.len(), layout.visible.max(0) as usize);
    // A notch up is the list come back toward its first row, which is the wheel's own reading of
    // the two everywhere else in the app.
    let scrolled = (playlist.scroll as i32 - rows).clamp(0, max as i32) as usize;

    let moved = scrolled != playlist.scroll;
    playlist.scroll = scrolled;
    moved
}

/// The Up and Down arrows over a band that is up: the row the hand has picked stepped one row at a
/// time inside it, held at both ends rather than wrapped, and the band scrolled to keep it in sight.
///
/// The highlight is the one a press writes, so a hand that picked a row and a hand that arrowed to it
/// are on the same row. It answers whether the band took the key, which is what keeps the arrows off
/// the walk of the pinned file wherever it did (see `pin_geometry::pin_command_request`).
pub(super) fn step_the_pick(direction: i32) -> bool {
    let Some((_, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return true;
    };

    let layout = pin_chrome::playlist_layout(width, height, scale, playlist.shown.len());
    // The step is a step of the positions the query lets through, and the row it lands on is the row
    // the pick is written in: a filter renumbers neither of the two.
    let Some((position, scroll)) = pin_chrome::playlist_step_pick(
        playlist
            .picked
            .and_then(|row| playlist.position_of_row(row)),
        playlist
            .showing
            .and_then(|row| playlist.position_of_row(row)),
        playlist.scroll,
        playlist.shown.len(),
        layout.visible.max(0) as usize,
        direction,
    ) else {
        return true;
    };
    let Some(picked) = playlist.row_at_position(position) else {
        return true;
    };

    let moved = (playlist.picked, playlist.scroll) != (Some(picked), scroll);
    playlist.picked = Some(picked);
    playlist.scroll = scroll;
    playlist.set_the_marquee(&layout, scale);

    if moved {
        // A wash that moved is a paint the tick owes, asked for by the one flag every other move of
        // the band asks it by (see `the_box_has_moved`).
        the_box_has_moved();
    }
    true
}

/// The file the band's highlight is on, shown in the window that drew it: what a `Return` in a pinned
/// card asks for (see `PinCommand::PlayRow`).
///
/// It answers whether the band took the key, which is the whole of what keeps a `Return` in a window
/// with no band up — or with a band holding no pick — a key that does nothing at all. It is the row
/// the hand has *picked* and never the one it is merely on: a hover is where the pointer is, and the
/// file to play is one the hand has named.
pub(super) fn play_the_pick() -> bool {
    // A band that is down is not a list the hand is in, whatever row it was left picked on.
    if the_band_of_the_pin().is_none() {
        return false;
    }

    let picked = match THE_PLAYLIST.lock() {
        Ok(playlist) => playlist
            .picked
            .and_then(|index| playlist.order.get(index).cloned()),
        Err(_) => None,
    };

    // The lock is let go of before the pick is made, as it is for a double press on a row: showing a
    // file is a message to the loop, and the loop's own tick is what reads these rows (see
    // `show_pinned_file`).
    let Some(path) = picked else {
        return false;
    };

    show_pinned_file(&path);
    true
}

/// The band's own state, as a test reads it: the row the hand is on, the row it has picked, and the
/// row standing first in sight.
#[cfg(test)]
pub(super) fn the_band_in_hand() -> Option<(Option<usize>, Option<usize>, usize)> {
    let playlist = THE_PLAYLIST.lock().ok()?;
    Some((playlist.hovered, playlist.picked, playlist.scroll))
}

/// The query the band's search field holds, as a test reads it.
#[cfg(test)]
pub(super) fn the_query_in_hand() -> Option<(String, bool)> {
    let playlist = THE_PLAYLIST.lock().ok()?;
    Some((playlist.query.text().to_string(), playlist.search_focused))
}

/// The rows the query lets through, as a test reads them: the row indices, in the order the band draws
/// them.
#[cfg(test)]
pub(super) fn the_rows_in_sight() -> Vec<usize> {
    THE_PLAYLIST
        .lock()
        .map(|playlist| playlist.shown.clone())
        .unwrap_or_default()
}

/// Whether the band's search field holds the caret, which is what the hook is told once a tick (see
/// `crate::app::seam::publish_playlist_search_up`).
pub(super) fn search_has_focus() -> bool {
    THE_PLAYLIST
        .lock()
        .map(|playlist| playlist.search_focused)
        .unwrap_or(false)
}

/// Let the search field go, which a press anywhere else on the pin asks for: the query stays, because
/// it is what the band is filtered by — a hand that has gone to the transport bar has not asked for the
/// list it can see to come back (see `pin_input::pinned_press_on`).
pub(super) fn blur_the_search() {
    if let Ok(mut playlist) = THE_PLAYLIST.lock() {
        playlist.search_focused = false;
        // **A caret let go is a field's target asked again**: with no pointer in the band and no query
        // in the field there is nothing holding it, so it starts its way out from here (see
        // `retarget_the_search`).
        playlist.retarget_the_search();
    }
}

/// The field's own box, top and bottom, in the window's rows — the number a test keeps from changing
/// under a query or a drag.
#[cfg(test)]
pub(super) fn the_field_rows() -> Option<(i32, i32)> {
    let (top, height, width, scale) = the_band_of_the_pin()?;
    let rows = THE_PLAYLIST
        .lock()
        .map(|playlist| playlist.shown.len())
        .unwrap_or_default();
    let layout = pin_chrome::playlist_layout(width, height, scale, rows);

    Some((top + layout.search_top, top + layout.search_bottom))
}

/// The field's slide and whether it is meant to be in sight, as a test reads them: the share of the
/// pill that is on screen, and the target the tick is travelling toward.
#[cfg(test)]
pub(super) fn the_search_slide() -> Option<(f32, bool)> {
    let playlist = THE_PLAYLIST.lock().ok()?;
    Some((playlist.search_slide, playlist.search_in_sight))
}

/// Move the field's slide on by a tick of the band's own loop, which is what
/// `the_tick_is_owed_a_repaint` asks of it, answering whether it moved.
#[cfg(test)]
pub(super) fn slide_the_search_on(millis: u64) -> bool {
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };

    playlist.slide_the_search_by(std::time::Duration::from_millis(millis))
}

/// The pointer taken out of the band, which is what a tick that finds the cursor outside it asks: the
/// band's own answer, taken through the one place a hover is written.
#[cfg(test)]
pub(super) fn the_pointer_has_left_the_band() -> bool {
    let Some((_, height, width, scale)) = the_band_of_the_pin() else {
        return false;
    };

    set_the_hover((height, width, scale), None)
}

/// The wheel's notches, stepped on the band, which is what a wheel over a band asks of it: the answer
/// is whether the band scrolled, and the ask is always the band's own.
#[cfg(test)]
pub(super) fn step_the_notches_for_a_test(notches: i32) -> bool {
    if the_band_of_the_pin().is_none() {
        return false;
    }

    let _ = step_the_notches(notches);
    true
}

/// The rows a band draws, and which of the list's positions the hand is scrolled to.
#[cfg(test)]
pub(super) fn the_rows_and_scroll() -> (usize, usize) {
    let playlist = match THE_PLAYLIST.lock() {
        Ok(playlist) => playlist,
        Err(_) => return (0, 0),
    };

    (playlist.shown.len(), playlist.scroll)
}

/// Type a character into the band's search field, which is the road `WM_CHAR` answers it down — the
/// field's own edit, the filter run over it, and the band owed the paint it asks for.
pub(super) fn type_into_the_search(character: char) {
    if let Ok(mut playlist) = THE_PLAYLIST.lock() {
        playlist.edit_the_query(|query| query.insert(character));
    }
}

/// Answer one key of the field's own, when the field has the caret: `Backspace`, `Delete`, the two
/// arrows, the two ends, and the two keys that leave it. It answers whether the field took the key,
/// which is the whole of what keeps it off the pin's own key table — and a key it does not name is the
/// pin's, so `↑`/`↓` over a field that has the caret still step the list under it.
pub(super) fn key_while_the_search_has_focus(vk: i32) -> bool {
    if !search_has_focus() {
        return false;
    }

    let held = text_find::Modifiers::polled();
    // `Esc` empties the query and `Enter` only leaves it: the two ways out of a field of a list that
    // is being filtered under it, and the query outlives both of them.
    if vk == VK_ESCAPE.0 as i32 {
        if let Ok(mut playlist) = THE_PLAYLIST.lock() {
            playlist.clear_the_query();
            playlist.search_focused = false;
        }
        return true;
    }
    if vk == VK_RETURN.0 as i32 {
        if let Ok(mut playlist) = THE_PLAYLIST.lock() {
            playlist.search_focused = false;
        }
        return true;
    }

    // **The field's table is the find bar's**, asked of the same rule: two query fields on two surfaces
    // of one app are one editor's widget, and a backspace that moved the caret one way in one and
    // another way in the other is a field a hand has to learn twice. Everything the bar answers that a
    // list's field has no use for — its history, its steps, its switches — is left to the pin.
    let Some(edit) = text_find::key_edit_with(vk, held) else {
        return false;
    };
    let takes = matches!(
        edit,
        FindEdit::Backspace
            | FindEdit::DeleteForward
            | FindEdit::DeleteWordLeft
            | FindEdit::DeleteWordRight
            | FindEdit::Move(..)
            | FindEdit::SelectAll
            | FindEdit::Paste
    );
    if !takes {
        return false;
    }

    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return true;
    };
    match edit {
        FindEdit::Backspace => playlist.edit_the_query(|query| query.backspace()),
        FindEdit::DeleteForward => playlist.edit_the_query(|query| query.delete()),
        FindEdit::DeleteWordLeft => playlist.edit_the_query(|query| query.delete_word_left()),
        FindEdit::DeleteWordRight => playlist.edit_the_query(|query| query.delete_word_right()),
        FindEdit::SelectAll => playlist.query.select_all(),
        // The clipboard's text as the field's `paste` tidies it, taken from the keyboard's own road:
        // what a `WM_CHAR` cannot carry is a run of text, and a paste is the one run a query gets
        // that no key typed (see `text_find_field::sanitize_pasted`).
        FindEdit::Paste => {
            if let Some(pasted) = unsafe { text_selection::clipboard_text(HWND::default()) } {
                playlist.edit_the_query(|query| query.paste(&pasted));
            }
        }
        FindEdit::Move(motion, extend) => {
            // A caret move changes no query, so it is not the filter's and owes the band no paint.
            playlist.query.move_caret(motion, extend);
        }
        _ => {}
    }

    true
}

/// Focus the band's search field, which is what `Ctrl`+`F` over a pin with its band up asks: the caret
/// goes to the end of the query that is there, where the next character typed lands.
pub(super) fn focus_the_search() -> bool {
    // A band that is down has no field on screen to put a caret in.
    if the_band_of_the_pin().is_none() {
        return false;
    }

    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };
    playlist.search_focused = true;
    // A field whose caret has just been asked for is a field in sight whole: `Ctrl`+`F` over a parked
    // band is the hand asking for the pill, and it is there rather than being typed into unseen.
    playlist.retarget_the_search();
    let end = playlist.query.text().len();
    playlist.query.place_caret(end, false);

    true
}

/// The rows a name is travelling across, as the row and how far its name has been carried — nothing
/// at all for a band with no marquee.
#[cfg(test)]
pub(super) fn the_marquees_in_hand() -> Vec<(usize, i32)> {
    let Ok(playlist) = THE_PLAYLIST.lock() else {
        return Vec::new();
    };
    playlist.scrolled()
}

/// Move the band's travelling names on by one tick, answering whether any of them moved: each rests
/// at either end for `NAME_HOLD` and ping-pongs between them, which is the card's own rule for the
/// name it draws (see `audio_preview::NameScroll`).
///
/// A name the band has let go of is not stopped where it stands: it travels straight back to its
/// resting place and is given up there, so a row a hand has left is a name that parks rather than one
/// that jumps home (see `set_the_marquee`).
fn advance_the_marquees() -> bool {
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };
    if playlist.marquees.is_empty() {
        return false;
    }

    let now = Instant::now();
    let mut moved = false;

    playlist.marquees.retain_mut(|marquee| {
        let cadence = now.saturating_duration_since(marquee.last);
        let before = marquee.scroll.offset();
        marquee.scroll.advance(now, cadence);
        marquee.last = now;
        moved |= marquee.scroll.offset() != before;

        // A name that has come home and is one the band no longer travels is given up: what is left
        // of it is the resting place it is drawn at anyway, and a marquee kept there would be a band
        // repainted every tick for nothing (see `audio_preview::NameScroll::moves`).
        marquee.scroll.moves()
    });

    moved
}

/// **The field's slide, by one tick.** The share of it that is in sight travels toward its target a
/// tick at a time, so that a hand that reaches into the band sees a pill slide up out of the window's
/// own bottom edge and a hand that leaves sees it park there again — never a field that blinks.
///
/// The target is the pointer in the band, the caret, and a query that is hiding rows (see
/// `PinPlaylist::should_the_search_be_in_sight`): a field that is filtering a list is a field the hand
/// has to be able to read and edit, and one that is parked under a pointer that has gone is a field
/// nothing is asking for. The ends are snapped to exactly, because a slide left a hair short of its
/// end is a tick asking for a paint for the rest of the run.
fn advance_the_search() -> bool {
    let Ok(mut playlist) = THE_PLAYLIST.lock() else {
        return false;
    };

    let now = Instant::now();
    let cadence = now.saturating_duration_since(playlist.search_last);
    playlist.search_last = now;

    playlist.slide_the_search_by(cadence)
}

/// A pointer over the band, answering whether the band changed for it: which row the hand is on is
/// the marquee's own answer, and a row a hand has come to is a name that starts moving.
///
/// `y` is the window's own row, which the band's begins some way down: a hand above the band is a
/// hand on no row at all, and one on the card's own bar is not one either.
pub(super) unsafe fn hover(y: i32) -> bool {
    let Some(band) = the_band_of_the_pin() else {
        return false;
    };

    hover_at_the_band(band, Some(y))
}

/// The band's answer to a pointer at a row of the window, or at no row at all — **the one place the
/// band's own top is taken off the window's row**, so that a move the window is sent and a tick that
/// reads the cursor off the screen cannot name two different rows of one band.
///
/// It is handed the band's own geometry rather than asking the pin for it because the tick stands
/// inside the pin's lock, which `the_band_of_the_pin` would block on (see `set_the_hover`).
fn hover_at_the_band((top, height, width, scale): (i32, i32, i32, f32), y: Option<i32>) -> bool {
    let local = y
        .map(|y| y - top)
        .filter(|local| (0..height).contains(local));

    set_the_hover((height, width, scale), local)
}

/// The row under the pointer the tick reads for itself, which is the one answer a move this window
/// is never sent cannot give: a hand that has left the window sends no more moves, and a row left
/// marqueed under a pointer that has gone is a name travelling for nobody (see
/// `pin_audio_hover_refresh`, which is the same question asked of the card's own buttons).
///
/// **The row is the window's own**, read off the client rect the pointer is in rather than counted
/// from the box the pin's content projects: a window rests on the screen and that box does not have
/// to be where it rests, so the subtraction this used to do named a row the move never had — a row
/// hand is on that the tick cleared again, marquee and all. What a window says about itself is
/// DPI-exact, that arithmetic is not.
///
/// A pointer outside the window is on no row of it, and that is the whole of what clears the hover
/// of a hand that has walked away; the picked row is not the hover and is not touched here.
pub(super) fn pin_playlist_hover_refresh(
    pin: &PinnedPreview,
    hwnd: HWND,
    cursor: Option<(i32, i32)>,
) -> bool {
    let Some((top, height)) = pin_playlist_band(pin) else {
        return false;
    };
    let Some((x, y, width)) = cursor.and_then(|at| the_point_in_the_window(hwnd, at)) else {
        return false;
    };

    let y = (0..width).contains(&x).then_some(y);

    hover_at_the_band((top, height, width, pin.dpi as f32 / 96.0), y)
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

    let layout = pin_chrome::playlist_layout(width, height, scale, playlist.shown.len());
    // A hover is a position of the drawn list, and the row the hand is on is the row that position
    // names: the marquees and the stripe are rows, the scroll and the pointer are positions.
    let hovered = local.and_then(|y| {
        pin_chrome::playlist_row_at(&layout, playlist.scroll, playlist.shown.len(), y)
            .and_then(|position| playlist.row_at_position(position))
    });

    let changed = playlist.hovered != hovered;
    playlist.hovered = hovered;
    playlist.set_the_marquee(&layout, scale);

    // **A pointer anywhere in the band is a hand reaching for its field.** The pointer is asked of the
    // band's own rows rather than of a row: the pads above and below a list are the band's too, and a
    // field that stayed parked under a hand that is in the band is a field the hand cannot reach — the
    // engagement is the whole of the band, and the row is the marquee's answer alone.
    //
    // **And the field is never taken away from a hand that has named it.** A field with the caret or
    // one filtering the list is in sight whole rather than mid-slide, whatever the pointer does: it is
    // the only thing on the band that explains rows missing, and the only way to put them back (see
    // `PinPlaylist::retarget_the_search`).
    let was = (playlist.search_in_sight, playlist.search_slide);
    playlist.pointer_in_the_band = local.is_some();
    playlist.retarget_the_search();

    changed || (playlist.search_in_sight, playlist.search_slide) != was
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
        let layout = pin_chrome::playlist_layout(width, height, scale, playlist.shown.len());

        // **The search field first, and it picks nothing.** A press on it is the caret's and the keys',
        // and the field stands where a row cannot: the press is put down before the scrollbar and the
        // rows are asked, or the field's own bottom row would be a row of a list that is not under it.
        //
        // **A field on its way out is not one yet.** The pill's strip is reserved whether the pill is
        // in it or not, so a press there while the slide is anywhere short of whole is a press on the
        // band's own bottom pad — the hand has reached for a field that is not there, and what it gets
        // is nothing rather than a caret in a pill it cannot see. The press is still a press *in* the
        // band, so the field it reached for comes down to meet it (see `advance_the_search`).
        let reached_for_the_field = pin_chrome::search_field_holds(&layout, x, local);
        if reached_for_the_field && playlist.search_slide < 1.0 {
            playlist.pointer_in_the_band = true;
            playlist.retarget_the_search();
            return true;
        }
        if reached_for_the_field {
            // **The caret is what holds the field open**, so it is written before the paint that
            // follows: a field focused is a field in sight whole whatever the pointer then does (see
            // `retarget_the_search`).
            playlist.search_focused = true;
            playlist.pointer_in_the_band = true;
            playlist.retarget_the_search();
            let at = pin_chrome::search_index_at(&layout, playlist.query.text(), x, scale);
            playlist.query.place_caret(at, false);
            None
        } else {
            // A press anywhere else on the band is the hand leaving the field, and the query it typed
            // stays: it is what the band is filtered by until `Esc` or the next press on the field.
            // Whether that lets the pill park is the target's own answer, not this one's: a query that
            // is still hiding rows keeps it in sight (see `retarget_the_search`).
            playlist.pointer_in_the_band = true;
            playlist.search_focused = false;
            playlist.retarget_the_search();

            // **The scrollbar, before the rows it stands beside.** The bar is a mark on the window's own
            // right edge and a press on it carries the list rather than picking the row it happens to be
            // level with — and a press that is not on the thumb takes hold of the middle of it, which is
            // the one gesture a bar answers by jumping to where the hand landed (see
            // `pin_chrome::playlist_thumb_grab`).
            if pin_chrome::playlist_track_holds(&layout, x, local) {
                let grab = pin_chrome::playlist_thumb_grab(
                    &layout,
                    playlist.scroll,
                    playlist.shown.len(),
                    local,
                );
                // The pointer is taken for the bar, as it is for the band's own edge: a hand carrying the
                // thumb out of the window would otherwise stop being followed the moment it left it.
                let _ = SetCapture(hwnd);
                playlist.pressed = true;
                playlist.thumb = Some(grab);
                playlist.scroll =
                    pin_chrome::playlist_scroll_at(&layout, playlist.shown.len(), local - grab);
                None
            } else {
                let Some(position) = pin_chrome::playlist_row_at(
                    &layout,
                    playlist.scroll,
                    playlist.shown.len(),
                    local,
                ) else {
                    // The pad between two rows and the room below the last of them are the band's own: a
                    // press there picks nothing and is still not a press the card answers.
                    return true;
                };
                let Some(index) = playlist.row_at_position(position) else {
                    return true;
                };

                playlist.pressed = true;
                playlist.picked = Some(index);
                playlist.set_the_marquee(&layout, scale);

                double_click
                    .then(|| playlist.order.get(index).cloned())
                    .flatten()
            }
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
        let layout = pin_chrome::playlist_layout(width, height, scale, playlist.shown.len());
        let scrolled =
            pin_chrome::playlist_scroll_at(&layout, playlist.shown.len(), y - top - grab);

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
    // The layout and the scroll are counted in the positions the query lets through, while every row
    // the band draws is asked for by its own index: the filter hides rows and renumbers nothing.
    let layout = pin_chrome::playlist_layout(width, paint.playlist, scale, playlist.shown.len());
    let first = playlist.scroll.min(playlist.shown.len());
    let search = playlist.the_search(playlist.search_focused);
    // A query that hides every row is the band saying so rather than an empty room: an empty walk
    // says nothing at all, because there is nothing there to have been hidden.
    let note = (playlist.shown.is_empty() && !playlist.rows.is_empty()).then_some(NO_MATCHES);
    // The names travelling, as the rows and how far each has been carried: the paint is handed them
    // beside the rows rather than told which row the hand is on, because a row the band has let go of
    // is still on its way back to its resting place (see `set_the_marquee`).
    let scrolled = playlist.scrolled();
    let band = pin_chrome::PlaylistBand {
        layout,
        rows: &playlist.rows,
        shown: &playlist.shown,
        first,
        total: playlist.shown.len(),
        showing: playlist.showing,
        picked: playlist.picked,
        scrolled: &scrolled,
        search: Some(search),
        note,
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

/// How long the band's search field takes to slide all the way in or out, in milliseconds: a fifth of
/// a second, which is long enough to read as travel and short enough that a hand reaching into the
/// band is not kept waiting for the field it reached for.
const SEARCH_SLIDE_MS: u64 = 200;

/// What the band says where a query has hidden every row, in the muted ink a facts line is drawn in.
const NO_MATCHES: &str = "No names match";

/// The share of the display a band opens at when nothing is remembered, and the most it may ever
/// be: a band is for reading a folder's worth of names, and a band that has taken the whole screen
/// has taken the window.
const DEFAULT_SHARE_OF_THE_DISPLAY: f32 = 0.6;
const MOST_OF_THE_DISPLAY: f32 = 0.9;

/// The shortest a band may be dragged: two rows, because a band of one is a row a hand cannot
/// scroll to the next one of.
const TWO_ROWS: i32 = 2;

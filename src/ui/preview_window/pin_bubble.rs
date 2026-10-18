//! The pin's bubble: the smaller window that stands beside it for a collapsed pin, what is drawn
//! inside it, and the drag that carries it.
//!
//! What is drawn inside it is the pin's own media where a bubble can show it — the frame a film is
//! on, the first plate of a page, the picture a picture is — and a mark saying what kind of thing
//! it stands for where it cannot (see `bubble_art` and `pin_chrome::BubbleMark`). A film a player of
//! somebody else's is drawing is the one kind that is both: it holds no frame on this side at all, so
//! one is taken out of the file for the still a bubble goes up with (see `grab_video_frame`), and the
//! frames of the film *moving* are read beside that player (see `bubble_stream`).

use super::*;

/// The class the round bubble a collapsed pin leaves is created from.
pub(super) const PIN_BUBBLE_CLASS: PCWSTR = w!("RustHoverPreviewPinBubble");

/// The bubble's window, or zero while no pin is collapsed.
///
/// It is read where a press is told apart from a press that landed on the bubble rather
/// than somewhere else, out on the Explorer hook's side: a collapsed pin's window is
/// hidden, so the bubble is the only window of the pin's own a press can land on while
/// a pin is one (see `bubble_window_is_at`).
pub(crate) static PIN_BUBBLE_HWND: AtomicIsize = AtomicIsize::new(0);

/// One frame of a film a player of somebody else's is drawing, as the bubble shows it: the picture,
/// already scaled to cover a square and cropped to the middle of it, and the side of that square.
///
/// It is held as one thing rather than as a pair of statics because the two are one fact — a
/// picture of `side × side` BGRA is what the bubble paints — and a reader that could see the
/// picture without the size it was cropped for would be reading a frame of nothing.
#[derive(Clone)]
pub(super) struct BubbleFrame {
    pub(super) pixels: Vec<u8>,
    pub(super) side: u32,
}

/// The frame taken for the bubble a collapsed film left, if it has landed yet — or the frame a
/// pause of that film handed over in its place (see `keep_bubble_frame`).
///
/// It is a slot beside the bubble rather than a part of the pin, because that is what it is: the
/// frame belongs to the *collapsed* window, whose own media is off the screen and whose process is
/// drawing somewhere else — and a pin put back up has no bubble for it to be shown in.
pub(super) static BUBBLE_FRAME: Lazy<Mutex<Option<BubbleFrame>>> = Lazy::new(|| Mutex::new(None));

/// Whether a frame has landed in the slot above since the bubble was last painted, **or a frame has
/// been read out of a film beside its player** (see `bubble_stream`).
///
/// The grab is a process on a thread of its own, so the paint that put the bubble up cannot be the
/// one that shows what the grab found: this is the note it leaves, and the loop's next tick reads
/// it and paints the bubble again (see the bubble repaint in `run_preview_window`). It is raised
/// whether or not a frame came back — a file with no frame to give is a bubble that stays as it
/// was, and a flag left standing for a grab that answered nothing is a repaint every tick. The same
/// note is what a live frame leaves, and for the same reason: nothing else about a bubble over a
/// film at a level of nothing moves (see `bubble_bars_are_moving`).
pub(super) static BUBBLE_FRAME_READY: AtomicBool = AtomicBool::new(false);

/// How fast the wave in the bars a mixer's peak is drawn as travels, in radians a second.
///
/// A rate of its own rather than a period counted in repaints, because the phase is worked out from
/// the wall clock and nothing else: the loop's cadence is what makes that row *move*, and this is
/// what says how far a bar has moved between two of its ticks (see `render_pin_bubble` and
/// `pin_chrome::peak_bars`).
const BUBBLE_PHASE_RADIANS_PER_SECOND: f32 = 2.4;

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

/// Whether the press that is down landed on a bubble already holding the keyboard: the first
/// click focuses, the second restores (see `pin_bubble_proc`).
pub(super) static BUBBLE_PRESS_FOCUSED: AtomicBool = AtomicBool::new(false);

/// Whether the bubble is the window the user is in: the gate on the SPACE below, so a SPACE typed
/// into whatever is behind the bubble stays that window's key.
pub(super) fn bubble_is_focused() -> bool {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return false;
    }
    unsafe { GetFocus() == HWND(hwnd as *mut _) }
}

/// Whether a SPACE with the bubble focused holds anything: a film or a sound, and nothing else.
pub(super) fn bubble_space_toggles(kind: Option<MediaType>) -> bool {
    matches!(
        kind,
        Some(MediaType::Video) | Some(MediaType::NativeVideo) | Some(MediaType::Audio)
    )
}

/// Whether a press that is up puts the window back: a click — no move — on a bubble the hand was
/// already in. The focusing click stays a bubble, so the SPACE below has a window to hold.
pub(super) fn bubble_click_restores(was_focused: bool, moved: bool) -> bool {
    was_focused && !moved
}

/// One row of the menu a bubble opens — the same rows a pinned window's own menu carries, and
/// answered the same way on both wherever the two meet (see `shared_menu_row_press`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum BubbleMenuRow {
    /// Hold what is playing where it stands, or set it going again.
    Toggle,
    /// The file after the one pinned, the same walk the pin's own arrow takes.
    NextFile,
    /// The file before the one pinned.
    PreviousFile,
    /// The rule between the walk and the settings under it, between the switch and the two picks
    /// below it, and between those picks and the ways out of the menu.
    Separator,
    /// Whether a pin the minimize button puts away holds what it is playing where it stands: the
    /// switch the kind's own `Pin Mode` in the tray carries, on the bubble a minimize is what
    /// left here (see `pin_pause_video` and `pin_pause_audio`).
    PauseOnMinimize,
    /// Where a pin the minimize button puts away goes: the round bubble on the desktop. One of
    /// the two picks a group only one of which can be in force carries (see `PinMinimizeTo`).
    MinimizeToBubble,
    /// And the other: the app's own stand-in in the taskbar.
    MinimizeToTaskbar,
    /// Put the pin away into the taskbar rather than the bubble: the stand-in the pick above
    /// names, gone to now instead of at the next press of the minimize button — and on a bubble,
    /// which is a pin already away, the only thing the two picks beside it can mean.
    ToTaskbar,
    /// Take the pin down.
    Close,
}

/// The rows the bubble's menu holds, from the top of the panel to the bottom: the playback of
/// the file behind the bubble, the walk of the folder it stands in, the switch that decides what
/// a minimize holds back, the two places a minimize can put the pin, and the two ways out of any
/// bubble. The hold and the switch are rows only where the kind behind the bubble has playback
/// for them to answer about; the two places are every menu's, because a minimize is a thing any
/// pin has (see `bubble_pause_on_minimize` and `PinMinimizeTo`).
pub(super) fn bubble_menu_rows(kind: Option<MediaType>) -> Vec<BubbleMenuRow> {
    let mut rows = Vec::with_capacity(12);
    if bubble_space_toggles(kind) {
        rows.push(BubbleMenuRow::Toggle);
    }
    rows.push(BubbleMenuRow::NextFile);
    rows.push(BubbleMenuRow::PreviousFile);
    rows.push(BubbleMenuRow::Separator);
    if bubble_pause_on_minimize(kind).is_some() {
        rows.push(BubbleMenuRow::PauseOnMinimize);
        rows.push(BubbleMenuRow::Separator);
    }
    rows.push(BubbleMenuRow::MinimizeToBubble);
    rows.push(BubbleMenuRow::MinimizeToTaskbar);
    rows.push(BubbleMenuRow::Separator);
    rows.push(BubbleMenuRow::ToTaskbar);
    rows.push(BubbleMenuRow::Close);
    rows
}

/// What the kind behind the bubble has the `Pause on Minimize` switch at, or nothing where the
/// kind has no switch: whether a pin the minimize button puts away holds what it is playing
/// where it stands is a question a film and a sound have an answer to and nothing else has.
fn bubble_pause_on_minimize(kind: Option<MediaType>) -> Option<bool> {
    match kind {
        Some(MediaType::Video) | Some(MediaType::NativeVideo) => Some(pin_pause_video()),
        Some(MediaType::Audio) => Some(pin_pause_audio()),
        _ => None,
    }
}

/// Turn the `Pause on Minimize` switch over for the kind behind the bubble: written under the
/// configuration's lock and saved, the way every setting this app's own menus write is (see
/// `pin_menu::toggle_pin_audio_shuffle`). A kind with no switch of its own writes nothing,
/// because its menu carries no row to press.
fn toggle_pause_on_minimize(kind: Option<MediaType>) {
    if let Ok(mut config) = CONFIG.lock() {
        match kind {
            Some(MediaType::Video) | Some(MediaType::NativeVideo) => {
                config.pin_pause_video = !config.pin_pause_video
            }
            Some(MediaType::Audio) => config.pin_pause_audio = !config.pin_pause_audio,
            _ => return,
        }
        config.save();
    }
}

/// Where a pin the minimize button puts away goes, written the way the switch above it is. The
/// two picks are one group — a minimize leaves a pin either on the desktop as its bubble or in
/// the app's own place in the taskbar, never both — so the setting one of them writes is also
/// what the disc on the row is read from the next time the menu is painted (see
/// `bubble_menu_row_mark`), and nothing on screen moves: the pin already away is left where it
/// was put, and the answer is the next minimize's (see `pin_minimize_to`).
fn set_pin_minimize_to(target: PinMinimizeTo) {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_minimize_to = target;
        config.save();
    }
}

/// What a press on one of the rows a bubble's menu and a pinned window's menu both hold does,
/// where the two answer it the same way — which is every row but `ToTaskbar`, whose road is
/// handed in because the two menus walk it differently: a pin already collapsed moves the bubble
/// it left on the desktop into the app's stand-in, and one still standing puts itself away from
/// its own window (see `bubble_menu_press` and `pin_menu`). Answers whether the panel stays up.
///
/// A row that picks one of two answers goes away with the pick, the way a menu a choice was made
/// in does; the switch above them stays up with the mark it now carries, because there is
/// nothing to pick from and the mark is the answer.
pub(super) fn shared_menu_row_press(row: BubbleMenuRow, to_taskbar: impl FnOnce()) -> bool {
    match row {
        BubbleMenuRow::Toggle => {
            ask_pin(PinCommand::TogglePlayback);
            false
        }
        BubbleMenuRow::NextFile => {
            ask_pin(PinCommand::Next);
            false
        }
        BubbleMenuRow::PreviousFile => {
            ask_pin(PinCommand::Previous);
            false
        }
        BubbleMenuRow::Close => {
            ask_pin(PinCommand::Close);
            false
        }
        BubbleMenuRow::PauseOnMinimize => {
            toggle_pause_on_minimize(current_media_type());
            true
        }
        BubbleMenuRow::MinimizeToBubble => {
            set_pin_minimize_to(PinMinimizeTo::Bubble);
            false
        }
        BubbleMenuRow::MinimizeToTaskbar => {
            set_pin_minimize_to(PinMinimizeTo::Taskbar);
            false
        }
        BubbleMenuRow::ToTaskbar => {
            to_taskbar();
            false
        }
        BubbleMenuRow::Separator => true,
    }
}

/// What a row of the bubble's menu says: the hold named for what it will do, the two walks
/// named for the file they step to, the switch spelled the way the tray spells it, the two
/// places a minimize can put the pin under the words the panel has room for, and the two ways
/// out the same way. A rule carries nothing at all, because there is nothing on one to read.
pub(super) fn bubble_menu_row_label(row: BubbleMenuRow, playing: bool) -> &'static str {
    match row {
        BubbleMenuRow::Toggle if playing => "Pause",
        BubbleMenuRow::Toggle => "Play",
        BubbleMenuRow::NextFile => "Next File",
        BubbleMenuRow::PreviousFile => "Previous File",
        BubbleMenuRow::Separator => "",
        BubbleMenuRow::PauseOnMinimize => "Pause on Minimize",
        BubbleMenuRow::MinimizeToBubble => "Minimize to Bubble",
        BubbleMenuRow::MinimizeToTaskbar => "Minimize to Taskbar",
        BubbleMenuRow::ToTaskbar => "To Taskbar",
        BubbleMenuRow::Close => "Close",
    }
}

/// The mark a row of the bubble's menu carries: nothing anywhere, except the hairline on the
/// rule, the check the switch stands for, and the disc on the one of the two picks that is the
/// setting in force (see `pin_chrome::MenuMark`).
fn bubble_menu_row_mark(
    row: BubbleMenuRow,
    pause_on_minimize: Option<bool>,
    minimize_to: PinMinimizeTo,
) -> pin_chrome::MenuMark {
    match row {
        BubbleMenuRow::Separator => pin_chrome::MenuMark::Separator,
        BubbleMenuRow::PauseOnMinimize => {
            pin_chrome::MenuMark::Check(pause_on_minimize.unwrap_or(false))
        }
        BubbleMenuRow::MinimizeToBubble if minimize_to == PinMinimizeTo::Bubble => {
            pin_chrome::MenuMark::Bullet
        }
        BubbleMenuRow::MinimizeToTaskbar if minimize_to == PinMinimizeTo::Taskbar => {
            pin_chrome::MenuMark::Bullet
        }
        _ => pin_chrome::MenuMark::None,
    }
}

/// The rows the bubble's menu is painted with: the same panel a pinned sound's card opens, so a
/// bubble's menu reads as the same menu (see `pin_chrome::paint_menu_popup`) — and the same rows
/// a pinned window's own menu is painted with (see `pin_menu::pinned_menu_paint`).
pub(super) fn bubble_menu_chrome_rows(
    kind: Option<MediaType>,
    playing: bool,
) -> Vec<pin_chrome::MenuRow> {
    let pause_on_minimize = bubble_pause_on_minimize(kind);
    let minimize_to = pin_minimize_to();
    bubble_menu_rows(kind)
        .iter()
        .map(|row| pin_chrome::MenuRow {
            label: bubble_menu_row_label(*row, playing).to_string(),
            mark: bubble_menu_row_mark(*row, pause_on_minimize, minimize_to),
        })
        .collect()
}

/// The bubble's menu while it is up: the rows it holds, the panel they were placed in, the row
/// the pointer is on, and the scale both were worked out at.
///
/// It is a slot beside the bubble rather than a part of the pin for the reason the bubble's
/// frame is (see `BUBBLE_FRAME`): it belongs to the collapsed window, and a pin put back up has
/// no menu for it to be shown in.
pub(super) struct BubbleMenu {
    pub(super) rows: Vec<pin_chrome::MenuRow>,
    pub(super) popup: pin_chrome::MenuPopup,
    pub(super) hover: Option<usize>,
    pub(super) dpi: u32,
}

pub(super) static BUBBLE_MENU: Lazy<Mutex<Option<BubbleMenu>>> = Lazy::new(|| Mutex::new(None));

/// Place the bubble's menu panel at a point in the window: the shared placement, with the flyout
/// arrow cleared — no flyout hangs off this menu, and none of its rows is a door, so no row of it
/// carries an arrow (see `MenuArrow`).
pub(super) fn bubble_menu_popup(
    point: (i32, i32),
    width: i32,
    height: i32,
    dpi: u32,
    rows: &[pin_chrome::MenuRow],
) -> pin_chrome::MenuPopup {
    let mut popup = pin_chrome::menu_popup_from_point(point, width, height, dpi, rows);
    popup.arrow = pin_chrome::MenuArrow::None;
    popup
}

/// Whether the bubble's menu is up.
pub(super) fn bubble_menu_is_open() -> bool {
    BUBBLE_MENU
        .lock()
        .map(|menu| menu.is_some())
        .unwrap_or(false)
}

/// Open the bubble's menu at the pointer: the window grows down to hold the circle and the panel
/// beneath it, so the menu is painted with the theme's own panel the way a card's menu is
/// rather than with a menu of the Shell's (see `paint_menu_popup`).
///
/// A second right-click while the panel is up moves it to the new point rather than putting it
/// away, the way a card's menu does (see `open_pin_menu`).
pub(super) unsafe fn open_bubble_menu(hwnd: HWND) {
    let Some(cursor) = cursor_screen_point() else {
        return;
    };
    let Some(origin) = window_origin(hwnd).map(|rect| (rect.0, rect.1)) else {
        return;
    };
    let dpi = dpi_at(cursor.0, cursor.1);
    let side = logical_px(dpi, PIN_BUBBLE_PIXELS).max(16);
    let rows = bubble_menu_chrome_rows(current_media_type(), bubble_is_playing());

    // The panel's natural size first, so the window grows only as big as the circle and the
    // panel beneath it: measured with room to spare, then held inside the grown window below.
    let natural = bubble_menu_popup((0, side), 10_000, 10_000, dpi, &rows);
    let width = side.max(natural.panel.right);
    let height = natural.panel.bottom.max(side);
    let grown = clamp_pinned_box(
        (origin.0, origin.1, origin.0 + width, origin.1 + height),
        dpi,
        &DESKTOPS,
    );
    let (x, y) = (grown.0, grown.1);
    let (width, height) = (grown.2 - grown.0, grown.3 - grown.1);
    let popup = bubble_menu_popup((cursor.0 - x, cursor.1 - y), width, height, dpi, &rows);

    if let Ok(mut menu) = BUBBLE_MENU.lock() {
        *menu = Some(BubbleMenu {
            rows,
            popup,
            hover: None,
            dpi,
        });
    }
    // A press the menu answers is not a press on the bubble: the focusing click it could have
    // been is spent opening the menu instead.
    BUBBLE_PRESS_FOCUSED.store(false, Ordering::Release);
    let _ = MoveWindow(hwnd, x, y, width, height, true);
    // The menu is this window's own, so the window takes the pointer while it is up: a press
    // anywhere else arrives here first, and the first one puts the menu away rather than reaching
    // through to whatever was aimed at (see `open_pin_menu`).
    let _ = SetCapture(hwnd);
    render_pin_bubble(hwnd);
}

/// Put the bubble's menu away where it is up, answering whether it came down: the window shrinks
/// back to the circle, and the pointer with it, while the bubble itself is left standing.
pub(super) unsafe fn close_bubble_menu(hwnd: HWND) -> bool {
    let was_up = BUBBLE_MENU
        .lock()
        .map(|mut menu| menu.take().is_some())
        .unwrap_or(false);
    if !was_up {
        return false;
    }

    let (x, y) = window_origin(hwnd)
        .map(|rect| (rect.0, rect.1))
        .unwrap_or((0, 0));
    let dpi = dpi_at(x, y);
    let side = logical_px(dpi, PIN_BUBBLE_PIXELS).max(16);
    release_pin_capture(hwnd);
    let _ = MoveWindow(hwnd, x, y, side, side, true);
    render_pin_bubble(hwnd);
    true
}

/// Put the bubble's menu away from the preview loop's own tick, where one is up: the
/// answer the hook's ask for a press that landed on another window gets.
///
/// The bubble is `WS_EX_NOACTIVATE` and is the foreground window only once a hand has
/// pressed it, and only the foreground window's capture carries a press that lands
/// elsewhere to it — so a press on the browser under the bubble, or on the desktop,
/// arrives at the bubble's own procedure not at all. The hook's poll is the one side
/// that reads every press, and this is its answer on the bubble's own road, the way
/// `take_pin_menu_dismiss` is the card's (see `pin_menu_dismiss_ask`).
pub(super) unsafe fn close_pin_bubble_menu() -> bool {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return false;
    }
    close_bubble_menu(HWND(hwnd as *mut _))
}

/// A press with the bubble's menu up, answering whether it was the menu's: a press on a row is
/// that row's own and the menu goes away with the answer, while a press anywhere else is a press
/// that has left the menu, which puts it away (see `pinned_menu_press`).
///
/// Nothing is held, because a menu is a thing a hand reads rather than a thing it carries: the
/// press is answered where it landed, and the panel is one paint away from being gone.
pub(super) unsafe fn bubble_menu_press(hwnd: HWND, x: i32, y: i32) -> bool {
    if !bubble_menu_is_open() {
        return false;
    }

    // The row the press landed on, read before the menu goes: the panel a press is answered
    // against is the panel that is drawn, and closing first is what takes the answer away.
    let row = BUBBLE_MENU.lock().ok().and_then(|menu| {
        let menu = menu.as_ref()?;
        let index = pin_chrome::menu_row_at(&menu.popup, x, y)?;
        bubble_menu_rows(current_media_type()).get(index).copied()
    });

    // The chrome rows above are built from the same list in the same order, so an index into the
    // panel is an index into the rows the menu holds.
    //
    // A rule is not a row a hand answers, so a press on one is nothing at all — and the menu
    // stays up, the way a press on a rule anywhere else does.
    if row == Some(BubbleMenuRow::Separator) {
        return true;
    }

    // The switch the `Pause on Minimize` row carries is turned over where it is pressed, and the
    // panel stays up with the mark it now carries — the way the card's own `Shuffle` row answers
    // (see `pin_menu::toggle_pin_audio_shuffle`). Nothing is asked of the pin: what a collapse
    // holds back is read by the tick that keeps a collapsed pin's playback in step, so a switch
    // turned over here is answered on that tick's next pass (see `settle_bubble_playback`).
    if row == Some(BubbleMenuRow::PauseOnMinimize) {
        toggle_pause_on_minimize(current_media_type());
        // The rows the panel is painted from are the ones the menu was opened with, so the check
        // it now carries is written into them here rather than being read again at the paint.
        let rows = bubble_menu_chrome_rows(current_media_type(), bubble_is_playing());
        if let Ok(mut menu) = BUBBLE_MENU.lock() {
            if let Some(menu) = menu.as_mut() {
                menu.rows = rows;
            }
        }
        render_pin_bubble(hwnd);
        return true;
    }

    close_bubble_menu(hwnd);
    // Every row but the taskbar's is the same row a pinned window's own menu carries, and is
    // answered the same way there (see `shared_menu_row_press`); the rule and the switch above
    // were answered where they were pressed, with the panel left standing.
    match row {
        Some(row) => {
            shared_menu_row_press(row, bubble_to_taskbar);
        }
        // A press outside both the rows and the circle is a press that has left the menu.
        None => {}
    }
    true
}

/// Ask what the bubble's menu is saying about the pointer, answering whether anything about it
/// changed — which is whether the bubble owes a repaint for it.
///
/// The pointer is read as a place on the screen and taken against the window's own box, because a
/// pointer that has left the bubble sends it nothing more, so a wash left up under one is a mark
/// on a hand that is gone (see `refresh_pin_menu`).
pub(super) unsafe fn bubble_menu_hover(hwnd: HWND) -> bool {
    let Some(cursor) = cursor_screen_point() else {
        return false;
    };
    let Some(origin) = window_origin(hwnd).map(|rect| (rect.0, rect.1)) else {
        return false;
    };
    let (x, y) = (cursor.0 - origin.0, cursor.1 - origin.1);

    let changed = BUBBLE_MENU
        .lock()
        .map(|mut menu| {
            let Some(menu) = menu.as_mut() else {
                return false;
            };
            let hover = pin_chrome::menu_row_at(&menu.popup, x, y);
            if hover == menu.hover {
                return false;
            }
            menu.hover = hover;
            true
        })
        .unwrap_or(false);

    if changed {
        render_pin_bubble(hwnd);
    }
    changed
}

/// Put the pin away into the taskbar rather than the bubble: the bubble goes, the stand-in comes
/// up, and the pin stays collapsed until it is brought back (see `show_pin_taskbar`).
///
/// The picture the hover is to hold is taken *before* the bubble goes, and the bubble is what it is
/// taken from: the pin's own windows are off the screen behind a bubble, the bubble is the pin's
/// picture of itself, and a frame taken after it would be a frame of nothing (see
/// `keep_the_pin_for_the_taskbar` and `pin_taskbar_mirror::mirror_source`).
pub(super) fn bubble_to_taskbar() {
    unsafe {
        keep_the_pin_for_the_taskbar();
    }
    hide_pin_bubble();
    unsafe {
        show_pin_taskbar();
    }
}

/// Collapse a pinned window into the place the configuration has it put: the round bubble it
/// leaves on the desktop, or the app's own stand-in in the taskbar (see `pin_minimize_to`).
///
/// This is what a minimize button and a `Pin Command → Minimize` ask for; a place chosen by
/// what asked for it goes through `collapse_pin_to` (see `PinCommand::MinimizeTo`).
pub(super) fn collapse_pin() {
    collapse_pin_to(pin_minimize_to());
}

/// Collapse a pinned window into a place of the asker's own choosing: the window and
/// everything standing in it come off the screen, and a small circle takes their place — or the
/// app is left standing in the taskbar with the pin off the screen instead.
///
/// A collapsed pin is still a pin — previews are still held back until it is restored or
/// closed, and what it is playing is held where it is by the two `Pin Mode → Pause on
/// Minimize` switches (see `PIN_COLLAPSED` and `settle_bubble_playback`).
///
/// Everything but the last step is the same either way — which is why only the last step of a
/// collapse differs between the two places (see `pin_minimize_to`).
pub(super) fn collapse_pin_to(target: PinMinimizeTo) {
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

    // Where the pin goes is the target it was collapsed into: the
    // round bubble stands on the desktop where the minimize button was, and the taskbar
    // stand-in leaves the app in the taskbar instead with the pin off the screen. Everything
    // above is the same either way — which is why only the last step of a collapse differs.
    unsafe {
        if target == PinMinimizeTo::Taskbar {
            // The button's hover keeps the pin's own picture, and a picture of a window that has been
            // taken off the screen is a picture of nothing: what the hover is to hold is taken while
            // the windows it is made of are still up (see `keep_the_pin_for_the_taskbar`).
            keep_the_pin_for_the_taskbar();
        }

        hide_pinned_windows();
        if target == PinMinimizeTo::Taskbar {
            show_pin_taskbar();
        } else {
            // A pin put away into the bubble is a pin on the desktop, and the button a pin on the
            // screen was shown in is not a button for it: what the stand-in was showing is of a window
            // that has gone, and a pin put away into the bubble is reached by the bubble rather than by
            // the taskbar. So the two forms of it go down together with the pin, and the taskbar is
            // left as it was before it came up (see `pin_taskbar`).
            hide_pin_taskbar();
            show_pin_bubble(anchor);

            // And the frame the bubble of a film is to show until its player is read: every player
            // a film here is drawn by is somebody else's process and holds no frame on this side, so
            // it is read out of the file at the second the pin's bar was showing, on a thread of its
            // own, and the bubble goes up on the theme's face until it lands (see
            // `grab_bubble_frame`). It is also the colour the ring is drawn in, and the picture the
            // bubble falls back to wherever the frames read beside the player are not there (see
            // `bubble_stream`).
            grab_bubble_frame(bubble_box(anchor));
        }
    }
}

/// Take the frame the bubble a collapsed film leaves is to show, and leave it in the slot above.
///
/// **A film a player of somebody else's is drawing holds no frame on this side**, which is the whole
/// of why this exists: that player draws in a window of its own and what this app has of the film is
/// a process and a second on a bar, so the one picture a bubble can show of it is a frame read out
/// of the file (see `grab_video_frame`). Every other kind has its picture already — a
/// picture, a page and an engine's document are decoded or drawn by this app and the bubble samples
/// what the pin is holding, a film the media engine plays keeps its frame in memory, and a sound is
/// drawn as a meter rather than as a picture (see `bubble_art`).
///
/// It is spawned rather than waited for because a grab is a process launch: the bubble is up before
/// the launch has finished, and what the frame changes is the picture inside the circle rather than
/// whether there is a bubble at all. The thread is deliberately short and unnamed — it spawns a
/// child, waits for it under a bound, and leaves one frame in a slot (see `BUBBLE_FRAME_READY`).
///
/// **The pin's lock is not held across the spawn and neither is the media's.** Everything the grab
/// needs — the kind, the file, the second the bar is showing and the size the bubble was made — is
/// read here, one ask at a time, and what the thread is handed is the answers rather than the pin.
///
/// It is asked for once per collapse, and the slot is emptied before the ask rather than when the
/// frame lands: a grab that never answers — no FFmpeg on the machine, a file it will not open —
/// leaves a bubble with nothing in it rather than one still showing a frame of the film before this
/// one.
pub(super) fn grab_bubble_frame(box_: ScreenRegion) {
    // A frame from the last collapse is a frame of another file: the slot is emptied here rather
    // than when the frame lands, so a grab that never answers leaves a bubble with nothing in it
    // rather than one showing the film before this one.
    if let Ok(mut frame) = BUBBLE_FRAME.lock() {
        *frame = None;
    }
    BUBBLE_FRAME_READY.store(false, Ordering::Release);

    if current_media_type() != Some(MediaType::Video) {
        return;
    }

    let Some((path, _)) = pinned_media_owner() else {
        return;
    };

    let at = pinned_playhead().unwrap_or(0.0);
    let side = bubble_side(box_);

    // **The suite starts no process at all**, and a grab is a process: what a test build answers is
    // the ask — the film, the second and the side — and the frame is left exactly as it stands (see
    // `mpv_ipc::deliver`, where the same rule is kept for the pipe).
    #[cfg(test)]
    if let Ok(mut noted) = NOTED_GRABS.lock() {
        noted.push((path, at, side));
        return;
    }

    std::thread::spawn(move || {
        if let Some(pixels) = grab_video_frame(&path, at, side) {
            if let Ok(mut frame) = BUBBLE_FRAME.lock() {
                *frame = Some(BubbleFrame { pixels, side });
            }
        }

        BUBBLE_FRAME_READY.store(true, Ordering::Release);
    });
}

/// Keep a frame read beside a film's player as the picture the bubble falls back to, in place of
/// the still the collapse grabbed.
///
/// **The picture a bubble falls back to is the newest frame of the film this app has**, and a pause
/// is a newer one than the grab: a film held behind a bubble is one whose read stopped on a frame
/// of the film at the second it stopped at, while the grab is a frame of the second the pin went
/// down. A bubble that fell back to the grab the moment the film was let go of again would show the
/// second the pin was collapsed at, for as long as the read it has just begun takes to land its
/// first frame (see `bubble_stream::settle`).
pub(super) fn keep_bubble_frame(pixels: &[u8], side: u32) {
    if let Ok(mut frame) = BUBBLE_FRAME.lock() {
        *frame = Some(BubbleFrame {
            pixels: pixels.to_vec(),
            side,
        });
    }
}

/// The films a still has been asked for, with the second and the side it was asked at, for a test
/// build (see `grab_bubble_frame`).
#[cfg(test)]
static NOTED_GRABS: Mutex<Vec<(PathBuf, f64, u32)>> = Mutex::new(Vec::new());

/// The grabs noted so far (see `NOTED_GRABS`).
#[cfg(test)]
pub(super) fn take_noted_grabs() -> Vec<(PathBuf, f64, u32)> {
    NOTED_GRABS
        .lock()
        .map(|mut noted| std::mem::take(&mut *noted))
        .unwrap_or_default()
}

/// One frame of a film a player of somebody else's is drawing, taken for the bubble the pin
/// collapses into.
///
/// The frame is *read* rather than copied, and that is the whole of the difference between this and
/// every other kind's picture: what a bubble shows of a film the media engine plays is the frame
/// already in memory, where this engine's picture is on the screen and nowhere else. So the file is
/// opened at the second the bar is showing, one frame is decoded, and it comes back scaled to cover
/// a `side × side` square and cropped to the middle of it — the same shape, and from the same
/// corner, that every other bubble's picture is cut to (see `bubble_thumbnail`).
///
/// The output is raw BGRA on stdout, which is what the preview's own frame is in: a picture is not
/// encoded and decoded again for a process boundary when the bytes are what is wanted either side
/// of it. `-pix_fmt bgra` is asked for by name rather than left to the file, because what the file
/// holds is a codec's own arrangement of the colour and this side wants the one the compositor
/// reads.
///
/// It is bounded at three seconds and it is a *bound* on the wait rather than a budget for it
/// (`video_probe::wait_bounded`): a grab that outlasts it is a process killed rather than a thread
/// held, and the answer is the same `None` a file FFmpeg cannot open gives — the bubble keeps the
/// face it was painted with. The child is hidden and adopted like every other process this app
/// starts, so a machine closed mid-grab leaves no `ffmpeg` behind.
///
/// **The frame is drained on a thread of its own, and the bound is what that thread is for.** A
/// child has to write the whole of its frame before it can leave, and a pipe that nobody is
/// reading takes only its own buffer's worth: a frame bigger than that — a bubble at a high display
/// scale is a few tens of kilobytes, which is at or past the buffer — blocks `ffmpeg` in the middle
/// of its write, so the wait below would expire on a grab that was never given the chance to
/// finish. Reading as the child writes is what lets it reach the end of its one frame and exit; the
/// answer is the first `side × side × 4` bytes of what was read.
pub(super) fn grab_video_frame(path: &Path, at: f64, side: u32) -> Option<Vec<u8>> {
    let expected = side as usize * side as usize * 4;
    if expected == 0 {
        return None;
    }

    let mut command = engine_processes::hidden_command("ffmpeg");
    command
        .arg("-v")
        .arg("error")
        .arg("-nostdin")
        .arg("-ss")
        .arg(format!("{at:.3}"))
        .arg("-i")
        .arg(path)
        .arg("-frames:v")
        .arg("1")
        .arg("-vf")
        .arg(format!(
            "scale={side}:{side}:force_original_aspect_ratio=increase,crop={side}:{side}"
        ))
        .arg("-f")
        .arg("rawvideo")
        .arg("-pix_fmt")
        .arg("bgra")
        .arg("-")
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::null());

    let mut child = command.spawn().ok()?;
    let _ = engine_processes::adopt(child.id());

    // The frame is read as the child writes it, on a thread of its own: the child cannot exit
    // until the whole of it is out, and a pipe read only after the wait would leave the child
    // blocked mid-write on any frame larger than the pipe's own buffer (see the note above).
    let mut stdout = child.stdout.take()?;
    let reader = std::thread::spawn(move || {
        use std::io::Read;

        let mut bytes = Vec::new();
        let _ = stdout.read_to_end(&mut bytes);
        bytes
    });

    let landed = video_probe::wait_bounded(child, Duration::from_secs(3));
    let bytes = reader.join().unwrap_or_default();

    // A grab that ran past its bound answered nothing, whatever the reader had taken by the time
    // the child was killed.
    landed?;
    (bytes.len() >= expected).then(|| bytes[..expected].to_vec())
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
    // Whether the pin was put away into the taskbar is read before anything is hidden, because
    // hiding the stand-in is what takes the answer away. A pin restored from the taskbar goes
    // back exactly where it stood: there is no bubble on the desktop to place it beside, and the
    // pointer the button was clicked with is on the taskbar rather than on the work area.
    let from_taskbar = pin_taskbar_is_showing();

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
        // The playback goes back to being the switches' rather than the hand's: a pin up again is
        // a window a park would be written for, and a claim carried up out of the bubble would be
        // one the next collapse is answered with (see `PinTransport::drag_held`).
        pin.bubble_run = false;

        // The window goes back up where a preview of it would have been put at the pointer the
        // bubble was clicked with, rather than where it stood when it was collapsed: the bubble is
        // the bit of the pin the hand is on, and what the hand gets back is a window placed beside
        // it the way this app places everything else (see `placed_pin_box`).
        if !from_taskbar {
            if let Some(window) = placed_pin_box(pin, &DESKTOPS) {
                pin.content =
                    content_box_of(window, pin.dpi, pin.transport_bar, pin.overlay, pin.caption);
            }
        }
    }

    hide_minimized_pin();

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
/// — each kind behind the switch of its own, a film behind the one the video's own `Pin Mode →
/// Pause on Minimize` carries and a sound behind the audio's — and put it back the way it was
/// when the pin comes up again.
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
    let (collapsed, parked, taken_over) = {
        let Some(pinned) = pin_state() else {
            return;
        };
        let Some(pin) = pinned.pin() else {
            return;
        };

        (pin.collapsed, pin.bubble_pause, pin.bubble_run)
    };

    if collapsed {
        // What the switches parked is not theirs to hold a second time: a press on the bubble
        // gave it back, and the pin is down because a *hold* is a thing the user asked for rather
        // than a park the collapse made. The switches read the other way round too — one thrown
        // while the pin is a bubble is answered on the next tick — so the claim is written down
        // where the park was and stands until the pin is up again (see `toggle_pinned_by_key`).
        if parked.is_some() || taken_over {
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
            } else if video_player::is_playing() && !video_player::is_paused() {
                // What is not going is not parked, which is the same rule the sound's own key
                // reads by: a sound the engine is already holding is one a restore must put back
                // the way it found it, and a park of it would be a sound played again by the
                // press that brought its window back (see `toggle_pinned_audio`).
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

    // The cached handle alone is not the player: it is stale where the window was recreated and
    // zero where it was never found, and mpv may hold more than one window — so every visible
    // window of the player's process goes too, or a collapse leaves a chrome-less film on screen.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid != 0 {
        let _ = EnumWindows(Some(hide_player_windows_callback), LPARAM(pid as isize));
    }

    if webview_preview::is_showing() {
        webview_preview::hide();
    }
}

/// Hide every visible window belonging to `pid`: what `hide_pinned_windows` enumerates with.
unsafe extern "system" fn hide_player_windows_callback(
    hwnd: HWND,
    lparam: LPARAM,
) -> windows::Win32::Foundation::BOOL {
    let pid = lparam.0 as u32;
    let mut window_pid: u32 = 0;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));
    if window_pid == pid && IsWindowVisible(hwnd).as_bool() {
        let _ = ShowWindow(hwnd, SW_HIDE);
    }
    windows::Win32::Foundation::BOOL(1)
}

/// Take the round bubble down, if one is up.
pub(super) fn hide_pin_bubble() {
    // No bubble carries a menu: a pin put back up or taken down has no menu for one to be shown
    // in, and the next bubble starts without one.
    if let Ok(mut menu) = BUBBLE_MENU.lock() {
        *menu = None;
    }
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
    // A bubble that goes up goes up without a menu, whatever the one before it held.
    if let Ok(mut menu) = BUBBLE_MENU.lock() {
        *menu = None;
    }
    let (x, y, right, _) = bubble_box(anchor);
    let side = (right - x).max(1);

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

/// The box the round bubble a collapsed pin leaves stands in: the square that takes the place of
/// the button the hand went for, on the display that button is on, clamped to that display's own
/// work area so that a bubble at an edge comes back up inside it (see `show_pin_bubble`).
///
/// It is one function rather than a computation inside the window's own placement because the
/// frame a film's player is decoded into is cut to the very same square: a box that disagreed with
/// the bubble's would be a picture whose edges the ring does not reach (see `bubble_stream`).
pub(super) fn bubble_box(anchor: ScreenRegion) -> ScreenRegion {
    let dpi = dpi_at(anchor.0, anchor.1);
    let side = logical_px(dpi, PIN_BUBBLE_PIXELS).max(16);

    let centre_x = (anchor.0 + anchor.2) / 2;
    let centre_y = (anchor.1 + anchor.3) / 2;
    let (x, y) = (centre_x - side / 2, centre_y - side / 2);
    let clamped = clamp_pinned_box((x, y, x + side, y + side), dpi, &DESKTOPS);
    (clamped.0, clamped.1, clamped.0 + side, clamped.1 + side)
}

/// The side a bubble is drawn at on the display `box_` stands on: the circle's own, whatever the
/// square it is painted in.
///
/// It is one function because two things are made at that size — the still a bubble is ringed in and
/// falls back to, and the frames a film's player is decoded beside (see `grab_bubble_frame` and
/// `bubble_stream`) — and a pair of them that disagreed would be a picture of one size painted into a
/// circle of another.
pub(super) fn bubble_side(box_: ScreenRegion) -> u32 {
    logical_px(dpi_at(box_.0, box_.1), PIN_BUBBLE_PIXELS).max(16) as u32
}

/// The box the bubble's own window is standing in, where one is **up**: the square the circle is
/// painted in.
///
/// It is read off the window rather than remembered, because a bubble the hand has carried is a
/// bubble standing somewhere else — and the frames read beside a film have to follow it (see
/// `bubble_stream`).
/// **The circle's own side and not the window's height**, because the window grows downwards while
/// its menu is up (see `open_bubble_menu`).
///
/// **A bubble that is not on screen has no box**, which is the whole of what a pin put away into the
/// taskbar is: what a stream would do with the box it still has is decode a film for a circle nobody
/// is looking at.
pub(super) fn bubble_box_on_screen() -> Option<ScreenRegion> {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return None;
    }

    let hwnd = HWND(hwnd as *mut _);
    if !unsafe { IsWindowVisible(hwnd).as_bool() } {
        return None;
    }

    bubble_box_of(hwnd)
}

/// The box the circle of `hwnd` is standing in (see `bubble_box_on_screen`).
fn bubble_box_of(hwnd: HWND) -> Option<ScreenRegion> {
    let (x, y, ..) = window_origin(hwnd)?;
    let side = logical_px(dpi_at(x, y), PIN_BUBBLE_PIXELS).max(16);
    Some((x, y, x + side, y + side))
}

/// On the screen as a bubble, where it is: the window the pin is standing in while it is put away
/// into the round bubble, or nothing while no bubble is up.
///
/// The question is the mirror's rather than the pin's — what a taskbar hover is a picture of — and
/// it is asked of the same two facts `bubble_box_on_screen` is: a handle that was never made is not
/// a window, and a bubble that is hidden is not on the screen to be mirrored (see
/// `pin_taskbar_mirror`).
pub(super) fn showing_bubble_window() -> Option<HWND> {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return None;
    }

    let hwnd = HWND(hwnd as *mut _);
    unsafe { IsWindowVisible(hwnd).as_bool().then_some(hwnd) }
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

/// Paint the bubble: the round window a collapsed pin leaves, with what its media has to show
/// inside the circle where it has anything (see `pin_chrome::paint_bubble`).
///
/// The bars of a sound are drawn from the bands of this app's own sound, read where a bubble is
/// painted and not kept anywhere — and the phase the rows of a *level* are drawn at is taken from
/// the wall clock rather than kept beside the bubble, which is what keeps that wave moving while
/// nothing else about the bubble changes (see `BubbleArtOwned::Level`). Every other kind's art is a
/// picture and does not care what the phase says.
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

    // The circle keeps its own size whatever the window has grown into: the menu below it owns
    // the extra room, and a circle painted across the whole of it would sit in the middle of its
    // own window rather than where the bubble stood.
    let side = (logical_px(dpi_at(rect.left, rect.top), PIN_BUBBLE_PIXELS).max(16) as u32)
        .min(width)
        .min(height)
        .max(1);

    let (art, mark) = bubble_art(side);
    let phase = SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|since| since.as_secs_f32() * BUBBLE_PHASE_RADIANS_PER_SECOND)
        .unwrap_or(0.0);

    // What the bars are drawn from, which is the one thing the bubble's own art cannot say: seven
    // band levels where this machine reads them, the peak through the shape a level has always been
    // drawn as where it does not, and nothing where there is no sound being made at all.
    //
    // **A film whose sound is playing gets them too, over its own frame**, and the guard for that is
    // the play mark rather than the kind: a picture with a play mark is a film that is playing, and
    // one with a picture mark is a page or a still. A film the app plays at its own default volume
    // of nothing is a film making no sound, which is a row of no bars and its frame left exactly as
    // it is — the app never turns a burst of music on to make a meter move (see `DEFAULT_VIDEO_VOLUME`
    // and `audio_meter::output_bands`).
    let bars = match &art {
        BubbleArtOwned::Bands(bands) => *bands,
        BubbleArtOwned::Level(level) => pin_chrome::peak_bars(*level, phase),
        // A film's bars are drawn over its own picture — a frame this app painted — and only while
        // it is playing: a picture with a play mark is a film that is playing, and one with a picture
        // mark is a page or a still.
        BubbleArtOwned::Picture(..) if mark == pin_chrome::BubbleMark::Play => {
            audio_meter::output_bands().unwrap_or([0.0; pin_chrome::BUBBLE_BARS])
        }
        _ => [0.0; pin_chrome::BUBBLE_BARS],
    };

    // The owned art is what outlives the call that made it — a frame read out of a file, a
    // thumbnail cut down to the bubble's own size — and the chrome paints from a picture it is
    // lent, so this is the one place the two meet (see `BubbleArtOwned`).
    let art = match &art {
        BubbleArtOwned::Picture(pixels, art_width, art_height, ring) => {
            pin_chrome::BubbleArt::Picture {
                pixels,
                width: *art_width,
                height: *art_height,
                ring: *ring,
            }
        }
        _ => pin_chrome::BubbleArt::None,
    };

    out.fill(0);

    // The circle is painted at its own size and carried into the top of the window: the chrome
    // paints a circle that fills whatever buffer it is handed, so handing it the whole window
    // would centre it over the menu (see `paint_bubble`).
    let mut circle = vec![0u8; side as usize * side as usize * 4];
    pin_chrome::paint_bubble(&mut circle, side, side, &palette, art, bars, mark);
    for (row, chunk) in circle.chunks(side as usize * 4).enumerate() {
        let start = row * width as usize * 4;
        out[start..start + chunk.len()].copy_from_slice(chunk);
    }

    // The menu, below the circle: the theme's own panel, washed under the row the pointer is on,
    // the way a card's menu is (see `paint_menu_popup`).
    if let Ok(menu) = BUBBLE_MENU.lock() {
        if let Some(menu) = menu.as_ref() {
            // A surface of the window's own size, because the rows' labels are measured and drawn
            // through GDI and GDI needs a device context of its own — kept between paints and
            // blanked before each use, the way the pin's own menu surface is (see `MENU_SURFACE`).
            let wanted = (width.max(1), height.max(1));
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

                let blanked = unsafe {
                    std::slice::from_raw_parts_mut(
                        surface.bits(),
                        surface.width as usize * surface.height as usize * 4,
                    )
                };
                blanked.fill(0);

                pin_chrome::paint_menu_popup(
                    out,
                    width as i32,
                    &palette,
                    &menu.popup,
                    &menu.rows,
                    menu.hover,
                    surface,
                    menu.dpi as f32 / 96.0,
                );
            });
        }
    }

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

/// Paint the bubble again, if one is up.
///
/// The bubble is a window of its own with a surface of its own, so a repaint is the whole circle:
/// the loop asks for one when the playback the bubble stands for has turned over, when a frame
/// grabbed for it has landed, and once a frame while its bars animate (see the bubble repaint in
/// `run_preview_window`). Nothing is painted where no bubble is up, which is every pin that is
/// showing its window.
pub(super) unsafe fn repaint_pin_bubble() {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd != 0 {
        render_pin_bubble(HWND(hwnd as *mut _));
    }
}

/// What the bubble is drawn with, owned: the chrome paints from a picture it is lent, and this is
/// the side that makes one.
///
/// The two are deliberately not the same type. What the chrome takes is a picture somebody else is
/// holding — a frame in the media, alive for the length of a paint — and what a bubble *has* is a
/// picture made for it: a frame read out of a file by a process, or a thumbnail cut down to the
/// bubble's own size. One of those outlives the call that asked for it and the other does not, and
/// the difference is the whole of what the two types are (see `bubble_art`).
pub(super) enum BubbleArtOwned {
    /// A picture of `side × side` BGRA, and the side it was cut to, together with the colour the
    /// bubble's ring is drawn in where the caller has one to give — which is the still's own average
    /// for the frames of a film read beside its player, so that the ring over a moving picture is one
    /// colour rather than the average of every frame that passes. A still passes `None` and is
    /// ringed in its own colour, which is the same thing it would have been given here.
    Picture(Vec<u8>, u32, u32, Option<[u8; 3]>),
    /// The seven bands of this app's own sound, 0.0..=1.0 each (see `audio_meter::output_bands`).
    Bands([f32; pin_chrome::BUBBLE_BARS]),
    /// The level the sound on screen is being heard at, 0.0..=1.0, for the machine whose bands
    /// will not be read (see `audio_meter::output_peak`).
    Level(f32),
    /// Nothing to show inside the circle: the face, and the mark on it.
    None,
}

/// What the bubble is drawn with: the picture the whole pin has to show inside the circle, the
/// bands a sound is doing, the level it is being heard at, or nothing — together with the mark its
/// kind carries.
///
/// The kind decides which, and each arm answers the question its own way: a film a player of
/// somebody else's is drawing shows the frame read out of the file beside that player where one has
/// landed and the frame grabbed for it where none has, a film the media engine plays shows the frame
/// the engine is holding, a sound shows the bands of this app's own sound — or the mixer's own peak
/// through the shape a level has always been drawn as, where this machine will not read bands at
/// all. A sound is answered **while it is playing** and not before: a histogram is a picture of a
/// sound being made, and a paused sound has none to draw, so it is answered with the pause glyph
/// instead. Everything else whose kind *is* a picture shows the same thumbnail a page's first
/// plate is cut to.
///
/// **The playback is asked before the media is locked, and never while it is held.** The question
/// is answered out of the media itself — a sound's player process is a field of it, and a film's
/// transport is the pin's — so asking it under this lock would be this thread asking for a lock it
/// already holds, which is not a wait but a stop. It is also why the answer cannot be asked again
/// further down: everything below runs with the media in hand (see `bubble_is_playing`).
pub(super) fn bubble_art(side: u32) -> (BubbleArtOwned, pin_chrome::BubbleMark) {
    let playing = bubble_is_playing();

    let Ok(media) = CURRENT_MEDIA.lock() else {
        return (BubbleArtOwned::None, pin_chrome::BubbleMark::Picture);
    };
    let Some(media) = media.as_ref() else {
        return (BubbleArtOwned::None, pin_chrome::BubbleMark::Picture);
    };

    let mark = media.media_type.bubble_mark(playing);
    let art = match media.media_type {
        MediaType::Video => {
            let frame = BUBBLE_FRAME.lock().ok().and_then(|frame| frame.clone());

            // **A film a player of somebody else's is drawing is painted by this app**, out of the
            // frames read beside that player: the film's own window is hidden with the pin and never
            // moved into the circle, because a window of somebody else's cannot be cut to a circle
            // and carried about without a fight with its own region, its own topmost band and its own
            // resizes (see `bubble_stream`). The still is kept either way: its colour is the ring's,
            // and it is what the bubble shows wherever there is no frame of the film to hold.
            if let Some((pixels, side)) = bubble_stream::latest_frame() {
                // The ring's colour is the still's and not the frame's: a hairline averaged from
                // every frame that passes is a hairline that changes twenty-five times a second, and
                // a ring is the file's colour rather than a reading of it.
                let ring = frame.and_then(|frame| pin_chrome::average_color(&frame.pixels));
                BubbleArtOwned::Picture(pixels, side, side, ring)
            } else {
                frame
                    .map(|frame| {
                        BubbleArtOwned::Picture(frame.pixels, frame.side, frame.side, None)
                    })
                    .unwrap_or(BubbleArtOwned::None)
            }
        }
        MediaType::NativeVideo => bubble_thumbnail(media, side)
            .map(|pixels| BubbleArtOwned::Picture(pixels, side, side, None))
            .unwrap_or(BubbleArtOwned::None),
        // The bands of this app's own sound where this machine reads them, and the mixer's own
        // peak where it will not: the same row of bars, drawn from the answer this machine has
        // (see `audio_meter` and `BubbleArtOwned`).
        MediaType::Audio if playing => match audio_meter::output_bands() {
            Some(bands) => BubbleArtOwned::Bands(bands),
            None => BubbleArtOwned::Level(audio_meter::output_peak().unwrap_or(0.0)),
        },
        kind if kind.has_bubble_picture() => bubble_thumbnail(media, side)
            .map(|pixels| BubbleArtOwned::Picture(pixels, side, side, None))
            .unwrap_or(BubbleArtOwned::None),
        _ => BubbleArtOwned::None,
    };

    (art, mark)
}

/// Whether the media a pin stands for is playing now, read off the pin a reader already has in
/// hand.
///
/// A film is the pin's own transport, asked the way its bar asks it (see `pin_is_playing`). A sound
/// is the card's own question, asked the way the card asks it: neither of the two players behind a
/// sound's card is a transport this app keeps, so the question is put to the players themselves — a
/// process of this app's is one this app started and can see, and the engine's own is a session in
/// this process that is read whether it is *held* rather than whether it is playing, so a sound
/// sitting under a hold is not one a menu claims is running (see `a_sound_is_playing`). A sound
/// with neither behind it is a sound that is not playing rather than an answer that could not be
/// given.
///
/// **It is asked before the media's lock is taken and never during it**: a sound's arm reaches for
/// that lock itself (the player process is a field of the media), and a lock a thread already holds
/// is a stop rather than a wait (see `bubble_art`). It is asked of the pin rather than looked up in
/// the state the pin lives in for the same reason the other way round: the readers that paint a
/// menu are already under that state's lock, and one lock asked twice on one thread is a hang (see
/// `pinned_menu_paint`).
pub(super) fn pin_media_is_playing(pin: &PinnedPreview) -> bool {
    match current_media_type() {
        Some(MediaType::Video) | Some(MediaType::NativeVideo) => pin_is_playing(&pin.transport),
        Some(MediaType::Audio) => a_sound_is_playing(),
        _ => false,
    }
}

/// The same question for a reader with no pin in hand: nothing is playing where there is no pin.
pub(super) fn bubble_is_playing() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(pin_media_is_playing))
        .unwrap_or(false)
}

/// Whether the bars a bubble is drawn with are read from sound that is moving.
///
/// A sound's own bands move every few milliseconds while it plays. A film's bars are read from the
/// same capture of this app's own sound, drawn over the film's own frame, and they move the same
/// way — but only while the film is making any: a film the app plays at a level of nothing is begun
/// with no audio track at all, so its bars are a row of nothing and there is nothing for the bubble
/// to paint again for (see `render_pin_bubble` and `BUBBLE_REPAINT`).
///
/// The playback is not part of the question: whether a film is playing at all is the caller's own
/// ask, so that a bubble drawn while a film is held is not one this claims is making sound (see the
/// bubble repaint in `run_preview_window`).
pub(super) fn bubble_bars_are_moving() -> bool {
    match current_media_type() {
        Some(MediaType::Audio) => true,
        Some(MediaType::Video) | Some(MediaType::NativeVideo) => pinned_volume_level() > 0,
        _ => false,
    }
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

/// The bubble's own window procedure: a click focuses it and a second one puts the pinned window
/// back, a drag carries it anywhere on the desktop, SPACE holds a film or a sound while the bubble
/// holds the keyboard, and a right click opens its menu — the things a collapsed pin can be asked
/// for, on the one piece of it that is still on screen.
pub(super) unsafe extern "system" fn pin_bubble_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_LBUTTONDOWN => {
            // With the menu up a press is the menu's before it is anything else's, and the drag
            // below never starts under one (see `bubble_menu_press`).
            if bubble_menu_is_open() {
                if let (Some(cursor), Some(origin)) = (
                    cursor_screen_point(),
                    window_origin(hwnd).map(|rect| (rect.0, rect.1)),
                ) {
                    bubble_menu_press(hwnd, cursor.0 - origin.0, cursor.1 - origin.1);
                }
                return LRESULT(0);
            }
            // Whether the bubble already had the keyboard when the hand landed: the first click
            // focuses a bubble nothing is in, and only a click on one already focused puts the
            // window back up — so a focused bubble stays up for the SPACE below to hold.
            let was_focused = bubble_is_focused();
            BUBBLE_PRESS_FOCUSED.store(was_focused, Ordering::Release);
            if !was_focused {
                // One way: a bubble is shown non-activating and keeps refusing the keyboard until
                // the hand presses it, exactly as a pin does (see `pin_take_focus`).
                pin_set_focusable(hwnd, true);
                let _ = SetForegroundWindow(hwnd);
                let _ = SetFocus(hwnd);
            }

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

            // A film's window is hidden with the rest of the pin at the collapse, and this is the
            // press that carries the bubble about rather than the film: what the frames read beside
            // the player follow is the box the circle is standing in, which is what the drag's own
            // moves write (see `bubble_stream` and `drag_pin_bubble`).

            // And the drag is carried here rather than left to the preview loop's next tick: what
            // has hold of the bubble is the pointer, and the pointer does not wait for a tick of
            // anything (see `carry_bubble_drag`).
            carry_bubble_drag(hwnd);
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            // The wash follows the pointer while the menu is up; a drag never starts under one.
            if bubble_menu_is_open() {
                bubble_menu_hover(hwnd);
                LRESULT(0)
            } else {
                // Where the pointer is *now*, read from the cursor rather than from the message: the
                // coordinates a mouse message carries are measured in the window this drag moves.
                drag_pin_bubble(hwnd);
                LRESULT(0)
            }
        }
        WM_LBUTTONUP => {
            // A release with the menu up ends no drag: the press it answers was the menu's, and
            // the pointer the menu took stays taken until the menu goes.
            if bubble_menu_is_open() {
                drop_bubble_drag();
                return LRESULT(0);
            }
            let dragging = PIN_BUBBLE_DRAG
                .lock()
                .map(|mut drag| drag.take().is_some())
                .unwrap_or(false);

            if dragging {
                let _ = ReleaseCapture();
                let was_focused = BUBBLE_PRESS_FOCUSED.swap(false, Ordering::AcqRel);
                if bubble_click_restores(was_focused, PIN_BUBBLE_MOVED.load(Ordering::Acquire)) {
                    // A press that did not move is a click, and a click on a bubble the hand is
                    // already in is what puts the window back up. The focusing click stays a
                    // bubble, so the SPACE below has a window to hold.
                    ask_pin(PinCommand::Restore);
                }
            }
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_CAPTURECHANGED => {
            // A capture taken away is a drag that is over wherever the bubble stood: a window
            // that is not holding the pointer is not being carried by it, and the state is what
            // the drag's own wait is ended by (see `carry_bubble_drag`). A menu that held the
            // pointer goes with it, or it would stand past the hand that left it.
            drop_bubble_drag();
            close_bubble_menu(hwnd);
            LRESULT(0)
        }
        WM_RBUTTONUP => {
            // The bubble's own menu rather than an immediate close: what the hand chose is what
            // the loop answers, the way a caption button is (see `open_bubble_menu`).
            open_bubble_menu(hwnd);
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_LBUTTONDBLCLK => {
            ask_pin(PinCommand::Restore);
            LRESULT(0)
        }
        WM_KEYDOWN | WM_SYSKEYDOWN => {
            // A SPACE the bubble holds the keyboard for is the pause its bar offers, asked by the
            // keyboard rather than by a press on it — the same hold a pin's own window answers
            // (see `toggle_pinned_by_key`). Only a film or a sound has one, and a repeat is a hand
            // resting on the key rather than another hold.
            // Escape puts the menu away where it is up, and is nothing at all elsewhere.
            if wparam.0 as i32 == VK_ESCAPE.0 as i32 && bubble_menu_is_open() {
                close_bubble_menu(hwnd);
                LRESULT(0)
            } else if wparam.0 as i32 == VK_SPACE.0 as i32 {
                if lparam.0 & KEY_REPEAT == 0 && bubble_space_toggles(current_media_type()) {
                    ask_pin(PinCommand::TogglePlayback);
                }
                LRESULT(0)
            } else {
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
        }
        WM_CHAR => {
            // The character the SPACE above was already answered as: swallowed so the default
            // procedure does not beep for a key the bubble holds (see the pin's own `WM_CHAR`).
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

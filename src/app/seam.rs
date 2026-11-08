//! The facts the shell side and the preview-window side exchange, in one module.
//!
//! The tray, the explorer hook, the wheel watcher and the keyboard watcher are the main thread and the
//! threads Windows made for the two low-level hooks; the preview loop is a thread of its own drawing
//! layered windows. What crosses between them is small and deliberate, and it is of two kinds. Values:
//! a channel a `PreviewMessage` is sent on and an event a wait is woken by; the wheel notches a hook
//! counts and the loop drains; a pin's two window handles, whether it is Always On Top, whether its
//! find bar is up, and whether the preview is the surface the pointer stands on. Calls: whether
//! Explorer is in front, a tray toggle, a window class match. Only the first kind is here, because only
//! a value can be published where a caller cannot be. A call is behaviour, and moving one moves the
//! behaviour with it.
//!
//! They used to live wherever the code that first needed one happened to be — the channel and its
//! event in `tick`, the wheel counts in `wheel_input`, the pin's published half in `key_input` — so
//! each side reached into the other side's module for a fact neither owned: `ui::preview_window` into
//! `shell` for a notch it did not count, `shell` into `ui` for the message it posts. That a module
//! line is not a thread line is exactly what read wrong there, and it is why a hook callback could be
//! found holding a type the drawing thread had defined.
//!
//! Nothing here is new. It is the same atomics, the same orderings and the same timing, and no reader
//! and no writer moves off the thread it was on: the hook still counts a notch, the loop still swaps
//! it, the loop still publishes a handle, the hook still reads it, and neither takes a lock on the
//! hook path — the one `Mutex` is the channel's own, taken to send and never to read. Each accessor
//! stays beside its static for the reason the static is here at all: what is read and written together
//! is one fact, and a count kept in one file with its reader in another is a fact with two homes.
//!
//! The step after this one gives these the same shape across a process boundary rather than within
//! one, at which point what is in here is the whole of what a second process is told, and the modules
//! this came out of have no second thing to change.

use crate::shell::key_input::PIN_COMMAND_COUNT;
use crate::ui::preview_window::{AvoidRegion, PinPlanned, ScreenRegion};
use once_cell::sync::Lazy;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, AtomicI32, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc::{Receiver, Sender};
use std::sync::Mutex;
use windows::Win32::Foundation::HANDLE;
use windows::Win32::System::Threading::{CreateEventW, ResetEvent, SetEvent};
use windows::Win32::UI::WindowsAndMessaging::{
    MsgWaitForMultipleObjectsEx, PeekMessageW, MSG, MWMO_INPUTAVAILABLE, PM_NOREMOVE, QS_ALLINPUT,
};

#[derive(Clone)]
pub(crate) enum PreviewMessage {
    /// A preview of the file the pointer hovers, opened from the cursor it was
    /// hovered at.
    ///
    /// The region that comes with it is what the `Avoid` setting measured off the
    /// hovered item — its name, and the columns beside it at `Avoid Details` — which
    /// the placement is kept off at either way of avoiding, and whether that region is
    /// a column of the item's view, which decides *how* a placement is kept off it: a
    /// column is stepped off to one of its sides, an item's own text in whichever
    /// direction asks the least. See `AvoidRegion` and `avoiding_text`. It is `None`
    /// when the setting is off, when the view reported no text for the item, or when
    /// the walk found no item at the cursor.
    Show(PathBuf, i32, i32, Option<AvoidRegion>),
    /// A preview of the focused item, whose box comes with it, and the region the
    /// `Avoid` setting keeps it off — the item's own text, or the name alone, or the
    /// column the name sits in, depending on the way the setting is on. That region is
    /// where its preview is placed from as well as what it is kept off, the way a
    /// hovered item's is, for an item that draws its text as a row of its view: the
    /// last field says whether it does, read off the item's own text by the hook (see
    /// `explorer_hook::ItemText`). See `compute_keyboard_layout`.
    ShowKeyboard(PathBuf, i32, i32, i32, i32, Option<ScreenRegion>, bool),
    Hide,
    Refresh,
    /// A preview type was switched on or off. Only a preview whose own kind is
    /// now off is rebuilt, and it is rebuilt from the hover it came from, so it
    /// goes away on the spot rather than at the next pointer move.
    RefreshTypes,
    /// The pin key was pressed while a preview was on screen: what is up becomes a
    /// window of its own — captioned, movable, always on top — and the hover machinery
    /// stays quiet behind it until it is closed. The rect is the box the preview is
    /// already in, so nothing about the picture moves when it is pinned.
    ///
    /// The same message installs the file a pin that is up is shown next, and `arriving` says
    /// which of the two it is: the take-up out of a hover is the pin *arriving*, and it comes
    /// with no pin standing because that is what arriving is, while a file stepping onto a pin
    /// that is up is one the pin can have ended before the install landed — the watchdog ends a
    /// pin from a thread of its own (see `pin_lifecycle::the_pin_takes_the_file_up`).
    Pin {
        path: PathBuf,
        rect: ScreenRegion,
        arriving: bool,
    },
    /// The pinned window was given another box — maximized, restored, or resized by an
    /// edge — and the media in it is laid out again for that box. It is the same
    /// question a hover asks, with the box already answered.
    PinBox(ScreenRegion),
    /// The file the user picked while a preview was pinned, which the pin is asked to show
    /// instead of the one it has: the window keeps its place and its size, and the media in it
    /// becomes the new file's, fitted to the box the pin already has (see `Pin Mode → Update
    /// Preview`). The Explorer hook sends it — for a click, for a key, and — where the setting
    /// asks for it — for the pointer settling on another file (see `update_pinned_preview`).
    PinUpdate(PathBuf),
    /// A press landed on the listing behind the pin, before the file it picked is known: the
    /// click activates the listing, which puts it over the pin for as long as the shell takes to
    /// answer for the file, so the pin comes forward here rather than after that answer. `window`
    /// is the frame of the listing the press was made in, and it travels because the activation
    /// that lands after this loop reads the press is *that* window being put over the pin: the
    /// answer to it is asked of the window, not of whatever is in front by then (see
    /// `PinFrontOwed`). The answer is a [`PinUpdate`](Self::PinUpdate); a press that names no file
    /// is put back behind the listing by [`PinNoPick`](Self::PinNoPick) (see
    /// `press_pinned_preview`).
    PinPress {
        window: isize,
    },
    /// A press that was read as a pick named no file — a folder, a file no kind previews, the gap
    /// between the rows, the toolbar above them — and the pin goes back behind the window the press
    /// was made in, which is the frame `window` names: the foreground is still the window the user
    /// came from where the click has not been dispatched yet. What the press above raised is undone
    /// here; a pick that lands after it still raises the pin again (see
    /// `pin_press_was_not_a_pick`).
    PinNoPick {
        window: isize,
    },
    /// Whether the pin key is watched was changed in the tray, or previews themselves
    /// were turned off. A pin on screen is ended here rather than left as a window
    /// nothing would ever take down again.
    PinChanged,
    /// One of the tray's two `Volume` rows was moved, and the one it is: the film's half or
    /// the sound's. A pin already playing at a level of its own is told the new one here,
    /// on the thread its player belongs to, rather than being left at what it was taken up
    /// at (see `publish_tray_volume`).
    TrayVolume {
        audio: bool,
    },
    /// The tray's `Text → Font Size` row was picked, and the size it names. A pin already showing a
    /// page is put to that size on the thread the page is drawn on — the reader's own size on it
    /// given up, the way closing the pin gives it up — while a hover of a page is rebuilt from its
    /// own load at the new size and is owed nothing here (see `publish_text_font_size` and
    /// `redraw_pinned_page`).
    TrayTextSize {
        percent: u32,
    },
    /// The render tier is done with a document: a page is waiting in the cache,
    /// or there is no page. The generation is the hover that asked for it, so a
    /// render landing after the pointer has moved on is ignored — the page is
    /// still cached for the next hover either way.
    OfficeRenderReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// A video's probe is done: the geometry is waiting in the cache, or the answer is
    /// that there is none. The generation is the hover that was waiting on it, so a probe
    /// landing after the pointer has moved on is ignored — the answer is held for the next
    /// hover either way (see `video_probe`).
    VideoProbed {
        path: PathBuf,
        generation: u64,
    },
    /// A measure that reads a file is done: the box it answers with is held, or the answer is
    /// that the reader has none for the file — which is not a wait that can be answered, so it
    /// is the one answer a wait comes down on (see `measured_off_the_tick`).
    ///
    /// What was waiting on it is the hover that is on screen, and that is the whole of what
    /// this answer has to be matched to: a box is measured per version of the file, and a hover
    /// of a file whose version has changed is measured again rather than answered with the box
    /// of the version before it.
    MeasureProbed {
        path: PathBuf,
        size: Option<(u32, u32)>,
    },
    /// The planner is done with a question the pin asked — a walk along the folder, or the
    /// name of the program a file would open with. Neither question can be asked on the
    /// thread that draws the pin, so both are asked elsewhere and answered here (see
    /// `PinPlanner`).
    ///
    /// A walk is answered whatever it found: a folder that could not be read and a folder
    /// with nothing in it to step onto are both answers, not the absence of one, so the pin's
    /// arc comes down when they land rather than waiting out a bound (see `answer_pin_job`).
    PinAnswered(PinPlanned),
    /// The small subtitle files a film's own tracks were copied into are ready, sent the
    /// moment the extraction's one pass finishes (see `subtitle_files`).
    ///
    /// Only a pinned window acts on it: a pin has its film on screen already, so what the
    /// copy landing changes is what its *next* relaunch is drawn with — the copy is preloaded
    /// for that rather than the picture being changed hands for a subtitle (see
    /// `reload_pinned_subtitles`). A hover is deliberately not begun again under the pointer,
    /// and this message is not for one (see the note on the first hover in
    /// `video_launch::subtitle_filter`).
    VideoSubtitlesReady(PathBuf),
    /// The ImageMagick engine is done with a file: the picture it developed is in hand, or
    /// there is none — a file it cannot read is remembered as one it will not draw. The
    /// generation is the hover that was waiting on it, so a conversion landing after the
    /// pointer has moved on is ignored; the picture itself is held for the hover that asked
    /// either way (see `magick_render_is_due`).
    MagickReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// The PeaZip engine is done with a file: the archive's table of contents is in hand, read
    /// into the listing cache, or there is none — a file it cannot open is remembered as one it
    /// will not list. The generation is the hover that was waiting on it, so a listing landing
    /// after the pointer has moved on is ignored; the listing itself is held under the file's own
    /// key either way, so the next hover of it is a read rather than a launch (see
    /// `peazip_render_is_due`).
    PeazipReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
}

/// The event the preview channel's sends are announced on, so that a thread waiting for one
/// sleeps on the channel as well as on the message queue.
///
/// The wait the loop does was a wait on the queue alone, polled in eight-millisecond slices
/// because nothing would otherwise wake it: a thread with nothing on screen spent the whole of
/// `IDLE_WAIT_MS` waking to ask an empty channel whether it had anything, several times a
/// second, for the whole life of the process. Signalling this on the send is what lets the
/// wait block on the channel too — so a hover is answered when it is sent rather than at the
/// end of a slice, and an idle loop costs one kernel wait per tick instead of one per slice.
pub(crate) fn preview_channel_event() -> HANDLE {
    // Kept as the number a handle is rather than as the handle, for the reason
    // `NOACTIVATE_WAKE` is: a handle is a raw pointer, and this is shared across threads.
    static EVENT: Lazy<isize> = Lazy::new(|| {
        // Safety: an event with no name and no security descriptor, manual-reset, which is
        // what a wait that may be woken several times before it is read wants.
        unsafe { CreateEventW(None, true, false, None) }
            .map(|handle| handle.0 as isize)
            .unwrap_or_default()
    });

    HANDLE(*EVENT as *mut core::ffi::c_void)
}

/// Wake the preview loop where it stands, without a message: what a wheel notch over a pinned
/// picture, or a pin button press, asks for.
///
/// A notch or a press is counted on a hook's own thread — an atomic written there and nothing else —
/// and read by the loop on its next tick, so a loop waiting for a pin is a loop that will not see it
/// until its own wait runs out: `STATIC_PIN_WAIT_MS` rounded up to the system timer's own 15.6 ms
/// tick, which measured 62.2 ms over forty waits, worst 63.7 — four frames of a 144 Hz display, and
/// the whole of what makes a zoom start unevenly. The same wait signalled where the notch is counted
/// came back in 5.4 ms with the signal sent five in. A middle click shorter than that wait would
/// otherwise be lost outright.
///
/// It is the channel's own event rather than one of its own, so nothing new is waited on: the loop is
/// already armed on it, and a signal to an event nobody is waiting on is read by the next wait — it is
/// manual-reset, and reset after the wait rather than at the send (see `wait_preview_channel`).
pub(crate) fn wake_preview_loop() {
    // Safety: an event created once for the process, signalled: nothing here is dereferenced.
    unsafe {
        let _ = SetEvent(preview_channel_event());
    }
}

/// Send a message to the preview loop and wake it, so that a wait in progress ends at once.
pub(crate) fn send_preview(message: PreviewMessage) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(tx) = sender.as_ref() {
            let _ = tx.send(message);
            // Signalled after the send rather than before, so an armed wait can only be woken
            // by a message that is already in the channel to be read.
            unsafe {
                let _ = SetEvent(preview_channel_event());
            }
        }
    }
}

pub(crate) static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

/// Wait for a preview message or for window input, whichever comes first.
///
/// This is what makes a pinned window follow the hand at the pointer's own
/// pace rather than at the pace of the repaint: waiting on the message queue instead of on the
/// channel wakes the moment the pointer moves, and costs nothing extra when idle.
///
/// The wait is on two things at once: the message queue, and the channel's own event. The event
/// is what the slicing this used to do was standing in for. A wait that could not see the channel
/// had to wake on a timer to ask it, so `wait_ms` was cut into eight-millisecond slices and each
/// one was a kernel wait followed by a poll of an empty channel — which for a loop with nothing
/// on screen, waiting `IDLE_WAIT_MS`, was a hundred kernel transitions a second for the whole
/// life of the process, doing nothing. A send now signals the event, so the wait ends where the
/// message lands, and an idle loop costs one wait per tick.
///
/// The event is manual-reset and is reset after the wait rather than at the send, which is what
/// makes the arm-then-send race safe: a message that arrived between the poll above and the wait
/// below has already signalled it, so the wait returns at once instead of sleeping through a
/// message that is in the queue.
pub(crate) fn wait_preview_channel(
    rx: &Receiver<PreviewMessage>,
    wait_ms: u64,
) -> Option<PreviewMessage> {
    // A message sent just before the wait is answered without waiting at all.
    if let Ok(message) = rx.try_recv() {
        return Some(message);
    }

    let event = preview_channel_event();
    let handles = [event];
    unsafe {
        let _ = MsgWaitForMultipleObjectsEx(
            Some(&handles),
            wait_ms as u32,
            QS_ALLINPUT,
            MWMO_INPUTAVAILABLE,
        );
        let _ = ResetEvent(event);
    }

    // Woken by input, by a message, or by the timeout. Either way the channel is only ever
    // polled, never blocked on, so the queue is never left unpumped while a drag is going.
    if let Ok(message) = rx.try_recv() {
        return Some(message);
    }

    // Input alone is a reason to go round the loop now rather than to sit out the rest of
    // the tick: the queue holds the move the wait woke for.
    let mut msg = MSG::default();
    unsafe {
        if PeekMessageW(&mut msg, None, 0, 0, PM_NOREMOVE).as_bool() {
            return None;
        }
    }

    None
}

/// Wheel notches handed to a text preview. Signed: up is positive. Swapped to
/// zero by the preview thread, so a notch is counted once.
pub(crate) static TEXT_SCROLL_DELTA: AtomicI32 = AtomicI32::new(0);

/// Wheel notches handed to a pinned picture, in the same units and on the same terms as
/// the text preview's own count above: signed, swapped to zero by the reader, and added
/// to only where the pointer is inside the box the pin published.
pub(crate) static PIN_ZOOM_DELTA: AtomicI32 = AtomicI32::new(0);

/// Wheel notches handed to a pin's volume control, on the same terms as the counts
/// above: signed (up is louder), swapped to zero by the reader, and added to only
/// where the pointer is over the volume popup or the button that opens it.
pub(crate) static PIN_VOLUME_DELTA: AtomicI32 = AtomicI32::new(0);

/// Wheel notches handed to a pinned page's own text size, on the same terms again: a notch turned
/// over a page with `Ctrl` held is how large that page is drawn rather than where it is (see
/// `text_size_after_notches`).
pub(crate) static PIN_TEXT_SIZE_DELTA: AtomicI32 = AtomicI32::new(0);

/// Whether the preview on screen is the surface the pointer is standing on, published by the
/// preview thread once a tick.
///
/// It is what makes a pin *scrollable the moment it is raised*: a notch over a preview used to be
/// this app's only while the pin was the window the keyboard was in — Always On Top, or the window
/// the user had clicked into — so a pin raised over the listing they were working in took no notch
/// and had to be clicked once before it would scroll. The pointer is over the pin the whole time;
/// the keyboard is in the listing behind it, which is the point of the raise (see
/// `Pin Mode → Update Preview → Move To Top`).
///
/// It is published rather than asked here because a `WH_MOUSE_LL` callback may not walk the
/// desktop to find out what is in front — the same reason the pin's key owners are published by
/// the same tick (see `key_input::publish_pin_key_owner`). One tick of staleness is all it can be
/// wrong by, and a notch is turned by a hand that is not moving.
pub(crate) static PREVIEW_UNDER_POINTER: AtomicU32 = AtomicU32::new(0);

/// Whether the preview on screen is the surface the pointer is standing on, as the preview thread
/// last saw it (see `PREVIEW_UNDER_POINTER`).
pub(crate) fn preview_is_under_the_pointer() -> bool {
    PREVIEW_UNDER_POINTER.load(Ordering::Acquire) != 0
}

/// Publish what the pointer is on, from the tick that is allowed to ask the desktop (see
/// `PREVIEW_UNDER_POINTER`).
pub(crate) fn publish_preview_under_the_pointer(under: bool) {
    PREVIEW_UNDER_POINTER.store(u32::from(under), Ordering::Release);
}

/// Wheel notches a text preview has been given since this was last called, in
/// wheel units (120 to a notch).
pub(crate) fn take_text_scroll_delta() -> i32 {
    TEXT_SCROLL_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pinned page has been given for its own text size since this was last called, in
/// wheel units (120 to a notch). Signed, up being larger.
pub(crate) fn take_text_size_delta() -> i32 {
    PIN_TEXT_SIZE_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pinned picture has been given since this was last called, in wheel
/// units (120 to a notch). Signed, up being positive — which is a zoom in.
pub(crate) fn take_pin_zoom_delta() -> i32 {
    PIN_ZOOM_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pin's volume control has been given since this was last called, in
/// wheel units (120 to a notch). Signed, up being louder.
pub(crate) fn take_pin_volume_delta() -> i32 {
    PIN_VOLUME_DELTA.swap(0, Ordering::AcqRel)
}

/// Whether a volume notch is waiting to be read, so the preview loop can keep its frame cadence
/// for it rather than let the level land on its own slow wait (see `wait_before_the_next_tick`).
pub(crate) fn pin_volume_delta_pending() -> bool {
    PIN_VOLUME_DELTA.load(Ordering::Acquire) != 0
}

/// The two window handles a standing pin answers keys for, as raw pointers, or zero.
///
/// Published by the preview loop on every tick rather than asked for from inside the hook, which
/// may not take a lock and may not walk the desktop: two atomics and a compare are all a
/// `WH_KEYBOARD_LL` callback may afford, because every keystroke on the machine waits for it to
/// return and Windows removes a hook that is slow.
pub(crate) static PIN_KEYS_OWNER: AtomicU64 = AtomicU64::new(0);
pub(crate) static PIN_KEYS_VIDEO: AtomicU64 = AtomicU64::new(0);

/// Whether the standing pin is `Pin Mode → Always On Top`, published beside the two handles for
/// the same reason: a window that is above everything is the window the pointer is over, so the
/// notch a hand turns on it is the pin's whether or not the keyboard happens to be in it (see
/// `pin_owns_foreground`).
pub(crate) static PIN_KEYS_TOPMOST: AtomicBool = AtomicBool::new(false);

/// Tell the hook which windows a standing pin is answering for, or that there is no such pin.
///
/// Called by the preview loop on the tick it settles the pin's chrome, which is the same tick that
/// already knows whether a pin is up and which player it is showing. Passing zero for the video
/// handle is the normal case — a pin showing an image has no player to lose keys to.
pub(crate) fn publish_pin_key_owner(owner: u64, video: u64) {
    PIN_KEYS_OWNER.store(owner, Ordering::Release);
    PIN_KEYS_VIDEO.store(video, Ordering::Release);
}

/// Tell the hooks whether the standing pin is above every other window.
///
/// It is published rather than asked for where it is read for the reason the two handles above
/// are: a low-level callback may not take the configuration's lock, and every notch on the
/// machine waits for the answer. No pin up is the same answer as a pin that is not on top —
/// a hover preview stands in front of everything and needs no exemption (see
/// `pin_owns_foreground`).
pub(crate) fn publish_pin_always_on_top(always_on_top: bool) {
    PIN_KEYS_TOPMOST.store(always_on_top, Ordering::Release);
}

/// Pin-key commands seen since this was last called, one per press, in the order pressed.
///
/// Drained by the preview loop on the same tick that drains the pin key's own counter, so a
/// command travels the same path a press of the pin's own caption button does: it is queued, and
/// the loop acts on it (see `ask_pin`).
pub(crate) static PIN_KEY_PRESSES: [AtomicU32; PIN_COMMAND_COUNT] = [
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
];

/// Take the pin's own key commands, one count per command and per press since the last call.
pub(crate) fn take_pin_key_presses() -> Vec<u8> {
    PIN_KEY_PRESSES
        .iter()
        .enumerate()
        .flat_map(|(command, count)| {
            std::iter::repeat_n(command as u8, count.swap(0, Ordering::AcqRel) as usize)
        })
        .collect()
}

/// How many commands this hook has counted since the process started, which is the one thing a
/// reader of the two roads a pin's key can travel can tell apart: a key this hook counted never
/// became a message, and a key the pin's own window answered never came here (see `pin_owns_caret`).
#[cfg(test)]
pub(crate) static PIN_KEY_COUNTED: AtomicU64 = AtomicU64::new(0);

/// The number of commands this hook has counted (see `PIN_KEY_COUNTED`).
#[cfg(test)]
pub(crate) fn pin_key_counted() -> u64 {
    PIN_KEY_COUNTED.load(Ordering::Relaxed)
}

/// Whether a find bar is up over the standing pin's page, published beside the two handles and for
/// the same reason: `text_find::is_open` takes the media lock, and a `WH_KEYBOARD_LL` callback may
/// take none. It is also the one question that changes what this hook may do at all — a bar that is
/// up takes every key the pin would have answered, by passing it on.
pub(crate) static PIN_FIND_BAR_UP: AtomicBool = AtomicBool::new(false);

/// Tell the hook whether a find bar is up over the pin's page.
///
/// **A swallowed key never becomes a character.** `TranslateMessage` runs after every low-level hook
/// has returned, so the `WM_CHAR` a text field types from exists only for a key-down the hook let
/// through. That is why a bar is not another reader of this hook's commands: while it is up the
/// pin's keys are the field's, and the field's road is a message (see `text_find::key_edit_with`).
pub(crate) fn publish_find_bar_up(open: bool) {
    PIN_FIND_BAR_UP.store(open, Ordering::Release);
}

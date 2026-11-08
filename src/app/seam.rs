//! The facts the shell side and the preview-window side exchange, in one module.
//!
//! The tray, the explorer hook, the wheel watcher and the keyboard watcher are the main thread and the
//! threads Windows made for the two low-level hooks; the preview loop is a thread of its own drawing
//! layered windows. What crosses between them is small and deliberate, and it is of three kinds.
//! Values: a channel a `PreviewMessage` is sent on and an event a wait is woken by; the wheel notches a
//! hook counts and the loop drains; a pin's two window handles, the boxes a notch belongs to, whether a
//! pin is up, whether its find bar is, and whether the preview is the surface the pointer stands on.
//! Notices: the same, the other way — a setting the pin's own chrome asks the tray to turn, a shortcut
//! pressed at a preview, a line for the hook's trace — on a channel of its own.
//!
//! And the third kind, which is the reason this module is a doorway as much as a shelf: **every call one
//! side makes of the other is named from here**, even the ones that are calls today. A value can be
//! published where a caller cannot be, so the first two kinds are the whole of what a second process is
//! told; a call is not, and a call named from here is one whose body can become a message, a published
//! fact or a round trip without a call site to follow it. That is what the list of names at the end of
//! this module is for, and it is the one thing here that is not a fact either side owns: the folder walk
//! a pin steps through, the sort the listing is showing, the sizes a page can be drawn at.
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
//! it, the loop still publishes a handle, the hook still reads it. **No fact here is read through a
//! lock**: the three `Mutex`es are the two channels' own — a sender and two ends — taken to send a
//! notice and to drain one on the thread that already owns the drain, and never inside a low-level
//! callback, which is the one place a lock is forbidden. Each accessor stays beside its static for the
//! reason the static is here at all: what is read and written together is one fact, and a count kept in
//! one file with its reader in another is a fact with two homes.
//!
//! The step after this one gives these the same shape across a process boundary rather than within
//! one, at which point what is in here is the whole of what a second process is told, and the modules
//! this came out of have no second thing to change.

use crate::ui::preview_window::{AvoidRegion, PinPlanned, ScreenRegion};
use once_cell::sync::Lazy;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicI32, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc::{Receiver, Sender};
use std::sync::Mutex;
use windows::Win32::Foundation::{HANDLE, HWND};
use windows::Win32::System::Threading::{
    CreateEventW, GetCurrentProcessId, ResetEvent, SetEvent, WaitForSingleObject,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetWindowThreadProcessId, MsgWaitForMultipleObjectsEx, PeekMessageW, MSG, MWMO_INPUTAVAILABLE,
    PM_NOREMOVE, QS_ALLINPUT,
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
    /// Every probed geometry is given up: the engine choice moved, and a geometry read by FFprobe is
    /// not the geometry a media-engine preview wants, so the next hover probes again. Given up on the
    /// loop because the cache is that side's, and a row moved in the tray is not one that reaches a
    /// preview already up — the box it is being placed by has to stand until the next one begins (see
    /// `forget_video_geometry`).
    ForgetVideoGeometry,
    /// Only the extraction of `path` is kept running, every other one being dropped; `None` drops all
    /// of them. Sent by the tray's `Subtitles` switch, to drop an extraction running for a film no
    /// subtitle will be drawn from (see `keep_extraction_for`).
    KeepExtraction {
        path: Option<PathBuf>,
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

/// What the preview side tells the tray side, on a channel of its own.
///
/// Three things, and all three are notices rather than questions: a setting the pin's own chrome can
/// be asked to turn, a shortcut pressed at a preview, and a line for the hook's click trace. None of
/// them is answered, which is why they travel one way and why the sender never waits for one to be
/// read — a preview thread that blocked on the tray would be a picture that stopped while a menu was
/// open.
pub(crate) enum TrayMessage {
    /// The pin's **Always On Top** was turned over from the pin itself — its caption button, or the
    /// key bound to the same command. The setting is a row of the tray's with a file behind it, so the
    /// flip is asked for here and made there; the pin is told again by the message the flip sends
    /// (see `PreviewMessage::PinChanged`), which is the whole of what makes the answer arrive once the
    /// file says what it says rather than once the ask was made.
    TogglePinAlwaysOnTop,
    /// The preview shortcut was pressed over Explorer, and previews are switched off or on. It is
    /// counted by the low-level keyboard hook and was acted on by the preview loop only because that
    /// loop is where it happened to be drained; the switch is the tray's own, and the loop's business
    /// is not to hold a setting.
    TogglePreviewEnabled,
    /// One line for the hook's click trace, already built and not yet written (see
    /// `note_a_pin_click`). It is a `String` because the side that writes it also decides the file, the
    /// gate and the words around it, and because a line built on the asking side is a line built
    /// whether or not the run asked for any.
    PinClick(String),
}

/// The channel the two, and the one event its sends are announced on.
///
/// Both ends are made here rather than handed in by whoever starts first: the tray side's reader is a
/// loop that begins with the hook and the preview thread's writer begins with the window, and neither
/// is a place to be told about the other's start-up. It is a queue rather than an atomic because a
/// trace line does not fit one and because the order the asks are made in is the order they must be
/// answered in; nothing here is on a hook's own callback, which is the only place a lock is forbidden.
struct TrayChannel {
    sender: Mutex<Sender<TrayMessage>>,
    receiver: Mutex<Receiver<TrayMessage>>,
}

static TRAY_CHANNEL: Lazy<TrayChannel> = Lazy::new(|| {
    let (sender, receiver) = std::sync::mpsc::channel();
    TrayChannel {
        sender: Mutex::new(sender),
        receiver: Mutex::new(receiver),
    }
});

/// The event a tray message's send is announced on, so a thread waiting for one sleeps on the channel
/// as well as on its own clock.
///
/// It is an event of its own rather than the preview channel's: the two waits are on different threads
/// — the preview loop waits for a hover, the hook for a pin — and an event one of them resets after
/// its wait would be an event the other could be woken by and find nothing on the channel it read.
pub(crate) fn tray_channel_event() -> HANDLE {
    static EVENT: Lazy<isize> = Lazy::new(|| {
        // Safety: an event with no name and no security descriptor, manual-reset, which is what a wait
        // that may be woken several times before it is read wants.
        unsafe { CreateEventW(None, true, false, None) }
            .map(|handle| handle.0 as isize)
            .unwrap_or_default()
    });

    HANDLE(*EVENT as *mut core::ffi::c_void)
}

/// Tell the tray side something, and wake it where it stands.
///
/// The send is signalled after the message is in the channel rather than before, so a wait that is
/// already armed can only be woken by something there is to read.
pub(crate) fn send_tray(message: TrayMessage) {
    if let Ok(sender) = TRAY_CHANNEL.sender.lock() {
        let _ = sender.send(message);
        unsafe {
            let _ = SetEvent(tray_channel_event());
        }
    }
}

/// Everything the preview side has told the tray side since this was last called, in order.
pub(crate) fn take_tray_messages() -> Vec<TrayMessage> {
    let Ok(receiver) = TRAY_CHANNEL.receiver.lock() else {
        return Vec::new();
    };

    std::iter::from_fn(|| receiver.try_recv().ok()).collect()
}

/// Wait for a tray message, or for the same time out, whichever comes first.
///
/// This is the hook loop's sleep, and it is a wait on the channel's event rather than a plain sleep for
/// the reason the preview loop's own wait is: a message is answered where it lands rather than at the
/// end of whatever a loop had decided to sit out, and the interval is still exactly the interval when
/// nothing arrives. There is no message queue to wake for as well — the hook's thread pumps no window
/// messages, its hook callbacks arriving on the two threads that pump one — so this is one kernel wait
/// and nothing else.
pub(crate) fn wait_for_the_tray(wait_ms: u64) {
    let event = tray_channel_event();
    unsafe {
        let _ = WaitForSingleObject(event, wait_ms as u32);
        let _ = ResetEvent(event);
    }
}

/// Write one line about what a click in the listing did to a pinned window, from the preview side.
///
/// The line belongs to the hook's own trace file and is written by the loop that owns it, so what
/// travels is the line and not a call: the gate is read here, on the side that pays for building it,
/// and a run that asked for no trace sends nothing at all.
pub(crate) fn note_a_pin_click(line: String) {
    if std::env::var_os("RHP_HOOK_TRACE").is_some() || std::env::var_os("RHP_PIN_TRACE").is_some() {
        send_tray(TrayMessage::PinClick(line));
    }
}

/// Whether a preview is pinned: a window of its own with a caption, which stays until it is closed,
/// collapsed into its bubble included.
///
/// It is the fact the hook's whole tick is arranged around — while it is true the hover machinery is
/// quiet, the stray-player sweep is held, and the trigger key means nothing — and it used to be asked
/// of the preview thread through a lock, which is a hook that stops answering Explorer behind a preview
/// thread that has stopped turning. This is the published copy written beside the state it copies (see
/// `pin_window::PIN_UP`, where it is written).
pub(crate) static PIN_UP: AtomicBool = AtomicBool::new(false);

/// Whether a preview is pinned, answered without a lock and without waiting for one (see `PIN_UP`).
pub(crate) fn the_pin_is_up() -> bool {
    PIN_UP.load(Ordering::Acquire)
}

/// The window a collapsed pin leaves on screen: the circle the whole of a pin is hidden behind, which
/// carries the pin's own menu in the same window.
///
/// Published by the preview thread and read by the hook, for the same reason `PIN_UP` is: a press the
/// hook reads on the tick after a row of that menu was pressed landed on the panel, and the window under
/// the pointer is how that press is told from one on the file the panel covers (see
/// `window_state::bubble_window_is_at`).
pub(crate) static PIN_BUBBLE_HWND: AtomicIsize = AtomicIsize::new(0);

/// The window a standing pin's own menu is drawn in, published beside the bubble and read by the same
/// press (see `PIN_BUBBLE_HWND`).
pub(crate) static PIN_MENU_HWND: AtomicIsize = AtomicIsize::new(0);

/// The file a pin has been asked for and is not showing yet, as a fingerprint, or `0` for none.
///
/// **The file itself cannot be published here.** A `PathBuf` behind a lock is an allocation and a wait
/// on a path the hook asks every time it offers a pick, and the hook may have neither; so what crosses
/// is the number a hash of the path comes to, and the hook hashes the path it is holding and compares
/// it. What the answer is for is one press: a press on the file the pin is already on its way to is not
/// another file, and read as one it restarts the swap the pin is waiting on (see
/// `update_pinned_preview`).
///
/// shortcut: two spellings of one file — a case difference, a short name — hash apart, and the cost is a
/// press answered as a new pick, which is what it was before this was published at all. Upgrade to a
/// published path if a user ever meets it.
pub(crate) static PIN_FILE_IN_FLIGHT: AtomicU64 = AtomicU64::new(0);

/// Note the file a pin is now on its way to, or `None` where nothing is in flight any more.
pub(crate) fn publish_the_pin_file_in_flight(path: Option<&Path>) {
    PIN_FILE_IN_FLIGHT.store(path.map_or(0, fingerprint_of), Ordering::Release);
}

/// Whether `path` is the file the pin is already on its way to (see `PIN_FILE_IN_FLIGHT`).
pub(crate) fn pin_is_on_its_way_to(path: &Path) -> bool {
    PIN_FILE_IN_FLIGHT.load(Ordering::Acquire) == fingerprint_of(path)
}

/// The number a path is published as, asked the same way on both sides so that the two answers are the
/// same answer (see `PIN_FILE_IN_FLIGHT`).
///
/// A path that hashes to nothing is published as nothing in flight: the press that costs is a pick
/// offered twice, and the other way round — nothing in flight read as a file — is a pick dropped.
fn fingerprint_of(path: &Path) -> u64 {
    use std::hash::{Hash, Hasher};

    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    path.hash(&mut hasher);
    hasher.finish().max(1)
}

/// A box on the screen published whole, and read without a lock: the one fact here that does not fit an
/// atomic on its own.
///
/// Four edges are 128 bits and an atomic is 64, so the box is packed — four signed 16-bit fields in one
/// word — rather than published edge by edge. The reader is why: the wheel hook runs inside a
/// system-wide callback that every wheel message on the desktop waits for, and a box read half from one
/// publish and half from the next could claim a notch over a listing no preview covers, or refuse one
/// over a preview that does. The coordinate space is Windows' own, which is a 16-bit one in each axis —
/// the virtual screen is at most 32768 pixels across and its origin no further out than that — so a box
/// that somehow did not fit is published as no box at all, which costs a notch and never claims one.
pub(crate) struct PublishedRegion(AtomicU64);

/// What a `PublishedRegion` holds where it names no box. It decodes to `(-1, -1, -1, -1)`, an empty box
/// no point is inside, so even a box of its own that happened to pack to this answers the same way.
const NO_REGION: u64 = u64::MAX;

impl PublishedRegion {
    pub(crate) const fn new() -> Self {
        Self(AtomicU64::new(NO_REGION))
    }

    /// Publish `region` whole, or withdraw the box where there is none.
    pub(crate) fn publish(&self, region: Option<(i32, i32, i32, i32)>) {
        let packed = region
            .and_then(|(left, top, right, bottom)| {
                let edge = |value: i32| i16::try_from(value).ok().map(|value| value as u16 as u64);
                Some(edge(left)? | edge(top)? << 16 | edge(right)? << 32 | edge(bottom)? << 48)
            })
            .unwrap_or(NO_REGION);

        self.0.store(packed, Ordering::Release);
    }

    /// The box as it was last published, or `None` where there is none.
    pub(crate) fn get(&self) -> Option<(i32, i32, i32, i32)> {
        let packed = self.0.load(Ordering::Acquire);
        if packed == NO_REGION {
            return None;
        }

        let edge = |shift: u32| ((packed >> shift) as u16 as i16) as i32;
        Some((edge(0), edge(16), edge(32), edge(48)))
    }
}

/// The box a wheel notch over a pin's volume control belongs to: the open popup's panel and the button
/// it hangs from, or just the button while it is closed.
pub(crate) static PIN_VOLUME_REGION: PublishedRegion = PublishedRegion::new();

/// The box a wheel notch over a scrollable text preview belongs to: the journey from where the hover was
/// opened to the preview itself, or the pinned window's box where what is pinned is a page.
pub(crate) static TEXT_SCROLL_REGION: PublishedRegion = PublishedRegion::new();

/// The same box for a page that fits the one it is drawn in: there is no scroll there for Explorer to
/// lose, so a `Ctrl` and a wheel over one is the page's own size rather than its place (see
/// `text_size_after_notches`).
pub(crate) static TEXT_SIZE_REGION: PublishedRegion = PublishedRegion::new();

/// The box of a pinned picture, which a notch zooms.
pub(crate) static PIN_ZOOM_REGION: PublishedRegion = PublishedRegion::new();

/// The band a pinned model is drawn in, which a notch is handed to the model's own camera for.
pub(crate) static MODELS_BAND: PublishedRegion = PublishedRegion::new();

/// Notches turned over that band and not yet handed to the model: how many, and the screen point the last
/// one was turned at, packed as two 16-bit fields.
///
/// **A value rather than a message, and the one place here where that matters**: a notch is counted inside
/// `wheel_hook_proc`, the system-wide mouse callback every mouse message on the desktop waits for, so what
/// is written there is an atomic and nothing else — a channel send would be the first lock and the first
/// allocation on that path. The two are one fact in two atomics, and only the count is exact: the point is
/// the last one turned and is asked of the engine only to clamp the notch into its own client, so a point
/// paired with a notch turned a hair earlier answers the same way.
pub(crate) static MODELS_NOTCHES: AtomicI32 = AtomicI32::new(0);
pub(crate) static MODELS_NOTCH_POINT: AtomicU64 = AtomicU64::new(0);

/// Hand a notch turned over a pinned model's band to the model — f3d's wheel is its camera's zoom, so the
/// notch is the engine's and nothing of this app's (see `MODELS_BAND` and `f3d_player`).
///
/// The count is added and the loop woken; what the engine's window is and where its client is are asked on
/// the loop's own thread, where the answer may be looked up (see `take_the_models_notch`).
pub(crate) fn post_a_notch_to_the_model(x: i32, y: i32, notches: i32) {
    let (Ok(x), Ok(y)) = (i16::try_from(x), i16::try_from(y)) else {
        // A point that is not inside a screen this app can name is not a notch over anything of its own:
        // the box it was turned in was read off the same desktop, so this is a display that moved between
        // the two readings rather than a notch to post anywhere.
        return;
    };

    MODELS_NOTCH_POINT.store(x as u16 as u64 | (y as u16 as u64) << 16, Ordering::Release);
    MODELS_NOTCHES.fetch_add(notches, Ordering::AcqRel);
    wake_preview_loop();
}

/// The notches waiting to be handed to the model, and the point the last of them was turned at (see
/// `MODELS_NOTCHES`).
pub(crate) fn take_the_models_notch() -> Option<((i32, i32), i32)> {
    let notches = MODELS_NOTCHES.swap(0, Ordering::AcqRel);
    if notches == 0 {
        return None;
    }

    let point = MODELS_NOTCH_POINT.load(Ordering::Acquire);
    let edge = |shift: u32| ((point >> shift) as u16 as i16) as i32;
    Some(((edge(0), edge(16)), notches))
}

/// Pin-key presses seen since this was last drained, one per press of the key or of the mouse button it
/// is bound to.
///
/// Both low-level hooks count it — the keyboard hook for keys, and the mouse hook for `mouse3` through
/// `mouse5`, which never reach the keyboard — and the preview loop drains it once a tick. It is here
/// rather than beside the counter it is added to because its reader is the other side.
pub(crate) static PIN_PRESSES: AtomicU32 = AtomicU32::new(0);

/// Preview-shortcut presses seen over Explorer since this was last drained.
pub(crate) static TOGGLE_PRESSES: AtomicU32 = AtomicU32::new(0);

/// Pin-key presses since this was last called (see `PIN_PRESSES`).
pub(crate) fn take_presses() -> u32 {
    PIN_PRESSES.swap(0, Ordering::AcqRel)
}

/// Preview-shortcut presses since this was last called (see `TOGGLE_PRESSES`).
pub(crate) fn take_preview_toggle_presses() -> u32 {
    TOGGLE_PRESSES.swap(0, Ordering::AcqRel)
}

/// The pin key's main keys that are held, as a bit per chord slot.
///
/// It is what the preview side reads as "the hand is still on the press that took this pin up": the
/// activation Explorer dispatches for a middle press lands after this app's raise as often as before it,
/// so the band a raise gives back is held for the hand rather than for a duration (see
/// `pin_lifecycle::A_PRESS_BAND_IS_HELD_FOR`).
pub(crate) static PIN_MAIN_DOWN: AtomicU32 = AtomicU32::new(0);

/// Whether a key of the pin's own is down right now — a mouse button it is bound to included (see
/// `PIN_MAIN_DOWN`).
pub(crate) fn the_pin_key_is_down() -> bool {
    PIN_MAIN_DOWN.load(Ordering::Acquire) != 0
}

/// The keys a standing pin answers, as the commands they mean.
///
/// The table is here because the two sides have to agree about it and neither owns it: the hook reads it
/// to know which keys are the pin's before it may swallow one — `TranslateMessage` runs after every
/// low-level hook has returned, so a swallowed key never becomes a character — and the pin's own window
/// procedure reads it to answer the keys that reach it (see
/// `preview_window::pinned_key_message_command`). A key listed on one side and not the other is a key one
/// of them acts on and the other one throws away.
pub(crate) const PIN_KEY_COMMANDS: [(i32, u8); 12] = [
    (0x25, 0),  // VK_LEFT   -> Previous
    (0x27, 1),  // VK_RIGHT  -> Next
    (0x1B, 2),  // VK_ESCAPE -> Close
    (0x20, 3),  // VK_SPACE  -> TogglePlayback
    (0x54, 4),  // VK_T      -> NextSubtitle
    (0x26, 5),  // VK_UP     -> CaretUp
    (0x21, 6),  // VK_PRIOR  -> PageBack
    (0x28, 7),  // VK_DOWN   -> CaretDown
    (0x22, 8),  // VK_NEXT   -> PageOn
    (0x24, 9),  // VK_HOME   -> PageTop
    (0x23, 10), // VK_END    -> PageBottom
    (0x72, 11), // VK_F3     -> FindNext
];

/// How many commands `PIN_KEY_COMMANDS` names, which is how many counters are kept for them.
///
/// Named rather than written as the `12` of an array length so that a command added to the table above is
/// a command this count has to be told about, rather than one whose presses are counted into a counter
/// that was never reserved for it.
pub(crate) const PIN_COMMAND_COUNT: usize = 12;

/// Which command of `PIN_KEY_COMMANDS` a virtual key is listed under, or `None` for a key this app does
/// not answer.
pub(crate) fn pin_key_command(vk: i32) -> Option<u8> {
    PIN_KEY_COMMANDS
        .iter()
        .find(|(key, _)| *key == vk)
        .map(|(_, command)| *command)
}

/// Whether a message about a key a standing pin answers is one the pin's own window would act on.
///
/// The class of the message is part of the question rather than an accident of which arm of a window
/// procedure it arrived on: a pinned window answers a key pressed plainly and answers no chord at all
/// (`WM_SYSKEYDOWN` is swallowed there so that `Alt+F4` cannot close it and `Alt+Tab` cannot leave it),
/// so **the hook must not count a system key either**. It did, and the difference was a chord acted on for
/// a caret this app is in: `Ctrl`+Left walked the pin's folder and `Ctrl`+Space held its film, each of them
/// for a chord the pin's own window had just thrown away.
pub(crate) fn pin_key_message_acts(message: u32) -> bool {
    message == windows::Win32::UI::WindowsAndMessaging::WM_KEYDOWN
}

/// Whether a window belongs to a viewer this tray started: the hover's window, the window a pin stands
/// in, the host a film is drawn in, a bubble or the menu it carries, a taskbar stand-in, a page's source.
///
/// **A question about whose a window is rather than a list of the ones we made.** A list would be
/// published by the side that makes windows, and that side makes many of them, so a list is a leak every
/// time one is made and a set owned by the wrong side besides. The process the window was made in is the
/// one thing about it that does not change as the preview side changes what it draws.
///
/// It is asked of a press the hook reads, so that a click on a pinned window is read out of the listing
/// that window is standing on rather than taken as a click on something that is not a listing at all (see
/// `click_is_over_a_listing`).
pub(crate) fn window_is_a_viewers(window: isize) -> bool {
    if window == 0 {
        return false;
    }

    let mut pid = 0u32;
    unsafe {
        GetWindowThreadProcessId(HWND(window as *mut core::ffi::c_void), Some(&mut pid));
    }

    pid != 0 && pid == the_viewer_process()
}

/// The process the viewers are, and the whole of what is still a stub about them.
///
/// A second process is what the step after this one makes, and it is where this answer comes from the
/// pids this one started. Until then the viewer is this process, which is what the call this replaces
/// answered: `is_preview_window` compared a handle against the three windows the preview side could be
/// drawn in, and every one of those is a window of this process.
fn the_viewer_process() -> u32 {
    unsafe { GetCurrentProcessId() }
}

// What the two sides ask each other for by name, in one doorway rather than across the modules that
// happen to define them. A symbol here is a crossing: the tray side calls the preview side through this
// list, and the preview side calls the tray side through the rest of this module. What makes the step
// after this one a change of one module is that every one of these is named from here, so a body that
// becomes a message, a published fact or a round trip has no call site to follow it.
pub(crate) use crate::shell::explorer_hook::{view_sort_of, SortKey, ViewSort};
pub(crate) use crate::shell::tray::FONT_SIZE_CHOICES;
pub(crate) use crate::ui::preview_window::{
    cursor_preview_hover, hide_preview, kill_stray_video_process, monitor_dpi_from_point,
    note_engine_page_drag, pin_button_covers, pin_press_was_not_a_pick, pinned_path,
    pointer_item_box, pointer_item_holds, press_pinned_preview, preview_pointer_hold,
    preview_screen_rect, preview_stall_ms, preview_window_rect, publish_pin_media_press,
    publish_pointer_item_box, refresh_pin, refresh_preview, refresh_preview_types,
    refresh_render_html, request_pin_end, request_pin_menu_dismiss, show_preview,
    show_preview_keyboard, take_pin_handed_off, take_pin_menu_pressed, take_pin_resumed,
    trim_image_cache, trim_subtitle_cache, update_pinned_preview, PreviewCursorHover,
};

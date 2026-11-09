//! Files dropped onto a pinned window: what a drop of them is, and the one message it becomes.
//!
//! A pin is a window of its own, and the file the user wants to look at may be one they are carrying —
//! dragged out of Explorer, out of a browser, off another window of this machine. This module is the
//! whole of what that app does about such a drop: the two windows it can land on say they take files,
//! and every drop is read here, on the thread that draws the pin, as the one thing it is — a file the
//! pin is asked to show.
//!
//! **The rest of the road is the pin's own.** What comes out rests in the slot a pick in the listing
//! rests in (`PIN_DROP`), so the loop takes it with the same take, plans it with the same plan and
//! installs it with the same install: the box it is measured for, the engine chosen for its kind, the
//! load, the paint, the take-up. And so does the walk, which is asked for the folder the *file* is in
//! rather than the one the pin came from — the next press of a caption's next is a step of the folder
//! that was dropped out of (see `pin_navigation::list_for`).
//!
//! **A drop is not a pick behind a window, and the setting that gates those does not gate this.** The
//! file is handed to the window the user dropped it on, and nobody else's hand put it there: `Pin Mode
//! → Update Preview` is the row that says whether a pick made in the listing behind a pin is followed,
//! and a drop answers for itself (see `pin_walk::take_pin_pick`, where a picked file's gate is read).
//!
//! **Nothing is done about a drop that cannot be shown, and nothing is said about it.** A folder, a
//! file of a kind no preview claims, a list that is not files — every one of them is a drop the pin
//! answers by not answering: no pick, no load, no walk, no paint, and the file it is showing is exactly
//! the file it was showing. That rule is asked *here* rather than where a pick is taken up, because a
//! pick's refusal is a walk's step over it (see `pin_step_off`) and a drop is nobody's step. What is
//! asked of the file is its name and not its content, because this runs on the thread that draws the
//! pin (see `the_pin_can_show`).

use super::pin_trace::note_pin_trace;
use super::*;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;

/// The one file a drop names that the pin is to be shown next, and how it got into the window.
///
/// It is a slot of the viewer's rather than a value passed along, because the two windows a drop lands
/// on are one thread away from the loop that shows a file: a window procedure answers a message and
/// returns, and the tick that plans a pick is the loop's own (see the read beside `pin_pick` in
/// `run_preview_window`).
pub(super) static PIN_DROP: Lazy<Mutex<Option<DroppedOn>>> = Lazy::new(|| Mutex::new(None));

/// **A dropped file is one the user named**, which is what the pick behind it reads for the one setting
/// an update is gated by, and what the window answers with once it is shown.
pub(super) struct DroppedOn {
    pub(super) path: PathBuf,
    /// Whether the window the file was dropped on was put away when it was dropped — a collapsed pin is
    /// a bubble, and a file dropped on it is the pin being asked to come back on that file (see
    /// `pin_was_put_away`).
    pub(super) put_away: bool,
}

/// Take the file a drop left for the loop, if there is one.
pub(super) fn take_dropped_on() -> Option<DroppedOn> {
    PIN_DROP.lock().ok().and_then(|mut slot| slot.take())
}

/// One drop, read: the file it names if the pin can show it and nothing at all if it cannot.
///
/// Called where the drop arrives rather than where the file is shown, so that the shell's own answer
/// for the list — `DragFinish`, without which the list it handed over is never given back — is given in
/// the same breath on every road, including the ones that change nothing (see `read_dropped_file`).
///
/// Safety: `drop` is the handle a `WM_DROPFILES` carried into this thread, which is owned here and
/// released once.
pub(super) unsafe fn take_drop(drop: HDROP) -> Option<PathBuf> {
    let named = read_dropped_file(drop);
    DragFinish(drop);

    let path = named?;
    if !the_pin_can_show(&path) {
        if pin_trace::asked() {
            note_pin_trace!(
                "drop names a file no kind previews, left alone: {}",
                plain_path(&path)
            );
        }
        return None;
    }

    if pin_trace::asked() {
        note_pin_trace!("drop names {}", plain_path(&path));
    }
    Some(path)
}

/// Whether the file a drop names is one a pin can be shown, asked off the name alone.
///
/// **The same question a pick is planned from**, so a file this answers yes to is one the plan behind
/// the pick answers with a box or a wait rather than with nothing: a name `formats::routing` claims has
/// a kind, and a kind is what every reader, engine and measurement below is chosen by (see
/// `pin_update_plan`).
///
/// **It is asked of the name and not of the file**, which is the difference between this and a hover's
/// question: `routing::kind_of` opens the file to settle the two names the video and text lists share,
/// and this runs on the thread that draws the pin and pumps its messages — a file on a share that is
/// away, a synchronising provider's placeholder, or a `.ts` of a few gigabytes would be the pin's
/// window standing still for as long as the read takes, to answer a question whose only use is whether
/// to say nothing about the drop. A name the lists carry is the whole of what is asked here, so nothing
/// is opened and nothing is read: a `.ts` is answered by the name's own claim, which is what the walk
/// scan behind a caption asks of every file in a folder for the same reason (see
/// `routing::kind_of_name` and [`crate::formats::routing::Asked::Name`]).
fn the_pin_can_show(path: &Path) -> bool {
    if !path.is_file() {
        // A folder is the one entry a drop may name that a kind claims nothing about, and a folder is
        // not this window's to show: the listing is the answer to a drop of one, and it is not here.
        return false;
    }

    CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::routing::kind_of_name(path, &config).is_some())
        .unwrap_or(false)
}

/// The first file one drop's list names, or nothing where the drop names no file at all.
///
/// **The first entry is the whole of the answer**, and the others are read by nobody: the hand put one
/// list on one window, and the file it dragged first is the file it meant. A first entry that is a
/// folder or a file nothing previews is a drop nothing is done about — not the second entry, and not
/// the third — because the pin is asked about what the user pointed at and not about what else was
/// carried with it.
///
/// Every entry has to be told its name's length before its name can be asked for: the Shell numbers a
/// path in the list by the number of characters it takes, and a buffer a path may be longer than is one
/// it writes past.
///
/// Safety: `drop` is the handle a `WM_DROPFILES` carried, held for as long as this is asked.
unsafe fn read_dropped_file(drop: HDROP) -> Option<PathBuf> {
    let length = DragQueryFileW(drop, 0, None);
    // `0xFFFFFFFF` is the Shell's answer for a handle that is not one, and `0` for a list that begins
    // with nothing: neither is a name to be read, and both are answered as no file.
    if length == 0 || length == u32::MAX {
        return None;
    }

    let mut units = vec![0u16; length as usize + 1];
    if DragQueryFileW(drop, 0, Some(&mut units)) == 0 {
        return None;
    }

    // Cut to the name it counted rather than kept to the buffer's length: the buffer is one longer
    // than the name, and its last unit is the null that says where the name ended.
    units.truncate(length as usize);
    let path = PathBuf::from(OsString::from_wide(&units));

    (!path.as_os_str().is_empty()).then_some(path)
}

/// Ask the window a drop landed on to answer for files, for as long as that window lives.
///
/// Asked once per window and asked again on every road that makes one, which is what makes it a call
/// rather than a style written at creation: a window is made once and kept, and a window of the pin's
/// own is made, hidden, and taken up again (see `pin_real_window_at` and `pin_bubble_window`).
///
/// A window that does not say it takes files is a window the Shell refuses the drop at — no cursor
/// change over it, no message to it, and the drop going to whatever is behind it — which is the whole of
/// what is wanted for the bubble and the pin: the drop lands on the window the user sees, and the file
/// is its own.
///
/// Safety: `hwnd` is this thread's own window, made and not yet taken apart.
pub(super) unsafe fn take_files_dropped_on(hwnd: HWND) {
    // A window nothing made is not one the Shell may be told about: asking it to take files is asking
    // a window that is not there, and a handle the loop has not got is the road a test's stand-in takes.
    if hwnd.0.is_null() {
        return;
    }

    DragAcceptFiles(hwnd, true);
}

//! A pinned PDF shown by the viewer the WebView2 runtime already carries.
//!
//! The hover's PDF road reads a raster of page one (`readers::pdf_preview`), and it stays that
//! road: a hover is a card the pointer is on, and one page of it is the whole of what a card can
//! say. A *pin* is the other thing — a document somebody works in — so a pinned PDF is handed to
//! Chromium instead: the runtime's own PDF viewer, pointed at the file itself, which is what makes
//! pages, continuous scroll, zoom, thumbnails, search and text selection Chromium's rather than
//! this app's (see ARCHITECTURE.md, "The two PDF roads").
//!
//! **Shape: one thread, one environment, one controller per viewer.** WebView2 wants the thread that
//! made a controller to keep retrieving messages for it, so every ask about a controller is made
//! here and nowhere else, and the rest of the app talks to this module rather than to a controller.
//! A viewer's window is a plain popup of this app's — no layered surface and no alpha, which is the
//! one visual difference from a picture pin and is deliberate: Chromium's pixels are the window's
//! own, and the Shell's DWM look comes for free. It is *owned* by the pin's window, so it stands
//! above it, is hidden and minimized with it, and needs no placement road of its own beyond the box
//! the pin's content band has.
//!
//! **Nothing here is the hover engine's.** `SHOWING`, `ENGINE_HWND`, `SHOWING_RECT` and
//! `webview_idle` describe the one hover page and are left alone: a PDF pin has its own refcount and
//! no TTL at all. The last pin closed closes the controller, closes the host window and ends the
//! thread; the next pin begins a fresh browser.
//!
//! **A viewer that never arrives shows nothing.** The window is created hidden and revealed only
//! when `NavigationCompleted` says the document loaded. What is on screen until then — and what
//! stays on screen for good where the engine or the navigation fails — is the static raster the pin
//! was already showing, so a pin is never a white window.
//!
//! **And it is revealed at a box that came with it.** The box an ask is made with is the band that
//! stood when the file was picked, and a pin taken up on another file moves its band after that: a
//! document that arrived at the box it was created with is a second window standing beside its pin
//! until the next box change snaps it in. So a viewer is `Arriving` until both the document has
//! arrived *and* a box has been published for its pin since it was created (see `Life::navigated`),
//! and the pin's roads keep publishing one while it is arriving (see `is_arriving_for`).
//!
//! **One pin holds one document.** A file asked for in a pin that was showing another takes the
//! standing viewer down before the new one is made, so a walk of two files can never leave two
//! windows in one band (see `Seat::replaced_by`).
//!
//! **And the document holds the caret, because the wheel is the document's.** A viewer is a window
//! the keyboard is not in until something puts it there, and a wheel turned over a window that is
//! not the one the caret is in is delivered elsewhere — so the controller is asked for the focus at
//! the reveal and at every resume (`Host::take_the_caret`), and from then on a notch over the page
//! scrolls it rather than being posted here and there by the seam. The seam's posting is what a
//! desktop without that focus falls back on, and the pin publishes no wheel box of its own for a
//! document at all (see ARCHITECTURE.md, "The two PDF roads").
//!
//! **Two windows, one box, and one frame.** The viewer stands in the band of a window this thread does
//! not own, so nothing but both moves landing inside one frame makes them one window to the eye: a box
//! is published into a cell per pin rather than queued as a command per pointer move (see
//! `PLACED_BOXES`), and a pin whose document is on the screen waits for that box to be taken up before
//! it goes on with the rest of its move (see `place`, `wait_for_the_take`).
//!
//! **And the frame around the document owes a paint the moment the document stands up.** What the pin
//! fills the band with is asked of `is_up_for` rather than of the file's kind (`a_documents_band_is_the_viewers`),
//! so a pin whose viewer has just been revealed is a band whose pixels have just changed — and the
//! viewer thread is the one that learns it, and the preview loop's thread the one that paints it (see
//! `the_band_owes_a_repaint`).

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use std::sync::mpsc::{self, Receiver, RecvTimeoutError, Sender};
use std::sync::{Condvar, Mutex};
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use webview2_com::Microsoft::Web::WebView2::Win32::{
    ICoreWebView2, ICoreWebView2Controller, ICoreWebView2Environment, ICoreWebView2_3,
    COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC,
};
use webview2_com::NavigationCompletedEventHandler;
use windows::core::{w, Interface, PCWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, POINT, RECT, WPARAM};
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::System::WinRT::EventRegistrationToken;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, GetAncestor, GetClassNameW,
    GetClientRect, IsIconic, IsWindowVisible, PeekMessageW, PostMessageW, RegisterClassW,
    SetWindowPos, ShowWindow, TranslateMessage, WindowFromPoint, GA_ROOT, MSG, PM_NOREMOVE,
    PM_REMOVE, SWP_NOACTIVATE, SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE, WHEEL_DELTA,
    WINDOW_EX_STYLE, WNDCLASSW, WS_POPUP,
};

use crate::app::engine_processes;
use crate::engines::webview_preview::{
    create_controller, create_environment, file_url, trace, wide, Area,
};

/// The window class every viewer window is made of. It is asked of the desktop by the wheel seam
/// (see `the_window_under`), which may take no lock, so the answer is a class name rather than a
/// table: one `GetClassNameW` on the root above the point.
const VIEWER_CLASS: PCWSTR = w!("RhpPdfViewer");

/// What the browser is given: the hover engine's promise that nothing a document names is ever
/// fetched — a PDF may name a remote resource as readily as a page may, and this app reads the file
/// under the pointer and nothing else — and **not** its `--hide-scrollbars`, which would take the
/// scrollbar out of a document the user is meant to scroll.
const BROWSER_ARGUMENTS: &str = "--host-resolver-rules=\"MAP * ~NOTFOUND\"";

/// How long a viewer may take to arrive before the pin is read as one that will not: a local PDF
/// through the runtime's own reader is a browser start and a file read, so what is past this is a
/// navigation that is not coming (see `Host::settle`).
const NAVIGATION_TIMEOUT: Duration = Duration::from_secs(15);

/// The pins that hold a viewer which is on the screen — revealed, or suspended with its window off
/// it — as the pin windows that own them.
///
/// A set of owners rather than one flag, because two pins are two documents: a pin whose viewer
/// never arrived keeps painting its static raster, and what another pin's live viewer says about
/// that pin's band is nothing at all (see `is_up_for`).
static LIVE_OWNERS: Lazy<Mutex<Vec<isize>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Whether any viewer at all is on the screen, which is the fast path: a paint of a kind that is
/// not a PDF — and of a PDF with no viewer — takes no lock, which is the shape `SHOWING` has for
/// the hover page.
static ANY_VIEWER_UP: AtomicBool = AtomicBool::new(false);

/// The pins whose viewer is still on its way, beside the ones it is up for (see `is_arriving_for`).
static ARRIVING_OWNERS: Lazy<Mutex<Vec<isize>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Whether any viewer at all is still on its way, for the same reason `ANY_VIEWER_UP` exists: a pin
/// that is up and owes nothing pays one read rather than a lock.
static ANY_VIEWER_ARRIVING: AtomicBool = AtomicBool::new(false);

/// Whether a pin's band owes a paint because what stands in it has changed — a document revealed in it,
/// or one let go of — which is the one fact the pin's colour there turns on (see `the_band_owes_a_repaint`).
static BAND_REPAINT_OWED: AtomicBool = AtomicBool::new(false);

/// The engine: the thread and the channel into it, begun by the first `open` and ended by the last
/// `close`.
static ENGINE: Lazy<Mutex<Option<Engine>>> = Lazy::new(|| Mutex::new(None));

/// The channel into the viewer thread. The handle to the thread is deliberately not kept: a thread
/// whose channel is dropped ends on its own — it closes what it still holds and returns — so a pin
/// closing never waits on a browser on the thread that draws the previews.
struct Engine {
    sender: Sender<Command>,
}

/// The newest box each pin has published for its viewer, and whether a command saying one is owed is
/// already on its way to the thread.
///
/// **One cell per pin rather than a queue, and one command rather than a command per box.** A hand
/// carrying the pin publishes a box per pointer move, and only the box it ends in is one to be in: what
/// a queue of them buys is a viewer that falls further behind the window it stands in for as long as
/// the hand keeps moving, and then goes on moving after the hand has stopped until it has drained —
/// which reads as the document tailing the pin rather than standing in it. The cell is what lets the
/// engine's thread decide how many moves a drag costs, and the flag is what makes a drag of a thousand
/// publishes one command to a thread that is a browser already busy.
static PLACED_BOXES: Lazy<Mutex<Vec<(isize, Area)>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Whether a command saying a box is published has been sent and not yet taken up (see `PLACED_BOXES`).
static PLACE_ASKED: AtomicBool = AtomicBool::new(false);

/// How many boxes the viewer thread has taken up, and the signal a pin waits on for one (see
/// `wait_for_the_take`).
///
/// A count rather than a flag, because a hand carrying a pin publishes a box per pointer move and each
/// of them is owed a take-up of its own: a flag would be cleared by the take-up of the box before it,
/// and the pin would be answered for a move that had not been made.
static TAKES: Lazy<(Mutex<u64>, Condvar)> = Lazy::new(|| (Mutex::new(0), Condvar::new()));

/// How long a pin waits for its viewer to take up the box it has just published: a frame of the fastest
/// display this app is likely to meet (240 Hz is 4.17 ms), and never more than one.
///
/// The wait is the pin's own thread — the one that draws it and pumps its window — so it is bounded by
/// what a frame is worth: a viewer in the middle of a navigation may not reach its pass for as long as
/// that takes, and a pin stopped for it would be a window that answers nothing while a browser thinks.
const VIEWER_CATCH_UP: Duration = Duration::from_millis(4);

/// What the app asks of a viewer.
enum Command {
    Open {
        path: PathBuf,
        /// The pin's window as a number: the viewer's owner, and the window whose content band the
        /// viewer fills. A handle rather than the type because this crosses a thread.
        owner: isize,
        area: Area,
    },
    /// **That a box is published for a viewer, and the box itself travels in the cells beside the
    /// channel**: a drag asks on every pointer move and only the last one is a box to be in (see
    /// `PLACED_BOXES`). The command is the wake and nothing else, so that a pass asleep in `recv_timeout`
    /// takes the newest box up at once rather than at the end of its wait. The owner is carried by the
    /// cell, because two pins are two boxes and a box of one is no box of the other.
    Place,
    /// One pin's viewer for one file, and never another pin's of the same name.
    Close {
        path: PathBuf,
        owner: isize,
    },
    Shutdown,
}

/// Whether a viewer stands in **this pin's** band, which is what stops that pin painting the raster
/// it still holds underneath it. Asked of the pin window rather than of the app: one pin's live
/// viewer says nothing about another pin's band.
pub(crate) fn is_up_for(owner: isize) -> bool {
    if owner == 0 || !ANY_VIEWER_UP.load(Ordering::Acquire) {
        return false;
    }

    LIVE_OWNERS
        .lock()
        .map(|owners| owners.contains(&owner))
        .unwrap_or(false)
}

/// Publish the pins a viewer is on the screen for. Written by the viewer thread once a pass, in
/// that order — the list first and the fast path after it — so a reader that sees the flag sees a
/// list that already names its own pin.
fn publish_the_live_owners(owners: Vec<isize>) {
    let any = !owners.is_empty();
    let changed = match LIVE_OWNERS.lock() {
        Ok(mut live) => {
            let changed = *live != owners;
            *live = owners;
            changed
        }
        Err(_) => false,
    };

    ANY_VIEWER_UP.store(any, Ordering::Release);

    // **A band whose middle is Chromium's owes a paint the moment that answer changes.** What the pin
    // fills that band with is asked of this list and not of the file's kind (see
    // `a_documents_band_is_the_viewers`): while no viewer stands in it the band is the raster of page
    // one the pin was taken up with, and once one does it is the caption's colour around the document.
    // So the pass a viewer is revealed in — and the one it is let go of in — is a paint owed to the pin
    // whose answer just changed: a frame left as it was is a document standing in the white of the page
    // it was taken up on, until something else happens to ask for a paint.
    if changed {
        the_band_owes_a_repaint();
    }
}

/// Owe the pins whose answer changed one paint of their band, and wake the loop that paints them.
///
/// The viewer thread is the one that learns a document has stood up or been let go of, and the preview
/// loop's own thread is the one that paints the frame around it — so the debt is a flag and the wake is
/// the channel's event. A reveal left to the loop's own tick would be a document standing in the white
/// of the page it was taken up with for as long as that tick is away (see `take_the_band_repaint`).
pub(crate) fn the_band_owes_a_repaint() {
    BAND_REPAINT_OWED.store(true, Ordering::Release);
    crate::app::seam::wake_preview_loop();
}

/// Take the band's repaint debt, if one is owed: the preview loop's tick spends it (see
/// `the_band_owes_a_repaint`).
pub(crate) fn take_the_band_repaint() -> bool {
    BAND_REPAINT_OWED.swap(false, Ordering::AcqRel)
}

/// Publish the pins whose viewer has not arrived yet, written the same way and in the same pass as
/// the live list above.
///
/// **It is what keeps a reveal off a box change the user has to cause**: a pin in this list is one
/// the preview loop publishes a box for on every tick, so the box a document arrives at is the one
/// its pin has now rather than the one the ask was made with (see `is_arriving_for`).
fn publish_the_arriving_owners(owners: Vec<isize>) {
    let arriving = !owners.is_empty();

    if let Ok(mut list) = ARRIVING_OWNERS.lock() {
        *list = owners;
    }

    ANY_VIEWER_ARRIVING.store(arriving, Ordering::Release);
}

/// Whether this pin is still waiting for its viewer, which is what the loop asks before publishing
/// a box for it (see `pin_lifecycle::place_pinned_siblings`).
pub(crate) fn is_arriving_for(owner: isize) -> bool {
    if owner == 0 || !ANY_VIEWER_ARRIVING.load(Ordering::Acquire) {
        return false;
    }

    ARRIVING_OWNERS
        .lock()
        .map(|owners| owners.contains(&owner))
        .unwrap_or(false)
}

/// The asks outstanding, one per file per pin: a viewer is one per pin, and a second ask for the
/// same file in the same pin is nothing. The list is also the refcount — the thread lives while it
/// has an ask and ends with the last one — and that is the whole lifetime rule: no TTL.
static ASKED: Lazy<Mutex<Vec<(PathBuf, isize)>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Whether this ask is the first for this file in this pin, which is what makes it one owed viewer
/// rather than another placement of the same one (see `open`).
fn the_ask_is_the_first(asks: &mut Vec<(PathBuf, isize)>, path: &Path, owner: isize) -> bool {
    if asks
        .iter()
        .any(|(asked, asked_owner)| asked == path && *asked_owner == owner)
    {
        return false;
    }

    asks.push((path.to_path_buf(), owner));
    true
}

/// Let go of this file's ask in this pin, and answer how many asks are left: the same document
/// pinned twice is two viewers, and one pin closing says nothing about the other (see `close`).
fn let_go_of_the_ask(asks: &mut Vec<(PathBuf, isize)>, path: &Path, owner: isize) -> usize {
    asks.retain(|(asked, asked_owner)| !(asked == path && *asked_owner == owner));
    asks.len()
}

/// Ask for a viewer on `path`, in the pin `owner`, at `area`, counting one more viewer owed.
///
/// The ask is taken as soon as there is a thread to take it: what it becomes is a browser start and
/// a navigation, both of which are the viewer thread's and neither of which the caller waits for.
/// A second ask for the file already asked for is nothing at all — the placement road asks on every
/// box change, and a viewer is one thing per pin rather than one per move.
pub(crate) fn open(path: &Path, owner: isize, area: Area) -> bool {
    if path.as_os_str().is_empty() || owner == 0 {
        return false;
    }

    // The placement road asks on every box change, and one pin's placement asks nothing of another
    // pin's file: the pair is what says whether a viewer is owed (see `the_ask_is_the_first`).
    match ASKED.lock() {
        Ok(mut asked) => {
            // **One pin holds one document**, so the ask for the file it was showing is the ask this
            // one replaces: let go of here rather than left behind, where it would be a browser kept
            // alive for a file nothing is showing (see `let_go_of_the_ask`).
            asked.retain(|(asked_path, asked_owner)| *asked_owner != owner || asked_path == path);

            if !the_ask_is_the_first(&mut asked, path, owner) {
                return true;
            }
        }
        Err(_) => return false,
    }

    let mut engine = match ENGINE.lock() {
        Ok(engine) => engine,
        Err(_) => return false,
    };

    if engine.is_none() {
        let (sender, receiver) = mpsc::channel();
        std::thread::spawn(move || viewer_thread(receiver));

        *engine = Some(Engine { sender });
    }

    let taken = engine.as_ref().is_some_and(|engine| {
        engine
            .sender
            .send(Command::Open {
                path: path.to_path_buf(),
                owner,
                area,
            })
            .is_ok()
    });

    if !taken {
        if let Ok(mut asked) = ASKED.lock() {
            let_go_of_the_ask(&mut asked, path, owner);
        }
    }

    taken
}

/// Move **this pin's** viewer to `area` — the pin's content band less the margin the viewer is inset
/// by, which is the whole of what a viewer is placed by. A drag of one pin is no ask at all for the
/// window of another (see `Place`).
///
/// **A box is published for a viewer that is still arriving as well as for one that is up.** The rect
/// a viewer is created with is never the rect it is shown at, so the box that answers an ask is the
/// newest one its pin has rather than the one that came with the ask (see `Seat::place`) — which is why
/// the arriving half is published and not waited on: there is no window on the screen to pair with yet.
///
/// **And a pin whose document is on the screen waits for it to arrive at the box.** The viewer is a
/// window of another thread's, so a box sent to it is a `SetWindowPos` made a thread-hop later, while
/// the pin's own window is put at that box in this one: a viewer that catches up afterwards is seen
/// standing where the pin was, one frame behind the frame the hand is watching — a document tailing the
/// window it stands in rather than being carried with it. What the wait buys is both moves committing
/// in the same frame, which is what makes the two read as one window (see `wait_for_the_take`).
pub(crate) fn place(owner: isize, area: Area) {
    if !a_viewer_is_owed(owner) {
        return;
    }

    let pairing = is_up_for(owner);
    // Read before the publish, so that the take the wait is for is one this box is inside of: a take
    // that has already happened is a box carried out before this one was published (see
    // `wait_for_the_take`).
    let taken_before = boxes_taken();

    publish_a_box(owner, area);

    // And a box nothing is going to take up is not a box to wait on: a thread that has gone leaves the
    // channel it was reading alive but deaf, and a pin waiting on it would pay the bound on every step
    // of a drag for a move that is never coming (see `ask_for_a_pass`).
    if ask_for_a_pass() && pairing {
        wait_for_the_take(taken_before);
    }
}

/// Publish this pin's newest box for the engine's thread to move its viewer into (see `PLACED_BOXES`).
/// The cell is replaced rather than added to, so a hand that has moved a hundred times has left one box.
fn publish_a_box(owner: isize, area: Area) {
    if let Ok(mut boxes) = PLACED_BOXES.lock() {
        match boxes.iter_mut().find(|(asked, _)| *asked == owner) {
            Some(box_) => box_.1 = area,
            None => boxes.push((owner, area)),
        }
    }
}

/// Ask the engine's thread for a pass, once however many boxes have been published, and answer whether one
/// is owed (see `PLACE_ASKED`).
///
/// A send that finds no engine, or one whose thread has gone, is a box nothing will take up, so the flag
/// is given back rather than left owed: a flag left owed is a viewer that stops following the hand for
/// the rest of the run.
fn ask_for_a_pass() -> bool {
    if PLACE_ASKED.swap(true, Ordering::AcqRel) {
        return true;
    }

    let Ok(engine) = ENGINE.lock() else {
        PLACE_ASKED.store(false, Ordering::Release);
        return false;
    };

    match engine.as_ref() {
        Some(engine) if engine.sender.send(Command::Place).is_ok() => true,
        _ => {
            PLACE_ASKED.store(false, Ordering::Release);
            false
        }
    }
}

/// The boxes the engine's thread is to move its viewers into, taken out of the cells whole — the flag
/// released with them, because a flag left owed over an empty cell is a viewer that never moves again.
fn take_the_boxes() -> Vec<(isize, Area)> {
    PLACE_ASKED.store(false, Ordering::Release);

    PLACED_BOXES
        .lock()
        .map(|mut boxes| boxes.drain(..).collect())
        .unwrap_or_default()
}

/// How many boxes the engine's thread has taken up (see `TAKES`).
fn boxes_taken() -> u64 {
    TAKES.0.lock().map(|taken| *taken).unwrap_or(0)
}

/// Note that a box has been taken up, which is what a pin waiting on one is woken for.
fn a_box_was_taken() {
    if let Ok(mut taken) = TAKES.0.lock() {
        *taken += 1;
        TAKES.1.notify_all();
    }
}

/// Wait for the take that carries the box published after `taken_before` was read, and answer whether it
/// came inside `VIEWER_CATCH_UP`.
///
/// **It is the whole of what makes a carried pin and its document one window.** A pinned document is a
/// window of the viewer thread's standing in the band of a window of this thread's, and the two moves
/// are one `SetWindowPos` each: the pin's is made in the pass that asks, and the viewer's is made by the
/// take-up that ask asked for, a thread-hop later. Nothing but both landing before the compositor's own
/// cut makes them travel together, so the pin waits for exactly that take — never for the browser, and
/// never for anything the document is doing.
///
/// Bounded, because this is the pin's thread: a viewer whose browser is busy may not reach its pass
/// inside the bound, and a step that passes without a take is the pair one frame apart for that step
/// rather than a pin stopped.
fn wait_for_the_take(taken_before: u64) -> bool {
    let mut taken = TAKES.0.lock().unwrap_or_else(|held| held.into_inner());
    let deadline = Instant::now() + VIEWER_CATCH_UP;

    while *taken <= taken_before {
        let left = deadline.saturating_duration_since(Instant::now());

        if left.is_zero() {
            return false;
        }

        let (guard, waited) = TAKES
            .1
            .wait_timeout(taken, left)
            .unwrap_or_else(|held| held.into_inner());
        taken = guard;

        // A wake at the very end of the bound is a take that did happen: it is read before the bound
        // rather than after it.
        if waited.timed_out() && *taken <= taken_before {
            return false;
        }
    }

    true
}

/// Whether this pin has a viewer owed or standing: an ask is what makes a box meaningful, and an ask
/// outlives the arrival it is about (see `open` and `place`).
fn a_viewer_is_owed(owner: isize) -> bool {
    owner != 0
        && ASKED
            .lock()
            .map(|asked| asked.iter().any(|(_, asked_owner)| *asked_owner == owner))
            .unwrap_or(false)
}

/// Let go of **this pin's** viewer for `path`, and — where it was the last one owed — of the browser
/// and the thread. The same document pinned twice is two viewers, and one pin closing says nothing
/// about the other (see `Close`).
pub(crate) fn close(path: &Path, owner: isize) {
    let left = match ASKED.lock() {
        Ok(mut asked) => let_go_of_the_ask(&mut asked, path, owner),
        Err(_) => return,
    };

    let mut engine = match ENGINE.lock() {
        Ok(engine) => engine,
        Err(_) => return,
    };

    let Some(held) = engine.as_mut() else {
        return;
    };

    let _ = held.sender.send(Command::Close {
        path: path.to_path_buf(),
        owner,
    });

    if left == 0 {
        // Dropping the channel is what ends the thread: it closes what it still holds — controller,
        // window, environment — on its own way out, and nothing here waits for it (see `Engine`).
        if let Some(engine) = engine.take() {
            let _ = engine.sender.send(Command::Shutdown);
        }
    }
}

/// Let go of every viewer at once: the run is ending, or a pin went away without naming a file.
pub(crate) fn close_all() {
    if let Ok(mut asked) = ASKED.lock() {
        asked.clear();
    }

    let mut engine = match ENGINE.lock() {
        Ok(engine) => engine,
        Err(_) => return,
    };

    if let Some(engine) = engine.take() {
        let _ = engine.sender.send(Command::Shutdown);
    }
}

/// The viewer's own lifecycle: the four states it can be in, and the call each transition owes.
///
/// It is a state machine rather than a handful of flags because its three questions are asked from
/// different roads — a navigation that arrives, a pin that loses the foreground, a pin that goes
/// away — and every one of them is a *transition* rather than a setting: a viewer is revealed once,
/// suspended while it is not the window the user is in, resumed rather than begun again, and closed
/// exactly once however many of those arrive behind it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Life {
    /// Asked for, created hidden, navigated: nothing of it is on screen.
    Arriving,
    /// The document arrived and the window is up, which is the only state a notch or a click is
    /// answered in.
    Shown,
    /// Up but put away (`SetIsVisible(false)` then `TrySuspend`): the window is hidden or minimized.
    Suspended,
    /// Closed, and nothing more is owed to it.
    Closed,
}

/// What a transition of `Life` owes, which is the whole of what the viewer thread does about it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Step {
    /// Put the window up: the document arrived.
    Reveal,
    /// Stop drawing: the document is not on the screen.
    Suspend,
    /// Draw again.
    Resume,
    /// Let go of it, which is also what a document that never arrived is owed: the pin's static
    /// raster is what the user is left with (see the module docs).
    Close,
    /// Nothing at all.
    Nothing,
}

impl Life {
    /// A navigation that answered, and whether a box has arrived for this viewer since it was
    /// created.
    ///
    /// **Only a document that arrived at a box of its pin's is revealed**, because the box a window
    /// was created with is the band that stood before its pin was taken up on the file it is showing:
    /// a viewer put up there is a second window standing beside the pin, and nothing but the next box
    /// change — the user dragging the caption — moves it in. A document that arrived with no box yet
    /// is still `Arriving`, which is a state the thread keeps publishing a box for (see
    /// `is_arriving_for`).
    fn navigated(&mut self, arrived: bool, placed: bool) -> Step {
        if *self != Life::Arriving {
            return Step::Nothing;
        }

        if !arrived {
            *self = Life::Closed;
            return Step::Close;
        }

        if !placed {
            return Step::Nothing;
        }

        *self = Life::Shown;
        Step::Reveal
    }

    /// What the document being on the screen decides, asked once a pass while it is up.
    fn settle(&mut self, the_document_is_on_the_screen: bool) -> Step {
        match (*self, the_document_is_on_the_screen) {
            (Life::Shown, false) => {
                *self = Life::Suspended;
                Step::Suspend
            }
            (Life::Suspended, true) => {
                *self = Life::Shown;
                Step::Resume
            }
            _ => Step::Nothing,
        }
    }

    /// Let go of: one close, and every state after it is the same state.
    fn closed(&mut self) -> Step {
        if *self == Life::Closed {
            return Step::Nothing;
        }

        *self = Life::Closed;
        Step::Close
    }

    fn is_arriving(self) -> bool {
        self == Life::Arriving
    }

    fn is_closed(self) -> bool {
        self == Life::Closed
    }

    fn is_up(self) -> bool {
        matches!(self, Life::Shown | Life::Suspended)
    }
}

/// What a viewer is apart from its window: the file it shows, the pin it stands in, the box that pin
/// has published for it, and how far the document has got.
///
/// It is a value of its own because everything the thread decides about a viewer is decided about
/// this, and none of it is a question about a window: a file asked for replaces the one its pin was
/// showing, a box is kept until the document arrives, and the reveal waits for both (see `Life`).
/// That is also what lets a walk of two documents be read as a list (see `a_walk_of_two_pdfs`).
struct Seat {
    path: PathBuf,
    owner: isize,
    life: Life,
    /// The newest box published for this viewer's pin.
    area: Area,
    /// Whether a `Place` has arrived **since this viewer was created**.
    placed: bool,
    /// What `NavigationCompleted` said, kept rather than re-read: it is one message (see `Host`).
    arrived: bool,
}

impl Seat {
    fn create(path: &Path, owner: isize, area: Area) -> Self {
        Self {
            path: path.to_path_buf(),
            owner,
            life: Life::Arriving,
            area,
            placed: false,
            arrived: false,
        }
    }

    /// The newest box this pin has. A box that arrives before the document does is the box it arrives
    /// in, which is the whole of what makes a reveal placement-atomic (see `Life::navigated`).
    fn place(&mut self, area: Area) {
        self.area = area;
        self.placed = true;
    }

    /// Whether this viewer is the one a file asked for in its pin replaces: **one pin holds one
    /// document**, so the file a walk has just left is taken down before the one it steps onto is
    /// made (see `Command::Open`).
    fn replaced_by(&self, path: &Path, owner: isize) -> bool {
        self.owner == owner && self.path != path
    }
}

/// One viewer: the seat above, its window, the controller over it, and where it has got to.
struct Host {
    seat: Seat,
    hwnd: HWND,
    controller: ICoreWebView2Controller,
    webview: ICoreWebView2,
    deadline: Instant,
    /// What `NavigationCompleted` said, sent by the handler on this same thread and taken by the
    /// thread's own pass (see `settle`).
    answer: Receiver<bool>,
    /// Whether the browser has been asked to suspend, so the ask is made once (see `suspend`).
    suspended: bool,
}

impl Host {
    fn create(
        environment: &ICoreWebView2Environment,
        path: &Path,
        owner: isize,
        area: Area,
    ) -> Option<Self> {
        register_class();

        let owner_window = HWND(owner as *mut _);
        let deadline = Instant::now() + NAVIGATION_TIMEOUT;

        // The window is created hidden and revealed only by a document that arrived: a window put
        // up first would be the pin's band gone white for as long as a browser takes to load a
        // file, which is the one thing a preview may never do.
        let hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                VIEWER_CLASS,
                w!("Rust Hover Preview"),
                WS_POPUP,
                area.x,
                area.y,
                area.width,
                area.height,
                owner_window,
                None,
                None,
                None,
            )
        }
        .ok()?;

        // A viewer is closed on every road that gives up on it, and the window it was made with is
        // the one it closes — so a failure below has one teardown rather than four.
        let Some(controller) = create_controller(environment.clone(), hwnd, deadline) else {
            trace("pdf: no controller");
            unsafe {
                let _ = DestroyWindow(hwnd);
            }
            return None;
        };

        let webview = unsafe { controller.CoreWebView2().ok() };

        let Some(webview) = webview else {
            unsafe {
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
            }
            return None;
        };

        configure(&webview);

        let Some(url) = file_url(path) else {
            unsafe {
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
            }
            return None;
        };

        let (sender, answer) = mpsc::channel::<bool>();

        unsafe {
            let handler =
                NavigationCompletedEventHandler::create(Box::new(move |_sender, args| {
                    let mut arrived = windows::Win32::Foundation::BOOL(0);
                    let _ = args.is_some_and(|args| args.IsSuccess(&mut arrived).is_ok());
                    let arrived = arrived.as_bool();
                    let _ = sender.send(arrived);
                    Ok(())
                }));

            let mut token = EventRegistrationToken::default();

            if webview
                .add_NavigationCompleted(&handler, &mut token)
                .is_err()
            {
                trace("pdf: no navigation handler");
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
                return None;
            }

            let url = wide(&url);
            let started = webview.Navigate(PCWSTR(url.as_ptr()));

            if started.is_err() {
                trace("pdf: would not navigate");
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
                return None;
            }
        }

        trace(&format!("pdf: navigating to {}", path.display()));

        Some(Self {
            seat: Seat::create(path, owner, area),
            hwnd,
            controller,
            webview,
            deadline,
            answer,
            suspended: false,
        })
    }

    /// What this viewer is owed on this pass: the document that arrived at a box of its pin's, the
    /// navigation that is not coming, and the suspension the pin's place on the screen decides.
    fn settle(&mut self) {
        if self.seat.life.is_arriving() {
            if !self.seat.arrived {
                match self.answer.try_recv() {
                    Ok(arrived) => self.seat.arrived = arrived,
                    Err(mpsc::TryRecvError::Empty) if Instant::now() >= self.deadline => {
                        trace(&format!(
                            "pdf: gave up on {} after {} s",
                            self.seat.path.display(),
                            NAVIGATION_TIMEOUT.as_secs()
                        ));
                        self.close();
                        return;
                    }
                    // A document still on its way is owed nothing yet: the box it is waiting for is
                    // the pin's own to publish (see `is_arriving_for`).
                    Err(mpsc::TryRecvError::Empty) => return,
                    Err(mpsc::TryRecvError::Disconnected) => {
                        self.close();
                        return;
                    }
                }
            }

            match self
                .seat
                .life
                .navigated(self.seat.arrived, self.seat.placed)
            {
                Step::Reveal => self.reveal(),
                Step::Close => {
                    trace(&format!(
                        "pdf: nothing arrived for {}",
                        self.seat.path.display()
                    ));
                    self.close();
                    return;
                }
                _ => {}
            }
        }

        if !self.seat.life.is_up() {
            return;
        }

        let running = self.the_document_is_on_the_screen();

        match self.seat.life.settle(running) {
            Step::Suspend => self.suspend(),
            Step::Resume => self.wake(),
            _ => {}
        }
    }

    /// Whether the document is on the screen at all, which is what decides between drawing and
    /// being suspended.
    ///
    /// **A viewer that is merely behind another window is not suspended.** A suspended controller
    /// stops drawing, and a document that is in front of the listing while the caret is in the
    /// listing would be the one thing this road may not show: an empty band where a page was. So the
    /// ask is whether the window is on the screen — a minimize, a taskbar visit, a pin collapsed to
    /// its bubble, a pin put away for a swap all hide the window the viewer is owned by, and Windows
    /// hides an owned window with its owner — rather than which window holds the keyboard. The
    /// suspension is still the whole of what a pin that goes away pays: the document is kept, and a
    /// pin that comes back is a resume rather than a browser begun again.
    fn the_document_is_on_the_screen(&self) -> bool {
        // The pin's own window is asked as well as this one: a viewer hidden *by* us is hidden for
        // the same reason an owner's hide would be, and asking both means the answer does not
        // depend on whether the Shell hides an owned window with its owner.
        unsafe {
            IsWindowVisible(self.hwnd).as_bool()
                && IsWindowVisible(self.owner()).as_bool()
                && !IsIconic(self.owner()).as_bool()
        }
    }

    /// The pin this viewer stands in, as the window that owns it.
    fn owner(&self) -> HWND {
        HWND(self.seat.owner as *mut _)
    }

    /// Put the window at the box its pin has published, which is the only box it is ever shown at.
    fn put_the_window(&self) {
        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                HWND::default(),
                self.seat.area.x,
                self.seat.area.y,
                self.seat.area.width,
                self.seat.area.height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }
    }

    fn reveal(&mut self) {
        // **The box first and the controller's own after it.** The client is what a page is laid out
        // at, so a controller told its bounds before the window moved is a page at the box before it,
        // and the document would be laid out twice (see `set_bounds`).
        self.put_the_window();

        unsafe {
            let _ = self.controller.SetIsVisible(true);
            let _ = ShowWindow(self.hwnd, SW_SHOWNOACTIVATE);
        }

        self.set_bounds();

        // **And the band this window stands in is a paint owed the pin, before the caret is asked for.**
        // What the pin fills that band with is asked of the list this reader is published in
        // (`a_documents_band_is_the_viewers`), so a frame drawn before this reveal is the page's white
        // still standing around a document that is now there — and nothing on the pin's own tick asks
        // for another. The debt is left here rather than painted here because the pin's frame is the
        // preview loop's, which is another thread (see `the_band_owes_a_repaint`).
        the_band_owes_a_repaint();

        // **And the caret goes into the document, which is what the wheel then belongs to.** A notch
        // is delivered to the window the keyboard is in, and a viewer nobody has clicked into is not
        // that window: the pin's chrome around it is, so a wheel turned over the page would scroll
        // nothing while the counter it shares with a hover went on ticking — a listing re-resolved
        // under a pin the user is reading in. Chromium is asked for the caret once the document is
        // on the screen, and from then on the wheel is the document's (see `take_the_caret`).
        self.take_the_caret();

        trace(&format!(
            "pdf: shown {} at {}x{}+{}+{}",
            self.seat.path.display(),
            self.seat.area.width,
            self.seat.area.height,
            self.seat.area.x,
            self.seat.area.y
        ));
    }

    /// Ask Chromium for the caret, so that the document is the window the wheel and the keys are in.
    ///
    /// **The controller's focus is what Windows' own focus follows**, so this is the one call that
    /// makes a page scrolled by a wheel turned over it: a mouse message goes to the window the pointer
    /// is over, but Chromium's own page only takes it once its controller holds the focus. Asking for
    /// it is what this app's chrome had otherwise taken on the press that raised the pin.
    ///
    /// The foreground window is left to Windows: a viewer is owned by a pin and the pin is the window
    /// the user put there, so raising the viewer over the desktop is not this road's to do.
    fn take_the_caret(&self) {
        unsafe {
            let _ = self
                .controller
                .MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC);
        }
    }

    /// The controller's own box, which is the window's client: a page is laid out at the size it is
    /// drawn in, so the browser is told the same box the window has.
    fn set_bounds(&mut self) {
        let mut client = RECT::default();

        if unsafe { GetClientRect(self.hwnd, &mut client) }.is_err() {
            return;
        }

        unsafe {
            let _ = self.controller.SetBounds(RECT {
                left: 0,
                top: 0,
                right: client.right.max(1),
                bottom: client.bottom.max(1),
            });
        }
    }

    fn place(&mut self, area: Area) {
        self.seat.place(area);

        if !self.seat.life.is_up() {
            // A box asked for before the document arrived is the one it arrives in (see `reveal`).
            return;
        }

        self.put_the_window();
        self.set_bounds();
    }

    /// Stop drawing, and keep the document: a suspended browser is one still there to be resumed,
    /// which is what makes a pin the user works away from cheap rather than begun again.
    fn suspend(&mut self) {
        if self.suspended {
            return;
        }

        unsafe {
            let _ = self.controller.SetIsVisible(false);
        }

        // `TrySuspend` states a precondition rather than a preference: the controller must already
        // be hidden (see above), and a runtime older than suspension is answered by the hidden
        // controller alone.
        if let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() {
            unsafe {
                let _ = webview.TrySuspend(&webview2_com::TrySuspendCompletedHandler::create(
                    Box::new(|_code, _successful| Ok(())),
                ));
            }
        }

        // The window goes off the screen with the controller, because a hidden document that a
        // suspended controller no longer draws is a blank band: the pin's band is owed nothing of
        // this app's while the pin is away (see `the_document_is_on_the_screen`).
        unsafe {
            let _ = ShowWindow(self.hwnd, SW_HIDE);
        }

        // The band is the pin's own again — it paints its raster there rather than showing the browser
        // through it — so the frame around a document that has just left is a paint owed (see
        // `the_band_owes_a_repaint`).
        the_band_owes_a_repaint();

        self.suspended = true;
        trace(&format!("pdf: suspended {}", self.seat.path.display()));
    }

    fn wake(&mut self) {
        if !self.suspended {
            return;
        }

        self.suspended = false;

        if let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() {
            unsafe {
                let _ = webview.Resume();
            }
        }

        unsafe {
            let _ = self.controller.SetIsVisible(true);
            let _ = ShowWindow(self.hwnd, SW_SHOWNOACTIVATE);
        }

        // The caret goes back into the document with the document, for the same reason `reveal` asks
        // for it: a resumed viewer whose focus was left in the pin's chrome is a page no wheel
        // scrolls.
        self.take_the_caret();

        // And the band is Chromium's again, which is the same paint a reveal owes it (see
        // `the_band_owes_a_repaint`).
        the_band_owes_a_repaint();

        trace(&format!("pdf: resumed {}", self.seat.path.display()));
    }

    fn close(&mut self) {
        if self.seat.life.closed() == Step::Nothing {
            return;
        }

        unsafe {
            let _ = self.controller.Close();
            let _ = DestroyWindow(self.hwnd);
        }

        trace(&format!("pdf: closed {}", self.seat.path.display()));
    }
}

/// The browser's own furniture for a document to be worked in: the developer tools, the status bar
/// and the messages a page could reach this app's threads through are all taken away, and nothing
/// else is. The context menu, the accelerators (`Ctrl+F`, `Ctrl+wheel`, `Ctrl+C`) and the zoom stay
/// Chromium's, because the whole point of this road is that the document is used as a document (see
/// the module docs).
fn configure(webview: &ICoreWebView2) {
    unsafe {
        if let Ok(settings) = webview.Settings() {
            let _ = settings.SetAreDevToolsEnabled(false);
            let _ = settings.SetIsStatusBarEnabled(false);
            let _ = settings.SetIsWebMessageEnabled(false);
        }
    }
}

fn register_class() {
    static REGISTERED: AtomicBool = AtomicBool::new(false);

    if REGISTERED.swap(true, Ordering::AcqRel) {
        return;
    }

    let class = WNDCLASSW {
        lpfnWndProc: Some(window_proc),
        lpszClassName: VIEWER_CLASS,
        ..Default::default()
    };

    unsafe {
        let _ = RegisterClassW(&class);
    }
}

/// A viewer window answers nothing of its own: the pin that owns it is the window with the buttons,
/// and every click, key and drag inside the document is Chromium's (see the module docs).
extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> windows::Win32::Foundation::LRESULT {
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

fn viewer_thread(commands: Receiver<Command>) {
    // The browser's own apartment, and its own thread: WebView2 must be created on a thread that is
    // pumping messages, and this is that thread for every viewer this app has open.
    let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

    // **No box is owed by a thread only just begun.** Whatever was published before this thread existed
    // aimed at windows it has never made, and a flag left set by a thread that has gone would stop the
    // next box from being asked for at all (see `PLACE_ASKED`).
    PLACE_ASKED.store(false, Ordering::Release);
    if let Ok(mut boxes) = PLACED_BOXES.lock() {
        boxes.clear();
    }

    // The thread's message queue, made before anything can be posted to it.
    {
        let mut message = MSG::default();
        unsafe {
            let _ = PeekMessageW(&mut message, None, 0, 0, PM_NOREMOVE);
        }
    }

    // A folder of its own for this thread, beside the hover engine's: one folder is one browser at a
    // time, and a viewer begun while the last one is still closing is a second browser (see
    // `webview_preview::user_data_folder`). The run's number is the name it begins with, so the
    // sweep that ends a crashed run's browsers and clears its folders reaches it either way.
    static INSTANCES: AtomicUsize = AtomicUsize::new(0);
    let instance = INSTANCES.fetch_add(1, Ordering::AcqRel) as u32;
    let folder = crate::engines::webview_preview::a_pins_profile_folder(instance);
    let _ = std::fs::create_dir_all(&folder);

    let mut environment: Option<ICoreWebView2Environment> = None;
    let mut hosts: Vec<Host> = Vec::new();

    loop {
        // While a document is up the thread polls, because a browser needs the thread that made it
        // to keep retrieving messages; with nothing open there is nothing to watch for and the wait
        // is the channel's own.
        let wait = if hosts.is_empty() {
            Duration::from_secs(60 * 60)
        } else {
            Duration::from_millis(5)
        };

        match commands.recv_timeout(wait) {
            Ok(Command::Shutdown) => break,
            Ok(Command::Open { path, owner, area }) => {
                // **The pin's other document goes before this one is made.** A walk asks for the file
                // it steps onto in the pin that was showing the file it stepped off, and one pin holds
                // one document: the outgoing window is destroyed here, before the incoming one exists
                // to be revealed, so a walk can never leave two windows standing in one band (see
                // `Seat::replaced_by` and `Command::Close`).
                for host in hosts
                    .iter_mut()
                    .filter(|host| host.seat.replaced_by(&path, owner))
                {
                    host.close();
                }
                hosts.retain(|host| !host.seat.life.is_closed());

                if environment.is_none() {
                    // The browser the runtime starts is this app's child whatever engine asked for
                    // it: one begun here is written down like the hover engine's, so a run that
                    // ends badly leaves no browser of its own behind (see `app::engine_processes`).
                    let before = engine_processes::processes_named_by_parent(
                        engine_processes::BROWSER_IMAGE,
                        std::process::id(),
                    );

                    let deadline = Instant::now() + NAVIGATION_TIMEOUT;
                    environment = create_environment(&folder, deadline, BROWSER_ARGUMENTS);

                    if environment.is_none() {
                        trace("pdf: no environment");
                    }

                    for pid in engine_processes::processes_named_by_parent(
                        engine_processes::BROWSER_IMAGE,
                        std::process::id(),
                    ) {
                        if !before.contains(&pid) {
                            trace(&format!("pdf: browser {pid}"));
                            engine_processes::record(engine_processes::BROWSER_IMAGE, pid);
                        }
                    }
                }

                let host = environment
                    .as_ref()
                    .and_then(|environment| Host::create(environment, &path, owner, area));

                if host.is_none() {
                    trace(&format!("pdf: no viewer for {}", path.display()));
                }

                if let Some(host) = host {
                    hosts.push(host);
                }
            }
            // The box travels in the cells beside the channel rather than in the command, so this arm
            // has nothing of its own to say: the command's whole job is to bring back the wait
            // `recv_timeout` is in, and the pass below takes every cell up with the rest of its work
            // (see `take_the_boxes`).
            Ok(Command::Place) => {}
            Ok(Command::Close { path, owner }) => {
                for host in hosts
                    .iter_mut()
                    .filter(|host| host.seat.owner == owner && host.seat.path == path)
                {
                    host.close();
                }
            }
            Err(RecvTimeoutError::Timeout) => {}
            Err(RecvTimeoutError::Disconnected) => break,
        }

        // **A box published while this pass was asleep is taken up before anything else the pass owes.**
        // It is the window a hand is carrying, and the pin that published it is waiting on this take
        // (see `wait_for_the_take`): a placement made after a navigation has been begun is a placement
        // made a document late, and one made after the pass's own settlement is a frame late. Only the
        // pin the box belongs to is moved — a drag of one pin is no box of another — and an arriving
        // viewer is given the box rather than moved to it (see `Seat::place`).
        let boxes = take_the_boxes();

        for (owner, area) in &boxes {
            for host in hosts.iter_mut().filter(|host| host.seat.owner == *owner) {
                host.place(*area);
            }
        }

        if !boxes.is_empty() {
            a_box_was_taken();
        }

        for host in &mut hosts {
            host.settle();
        }

        hosts.retain(|host| !host.seat.life.is_closed());

        // What the paint and placement roads read, published from the state the pass just settled:
        // the pins whose documents are on the screen and the pins still waiting for one, and no
        // others.
        publish_the_live_owners(
            hosts
                .iter()
                .filter(|host| host.seat.life.is_up())
                .map(|host| host.seat.owner)
                .collect(),
        );
        publish_the_arriving_owners(
            hosts
                .iter()
                .filter(|host| host.seat.life.is_arriving())
                .map(|host| host.seat.owner)
                .collect(),
        );

        pump_messages();
    }

    for host in &mut hosts {
        host.close();
    }

    drop(environment);

    publish_the_live_owners(Vec::new());
    publish_the_arriving_owners(Vec::new());

    unsafe {
        CoUninitialize();
    }
}

/// Retrieve and dispatch whatever is waiting, without blocking: the thread a controller was made on
/// is the thread its callbacks arrive on.
fn pump_messages() {
    let mut message = MSG::default();

    unsafe {
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

/// The window a notch over a viewer is posted to: Chromium's own child of it, or the viewer window
/// itself where the point is on its furniture.
///
/// Asked by the wheel hook, which may take no lock and runs for every wheel message on the desktop:
/// the answer is one `WindowFromPoint`, one `GetAncestor` and one class name, and nothing is kept in
/// a table this side would have to be told about. A window of another class is three calls and a
/// comparison, which is what a notch over anything else costs.
pub(crate) fn the_window_under(point: POINT) -> Option<HWND> {
    unsafe {
        let window = WindowFromPoint(point);

        if window.is_invalid() {
            return None;
        }

        let root = GetAncestor(window, GA_ROOT);
        let root = if root.is_invalid() { window } else { root };

        // Chromium's child is the window that scrolls: a notch posted to the container would be a
        // message nobody in the browser is listening for.
        class_is_the_viewers(root).then_some(window)
    }
}

unsafe fn class_is_the_viewers(window: HWND) -> bool {
    let mut name = [0u16; 32];
    let len = unsafe { GetClassNameW(window, &mut name) };

    if len <= 0 {
        return false;
    }

    let name = &name[..(len as usize).min(name.len())];

    name == VIEWER_CLASS.as_wide()
}

/// Whether `window` is a document's: a viewer's own window, or Chromium's child of one.
///
/// **The root above the window is what answers it.** Chromium's furniture — the find box, the page,
/// the scrollbar — is a tree of children whose classes are Chromium's, and a child of a viewer is the
/// document's exactly as the viewer is. So a key typed into a document's search box, and a press
/// anywhere in the document, are answered by the browser and claimed by nothing of this app's: our
/// chrome claims input only on its own strips, the caption row and the frame margin (see
/// `key_input::key_hook_proc` and `pin_input::press_the_pin`).
///
/// One atomic read before the two calls, because it is asked for every key typed on the desktop, and
/// the calls are what a desktop with no live viewer never pays for.
pub(crate) fn a_documents_window(window: HWND) -> bool {
    if window.is_invalid() || !ANY_VIEWER_UP.load(Ordering::Acquire) {
        return false;
    }

    unsafe {
        let root = GetAncestor(window, GA_ROOT);
        let root = if root.is_invalid() { window } else { root };

        class_is_the_viewers(root)
    }
}

/// Post one wheel notch to `window`, which is Chromium's, at `point` in screen coordinates.
///
/// The message is the one Windows would have handed the viewer had the keyboard been in it: a notch
/// is `WHEEL_DELTA` a line, and the point is in the message because that is what the document
/// scrolls under (`WM_MOUSEWHEEL`'s `lParam` is screen, not client). `message` is either wheel: a
/// document may be scrolled either way, and both are named here rather than by the caller.
pub(crate) fn post_a_notch(window: HWND, message: u32, notches: i32, point: POINT) -> bool {
    let wparam = WPARAM(a_notch_wparam(notches));
    let lparam = LPARAM(a_notch_lparam(point));

    unsafe { PostMessageW(window, message, wparam, lparam).is_ok() }
}

/// The delta a notch message carries, in the high word: what Windows itself puts there.
fn a_notch_wparam(notches: i32) -> usize {
    ((((notches * WHEEL_DELTA as i32) as i16 as u16) as u32) << 16) as usize
}

/// The screen point a notch message carries, in the packed form Windows uses.
fn a_notch_lparam(point: POINT) -> isize {
    (point.y as u32 as isize) << 16 | (point.x as u16 as isize)
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::UI::WindowsAndMessaging::WS_CHILD;

    /// The viewers' published lists are one set for the whole process, so the tests that write them
    /// run one at a time (see `publish_the_live_owners`).
    static THE_VIEWERS_ROOM: Mutex<()> = Mutex::new(());

    fn the_viewers_room() -> std::sync::MutexGuard<'static, ()> {
        THE_VIEWERS_ROOM
            .lock()
            .unwrap_or_else(|held| held.into_inner())
    }

    /// The road a pin's PDF is on may not become a white window: a viewer that never arrived is
    /// closed, and the pin's static raster is what stays on screen — with or without a box of its
    /// pin's in hand.
    #[test]
    fn a_viewer_that_never_arrives_is_closed_and_never_revealed() {
        for placed in [false, true] {
            let mut life = Life::Arriving;

            assert_eq!(life.navigated(false, placed), Step::Close);
            assert!(life.is_closed());
            assert!(!life.is_up());
        }
    }

    /// It is revealed exactly once, and only by a document that arrived **at a box its pin published
    /// after the viewer was created**: the box it was created with is the band that stood before the
    /// pin was taken up on this file, and a window put up there is a second window beside the pin.
    #[test]
    fn a_viewer_is_revealed_once_by_a_document_that_arrived_at_a_box() {
        let mut life = Life::Arriving;

        assert_eq!(
            life.navigated(true, false),
            Step::Nothing,
            "a document that arrived with no box of its pin's is not one to show"
        );
        assert!(!life.is_up(), "and it is still arriving");
        assert_eq!(life.navigated(true, true), Step::Reveal);
        assert_eq!(life.navigated(true, true), Step::Nothing);
        assert!(life.is_up());
    }

    /// A pin whose window goes off the screen is suspended rather than closed, and a pin that comes
    /// back is resumed: the document is kept, which is the whole of what a pin is for.
    #[test]
    fn a_viewer_the_user_leaves_is_suspended_and_the_one_they_return_to_is_resumed() {
        let mut life = Life::Arriving;
        let _ = life.navigated(true, true);

        assert_eq!(life.settle(false), Step::Suspend);
        assert_eq!(life.settle(false), Step::Nothing);
        assert_eq!(life.settle(true), Step::Resume);
        assert_eq!(life.settle(true), Step::Nothing);
        assert!(life.is_up());
    }

    /// Nothing is suspended before it is up, and a viewer is closed exactly once however many roads
    /// reach the close.
    #[test]
    fn nothing_is_suspended_before_it_is_up_and_a_close_is_owed_once() {
        let mut life = Life::Arriving;
        assert_eq!(life.settle(false), Step::Nothing);

        let mut life = Life::Suspended;
        assert_eq!(life.closed(), Step::Close);
        assert_eq!(life.closed(), Step::Nothing);
        assert!(life.is_closed());
    }

    /// A viewer is up for the pin that owns it and for no other: two pins are two documents, and a
    /// pin whose window went off the screen is still that pin's viewer, suspended rather than gone.
    #[test]
    fn a_viewer_is_up_for_the_pin_that_owns_it_and_for_no_other() {
        let _room = the_viewers_room();
        let (first, second) = (9_100_001, 9_100_002);

        publish_the_live_owners(vec![first, second]);
        assert!(is_up_for(first));
        assert!(is_up_for(second));
        assert!(!is_up_for(9_100_003), "a pin with no viewer of its own");

        // The second pin put away: its viewer is suspended, so it is still on the screen for its
        // own pin and says nothing about the first's.
        publish_the_live_owners(vec![first]);
        assert!(is_up_for(first));
        assert!(
            !is_up_for(second),
            "a viewer let go of is not up for its pin"
        );

        // And the last one closed: nothing is up at all, which is what the paint's fast path reads.
        publish_the_live_owners(Vec::new());
        assert!(!is_up_for(first));
        assert!(!ANY_VIEWER_UP.load(Ordering::Acquire));
    }

    /// **A band whose middle is Chromium's owes a paint the moment that answer changes, and not one
    /// before it.**
    ///
    /// What the pin fills the band with is asked of the viewer rather than of the file's kind (`a_documents_band_is_the_viewers`
    /// in the preview's own paint), so the pass a document is revealed in is the pass the pin's frame
    /// around it is owed — and the one it is let go of in. A pass that published the same list it
    /// published before changed nothing of what the band holds, and asks for none: a paint a tick for a
    /// document standing still is a frame nobody asked for (see `the_band_owes_a_repaint`).
    #[test]
    fn a_band_owes_a_paint_when_the_viewer_in_it_comes_or_goes() {
        let _room = the_viewers_room();
        let pin = 9_100_004;

        let _ = take_the_band_repaint();
        publish_the_live_owners(Vec::new());
        assert!(
            !take_the_band_repaint(),
            "a desktop with no viewer in it has changed nothing"
        );

        publish_the_live_owners(vec![pin]);
        assert!(
            take_the_band_repaint(),
            "the document stood up: the frame around it is a paint owed"
        );
        assert!(
            !take_the_band_repaint(),
            "and one paint is the whole of the debt, however often it is asked for"
        );

        publish_the_live_owners(vec![pin]);
        assert!(
            !take_the_band_repaint(),
            "the same viewer standing in the same band is nothing to paint"
        );

        publish_the_live_owners(Vec::new());
        assert!(
            take_the_band_repaint(),
            "and the band is the pin's own again once the document is let go of"
        );
    }

    /// **A box is published into a cell rather than queued, and a hand that has moved a hundred times
    /// has left one.**
    ///
    /// It is what keeps a viewer from tailing the window it stands in: the pin publishes a box per
    /// pointer move and every one of them is replaced by the one after it, so what a take-up carries
    /// out is where the hand has got to rather than where it has been (see `PLACED_BOXES`).
    #[test]
    fn a_drag_leaves_the_box_the_hand_is_at_now_rather_than_a_backlog() {
        let (first, second) = (9_100_005, 9_100_006);
        let area = |x: i32| Area {
            x,
            y: 40,
            width: 640,
            height: 480,
        };

        // A thread only just begun owes no box of a run before it (see `viewer_thread`).
        PLACE_ASKED.store(false, Ordering::Release);
        if let Ok(mut boxes) = PLACED_BOXES.lock() {
            boxes.clear();
        }

        // A hand carrying one pin: three boxes published, and the cell holds the last.
        publish_a_box(first, area(10));
        publish_a_box(first, area(20));
        publish_a_box(first, area(30));

        // And a second pin's box is a cell of its own — one pin's move is no box of another.
        publish_a_box(second, area(50));

        let taken = take_the_boxes();
        assert_eq!(
            taken,
            vec![(first, area(30)), (second, area(50))],
            "one box per pin, and the one the hand is at now"
        );
        assert!(
            PLACED_BOXES
                .lock()
                .map(|boxes| boxes.is_empty())
                .unwrap_or(false),
            "and taking them up leaves the cells empty"
        );
    }

    /// An ask is one viewer per file per pin: the same document pinned twice is two viewers, a
    /// second ask in the same pin is nothing, and one pin closing takes its own ask and no other.
    #[test]
    fn an_ask_is_one_viewer_per_file_per_pin() {
        let mut asks = Vec::new();
        let pdf = Path::new("a.pdf");
        let other = Path::new("b.pdf");

        assert!(the_ask_is_the_first(&mut asks, pdf, 11));
        assert!(
            !the_ask_is_the_first(&mut asks, pdf, 11),
            "one viewer per pin"
        );
        assert!(
            the_ask_is_the_first(&mut asks, pdf, 22),
            "another pin, another viewer"
        );
        assert!(the_ask_is_the_first(&mut asks, other, 11));
        assert_eq!(asks.len(), 3);

        assert_eq!(let_go_of_the_ask(&mut asks, pdf, 11), 2);
        assert!(
            !asks.iter().any(|(path, owner)| path == pdf && *owner == 11),
            "the pin that closed took its own ask"
        );
        assert!(asks.iter().any(|(path, owner)| path == pdf && *owner == 22));

        let_go_of_the_ask(&mut asks, other, 11);
        assert_eq!(
            let_go_of_the_ask(&mut asks, pdf, 22),
            0,
            "the last ask ends the thread"
        );
    }

    /// **Two PDFs walked in one pin leave one viewer, standing in the inset band.**
    ///
    /// The ask that steps onto the second file names the band that stood when it was made, and the pin
    /// is taken up at the box of the document it shows now *after* that: a viewer revealed at the box it
    /// was created with is a window of its own beside the pin until a box change snaps it in. So the box
    /// a document arrives at is the last one its pin published, the file a walk left is taken down before
    /// the one it stepped onto is made, and exactly one viewer is left (see `Life::navigated`,
    /// `Seat::place` and `Seat::replaced_by`).
    #[test]
    fn a_walk_of_two_pdfs_leaves_one_viewer_in_the_band() {
        let pin = 9_100_011;
        // The band of the page the pin was showing, and the band of the page it shows now: the second
        // is what the pin publishes once it is taken up on the file the walk stepped onto.
        let the_band_before = Area {
            x: 100,
            y: 140,
            width: 400,
            height: 300,
        };
        let the_inset_band = Area {
            x: 108,
            y: 148,
            width: 384,
            height: 284,
        };

        // What the thread's `Command::Open` answers, in that order: the pin's other document goes, this
        // one is created at the box the ask named.
        fn open(seats: &mut Vec<Seat>, path: &str, owner: isize, area: Area) {
            seats.retain(|seat| !seat.replaced_by(Path::new(path), owner));
            seats.push(Seat::create(Path::new(path), owner, area));
        }

        // What the pass answers once the document has arrived and the box has come with it.
        fn arrive(seat: &mut Seat) -> Step {
            seat.arrived = true;
            seat.life.navigated(seat.arrived, seat.placed)
        }

        let mut seats = Vec::new();

        // **A document that arrived before its pin's box did is not shown**: the box it was created with
        // is the band of the file before it.
        open(&mut seats, "a.pdf", pin, the_band_before);
        assert_eq!(
            arrive(&mut seats[0]),
            Step::Nothing,
            "a viewer is never revealed at the rect it was created with"
        );
        assert!(!seats[0].life.is_up());
        seats[0].place(the_inset_band);
        assert_eq!(arrive(&mut seats[0]), Step::Reveal);
        assert_eq!(
            seats[0].area, the_inset_band,
            "the box it arrived at is its pin's"
        );

        // And the walk: another file in the same pin, at the band that stood when it was asked for.
        open(&mut seats, "b.pdf", pin, the_band_before);
        assert_eq!(
            seats.len(),
            1,
            "the file the walk left is gone before this one is made"
        );
        assert_eq!(seats[0].path, Path::new("b.pdf"));

        seats[0].place(the_inset_band);
        assert_eq!(arrive(&mut seats[0]), Step::Reveal);
        assert_eq!(
            seats[0].area, the_inset_band,
            "the viewer stands in the band the pin has now"
        );
        assert!(seats[0].life.is_up());

        // Another pin's document is another document, and no file of this pin's takes it down.
        let other = Seat::create(Path::new("b.pdf"), 9_100_012, the_inset_band);
        assert!(!other.replaced_by(Path::new("c.pdf"), pin));
        assert!(!other.replaced_by(Path::new("b.pdf"), other.owner));
        assert!(other.replaced_by(Path::new("c.pdf"), other.owner));
    }

    /// **Our chrome claims input only on its own strips.** A viewer's window — or Chromium's own child
    /// of it, its find box and its page included — is the document's, and the app claims nothing it is
    /// sent there; a window of the pin's own, and a desktop with no viewer on it, are what they always
    /// were (see `a_documents_window`).
    #[test]
    fn a_viewer_window_and_its_chromium_child_are_the_documents() {
        let _room = the_viewers_room();
        register_class();

        unsafe {
            // A viewer, a child of it of another class — what Chromium's own furniture is, at whatever
            // class a browser spells it — and a window of the pin's own.
            let viewer = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                VIEWER_CLASS,
                w!(""),
                WS_POPUP,
                0,
                0,
                320,
                240,
                None,
                None,
                None,
                None,
            )
            .expect("a viewer window");
            let chromium = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!(""),
                WS_CHILD,
                0,
                0,
                100,
                100,
                viewer,
                None,
                None,
                None,
            )
            .expect("a child of the viewer");
            let pin = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!(""),
                WS_POPUP,
                0,
                0,
                320,
                240,
                None,
                None,
                None,
                None,
            )
            .expect("a window of the pin's own");

            // A desktop with no viewer up claims nothing, whatever the window is.
            publish_the_live_owners(Vec::new());
            assert!(
                !a_documents_window(viewer),
                "with no viewer up there is no document"
            );
            assert!(!a_documents_window(chromium));

            publish_the_live_owners(vec![9_100_012]);
            assert!(
                a_documents_window(viewer),
                "a viewer's window is the document's"
            );
            assert!(
                a_documents_window(chromium),
                "and Chromium's child of it is the document's too: a key typed into its search box \
                 is a key typed into the document"
            );
            assert!(
                !a_documents_window(pin),
                "a window of the pin's own is the pin's, which is what carries it and resizes it"
            );
            assert!(
                !a_documents_window(HWND::default()),
                "and no window at all is nobody's"
            );

            publish_the_live_owners(Vec::new());

            let _ = DestroyWindow(chromium);
            let _ = DestroyWindow(viewer);
            let _ = DestroyWindow(pin);
        }
    }

    /// The notch a viewer is handed is the message Windows would have made: the delta in the high
    /// word of `wParam`, and the point in screen coordinates in `lParam`.
    #[test]
    fn a_notch_carries_its_delta_and_the_point_it_was_turned_at() {
        let point = POINT { x: 300, y: 200 };

        assert_eq!(
            (a_notch_wparam(3) as u32 >> 16) as i16,
            3 * WHEEL_DELTA as i16
        );
        assert_eq!(
            (a_notch_wparam(-1) as u32 >> 16) as i16,
            -(WHEEL_DELTA as i16)
        );

        let packed = a_notch_lparam(point) as u32;
        assert_eq!((packed & 0xffff) as i16, 300);
        assert_eq!((packed >> 16) as i16, 200);
    }
}

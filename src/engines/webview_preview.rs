//! Drawing a document — or a font specimen — in the browser engine that is already on the
//! machine.
//!
//! Every document this app previews is drawn here — still, gzipped or animated: this app
//! rasterizes none of them, so what a document needs is the WebView2 runtime Windows 11
//! ships with, which is Chromium, and which is the only complete implementation of SVG
//! that is on the machine without installing anything. A machine without the runtime has
//! no SVG preview at all, which is the answer a file that will not decode gets.
//!
//! A font is drawn by the same engine for the same reason, and it is one of the two kinds
//! this app draws rather than rasterizes: the five formats a font goes by are one container
//! or another around the same outlines, and the browser reads all of them. What the page
//! carries for one is a specimen — the font's own name and the lines its character map
//! covers, worked out on this side by `font_preview` — and a collection is answered with the
//! face the setting names written out as a font of its own, because no page can name a face
//! inside one.
//!
//! What lives here is the engine and the window it draws in, not the preview loop: the
//! loop measures the document for the layout, hands it over with the box the layout came
//! out with, and this answers with a window of its own. That window is its own because
//! the preview window is a layered one, and a layered window has no window tree to put a
//! child in — the same shape as the video path, where the player's own window is the
//! preview.
//!
//! One engine is kept warm between documents and let go after `webview_idle`, ten
//! minutes by default: beginning one costs a browser start, and pointing a warm one at
//! another file costs a few milliseconds, so what a hover pays for a second document is
//! nothing worth measuring — and what a hover pays for the *first* one is a browser
//! start, which is what the waiting spinner is shown for. What is let go of is the
//! engine and not the thread that holds it, which stays parked on its channel for the
//! run and begins a new engine for the next document: a thread that ended with its
//! browser would leave every document after the first idle timeout with nobody to draw
//! it, which is no preview for the rest of the run. An app left alone has no browser
//! process and one thread asleep — and the settings it is given are the app's own rules
//! rather than a browser's: a document is drawn and not run, and nothing about it is a
//! way out of the preview. A page of HTML is the one exception to that, and the exception
//! is deliberate rather than a leak: it is handed to the browser as a page rather than as an
//! image, a page that draws itself is nothing without a run, and what a run is given stops
//! at everything that is a way out of the frame it is in (see `html_page`, `page_runs`).
//!
//! What is kept between documents is the browser, not the work it was doing, and
//! `Host::hide`/`Host::wake` are what stop and start it (see `suspend`).

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc::{self, Receiver, RecvTimeoutError, Sender};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use webview2_com::Microsoft::Web::WebView2::Win32::{
    CreateCoreWebView2EnvironmentWithOptions, GetAvailableCoreWebView2BrowserVersionString,
    ICoreWebView2, ICoreWebView2Controller, ICoreWebView2Environment,
    ICoreWebView2EnvironmentOptions, ICoreWebView2_3, COREWEBVIEW2_COLOR,
};
use webview2_com::{
    CoreWebView2EnvironmentOptions, CreateCoreWebView2ControllerCompletedHandler,
    CreateCoreWebView2EnvironmentCompletedHandler, NavigationCompletedEventHandler,
    TrySuspendCompletedHandler,
};
use windows::core::{w, Interface, PCWSTR, PWSTR};
use windows::Win32::Foundation::{E_POINTER, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::System::Com::{CoInitializeEx, CoTaskMemFree, COINIT_APARTMENTTHREADED};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::System::WinRT::EventRegistrationToken;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, GetMessageW,
    GetWindowLongPtrW, KillTimer, PeekMessageW, PostThreadMessageW, RegisterClassW, SetTimer,
    SetWindowLongPtrW, SetWindowPos, ShowWindow, TranslateMessage, GWL_EXSTYLE, HWND_TOPMOST, MSG,
    PM_NOREMOVE, PM_REMOVE, SWP_NOACTIVATE, SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE, WM_APP,
    WM_MOUSEACTIVATE, WNDCLASSW, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

use crate::app::engine_processes;
use crate::config::config::{
    EngineIdle, PreviewType, TransparentBackground, DEFAULT_WEBVIEW_IDLE_SECS,
};
use crate::formats::font_formats;
use crate::formats::text_formats;
use crate::paths::plain_path;
use crate::readers::font_preview;
use crate::{readers::svg_preview, CONFIG};

/// The window class the engine's window is made from. It exists to refuse activation, and
/// the refusal is a default rather than a fixed answer: a document and a specimen never
/// take the keyboard away from what the pointer is over, and a page that runs is the one
/// thing a browser is brought here for — so a window of this class asks whether it is
/// holding a page before it says no (see `mouse_activate_answers`).
const WEBVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWebView");

/// The browser arguments this app passes. The first answers for no host at all, so
/// nothing a document links to is fetched from anywhere (see `create_environment`); the
/// second keeps a scrollbar out of a preview if a document is ever a pixel larger than
/// the window it was given. Both are quoted because the browser's command line is a
/// command line: unquoted, an argument with a space in it arrives as several.
const BROWSER_ARGUMENTS: &str = "--host-resolver-rules=\"MAP * ~NOTFOUND\" --hide-scrollbars";

/// What one engine cost to begin and to point at a document, in milliseconds. It is
/// kept for the same reason the other probes exist: "the preview is slow" is answered
/// by a number, and this is the module the number belongs to.
#[derive(Clone, Copy, Debug, Default)]
pub struct Timings {
    pub environment_ms: u64,
    pub controller_ms: u64,
    pub navigate_ms: u64,
}

static LAST_TIMINGS: Lazy<Mutex<Timings>> = Lazy::new(|| Mutex::new(Timings::default()));

/// Whether the runtime is on this machine, answered once: the check reads the version
/// of the installed runtime, and that does not change while the app runs.
static RUNTIME: Lazy<Option<String>> = Lazy::new(runtime_version);

/// Whether the engine's window is on screen. The preview loop reads this to know when
/// to take its own window down, so it is an atomic rather than a message.
static SHOWING: AtomicBool = AtomicBool::new(false);

/// Whether the document the engine's window is showing right now is one that runs, and so may
/// take the pointer and the keyboard: the window is made `WS_EX_NOACTIVATE` and refuses
/// activation for `WM_MOUSEACTIVATE`, which is right for a document that is only looked at and
/// wrong for a page whose own scripts are reading keys and whose own pointer is dragging it
/// about, so both of those answers are about what the window is holding rather than about the
/// window.
///
/// It is a fact about the document on screen, kept by the thread that puts the window up, and
/// read by the window procedure, which is that thread's own — an atomic rather than a
/// message because the two are the same thread at different moments, and because a window
/// procedure that had to ask the engine's lock to answer a click would be a click answered
/// late. Nothing is set when the window comes down: a window that is not on screen cannot be
/// clicked into, so the next document to be put up is what sets it again.
static RUNNING_DOCUMENT: AtomicBool = AtomicBool::new(false);

/// Where the engine's window is while it is on screen, in screen coordinates: the box
/// the document was last handed over in. Kept as a rectangle rather than as the window
/// handle, because a rectangle is what the preview loop asks about a preview.
static SHOWING_RECT: Lazy<Mutex<Option<ScreenRect>>> = Lazy::new(|| Mutex::new(None));

/// A rectangle in screen coordinates, as the preview loop asks about one.
type ScreenRect = (i32, i32, i32, i32);

/// The engine's thread, once one has been started.
static ENGINE: Lazy<Mutex<Option<Engine>>> = Lazy::new(|| Mutex::new(None));

/// Where the engine is asked to put its window, in screen coordinates.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub struct Area {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

/// The document the preview loop wants drawn, as the one thing that is wanted: the file,
/// the backdrop its page is written for, the box the wait has ended up in, and the
/// generation this want was made under.
///
/// What this is for is a pointer that crosses a folder of documents or specimens: every
/// file it settles on is one the loop asks for, one after another, and an engine that
/// draws one document at a time and can be a tenth of a second about it would otherwise
/// draw each of them in turn — showing a file the hand has already left before the one it
/// is on, for as long as the hand keeps moving.
///
/// One cell rather than a queue, so what the engine takes up is the newest file rather than
/// the oldest — the rule `PLACED` and `PLACE_ASKED` follow for a box rather than a document.
/// A want is a generation of its own only where the *document* is another one: a box that
/// moved while the same document was being navigated to is the same want, and counting it as
/// another would throw away the navigation it is waiting for (see `show`).
static WANTED: Lazy<Mutex<Option<Wanted>>> = Lazy::new(|| Mutex::new(None));

/// The generation of the newest want. Every request carries the generation it was made
/// under, and one that is not this is not carried out — which is the question the engine's
/// thread asks of a navigation that is still running, in the middle of pumping a browser's
/// messages and unable to take a lock the loop may be holding.
static WANTED_GENERATION: AtomicU64 = AtomicU64::new(0);

/// The box a document the engine is *holding* is to be moved to, published under one cell
/// rather than a queue (see `WANTED`).
///
/// A drag is many placements and only the place the hand let go at is worth putting up. This
/// is what a placement is made of, and it is deliberately not a want: a want names a document
/// and is replaced only by another document, while this is the same document in another box.
/// Publishing a new one here therefore *replaces* what was there, which is the whole of the
/// coalescing — kept apart from `WANTED` because a document that has landed keeps its want
/// and a box is asked for far more often than a document is (see `place`, `ask_place`).
static PLACED: Lazy<Mutex<Option<Placement>>> = Lazy::new(|| Mutex::new(None));

/// Whether a placement has been asked of the engine's thread and not yet taken up, so that a
/// drag which outruns the engine leaves one placement owed rather than a queue of them.
///
/// The engine's thread takes one command per pass, and a document that animates is a browser
/// already busy compositing — so a queue of placements is a queue it can never drain. What
/// falls behind is not the box but the whole browser: the window stops following the hand,
/// and then the engine stops answering altogether, which is the same answer as one that hung
/// (see `NAVIGATION_TIMEOUT`). One outstanding placement is the bound, and the cell above is
/// what makes satisfying it cheap: the ask that finds one already outstanding publishes a
/// newer box and returns, and that newer box is what the outstanding one is answered with.
static PLACE_ASKED: AtomicBool = AtomicBool::new(false);

/// A box a document the engine is holding is to be moved to, with the two facts it is
/// published under: the file it is for, so a box left behind by a document the preview has
/// moved on from is not carried out, and the backdrop, which is the page's own as well as the
/// controller's and so can make a placement a page to write again (see `Host::needs_page`).
#[derive(Clone)]
struct Placement {
    path: PathBuf,
    background: TransparentBackground,
    area: Area,
}

/// The newest placement asked for, taken back for the engine's thread to carry out, and the
/// flag that says one is owed released — the two are taken together because a placement left
/// owed with nothing behind it is a window that would never move again.
fn take_placement() -> Option<Placement> {
    PLACE_ASKED.store(false, Ordering::Release);

    PLACED.lock().ok().and_then(|mut placed| placed.take())
}

/// Throw a published placement away without carrying it out, and give back the flag that says
/// one is owed.
///
/// This is what a hide and a shutdown both do: the window the box was for is off screen or
/// gone, so the box belongs to nothing, and leaving it would have it applied to the next
/// document put up rather than dropped. The flag goes back for the same reason it is given
/// back when a send fails — a flag left owed is a window that stops following the hand for
/// the rest of the run.
fn drop_placement() {
    PLACE_ASKED.store(false, Ordering::Release);

    if let Ok(mut placed) = PLACED.lock() {
        *placed = None;
    }
}

/// The engine thread's id, once it is running: what a want published while that thread is
/// parked on `GetMessage` is posted to, so that the wait a navigation is in the middle of
/// ends the moment what is wanted is not what it is waiting for.
static ENGINE_THREAD: AtomicU32 = AtomicU32::new(0);

/// What a wakeup to the engine's thread is: a message the window procedure ignores, whose
/// whole job is to bring `GetMessage` back so a wait can look at what is wanted now.
const ENGINE_WAKE: u32 = WM_APP;

/// Wake the engine's thread if it is waiting on something.
///
/// A want published while that thread is parked on `GetMessage` would otherwise be noticed
/// only when the browser happened to say something — which, for a navigation that never
/// lands, is never. A post is not lost by the thread being between two messages: it waits
/// in the queue and brings the next `GetMessage` back at once.
fn wake_engine_thread() {
    let thread = ENGINE_THREAD.load(Ordering::Acquire);

    if thread != 0 {
        unsafe {
            let _ = PostThreadMessageW(thread, ENGINE_WAKE, WPARAM(0), LPARAM(0));
        }
    }
}

/// The document the engine's window is showing, when one is: the file a landed navigation
/// was for, kept until the window comes down.
///
/// The preview loop reads it to tell its own document landing from another hover's. A
/// document that lands is what a wait ends on, and a wait is for one file: a page that
/// arrives for a hover the loop has already left would otherwise take down the wait it is
/// in the middle of and put a file the pointer has left on screen — see
/// `preview_window`'s handover.
static SHOWN: Lazy<Mutex<Option<PathBuf>>> = Lazy::new(|| Mutex::new(None));

/// The engine's window handle, while it has one: what a hit test needs to know whether the
/// pointer is on a document's preview, which is a preview of this app's in every way but
/// the window it is drawn in.
static HOST_HWND: AtomicIsize = AtomicIsize::new(0);

/// The document the loop wants drawn, as the engine's thread takes it up: the want and the
/// generation it was made under, or nothing when nothing is wanted.
#[derive(Clone)]
struct Wanted {
    generation: u64,
    path: PathBuf,
    background: TransparentBackground,
    area: Area,
}

/// What is wanted this moment, for the engine's thread to take up.
fn wanted() -> Option<Wanted> {
    WANTED.lock().ok().and_then(|wanted| wanted.clone())
}

/// Whether what was asked for under `generation` is still what is wanted.
fn is_wanted(generation: u64) -> bool {
    WANTED_GENERATION.load(Ordering::Acquire) == generation
}

/// The box the newest want for `generation` asks for, when that is still the newest want:
/// where a navigation that has just landed puts its window, since a pointer that moved
/// while the document was being drawn has moved the preview with it.
fn wanted_area(generation: u64) -> Option<Area> {
    WANTED
        .lock()
        .ok()
        .and_then(|wanted| {
            wanted
                .as_ref()
                .map(|wanted| (wanted.generation, wanted.area))
        })
        .filter(|(wanted, _)| *wanted == generation)
        .map(|(_, area)| area)
}

/// The version of the WebView2 runtime on this machine, or nothing when it is not
/// installed — which is what decides whether an SVG document can be previewed at all.
pub fn runtime_version() -> Option<String> {
    unsafe {
        let mut version = PWSTR::null();
        GetAvailableCoreWebView2BrowserVersionString(PCWSTR::null(), &mut version).ok()?;

        // A null string is a runtime that is not installed, and the buffer is freed either
        // way: it is allocated with `CoTaskMem` on this thread and does not outlive this call.
        let text = if version.is_null() {
            None
        } else {
            version.to_string().ok()
        };
        CoTaskMemFree(Some(version.0 as *const _));

        text
    }
}

/// Whether a document can be handed to the engine at all.
pub fn is_available() -> bool {
    RUNTIME.is_some()
}

/// Whether the engine has a window on screen.
pub fn is_showing() -> bool {
    SHOWING.load(Ordering::Acquire)
}

/// Where the engine's window is, when it has one on screen: the box the document was
/// last handed over in, in screen coordinates.
///
/// The preview loop reads it for the reason it reads its own window's rectangle: a
/// preview that is under the parked pointer is one a keyboard hover placed there, and a
/// document the engine draws is a preview of this app's even though the window is not.
pub fn screen_rect() -> Option<ScreenRect> {
    SHOWING_RECT.lock().ok().and_then(|rect| *rect)
}

/// The document the engine's window is showing, when it is showing one: the file the page
/// that landed was written for.
///
/// The preview loop asks this to tell its own document from another hover's: a page that
/// arrives is what a wait ends on, and a wait is for one file — one that lands for a hover
/// the loop has left is not a wait that has ended (see `SHOWN`).
pub fn showing_path() -> Option<PathBuf> {
    SHOWN.lock().ok().and_then(|shown| shown.clone())
}

/// The window the engine draws in, when it has one: the handle a hit test compares against
/// the window under the pointer, so that a document the engine draws is touched the way
/// this app's own preview is — the window is the engine's, and the preview is this app's.
pub fn showing_hwnd() -> isize {
    HOST_HWND.load(Ordering::Acquire)
}

/// Say what the engine's window is showing and where it is, or that it is showing nothing
/// and is nowhere: the three things the preview loop asks a preview about, published together
/// because they are one fact and are read together (see `is_showing`, `showing_path`,
/// `screen_rect`).
///
/// Nothing is published as the window comes down, which is what `hide` is for.
fn publish(shown: Option<(PathBuf, Area)>) {
    if let Ok(mut path) = SHOWN.lock() {
        *path = shown.as_ref().map(|(path, _)| path.clone());
    }

    SHOWING.store(shown.is_some(), Ordering::Release);

    publish_rect(shown.map(|(_, area)| area));
}

/// Say where the engine's window is — or that it is nowhere — for `screen_rect` to answer
/// with. The one half of `publish` a window being *moved* is also owed.
fn publish_rect(area: Option<Area>) {
    if let Ok(mut rect) = SHOWING_RECT.lock() {
        *rect = area.map(|area| (area.x, area.y, area.x + area.width, area.y + area.height));
    }
}

/// What the engine cost last time it was asked for a document. Read by the probe: it
/// is the number that says whether the engine is worth keeping warm at all.
#[cfg(test)]
pub fn last_timings() -> Timings {
    LAST_TIMINGS
        .lock()
        .map(|timings| *timings)
        .unwrap_or_default()
}

/// Whether a file is one the engine draws: the runtime is on the machine, the file is a
/// document or a font, and the engine is not in one of its own bad spells.
///
/// A page of HTML is the third of those, and only while `render_html` asks for it — which is
/// what `page_runs` answers. A machine with no runtime answers no for all three alike, which
/// is what leaves such a file its text preview (see `load_media_of_kind`).
pub fn draws(path: &Path) -> bool {
    can_draw()
        && (svg_preview::is_svg_file(path) || font_formats::is_font_file(path) || page_runs(path))
}

/// Whether `path` is a page this app hands to the browser as a page rather than as an image,
/// and runs the code in it.
///
/// It is the one kind of document that is given to a browser whole: a document with a size of
/// its own and a specimen of a font are both pictures, and a picture in a browser is never
/// anything but a picture — no code in it, no pointer of its own, nothing outside itself. A
/// page of HTML has none of that as a picture, and a page that draws itself with WebGL, or
/// lays itself out from a script, is a page that nothing but a run can show at all: run with
/// the page withheld and it comes up blank. So the script is given here and nowhere else, and
/// it is given to this one kind because it is the one kind this app ever opens in a browser's
/// hands rather than in a frame of its own making — the wrapper is a file of this run's
/// profile folder and the document is another file, so nothing a document runs can reach the
/// page that framed it.
///
/// The two names are the two the text lists already claim, read through the same helper they
/// are read by, and the switch is `render_html`: the same file is a page of text without it,
/// so the name alone does not make a page run — it is the tray, or `config.ini`, that says a
/// page of HTML is drawn at all, and a preview that is not being drawn has no engine to run
/// anything in. The question is about the document's kind and the switch, not about the
/// machine: whether a browser is here to draw it is `can_draw`, which is deliberately not
/// folded in, because the answer a caller wants here is the rule and not whether it can be
/// carried out.
pub fn page_runs(path: &Path) -> bool {
    page_runs_under(renders_html(), path)
}

/// The same question as `page_runs`, with the switch handed in rather than read.
///
/// The two names a page goes by and the switch that says a page is drawn at all are two
/// separate questions, and only one of them is about a name: which names count is a matter
/// of the text lists, and can be asked — and answered — of any file at any time. The switch
/// belongs to the tray, or to `config.ini`, and is kept behind this app's one configuration
/// lock, so handing it in is what lets the part of the rule that is about a name be asked
/// without the configuration at all. A test that took the lock to ask it would be waiting on
/// the lock it is already holding.
fn page_runs_under(renders_html: bool, path: &Path) -> bool {
    renders_html && text_formats::is_html_extension(path)
}

/// Whether the configuration asks for a page of HTML to be drawn rather than its markup.
pub fn renders_html() -> bool {
    CONFIG
        .lock()
        .map(|config| config.render_html)
        .unwrap_or(false)
}

/// Whether the engine still owes `path`: a want published for it that has not been
/// superseded, and the engine not in one of its own bad spells.
///
/// A *want* is what is owed rather than what is drawn, and the two are different for as
/// long as a browser takes to come up: `showing_hwnd` is zero until the host exists, so a
/// document that was asked for a moment ago is one the engine owes and has no window for
/// yet (see `is_behind`, which is the question a pinned preview asks).
pub fn owed(path: &Path) -> bool {
    !is_failing() && wanted().is_some_and(|wanted| wanted.path == path)
}

/// Whether the engine still stands behind this document: the window it draws in exists —
/// up, or hidden because the pin that shows it is a bubble — or the document is still on
/// its way and no failure has stood the engine down.
///
/// It is what a pinned window asks to know whether the thing it is a window onto is still
/// there, and the window alone is not the answer: a browser being begun has no window yet,
/// and reading that as "gone" closed a pin the moment it was shown a document the engine
/// had to start one for.
pub fn is_behind(path: &Path) -> bool {
    !is_failing() && (showing_hwnd() != 0 || owed(path))
}

/// Whether the engine can draw anything at all.
///
/// The runtime is what draws a document or a specimen, so a machine without it — or one the
/// engine has stood down on — has neither: this is the question the layout asks before it
/// measures one, and answering no is what keeps a hover from opening a box that nothing
/// would be drawn into.
pub fn can_draw() -> bool {
    is_available() && !is_failing()
}

/// How long the engine is left alone after it has failed to come up.
///
/// The reason it can fail is a folder, not the document: one user data folder is one
/// browser at a time, and a browser left behind by an earlier run — one whose app was
/// ended before it could take its browser with it — holds it until it goes. What that
/// must not cost is more than it has to: an engine that cannot be had draws no document
/// at all, so for this long afterwards an SVG hover opens nothing, and the engine is
/// asked again once the window has passed.
const ENGINE_RETRY_AFTER: Duration = Duration::from_secs(60);

/// When the engine last failed to come up.
static ENGINE_FAILED_AT: Lazy<Mutex<Option<Instant>>> = Lazy::new(|| Mutex::new(None));

/// Whether the preview loop has not yet been told that the engine has something to
/// answer for.
static FAILURE_NOTICE: AtomicBool = AtomicBool::new(false);

/// Whether the engine has failed since this was last asked.
///
/// One signal for the two ways it can: an engine that could not be started at all, which
/// stands it down for a while, and one that could not put a document up, which costs
/// that hover and nothing else. What the preview loop does with either is the same,
/// because this app draws no document itself: the wait a hover was given goes, rather
/// than staying a spinner over nothing.
pub fn take_failure_notice() -> bool {
    FAILURE_NOTICE.swap(false, Ordering::AcqRel)
}

fn is_failing() -> bool {
    ENGINE_FAILED_AT
        .lock()
        .ok()
        .and_then(|failed| *failed)
        .is_some_and(|failed| failed.elapsed() < ENGINE_RETRY_AFTER)
}

/// The engine could not be started at all: the runtime is on the machine, but no browser
/// could be had for it, which is the profile folder being held by one that is not this
/// app's.
///
/// It stands the engine down for `ENGINE_RETRY_AFTER` — no document is handed over in
/// that window, because every one of them would fail the same way — and tells the
/// preview loop, whose hover has nothing left to wait for.
fn note_failure() {
    if let Ok(mut failed) = ENGINE_FAILED_AT.lock() {
        *failed = Some(Instant::now());
    }

    FAILURE_NOTICE.store(true, Ordering::Release);
}

/// One document could not be put up: its page could not be written, or the engine did
/// not arrive at it.
///
/// The engine itself is fine, so this is not a failure to stand down for — the next
/// document is drawn as this one was meant to be. What it costs is the hover that was
/// waiting on it, and the notice is what takes that wait down.
fn note_document_failed() {
    FAILURE_NOTICE.store(true, Ordering::Release);
}

fn note_engine_up() {
    if let Ok(mut failed) = ENGINE_FAILED_AT.lock() {
        *failed = None;
    }
}

/// The page the engine is pointed at, written beside this run's state.
///
/// The document is not the page. A browser draws a standalone SVG at the size it asks
/// for — it does not stretch one to the window it was given — so a document given to the
/// engine as the page is drawn small in a large window, and the window this app lays out
/// is the share of the display `vector_scale` asked for: a document asking for 120 pixels is
/// a 120-pixel picture in a window half a display wide, or in a display-sized one at
/// fit-to-screen. What is given to the engine instead is this page: an image of the
/// document, in a box that is the whole page. An image *is* scaled to the box it is
/// given, whatever its own size is, which is the one thing that makes the window and the
/// document the same size at every scale.
///
/// Nothing is given up for that. An SVG drawn as an image is animated and not scripted,
/// and here nothing runs either — an image is a picture a browser shows, and a picture
/// cannot run code, take the pointer, or reach anything outside itself — not a file beside
/// it, not a URL — so a document that links to the world is drawn without it. A page of
/// HTML is the one document this engine runs, and it is not this page: that one is a frame
/// around the file (see `html_page`), with its own rule, which is the same reach and one
/// more thing given. Chromium parses a document as XML, in the mode built for animated
/// images, so the document is the document rather than a copy of it in a page of our own.
///
/// The version is the document's own modification time, which is what keeps an edited
/// file from being answered out of the browser's image cache: the URL changes when the
/// file does, and the same file at the same version is drawn again from memory.
///
/// The backdrop is the page's business for one of the four kinds: a checkerboard is
/// drawn by whatever composites the frame, and this engine composites its own — it can be
/// given a colour and nothing else — so the page paints the same squares this app's own
/// compositing draws, and the controller's colour stands behind them for the moment
/// before the page is up.
fn frame_page(
    path: &Path,
    version: u64,
    background: TransparentBackground,
) -> Option<(PathBuf, String)> {
    let document = file_url(path)?;
    // The backdrop is in the page's name as well as in its content, of the four kinds the
    // one that is a page's own — a checkerboard — is painted here rather than by the
    // controller: a change of backdrop is then a page the browser has not seen, rather than
    // the same URL answered out of its cache with the squares of the last one.
    let page = user_data_folder().join(format!("frame-{}.html", background.as_str()));

    let html = frame_html(&document, version, background);

    write_page(&page, &html, version)
}

/// What every page of this app's opens with: the doctype, the encoding, and a page that fills
/// the window it is given rather than scrolling inside it — the arrangement `html_page` exists
/// for, and the one a document drawn as an image needs as well.
///
/// The style is left open: each of the three kinds of page carries its own rules after it and
/// closes the same element.
const PAGE_HEAD: &str = "<!doctype html><meta charset=\"utf-8\"><title>preview</title>\
     <style>html,body{margin:0;padding:0;height:100%;overflow:hidden}";

/// The markup a document is drawn in, as a function of its URL rather than of its path, so
/// that what the refusal on it is can be asked of the markup without a page being written for
/// the question — one page per backdrop is one file, so two hovers of the same kind answer
/// from the same name by design, and two tests asking the same question of it at once would
/// be reading each other's.
///
/// The version is in the image's URL and in the page's, so neither is answered out of the
/// browser's cache with a document that has been written since it was read.
fn frame_html(document: &str, version: u64, background: TransparentBackground) -> String {
    format!(
        "{PAGE_HEAD}\
         img{{display:block;width:100%;height:100%;object-fit:contain}}</style>\
         <style>{no_interaction}</style>\
         {checkerboard}\
         <img src=\"{}?v={version}\" draggable=\"false\" alt=\"\">",
        escape_attribute(document),
        checkerboard = checkerboard_style(background),
        no_interaction = NO_INTERACTION_STYLE
    )
}

/// The page a page of HTML is drawn in, and the one it runs in: the file itself, whole, in a
/// frame that fills the window it is given — the arrangement `frame_page` reaches the same
/// end by with an image, for a document that has a size of its own.
///
/// The frame is given two allowances and nothing else. It is given its own origin, without
/// which the page itself would not arrive: a page's stylesheets and its pictures are relative
/// to it, so a frame loaded as a document of an opaque origin of its own comes up unstyled.
/// And it is given script, which is the one thing a page of HTML is handed the browser's own
/// engine for: a page that draws itself with WebGL, or lays itself out from a script, is a
/// page nothing but a run can show, and withheld it comes up a blank rectangle.
///
/// What is still withheld is everything that is a way *out* of the frame — popups, forms, and
/// any navigation of the top frame — which is the same reach a document drawn as an image has,
/// since a picture cannot pop up, submit, or navigate either, and which `BROWSER_ARGUMENTS`
/// reaches from the outside for the links a page keeps. Nor is the frame given the three
/// things a run would otherwise bring with it: a browser that plays sound without a gesture, a
/// page that reads the files beside it, and a page that takes the whole screen — no autoplay
/// policy is passed, no `--allow-file-access-from-files`, and no `allowfullscreen` for a page
/// asking to be shown full screen to be refused.
///
/// The two allowances together are not the loosening they would be for same-origin content,
/// where an allowance to keep one's own origin alongside one to run would let a framed document
/// reach out of itself and take the frame with it. They are not same-origin here: the wrapper
/// is a file this app wrote into this run's own profile folder and the document is another
/// file altogether, so the two are separate origins whatever the sandbox is told, and a
/// document that runs reaches its own file and no further out of it.
fn html_page(
    path: &Path,
    version: u64,
    background: TransparentBackground,
) -> Option<(PathBuf, String)> {
    let page_url = file_url(path)?;
    let page = user_data_folder().join(format!(
        "html-{}-{}.html",
        background.as_str(),
        path_identity(path)
    ));

    let html = format!(
        "{PAGE_HEAD}\
         iframe{{display:block;width:100%;height:100%;border:0}}</style>\
         {checkerboard}\
         <iframe src=\"{}?v={version}\" sandbox=\"allow-same-origin allow-scripts\" title=\"\"></iframe>",
        escape_attribute(&page_url),
        checkerboard = checkerboard_style(background)
    );

    write_page(&page, &html, version)
}

/// What a page of HTML is called in the page that draws it, from its own path.
///
/// A wrapper is one page however many targets it has been written for, so the target is
/// hashed into the name the browser caches by: a hover on one page and then on another is two
/// pages rather than one page rewritten under the browser that has already seen the first.
/// The version is what distinguishes one file from itself, not two files from each other.
fn path_identity(path: &Path) -> u64 {
    use std::collections::hash_map::DefaultHasher;
    use std::hash::{Hash, Hasher};

    let mut hasher = DefaultHasher::new();
    path.hash(&mut hasher);
    hasher.finish()
}

/// The page a font specimen is drawn in: the font itself, in the page through
/// `@font-face`, with the lines the file's own character map covers under a heading the
/// font's `name` table supplies.
///
/// One page per font, pointed at the file the engine can actually read — the font itself, or
/// the face `font_preview::browser_source` wrote out for a collection, which is the face the
/// specimen was read at — and every size in it is a viewport unit, so the window the layout
/// planned is the size the type is drawn at: the share of the display `font_scale` names is a
/// share of the specimen's size, the same relationship `object-fit: contain` gives a document.
fn font_page(
    path: &Path,
    version: u64,
    background: TransparentBackground,
    face: usize,
) -> Option<(PathBuf, String)> {
    let specimen = font_preview::probe_face(path, face)?;
    let source = font_preview::browser_source(path, &specimen, &user_data_folder())?;
    let font = file_url(&source)?;
    // The face is in the page's name as well as in its content, for the reason the backdrop
    // is: a specimen read at another face is a page the browser has not seen, rather than the
    // same URL answered out of its cache with the face before it.
    let page = user_data_folder().join(format!(
        "font-{}-{}.html",
        background.as_str(),
        specimen.face
    ));

    // The first line is the specimen's headline — the pangram wherever the font has Latin —
    // and the lines under it are the scripts the font also holds. Each line carries the
    // direction its own script is written in, which is what a browser settles it from anyway:
    // an Arabic or Hebrew line is then laid out from the right, its full stop ending it where
    // the script ends it rather than where a left-to-right page would, and a line of any
    // other script is drawn as it was.
    let html = specimen_html(
        &font,
        version,
        background,
        &specimen.title,
        &specimen.samples,
    );

    write_page(&page, &html, version)
}

/// The markup a specimen is drawn in, apart from what it is read from: the font's own URL,
/// the name its `name` table gave and the lines its character map covers. Split out for the
/// reason `frame_html` is — one page per backdrop and face is one file, so a question asked of
/// the markup should not be asked of a file two hovers are also writing.
fn specimen_html(
    font: &str,
    version: u64,
    background: TransparentBackground,
    title: &str,
    samples: &[String],
) -> String {
    let lines: String = samples
        .iter()
        .enumerate()
        .map(|(index, line)| {
            let class = if index == 0 { "pangram" } else { "script" };
            format!(
                "<p class=\"line {class}\" dir=\"auto\">{}</p>",
                escape_text(line)
            )
        })
        .collect();

    // A font that covers more of the sample lines than the box was shaped for — a pan-script
    // one, which is rare — is drawn smaller rather than past the bottom of it: the sizes above
    // are what a pangram and the handful of script lines a font usually holds want.
    let (pangram_size, script_size, script_margin) = if samples.len() > SPECIMEN_FULL_LINES {
        ("6vh", "3.6vh", "1vh")
    } else {
        ("8vh", "5vh", "1.6vh")
    };

    let (ink, shadow) = specimen_ink(background);
    format!(
        "{PAGE_HEAD}\
         body{{display:flex;flex-direction:column;justify-content:center;\
         padding:6vh 6vw;box-sizing:border-box;color:{ink};{shadow}}}\
         .title{{font-family:\"Segoe UI\",system-ui,sans-serif;font-size:2.4vh;\
         font-weight:600;opacity:.6;margin:0 0 2.4vh;white-space:nowrap;\
         overflow:hidden;text-overflow:ellipsis}}\
         .line{{font-family:\"RHPPreviewFont\",\"Segoe UI\",sans-serif;margin:0;\
         line-height:1.15}}\
         .pangram{{font-size:{pangram_size}}}\
         .script{{font-size:{script_size};margin-top:{script_margin}}}\
         </style>\
         <style>{no_interaction}</style>\
         <style>@font-face{{font-family:\"RHPPreviewFont\";\
         src:url(\"{font}?v={version}\")}}</style>\
         {checkerboard}\
         <div class=\"title\">{title}</div>{lines}",
        font = escape_attribute(font),
        title = escape_text(title),
        checkerboard = checkerboard_style(background),
        no_interaction = NO_INTERACTION_STYLE
    )
}

/// How many lines a specimen is drawn at the sizes above: the pangram and six lines under it
/// are what the specimen's box holds, and a font that covers more of the lines this app knows
/// than that — a pan-script font, which covers most of them — is drawn smaller instead.
const SPECIMEN_FULL_LINES: usize = 7;

/// Put a page where the engine will find it, and answer with the URL it is navigated to.
///
/// The version — the file's own modification time — goes into that URL, which is what keeps
/// an edited file from being answered out of the browser's cache: the URL changes when the
/// file does, and the same file at the same version is drawn again from memory.
fn write_page(page: &Path, html: &str, version: u64) -> Option<(PathBuf, String)> {
    std::fs::create_dir_all(user_data_folder()).ok()?;
    std::fs::write(page, html).ok()?;

    let url = format!("{}?v={version}", file_url(page)?);

    Some((page.to_path_buf(), url))
}

/// The squares a checkerboard backdrop is drawn as, for the one backdrop the engine cannot be
/// given: it takes a colour and nothing else, so the squares are painted by the page — over
/// the mid grey the controller is given, which is what the two square colours average to (see
/// `background_color`).
fn checkerboard_style(background: TransparentBackground) -> &'static str {
    const SQUARES: &str = "<style>html{background:#e0e0e0;background-image:\
         conic-gradient(#909090 25%,transparent 0 50%,#909090 0 75%,transparent 0);\
         background-size:32px 32px}</style>";

    match background {
        TransparentBackground::Checkerboard => SQUARES,
        _ => "",
    }
}

/// The refusal a looked-at document is drawn under: the page takes no pointer at all.
///
/// This is the second of the two refusals a picture in a browser needs, and the first is
/// `draggable="false"` on the drawing itself (`frame_page`). Neither is a setting of the
/// browser's — WebView2 has none for it — and both are the page's own, because the page is
/// the only part of this a hand can be kept away from.
///
/// What is refused is a *drag*, and the drag is what hangs a pinned preview. A browser drag
/// of an image is not a message this app is given: Chromium begins a drag session of its own
/// on the engine's thread, with the drawing's own silhouette following the pointer, and that
/// session is a modal loop holding the pointer. The pin takes the same pointer for its own
/// drag of the window at almost the same moment, and one pointer with two owners is a press
/// whose release reaches neither — the pin left holding a capture nothing will release, and
/// a browser's drag that never ends. So the page is drawn so that there is no press to
/// begin one with.
///
/// A specimen is refused the same way for the same reason, and is the one place the selection
/// matters: text a pointer can sweep across is text a pointer can begin a drag out of. It is
/// put here rather than written into either page because the two of them are the two kinds of
/// document that are *looked at*, and a page of HTML — the third kind, and the only one this
/// engine runs — is exactly the one that keeps its pointer (`html_page`).
const NO_INTERACTION_STYLE: &str = "*{pointer-events:none;user-select:none;\
     -webkit-user-select:none;-webkit-user-drag:none}";

/// The colour a specimen's text is drawn in over each backdrop, and the shadow that goes with
/// it.
///
/// Three of the four are a colour: light text on black, dark on white and on the checkerboard's
/// light squares. The one that is not is transparency, which has no colour to be read
/// against — so the glyphs are drawn light with a soft dark shadow behind them, and what a
/// specimen looks like over whatever the desktop happens to be is still readable.
fn specimen_ink(background: TransparentBackground) -> (&'static str, &'static str) {
    match background {
        TransparentBackground::Black => ("#f2f2f2", ""),
        TransparentBackground::White | TransparentBackground::Checkerboard => ("#1a1a1a", ""),
        TransparentBackground::Transparent => ("#f2f2f2", "text-shadow:0 0 .45vh rgba(0,0,0,.6);"),
    }
}

/// A URL as an attribute value: the two characters that would end it early.
fn escape_attribute(url: &str) -> String {
    url.replace('&', "&amp;").replace('"', "&quot;")
}

/// A string as the page's text: the characters that would end it, or open a tag in it. A
/// specimen's title is the font's own business rather than this app's — it is a name out of a
/// file — so it is escaped rather than trusted.
fn escape_text(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Ask the engine to draw `path` in a window at `area`.
///
/// The answer is immediate and says nothing about whether the document arrived: the
/// engine works on its own thread, and what it does with this is navigates, waits for
/// the document, and puts its window up — `showing_path` is what says which document got
/// there. What a caller puts on screen while it waits is the waiting spinner and nothing
/// else: this app draws no document, so there is no still frame to hold the place.
///
/// A file that is not a document at all is not the engine's to draw, and everything else
/// is the want in `WANTED`: the newest one wins, and what was asked for before it is not
/// drawn at all. The *same* document asked for a second time — which is what a wait that
/// follows the pointer does, sixty times a second — is the same want with the box it has
/// moved to, and it keeps the generation it was made under, so a pointer that keeps moving
/// while the document is on its way does not call off the navigation it is waiting for.
pub fn show(path: &Path, area: Area, background: TransparentBackground) {
    if !draws(path) {
        trace(&format!(
            "show({}): not a document the engine draws",
            path.display()
        ));
        return;
    }

    let generation = {
        let Ok(mut wanted) = WANTED.lock() else {
            trace("show: the wanted cell is poisoned");
            return;
        };

        let same_document = wanted
            .as_ref()
            .is_some_and(|wanted| wanted.path == path && wanted.background == background);

        let generation = if same_document {
            wanted.as_ref().map(|wanted| wanted.generation).unwrap_or(0)
        } else {
            WANTED_GENERATION.fetch_add(1, Ordering::AcqRel) + 1
        };

        *wanted = Some(Wanted {
            generation,
            path: path.to_path_buf(),
            background,
            area,
        });

        generation
    };

    let Ok(mut engine) = ENGINE.lock() else {
        trace("show: the engine's lock is poisoned");
        return;
    };

    let sender = engine.get_or_insert_with(Engine::start).sender.clone();

    if let Err(error) = sender.send(Command::Show { generation }) {
        trace(&format!(
            "show({}): the engine's thread is gone: {error}",
            path.display()
        ));
    }

    // A navigation the engine is in the middle of is waiting for a page that is no longer
    // wanted: the thread is woken rather than left to finish it (see `wake_engine_thread`).
    wake_engine_thread();
}

/// Take the engine's window down. The engine itself is kept warm: what it costs to
/// begin is a browser start, and what it costs to point at another document is a few
/// milliseconds, so a hover that follows another one pays almost nothing.
///
/// Kept warm is not the same as left working, and this is the other half of it: the
/// browser is told to stop while its window is off screen, so a document that runs costs
/// nothing until the next one is asked for (see `Host::hide`, `suspend`).
///
/// Nothing is wanted once this returns, and the generation goes with it: a navigation the
/// engine is in the middle of is one whose file the pointer has left, so it is dropped
/// rather than put up, and the window comes down without waiting for it.
pub fn hide() {
    if let Ok(mut wanted) = WANTED.lock() {
        *wanted = None;
    }
    WANTED_GENERATION.fetch_add(1, Ordering::AcqRel);

    let Ok(engine) = ENGINE.lock() else {
        return;
    };

    if let Some(engine) = engine.as_ref() {
        let _ = engine.sender.send(Command::Hide);
    }

    // Nothing is wanted any more, so a navigation in the middle of arriving is one nobody
    // is waiting for: the thread is woken rather than left to finish it.
    wake_engine_thread();
}

/// Take the engine's window down when the file it is drawing is a page of HTML: the switch
/// that asks for the page has just been turned off, and a page left standing would be a
/// window nothing takes down again until the pointer leaves (see `renders_html`).
///
/// The want is read as well as what is up: a switch turned while the engine is still
/// navigating to a page has no window yet, and the page that lands afterwards would stay.
pub fn hide_html_preview() {
    let watching_html = wanted().is_some_and(|want| text_formats::is_html_extension(&want.path));
    let showing_html = showing_path().is_some_and(|path| text_formats::is_html_extension(&path));

    if watching_html || showing_html {
        hide();
    }
}

/// Move a want that is already in hand to the box the wait has ended up in: what a pointer
/// that kept moving while the document was on its way asks for.
///
/// Nothing is sent to the engine — a document that is being navigated to is already drawn
/// in the box the want carries when it lands — and nothing is asked where the file is not
/// the one that is wanted: a box belongs to the hover that is waiting on it, and a hover
/// for another file is a `show`, which is the ask that takes the place of this want
/// altogether.
pub fn wanted_here(path: &Path, area: Area) {
    if let Ok(mut wanted) = WANTED.lock() {
        if let Some(wanted) = wanted.as_mut().filter(|wanted| wanted.path == path) {
            wanted.area = area;
        }
    }
}

/// Move the engine's window to another box, for a preview whose *own* window has changed
/// rather than for a pointer that moved: a pinned document is dragged, resized, maximized,
/// restored or carried to another display, and what stands in its media band has to travel
/// with the box.
///
/// It is both of the asks above in one, and either half applies (`box_change`). A document
/// still on its way has only a want to move — moving one sends no command, exactly as
/// `wanted_here` does — while one the engine is already holding has the window put in the new
/// box by a placement rather than by a `show`, which would navigate again for a document the
/// engine already has. The two are told apart by the window rather than by the want: a
/// document that has landed keeps its want, so a box that changed under one would be read as
/// a box still on its way and moved nowhere (see `PLACED`, `Host::place`).
///
/// Nothing is asked for another file: a box belongs to the preview it was measured for, and a
/// preview for another file is a `show`, which is the ask that takes this want's place
/// altogether. A box that is already the one asked for is nothing to do at all — a drag is
/// many of these.
pub fn place(path: &Path, area: Area, background: TransparentBackground) {
    // What is done with the box is read from the two cells the engine keeps and nothing else:
    // whether the document is the one *wanted*, whether it is the one *held*, and whether the
    // box is already the one on record for it. A drag is many of these, so the last of the three
    // is what keeps a box that has not moved from asking anything at all.
    let owed = owed(path);
    let holds = showing_path().is_some_and(|shown| shown == path);
    let same_box = WANTED
        .lock()
        .ok()
        .and_then(|wanted| {
            wanted
                .as_ref()
                .map(|wanted| wanted.path == path && wanted.area == area)
        })
        .unwrap_or(false);

    match box_change(owed, holds, same_box) {
        // A document still on its way has only a want to move, and moving one sends nothing: what
        // is drawn is drawn in the box the newest want asks for when it lands, so a box that
        // changed while a browser was coming up is a document that arrives in the right place and
        // an engine that is not asked again (see `wanted_here`).
        BoxChange::Want => wanted_here(path, area),
        // A document the engine is holding has a window that has to move with the box, and that
        // is a *move* and not a document being put on screen: the want is moved too, so a
        // navigation that comes later lands in the box the window is in, and the window itself is
        // asked for by a placement rather than by a `show` (see `ask_place`, `Host::place`).
        BoxChange::Window => {
            wanted_here(path, area);
            ask_place(path, area, background);
        }
        // A box that did not move, or a file the engine neither holds nor is owed: nothing to
        // move, and nothing asked (see `box_change`).
        BoxChange::Nothing => {}
    }
}

/// Publish a box for the engine's thread to move a held document's window into, and ask for it
/// once however many boxes have been published.
///
/// The cell is what makes a drag of any length cost one move: each of these replaces what was
/// there, so the placement the engine eventually takes up is the box the hand ended in. The
/// flag is what makes it cost that one move rather than one per pointer move — a drag can
/// publish boxes far faster than a single-threaded engine takes commands, and a document that
/// animates is a browser already busy enough, so an ask that finds one already outstanding
/// leaves the newer box in the cell and is answered by the ask already in the channel.
///
/// A send that finds no engine, or an engine whose thread has gone, is a placement nothing will
/// ever take up, so the flag is given back rather than left owed: a flag left owed is a window
/// that stops following the hand for the rest of the run (see `PLACE_ASKED`).
fn ask_place(path: &Path, area: Area, background: TransparentBackground) {
    if let Ok(mut placed) = PLACED.lock() {
        *placed = Some(Placement {
            path: path.to_path_buf(),
            background,
            area,
        });
    } else {
        return;
    }

    // One outstanding placement at a time, and the newest box travels in the cell above rather
    // than in the command, so this is the same one command however long the drag is.
    if PLACE_ASKED.swap(true, Ordering::AcqRel) {
        return;
    }

    let Ok(engine) = ENGINE.lock() else {
        PLACE_ASKED.store(false, Ordering::Release);
        return;
    };

    match engine.as_ref() {
        Some(engine) if engine.sender.send(Command::Place).is_ok() => {}
        _ => PLACE_ASKED.store(false, Ordering::Release),
    }
}

/// Which half of a `place` a box that changed under a document belongs to.
///
/// "Owed" cannot answer it on its own: a want is what the engine is owed, a document that
/// lands keeps it until the next file takes its place, and both a document still on its way and
/// the one being drawn are owed — so a box that changed under the second was answered as if it
/// were the first, and a pinned window left its document behind the moment it was dragged. What
/// separates them is the window, which is what `holds` is.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum BoxChange {
    /// Nothing to do: the box is the one on record, or the file is not the engine's.
    Nothing,
    /// The want is moved, and the engine is not asked: the document has not landed.
    Want,
    /// The engine's window is put in the new box, and the document is not navigated again.
    Window,
}

fn box_change(owed: bool, holds: bool, same_box: bool) -> BoxChange {
    if same_box {
        return BoxChange::Nothing;
    }

    if owed {
        // Both halves of what a window is: the document is on screen, and it is the one this
        // preview is for — a held document nobody wants is a want that has already moved on.
        return if holds {
            BoxChange::Window
        } else {
            BoxChange::Want
        };
    }

    BoxChange::Nothing
}

/// Publish a want for a document and take it back again, without asking for a browser.
///
/// What `owed` and `is_behind` answer is a question about the *want*, and a want is otherwise
/// only ever made by asking the engine for one — which in a test means a browser, and there is no
/// browser in a test. This is the want those questions are about and nothing else: what is owed,
/// and no engine behind it yet.
#[cfg(test)]
pub fn publish_want_for_test(path: &Path) {
    let generation = WANTED_GENERATION.fetch_add(1, Ordering::AcqRel) + 1;

    if let Ok(mut wanted) = WANTED.lock() {
        *wanted = Some(Wanted {
            generation,
            path: path.to_path_buf(),
            background: TransparentBackground::Black,
            area: Area {
                x: 0,
                y: 0,
                width: 8,
                height: 8,
            },
        });
    }
}

/// And the same want taken back, for the same reason (see `publish_want_for_test`).
#[cfg(test)]
pub fn clear_want_for_test() {
    if let Ok(mut wanted) = WANTED.lock() {
        *wanted = None;
    }

    WANTED_GENERATION.fetch_add(1, Ordering::AcqRel);
}

/// Let the engine go, window, browser process and thread together. Called when the app
/// ends.
pub fn shutdown() {
    let engine = ENGINE.lock().ok().and_then(|mut engine| engine.take());

    // A placement published for a window this call is about to destroy belongs to nothing, and
    // the flag has to go back with it — the thread that would have taken it up is joined below
    // and the next engine begins with no placement owed (see `drop_placement`).
    drop_placement();

    if let Some(engine) = engine {
        let _ = engine.sender.send(Command::Shutdown);
        let _ = engine.thread.join();
    }

    // The engine thread ends its browser as it goes. What this is for is the browser
    // that is still there anyway — the runtime's process is not one this app gets to
    // assume about — and the one this run started that could not be told apart from
    // an earlier engine's, and so was never recorded. A browser started by anything
    // else is not a child of this process, which is what keeps this from reaching
    // past this app's own.
    engine_processes::end_our_browsers();
}

/// What the preview thread asks the engine's thread to do.
enum Command {
    /// Draw what is wanted, as the want made under this generation: the document itself
    /// travels in `WANTED`, and the generation is what says whether this ask is still the
    /// newest one by the time the engine's thread takes it up (see `WANTED`).
    Show {
        generation: u64,
    },
    /// Move the window of a document the engine is already holding, to the newest box
    /// published in `PLACED`.
    ///
    /// The box is not carried in the command and that is the whole of it: a drag asks for a
    /// box on every pointer move, so carrying one per command would queue a move per move for
    /// a thread that takes one command per pass. The box travels in the cell instead, where
    /// asking again replaces what was there, so a drag of any length is answered by one move
    /// to where the hand ended up rather than by a backlog of where it was (see `PLACED`).
    Place,
    Hide,
    Shutdown,
}

struct Engine {
    sender: Sender<Command>,
    thread: std::thread::JoinHandle<()>,
}

impl Engine {
    /// Start the engine's thread. Nothing is created until the first document is asked
    /// for: a machine that never hovers one never starts a browser.
    fn start() -> Self {
        let (sender, receiver) = mpsc::channel();
        let thread = std::thread::spawn(move || engine_thread(receiver));

        Self { sender, thread }
    }
}

/// How long the engine is kept after its last document. It is read from the
/// configuration each time rather than captured, so an edit applies to the engine that
/// is already warm.
fn idle_timeout() -> Option<Duration> {
    CONFIG
        .lock()
        .map(|config| config.webview_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_WEBVIEW_IDLE_SECS))
        .as_duration()
}

/// Whether the browser is kept whatever the user is doing, which is the `Persistent` toggle
/// at the top of the same submenu.
///
/// Read rather than captured, for the reason the idle time is: it is asked while the engine
/// is up, so a click applies to the browser that is already warm.
fn engine_persistent() -> bool {
    CONFIG
        .lock()
        .map(|config| config.webview_persistent)
        .unwrap_or(false)
}

/// Write one line to the trace file when `RHP_WEBVIEW_TRACE` is set, for the same
/// reason the probes exist: an engine that does not come up says nothing on its own,
/// and this is what it says. Nothing is written when the variable is not set, so an
/// ordinary run leaves no file anywhere.
fn trace(message: &str) {
    if std::env::var_os("RHP_WEBVIEW_TRACE").is_none() {
        return;
    }

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(std::env::temp_dir().join("rhp-webview-trace.log"))
    {
        use std::io::Write;
        let _ = writeln!(file, "{message}");
    }
}

fn engine_thread(commands: Receiver<Command>) {
    // The engine's own apartment, and its own thread: WebView2 must be created on a
    // thread that is pumping messages, and this is that thread.
    let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

    // The thread's message queue, made before anything can be posted to it: a wakeup posted
    // to a thread that has none is dropped, and a want published while this thread is still
    // bringing its browser up would be the first thing that had to be waited for.
    {
        let mut message = MSG::default();
        unsafe {
            let _ = PeekMessageW(&mut message, None, WM_APP, WM_APP, PM_NOREMOVE);
        }
    }
    ENGINE_THREAD.store(unsafe { GetCurrentThreadId() }, Ordering::Release);

    // No placement is owed by a thread that has only just begun: whatever was published before
    // it was aimed at a window this engine has not made, and a flag left set by a previous
    // engine's thread would stop the next box from ever being asked for (see `PLACE_ASKED`).
    drop_placement();

    let mut host: Option<Host> = None;
    let mut idle_since = Instant::now();

    loop {
        // While there is nothing on screen the wait is a poll of the idle clock; while
        // a document is up it is a poll of the message queue, because a browser needs
        // the thread that made it to keep retrieving messages. With no engine at all
        // there is nothing to watch for, so the wait is long enough that only a document
        // wakes the thread, and an app left alone is one thread asleep on a channel.
        let wait = if SHOWING.load(Ordering::Acquire) {
            Duration::from_millis(5)
        } else if host.is_some() {
            Duration::from_millis(250)
        } else {
            Duration::from_secs(60 * 60)
        };

        match commands.recv_timeout(wait) {
            Ok(Command::Shutdown) => break,
            Ok(Command::Hide) => {
                if let Some(host) = host.as_mut() {
                    host.hide();
                }
                idle_since = Instant::now();
            }
            Ok(Command::Place) => carry_out_placement(&mut host),
            Ok(Command::Show { generation }) => {
                // What was asked for here may have been asked for after: the pointer moves
                // while a document is on its way, and the file this is about is then one
                // it has left. Such a want is not taken up at all — nothing is navigated
                // to, and nothing is shown — which is what keeps a folder of documents
                // from being drawn one after another at the speed of the hand crossing it.
                let ask = wanted().filter(|wanted| wanted.generation == generation);

                if ask.is_none() {
                    trace(&format!(
                        "engine: dropped a show the loop has moved on from (generation {generation})"
                    ));
                }

                if let Some(ask) = ask {
                    trace(&format!("engine: show {}", ask.path.display()));

                    if host.is_none() {
                        host = Host::create();

                        // An engine that could not be had is noted, so that hovers stop
                        // opening a box nothing will be drawn into until the folder it
                        // could not have is free again.
                        if host.is_some() {
                            note_engine_up();
                        } else {
                            note_failure();
                        }
                        trace(&format!("engine: host created: {}", host.is_some()));
                    }

                    if let Some(host) = host.as_mut() {
                        host.show(&ask);
                        trace(&format!("engine: shown: {}", is_showing()));
                    }

                    // A navigation the browser never answered is not a document to
                    // hand another one to: it stops answering rather than failing, and
                    // what it would put up next is the file before this one — so it is
                    // let go of here and begun again by the next document, which is
                    // what the app does with an engine that has stopped answering
                    // wherever else one is kept.
                    if host.as_ref().is_some_and(Host::is_hung) {
                        trace("engine: let go, the browser stopped answering");

                        if let Some(mut host) = host.take() {
                            host.close();
                        }
                    }
                    idle_since = Instant::now();
                }
            }
            Err(RecvTimeoutError::Timeout) => {}
            Err(RecvTimeoutError::Disconnected) => break,
        }

        pump_messages();

        // The engine draws three kinds, and every gate that could reach one switched off is a
        // browser held for nothing: while none can be shown no hover can be answered with a
        // document, a specimen or a page at all, so there is nothing warm to keep. It is read
        // here rather than being told because a gate can be closed either way — in the tray,
        // or in `config.ini` for the watcher to reload — and because the thread that would be
        // told is this one, parked on its channel. The browser's own children are its
        // business: ending it ends them.
        //
        // The text gate is one of them, but only while it asks for pages: a page of HTML is
        // a text file, so the kind's gate is what switches the browser off — and with the
        // kind on by default and the switch for pages off by default, the gate alone would
        // hold a browser for everyone who never previews a page (see `draws`).
        if host.is_some()
            && !PreviewType::Vector.enabled()
            && !PreviewType::Fonts.enabled()
            && !(PreviewType::Text.enabled() && renders_html())
        {
            trace("engine: let go, the kinds are switched off");

            if let Some(mut host) = host.take() {
                host.close();
            }
            idle_since = Instant::now();
        }

        // A document that has been off screen for as long as the engine is kept for is one
        // the engine is let go of, browser process and all.
        //
        // Which question is asked is the `Persistent` toggle at the top of the TTL submenu.
        // Persistent, it is the idle time, exactly as it was before the toggle existed; not
        // persistent — the way this starts — it is the AFK timer, and the idle time is not
        // consulted at all. What an idle time is for is the memory an engine holds while the
        // user is elsewhere, which is the question the AFK timer asks directly.
        //
        // The thread is not let go of with it, and that is the whole of it: the channel
        // it holds is the one a hover sends into, and a thread that ended here would
        // leave every document after the first idle timeout with a message nobody reads.
        // What the app would show is no document at all — the engine's window is the
        // whole of an SVG preview — for the rest of the run.
        let expired = match host.as_ref() {
            None => false,
            // A document on screen is the engine earning its keep: the window is the
            // preview, and the browser behind it is what is drawing it.
            Some(_) if SHOWING.load(Ordering::Acquire) => false,
            Some(_) if engine_persistent() => {
                idle_timeout().is_some_and(|limit| idle_since.elapsed() >= limit)
            }
            Some(_) => crate::app::afk::expired(),
        };

        if expired {
            trace("engine: let go after idle");

            if let Some(mut host) = host.take() {
                host.close();
            }
        }
    }

    if let Some(mut host) = host {
        host.close();
    }

    // The thread is going away, so nothing is posted to it any more: a want published after
    // this is one the next engine's thread takes up (see `ENGINE_THREAD`).
    ENGINE_THREAD.store(0, Ordering::Release);

    unsafe {
        windows::Win32::System::Com::CoUninitialize();
    }
}

/// Carry out the placement the engine's thread has been asked for, if it still belongs to a
/// window this engine has.
///
/// The box in hand is the newest one rather than the one the command was sent for: a drag
/// publishes a box per pointer move and only the last of them is a place worth putting the
/// window in (see `PLACED`).
fn carry_out_placement(host: &mut Option<Host>) {
    let Some(placement) = take_placement() else {
        return;
    };

    // A placement for a document the engine no longer holds is a drag of a
    // window that has been swapped or closed since, and moving that window
    // would put somebody else's document where the hand let go.
    if !host
        .as_ref()
        .is_some_and(|host| host.holds(&placement.path))
    {
        trace(&format!(
            "engine: dropped a placement for {} it does not hold",
            placement.path.display()
        ));
        return;
    }

    let Some(ask) = wanted().filter(|ask| ask.path == placement.path) else {
        return;
    };

    // A backdrop is the page's as well as the controller's, so a box
    // published under a backdrop the page was not written for is a page to
    // write again rather than a window to move — which is what `show` is
    // for, and it is asked for once rather than per pointer move, because
    // every box of a drag carries the backdrop that is on record.
    if host
        .as_ref()
        .is_some_and(|host| host.needs_page(&placement))
    {
        if let Some(host) = host.as_mut() {
            host.show(&ask);
        }

        return;
    }

    // And a move: nothing navigates, nothing is woken and nothing is
    // re-shown, all of which describe a document arriving rather than a
    // window being carried (see `Host::place`).
    if let Some(host) = host.as_mut() {
        host.place(&placement);
    }
}

/// Everything one engine is: the window it draws in, the environment, and the
/// controller over it.
struct Host {
    hwnd: HWND,
    environment: ICoreWebView2Environment,
    controller: ICoreWebView2Controller,
    webview: ICoreWebView2,
    /// The file the engine is holding, the backdrop its page was written for, and — for a font —
    /// the face of a collection it was read at, so a second hover on the same file is a window
    /// that is put back up rather than a navigation. A change of backdrop or of face is a page to
    /// write again, since three of the four backdrops are partly the page's own (the
    /// checkerboard's squares, a specimen's ink) and the page is what the browser caches, and a
    /// specimen of another face is another page and another font in it. The face is the index the
    /// page was written for, and `0` for a file that is not a collection.
    current: Option<(PathBuf, TransparentBackground, usize)>,
    /// The width and height the controller was last told to lay the page out at, and the
    /// window's own last box, so that a box which merely moved is not also re-laid-out.
    ///
    /// This is what makes a carried window cost one `SetWindowPos` rather than a bounds change
    /// as well: a page is laid out at the size of the window it is drawn in, so `SetBounds` is
    /// owed only by a box that changed *size*, and a drag that changes only where the window
    /// is has no reason to ask the browser to lay the document out again (see `place`).
    last_area: Option<(i32, i32)>,
    /// The browser process this engine started, when it could be told which one it
    /// was. The runtime owns the browser, but the process is this app's own child —
    /// started by the loader in this process — and it is what is ended if it is
    /// somehow still there when the engine goes.
    browser_pid: u32,
    /// Whether the browser has stopped answering a navigation for longer than any document
    /// takes. Such an engine is let go of by the thread that holds it rather than handed
    /// the next document: what it would answer with is the file before this one, and what
    /// its thread would do is wait on a page that is never coming (see `NAVIGATION_TIMEOUT`).
    hung: bool,
    /// Whether the browser has been asked to suspend and is to be resumed before the next
    /// document is drawn. It is kept here rather than read back from the browser because
    /// both questions are about this host's own last word: a runtime older than
    /// suspension is a host that never suspends, and a resume of a WebView that is not
    /// suspended is harmless (see `suspend`, `wake`).
    ///
    /// It is also the record of the invariant `hide` keeps: after a hide returns, the
    /// browser has been asked to stop, whatever was on screen and whatever was. Nothing
    /// is read to decide whether the ask is owed, so a hide that has already been had
    /// does not ask again, and a browser that is stopped is a browser `wake` knows to put
    /// back to work before it navigates.
    suspended: bool,
}

impl Host {
    fn create() -> Option<Self> {
        register_class();

        // What the browser is running as before this engine is begun, so the one it
        // starts can be told from one an earlier engine of this same run left
        // closing: both are children of this process, and only the one that was not
        // there a moment ago is this engine's.
        let before = engine_processes::processes_named_by_parent(
            engine_processes::BROWSER_IMAGE,
            std::process::id(),
        );

        let folder = user_data_folder();
        std::fs::create_dir_all(&folder).ok()?;

        let window = create_host_window();
        trace(&format!(
            "host: window {:?} (last error {})",
            window.is_some(),
            windows::core::Error::from_win32().code().0
        ));
        let hwnd = window?;

        // One deadline for the whole attempt at having an engine: what is waited on is the
        // runtime answering at all, and one that has not answered by then is not going to
        // (see `HOST_CREATION_TIMEOUT`).
        let deadline = Instant::now() + HOST_CREATION_TIMEOUT;

        let started = Instant::now();
        let environment = create_environment(&folder, deadline);
        trace(&format!(
            "host: environment {:?} in {} ms",
            environment.is_some(),
            started.elapsed().as_millis()
        ));
        let environment = environment?;
        let environment_ms = started.elapsed().as_millis() as u64;

        let started = Instant::now();
        // A controller is refused while the folder this engine keeps its state in is
        // held by another browser — which is what a browser left behind by an earlier
        // run looks like, and it is usually gone within a moment. Asking again a few
        // times is worth more than the wait it costs the engine's own thread, and the
        // caller falls back to this app's reader if even that comes to nothing. A
        // browser that is *silent* rather than refusing spends the deadline of the
        // attempt instead of being asked again on a fresh clock, which is why the
        // deadline is one of its own and not a per-call timeout.
        //
        // The first ask is made whatever the deadline says, since `create_controller` is
        // what spends it; the ones after it are put off a quarter of a second further
        // each time, and never past the end of the attempt, up to four asks.
        let mut controller = create_controller(environment.clone(), hwnd, deadline);
        let mut asked = 1;

        while controller.is_none() && asked < 4 && Instant::now() < deadline {
            std::thread::sleep(Duration::from_millis(250 * asked).min(remaining(deadline)));

            controller = create_controller(environment.clone(), hwnd, deadline);
            asked += 1;
        }

        trace(&format!(
            "host: controller {:?} in {} ms",
            controller.is_some(),
            started.elapsed().as_millis()
        ));
        let controller = controller?;
        let controller_ms = started.elapsed().as_millis() as u64;

        let webview = unsafe { controller.CoreWebView2().ok()? };
        configure(&webview);

        // The browser is the runtime's, but the process is this app's: the loader
        // started it here, which is what makes it findable by its parent, and what it
        // is put in the job and written down for is the same thing the Office engines
        // are — a process that must not outlive the app that started it, however the
        // app ends.
        let browser_pid = engine_processes::processes_named_by_parent(
            engine_processes::BROWSER_IMAGE,
            std::process::id(),
        )
        .into_iter()
        .find(|pid| !before.contains(pid))
        .unwrap_or(0);

        if browser_pid != 0 {
            engine_processes::record(engine_processes::BROWSER_IMAGE, browser_pid);
        }
        trace(&format!("host: browser {browser_pid}"));

        if let Ok(mut timings) = LAST_TIMINGS.lock() {
            timings.environment_ms = environment_ms;
            timings.controller_ms = controller_ms;
        }

        let host = Self {
            hwnd,
            environment,
            controller,
            webview,
            current: None,
            last_area: None,
            browser_pid,
            hung: false,
            suspended: false,
        };

        // The window a hit test compares against, for the preview loop and the Explorer
        // hook: it is the window a document is drawn in, so it is the one the pointer
        // touches when it is taken onto a document's preview.
        HOST_HWND.store(hwnd.0 as isize, Ordering::Release);

        Some(host)
    }

    /// Draw the document this want names in the box it asks for, or move the window to it
    /// where the engine already holds that document.
    ///
    /// A file the engine is not already holding, or one whose backdrop or face has changed,
    /// is navigated to *before* the window is put up: a window shown first would be the file
    /// before it, and what is on screen a moment ago belongs to another hover. What the
    /// navigation ends with is the whole of what happens next — the four ways it can end are
    /// [`Arrival`], and three of them put nothing up at all:
    ///
    /// - the page arrived, and the window is put up in the box the *newest* want asks for,
    ///   which is where a wait that followed the pointer ended up rather than where it
    ///   started;
    /// - a newer want took this one's place while the page was on its way: a file the
    ///   pointer has left is not shown, and the window comes down, because what the window
    ///   holds is the file before it and the loop is waiting for another one;
    /// - the page could not be written or the engine would not navigate: the engine is fine
    ///   and the hover waiting on this one has nothing left to wait for;
    /// - the browser never answered at all, which is the one that leaves the engine unfit to
    ///   be handed the next document (see `NAVIGATION_TIMEOUT`).
    fn show(&mut self, wanted: &Wanted) {
        let Wanted {
            generation,
            path,
            background,
            area,
        } = wanted;
        let (background, area) = (*background, *area);

        // Which face of a collection a specimen of this file is drawn from, read here rather
        // than inside the page: it is part of what the engine is holding, so a setting changed
        // between two hovers of one file is a page to write and navigate to again. A file that
        // is not a font is drawn from no face at all.
        let face = if font_formats::is_font_file(path) {
            font_preview::configured_face()
        } else {
            0
        };

        // A browser that was told to stop when the last document was taken down is put back
        // to work here: before anything is navigated to, and before the controller is told
        // it is on screen, which is the order the runtime documents for a resume. Either of
        // those two would wake it anyway — a `Navigate` resumes a suspended WebView, and so
        // does making it visible — so nothing about the document turns on the order. The
        // resume is made explicitly rather than left to either of them so that the state this
        // app keeps is the state the browser is in, and the flag on `Host` cannot come to
        // disagree with it (see `suspend`, `wake`).
        self.wake();

        unsafe {
            // The background is a setting of the controller rather than of the page,
            // and it belongs to the interface that added it.
            if let Ok(controller) = self
                .controller
                .cast::<webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Controller2>()
            {
                let _ = controller.SetDefaultBackgroundColor(background_color(background));
            }

            // The box the page is rendered in, set before it is navigated to: a document
            // is drawn at this size, so what arrives is what the layout asked for rather
            // than something drawn small and resized after the fact. It is set again
            // below from the box the *newest* want asks for, which is where a wait that
            // followed a moving hand ended up.
            self.set_bounds(area);
        }

        if self.current.as_ref() != Some(&(path.clone(), background, face)) {
            // What the window is holding is *kept* while the next document is on its way:
            // a window taken down first would be a preview that vanishes and comes back,
            // which is worse than the file before it standing there for the few
            // milliseconds a navigation takes — and the file before it is one the pointer
            // has left only if the hook has said so, which is a hide of its own that
            // arrives here and takes the window down without waiting for this navigation.
            let started = Instant::now();
            let arrival = self.navigate(path, background, face, *generation);
            trace(&format!(
                "engine: navigate {} {arrival:?} in {} ms",
                path.display(),
                started.elapsed().as_millis()
            ));

            match arrival {
                Arrival::Arrived => {
                    if let Ok(mut timings) = LAST_TIMINGS.lock() {
                        timings.navigate_ms = started.elapsed().as_millis() as u64;
                    }

                    self.current = Some((path.clone(), background, face));
                }
                Arrival::Superseded => {
                    // Another want has taken this one's place, so what was navigated to is
                    // a file the pointer has left: nothing of it goes on screen. What the
                    // engine holds is nothing either — the page this navigation was made
                    // against has been written over with another document's — so the next
                    // want is navigated to rather than moved to.
                    self.current = None;
                    return;
                }
                Arrival::Failed => {
                    self.current = None;
                    note_document_failed();
                    return;
                }
                Arrival::TimedOut => {
                    // A browser that never answered the navigation is not a browser to
                    // hand the next document to: the hover waiting on this one is answered
                    // with nothing, and the engine is let go of and begun again (see
                    // `NAVIGATION_TIMEOUT` and `is_hung`).
                    self.hung = true;
                    self.current = None;
                    note_document_failed();
                    return;
                }
            }
        }

        // The box the *newest* want asks for, which is the one the wait has ended up in:
        // a wait that followed a moving pointer keeps following it, and a document that
        // takes a moment to be drawn lands where the hand is rather than where it was.
        let area = wanted_area(*generation).unwrap_or(area);

        // What the window is being asked to show decides what it may be activated for: a
        // page of HTML that runs is a program the user clicks into, and every other
        // document is a picture the caret is not to be taken out of for. The style is set
        // here rather than at the window's making, because the window is made once and
        // drawn in for every document this engine is given — and it is set before the
        // window is put up, so a click that lands on the first frame already has the
        // answer it is going to get.
        let runs = page_runs(path);
        RUNNING_DOCUMENT.store(runs, Ordering::Release);

        // The bounds this window was given are recorded, so that a box which only moves from
        // here is not also re-laid-out (see `place`).
        self.last_area = Some((area.width, area.height));

        unsafe {
            self.set_bounds(area);

            let current = GetWindowLongPtrW(self.hwnd, GWL_EXSTYLE);
            let style = ex_style_for(current, runs);
            if style != current {
                SetWindowLongPtrW(self.hwnd, GWL_EXSTYLE, style);
            }

            let _ = self.controller.SetIsVisible(true);

            self.put_window(area);
            let _ = ShowWindow(self.hwnd, SW_SHOWNOACTIVATE);
        }

        publish(Some((path.clone(), area)));
    }

    /// Put the window in `area`, topmost, and without ever activating it.
    ///
    /// The placement never activates, whatever the document is: a hover is not a click, and a
    /// preview that appeared over a window being named would put the caret somewhere else
    /// than where it was.
    fn put_window(&self, area: Area) {
        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                HWND_TOPMOST,
                area.x,
                area.y,
                area.width,
                area.height,
                SWP_NOACTIVATE | SWP_SHOWWINDOW,
            );
        }
    }

    /// Tell the controller the box the page is rendered in, which is what the browser lays the
    /// page out at: a page is drawn at the size of the window it is drawn in, so this is owed
    /// by a box that changed *size* and by nothing else (see `place`).
    fn set_bounds(&self, area: Area) {
        unsafe {
            let _ = self.controller.SetBounds(RECT {
                left: 0,
                top: 0,
                right: area.width,
                bottom: area.height,
            });
        }
    }

    /// Whether the document this host is holding is `path`, which is what a placement is
    /// answered for: a box published for a document the engine has since swapped away from
    /// belongs to no window it has, and moving whatever window it does have would put a
    /// document the pointer has left where the hand let go.
    fn holds(&self, path: &Path) -> bool {
        self.current
            .as_ref()
            .is_some_and(|(held, _, _)| held == path)
    }

    /// Whether a placement under this backdrop is a page to write again rather than a window
    /// to move.
    ///
    /// Three of the four backdrops are partly the page's own — the checkerboard's squares, a
    /// specimen's ink — and the page is what the browser caches, so a backdrop the page was
    /// not written for is a different page and a different URL rather than a different colour
    /// on the window. A placement is published with whatever backdrop is on record at the time
    /// it is asked for, so a tray switch behind a pin arrives here as one placement that is a
    /// page to write, and as a move for every pointer move after it.
    fn needs_page(&self, placement: &Placement) -> bool {
        self.current
            .as_ref()
            .is_some_and(|(_, held, _)| *held != placement.background)
    }

    /// Move the window of the document already on screen into a box that changed under it.
    ///
    /// This is deliberately the *whole* of what a moved box costs, and the difference from
    /// `show` is the point of it. A `show` is a document being put up: it wakes a suspended
    /// browser, sets the controller's colour, sets its bounds, tells it the window is on
    /// screen and shows the window. A drag of a pinned window asked for that on every pointer
    /// move, and the engine's thread takes one command per pass — so the queue of full shows
    /// grew faster than it could be drained and never was. A document that animates is what
    /// made that fatal rather than merely slow: its compositor is already working, so the
    /// thread fell so far behind that the drawing stopped following the window, and then
    /// stopped answering at all.
    ///
    /// So a box that merely moved costs one `SetWindowPos` and nothing else, which is all a
    /// window being carried has ever cost on this side (see `apply_pin_drag` in
    /// `preview_window`). `SetBounds` is owed only by a box that changed *size*, because a
    /// page is laid out at the size of the window it is drawn in; and nothing here wakes,
    /// re-shows or re-asserts visibility, all of which describe a document arriving rather
    /// than a window being carried. The window is kept above the pin's own, which is what a
    /// drag needs — the pin is raised on every move and would otherwise cover the drawing it
    /// is carrying.
    fn place(&mut self, placement: &Placement) {
        let size = (placement.area.width, placement.area.height);
        let resized = self.last_area != Some(size);
        self.last_area = Some(size);

        if resized {
            self.set_bounds(placement.area);
        }

        self.put_window(placement.area);
        publish_rect(Some(placement.area));
    }

    /// Whether the browser has stopped answering, which is a host to be let go of rather
    /// than one to hand another document to.
    fn is_hung(&self) -> bool {
        self.hung
    }

    /// Take the window down, and tell the browser that nothing is looking at it.
    ///
    /// `ShowWindow(SW_HIDE)` is only half of what a window leaving the screen means to a
    /// browser: a controller still told that it is on screen keeps the compositor making
    /// frames for it, and a document that runs goes on running at it for the rest of the
    /// session, so the browser is asked to suspend as well (see `suspend`). That is asked
    /// of every kind of document and not only of a page of HTML, because the frames are
    /// the browser's rather than the page's: an animated SVG is composited just as hard.
    ///
    /// There is deliberately no check of whether the window was ever put up, because a
    /// browser can be awake with its window hidden: a `show` that puts the browser back to
    /// work and then ends as a navigation that was superseded or failed returns before the
    /// controller is told it is on screen, so what is left is a browser drawing a window
    /// nobody is looking at (see `show`). A hide that took the absence of a visible window
    /// as an answer to the question of what the browser is doing would skip the one ask
    /// that stops it, and the engine is kept warm between documents, so nothing else would
    /// ever come along and do it. What not guarding costs is two Win32 calls and the one
    /// `Interface::cast` in `suspend`.
    fn hide(&mut self) {
        unsafe {
            let _ = ShowWindow(self.hwnd, SW_HIDE);
            // `TrySuspend` states a precondition rather than a preference: the controller's
            // `IsVisible` must already be false when it is called, and otherwise the call
            // fails outright with `HRESULT_FROM_WIN32(ERROR_INVALID_STATE)` rather than
            // answering no. So the controller is hidden first, and this half of the hide is
            // never made conditional on a runtime that can be asked to suspend — it is the
            // half that stops the frames whether or not the ask is ever answered, including
            // on a runtime too old to be asked at all (see `suspend`).
            let _ = self.controller.SetIsVisible(false);
        }

        // These are what the preview loop reads as the preview still being there — through
        // `is_showing`, `showing_path` and `screen_rect` — so they are published before the
        // ask rather than after it: `suspend` waits, and can wait for up to
        // `SUSPEND_TIMEOUT`, and publishing first is what keeps a window in which the
        // preview has been taken down and the app is still reporting one under the pointer.
        publish(None);

        // The window is off screen, so the box a drag last asked for is a box nothing is in:
        // taking it back is what stops a placement published before this hide being carried out
        // against a window that is no longer showing, and it is the same reason the publishes
        // above happen first.
        drop_placement();

        self.suspend();
    }

    /// Ask the browser to stop working while nothing is looking at its window.
    ///
    /// What this is for is the engine being kept warm between documents (see the module
    /// docs): the browser is there to be pointed at the next one, and a browser left
    /// rendering a window that is off screen spends CPU and GPU on a picture no one is
    /// seeing. `TrySuspend` is the runtime's own answer to that — it halts rendering and
    /// throttles the page's script timers, which is what a page that runs needs, and it
    /// does it for an animated document as much as for a page.
    ///
    /// The ask is waited for, and waited for within `SUSPEND_TIMEOUT`, rather than left to
    /// complete whenever it completes (see `wait_for_suspend`). A runtime older than
    /// suspension is answered by leaving the controller hidden and doing nothing else —
    /// `SetIsVisible(false)` has already stopped the frames, and that half is the one that
    /// is not optional.
    ///
    /// It is called on every hide rather than only on a hide of something on screen, so the
    /// guard at the top of it is what keeps that cheap: a browser that has already been asked
    /// to stop is asked nothing further, however many hides arrive behind it and whether or
    /// not a window was ever put up in between.
    fn suspend(&mut self) {
        if self.suspended {
            return;
        }

        let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() else {
            trace("host: this runtime does not suspend; the controller is hidden and no more");
            return;
        };

        trace("host: asked the browser to suspend");
        let (sender, receiver) = mpsc::channel::<bool>();

        unsafe {
            // The handler is let go of when this returns, and that is safe: the browser
            // holds a reference of its own to it, so an answer that arrives after the wait
            // has given up goes into a channel nobody is reading, and a send into a
            // receiver that has been dropped is the error this ignores.
            let handler =
                TrySuspendCompletedHandler::create(Box::new(move |_code, is_successful| {
                    let _ = sender.send(is_successful);
                    Ok(())
                }));

            if webview.TrySuspend(&handler).is_err() {
                trace("host: the browser would not be asked to suspend");
                return;
            }
        }

        // Whether the browser said yes, said no, or said nothing at all, the WebView is
        // recorded as suspended: what the flag is for is not asking again while the browser
        // stays where this left it, however many hides arrive behind it, and `wake` resumes
        // a browser that may not really have suspended — harmlessly, its result being
        // ignored. A refusal is therefore retried on the next hide, which is deliberate: a
        // browser that refused because of something transient — a script dialog left open,
        // say — is quite likely to say yes the next time it is asked (see `wake`).
        match wait_for_suspend(&receiver) {
            Some(true) => trace("host: the browser is suspended"),
            Some(false) => trace("host: the browser refused to suspend"),
            None => trace("host: the browser did not answer the suspend in time"),
        }

        self.suspended = true;
    }

    /// Put the browser back to work, which is what a page that runs needs before the next
    /// document is drawn: a resume rather than a browser begun again, and a page that picks
    /// up where it left off rather than an engine thrown away and started (see `suspend`).
    ///
    /// The order it is called in — resume first, and the controller told it is on screen
    /// after it — is the runtime's documented one, and nothing turns on it here: `Navigate`
    /// resumes a suspended WebView of its own accord, and so does making it visible. The
    /// resume is made explicitly rather than left to either of those so that the state this
    /// app keeps is the state the browser is in, and the flag on `Host` cannot come to
    /// disagree with it (see `show`).
    ///
    /// The flag is cleared before the call rather than after it, so a runtime with nothing to
    /// resume cannot leave the host believing it has a suspended WebView to wake before every
    /// document — and a resume of a WebView that was never suspended is harmless, its result
    /// being ignored either way.
    ///
    /// Nothing asserts that order: `suspend`, this and `wait_for_suspend` all need a live
    /// COM object to be asked against, so the invariant is on whoever changes this next rather
    /// than on the suite.
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
    }

    fn close(&mut self) {
        self.hide();
        self.current = None;

        // A browser that stopped answering is one whose close may never return: the call is
        // a message to that browser, and a browser that has stopped taking messages is what
        // a hung host is. So the process is ended first — by the id it was recorded under,
        // which is the same verified end every other engine of this app gets — and the close
        // that follows is the close of something that is already gone rather than a wait on
        // it. A browser `hide` suspended is closed the same way as one that was not, and
        // nothing here waits for it to be woken first: a suspend is a browser that has
        // stopped drawing rather than one that has stopped answering, and either way the
        // close is this engine's last word — what did not take the browser down is ended
        // by `Drop`, which is the path a suspended browser is answered on as well.
        if self.hung && self.browser_pid != 0 {
            if engine_processes::is_running(self.browser_pid) {
                engine_processes::terminate_owned(self.browser_pid);
            }
            engine_processes::forget(self.browser_pid);
        }

        unsafe {
            let _ = self.controller.Close();
            let _ = DestroyWindow(self.hwnd);
        }

        // The window is gone, so nothing under the pointer is this engine's any more: a
        // hit test that still held the handle would be reading a window that was destroyed.
        HOST_HWND.store(0, Ordering::Release);

        // The environment goes with the host when it is dropped, and the browser
        // process it owns goes with the last controller over it.
        let _ = &self.environment;
    }

    /// Point the engine at `path` and wait for the page to arrive, pumping the thread's
    /// messages while it does.
    ///
    /// What it is pointed at is a page of this app's rather than the file itself, and which
    /// page is what the file is: a document's is the document as an image, which is what makes
    /// it the size of the window, a font's is the font in the page with its own lines under
    /// it — read at `face`, which is which of a collection's faces is written out; see
    /// `frame_page` and `font_page`. A page of HTML is the third of those kinds of document:
    /// its page is a frame of its own around the file, so the page is laid out by the browser
    /// rather than drawn by the app, and what it asks for is whatever box it is given (see
    /// `html_page`).
    ///
    /// The wait is asked under the generation this navigation was made for, and it ends on one
    /// of the four [`Arrival`]s: the page arriving, a newer want taking this one's place, the
    /// browser not answering within `NAVIGATION_TIMEOUT`, or the navigation not being made at
    /// all. What is *not* done is waiting without a bound — see `wait_for_navigation`.
    fn navigate(
        &self,
        path: &Path,
        background: TransparentBackground,
        face: usize,
        generation: u64,
    ) -> Arrival {
        let version = file_version(path);
        let page = if font_formats::is_font_file(path) {
            font_page(path, version, background, face)
        } else if text_formats::is_html_extension(path) {
            html_page(path, version, background)
        } else {
            frame_page(path, version, background)
        };

        let Some((page, url)) = page else {
            return Arrival::Failed;
        };

        trace(&format!(
            "engine: page {} for {}",
            page.display(),
            path.display()
        ));
        let url = wide(&url);
        let (sender, receiver) = mpsc::channel();

        // Whether the browser runs what it is about to be given is settled here, before the
        // navigation rather than after it, because the document is read as it loads: a page
        // that draws itself would be a blank canvas by the time a setting asked for after
        // the fact could reach it (see `set_scripts`).
        set_scripts(&self.webview, page_runs(path));

        unsafe {
            let handler =
                NavigationCompletedEventHandler::create(Box::new(move |_sender, _args| {
                    let _ = sender.send(());
                    Ok(())
                }));

            let mut token = EventRegistrationToken::default();
            if self
                .webview
                .add_NavigationCompleted(&handler, &mut token)
                .is_err()
            {
                return Arrival::Failed;
            }

            let started = self.webview.Navigate(PCWSTR(url.as_ptr()));
            let arrival = if started.is_err() {
                Arrival::Failed
            } else {
                wait_for_navigation(&receiver, generation)
            };

            let _ = self.webview.remove_NavigationCompleted(token);

            arrival
        }
    }
}

/// How a navigation ended, as the four things `Host::show` does something about.
#[derive(Clone, Copy, Debug)]
enum Arrival {
    /// The page arrived: the document is the one the engine is holding.
    Arrived,
    /// The page could not be written, or the engine would not navigate at all.
    Failed,
    /// A newer want took this one's place while the navigation ran: what is being navigated
    /// to is a file the pointer has left, so nothing of it goes up.
    Superseded,
    /// The browser stopped answering: the navigation was given longer than any document
    /// takes and nothing came back (see `NAVIGATION_TIMEOUT`).
    TimedOut,
}

/// How long a navigation may run before the browser is read as having stopped answering.
///
/// What the wait is for is `NavigationCompleted`, and that event is the one thing in this
/// module with no bound of its own: a navigation it never fires for leaves the engine's
/// thread waiting on it for the rest of the run, with every document after it queued behind
/// a wait that never ends, the idle timeout unable to fire because the thread that holds it
/// is the thread that is waiting, and the window standing there with the file before this
/// one on it. A page of this app's own is a file on the machine and a document a browser has
/// to lay out — milliseconds warm, a few hundred with the browser being begun — so what is
/// past this is not a document still being drawn: it is a browser that has stopped
/// answering, and it is let go of and begun again by the next document.
const NAVIGATION_TIMEOUT: Duration = Duration::from_secs(15);

/// How long the browser is given to answer a suspend before the engine reads it as having
/// stopped rather than as having answered.
///
/// What the wait is for is `TrySuspendCompleted`, which the runtime calls as soon as it has
/// stopped, so a quarter of a second is longer than a suspend takes on any machine and short
/// enough that a pointer moving on over a document never notices it. A browser that does not
/// answer is recorded as suspended anyway — the answer is what says it has stopped, not the
/// asking — and a resume of a WebView that is not suspended is harmless (see `pump_until`).
const SUSPEND_TIMEOUT: Duration = Duration::from_millis(250);

/// Why a wait ended without an answer: the bound ran out, or the message queue is gone.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Unanswered {
    /// The wait ran longer than the answer takes.
    TimedOut,
    /// There was no queue left to wait on.
    Gone,
}

/// Pump this thread's messages until `probe` answers, `timeout` runs out, or the queue is
/// gone — whichever comes first.
///
/// The messages have to be pumped rather than waited on: a controller is created on this
/// thread and stops drawing when the thread stops retrieving messages, and both of the events
/// this waits for arrive through that same queue — so the wait *is* a `GetMessage`, and an
/// answer is noticed the moment it lands, with no interval between and nothing polled.
///
/// The timer is armed on the thread rather than on a window, and under an id of its own so
/// that two waits cannot answer one another's timer: its whole job is to bring `GetMessage`
/// back so that the elapsed check is read again. The bound is therefore read on every pass,
/// which is what stops a browser that never answers from parking the engine's thread for the
/// rest of the run.
fn pump_until<T>(
    timeout: Duration,
    timer: usize,
    mut probe: impl FnMut() -> Option<T>,
) -> Result<T, Unanswered> {
    let started = Instant::now();
    let mut message = MSG::default();

    unsafe {
        let _ = SetTimer(HWND::default(), timer, timeout.as_millis() as u32, None);
    }

    let answer = loop {
        if let Some(answer) = probe() {
            break Ok(answer);
        }

        if started.elapsed() >= timeout {
            break Err(Unanswered::TimedOut);
        }

        let mut answered = false;
        unsafe {
            let retrieved = GetMessageW(&mut message, HWND::default(), 0, 0);
            if retrieved.0 > 0 {
                let _ = TranslateMessage(&message);
                DispatchMessageW(&message);
                answered = true;
            }
        }

        if !answered {
            // `GetMessage` answered -1, an error, or 0, a quit: there is no queue left to
            // wait on, and an answer that has not arrived by then is not coming.
            break Err(Unanswered::Gone);
        }
    };

    unsafe {
        let _ = KillTimer(HWND::default(), timer);
    }

    answer
}

/// Wait for the page to arrive, pumping the thread's messages while it does, and ending on
/// one of the three things that can end the wait.
///
/// The three are the page arriving, a want that is no longer this one — which
/// `wake_engine_thread` brings this thread back to look at — and a browser that has not
/// answered within `NAVIGATION_TIMEOUT` (see `pump_until`).
fn wait_for_navigation(receiver: &Receiver<()>, generation: u64) -> Arrival {
    const NAVIGATION_TIMER: usize = 1;

    match pump_until(NAVIGATION_TIMEOUT, NAVIGATION_TIMER, || {
        if receiver.try_recv().is_ok() {
            Some(Arrival::Arrived)
        } else if !is_wanted(generation) {
            Some(Arrival::Superseded)
        } else {
            None
        }
    }) {
        Ok(arrival) => arrival,
        Err(Unanswered::TimedOut) => Arrival::TimedOut,
        Err(Unanswered::Gone) => Arrival::Failed,
    }
}

/// Wait for the browser to answer a suspend, pumping the thread's messages while it does.
///
/// The wait is a `GetMessage` for the reason the one in `wait_for_navigation` is (see
/// `pump_until`). The answer is `Some` when the browser gave one, and `None` when it gave
/// nothing before `SUSPEND_TIMEOUT` or when the queue is gone, both of which the caller reads
/// as the same thing (see `suspend`).
fn wait_for_suspend(receiver: &Receiver<bool>) -> Option<bool> {
    const SUSPEND_TIMER: usize = 2;

    pump_until(SUSPEND_TIMEOUT, SUSPEND_TIMER, || receiver.try_recv().ok()).ok()
}

impl Drop for Host {
    fn drop(&mut self) {
        publish(None);
        HOST_HWND.store(0, Ordering::Release);

        // Closing the engine drops the environment, and the browser goes with the
        // last controller over it — usually. A browser that does not is the leftover
        // the profile folders are named for, still holding the folder of a run whose
        // app is gone, so the process is asked about here and ended if it is still
        // there: it is one this app started, and nothing of it should outlive the
        // app. Its own children are its business — ending it ends them.
        if self.browser_pid != 0 {
            if engine_processes::is_running(self.browser_pid) {
                engine_processes::terminate_owned(self.browser_pid);
            } else {
                engine_processes::forget(self.browser_pid);
            }
        }
    }
}

/// The window the engine draws into: a popup of its own, topmost, tool-windowed and
/// refusing activation to begin with. The refusal is put on here rather than answered here
/// alone because it is the style that makes a click refuse as well, and it comes off for the
/// one document that runs (see `ex_style_for`).
fn create_host_window() -> Option<HWND> {
    unsafe {
        CreateWindowExW(
            WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
            WEBVIEW_CLASS,
            w!("Rust Hover Preview"),
            WS_POPUP,
            0,
            0,
            1,
            1,
            None,
            None,
            None,
            None,
        )
        .ok()
    }
}

fn register_class() {
    static REGISTERED: AtomicBool = AtomicBool::new(false);

    if REGISTERED.swap(true, Ordering::AcqRel) {
        return;
    }

    let class = WNDCLASSW {
        lpfnWndProc: Some(window_proc),
        lpszClassName: WEBVIEW_CLASS,
        ..Default::default()
    };

    unsafe {
        RegisterClassW(&class);
    }
}

/// A preview never takes the keyboard by itself: the pointer may be over it, but what is
/// being worked in is Explorer — unless the thing on screen is a page that runs, which is
/// the one document whose whole reading is a script answering keys, and a page that is
/// never activated is a page nothing can be typed into (see `mouse_activate_answers`).
extern "system" fn window_proc(
    hwnd: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if message == WM_MOUSEACTIVATE {
        return mouse_activate_answers(RUNNING_DOCUMENT.load(Ordering::Acquire));
    }

    unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
}

/// What a click into the engine's window is answered with, given what the window is showing.
///
/// Two answers, because a preview is two things. A document and a specimen are looked at, and
/// the pointer being over one of them says nothing about the user having left the window they
/// are working in, so a click is refused activation and the caret stays where it was. A page
/// of HTML that runs is interacted with, and a click that refused to activate it would leave
/// a page on screen that no key could reach, which is a page the user is looking at a picture
/// of. So the one document that is a program rather than a picture takes the keyboard, and it
/// takes it only on a click — showing it is still `SW_SHOWNOACTIVATE`, so a hover never
/// steals the caret from what the hand is on.
fn mouse_activate_answers(runs: bool) -> LRESULT {
    const MA_ACTIVATE: LRESULT = LRESULT(1);
    const MA_NOACTIVATE: LRESULT = LRESULT(3);

    if runs {
        MA_ACTIVATE
    } else {
        MA_NOACTIVATE
    }
}

/// The window's extended style as it is for a document that runs, or for one that is only
/// looked at: `WS_EX_NOACTIVATE` off for the first and on for the second.
///
/// The style is the standing part of the refusal — a window that carries it cannot be
/// activated by anything, the click included, which is why it has to come off before
/// `WM_MOUSEACTIVATE` can answer `MA_ACTIVATE` for a page that runs. Nothing else in the
/// style is touched: the window stays a tool window so it is nothing in the taskbar or on
/// the alt-tab, and stays topmost, since it is a preview the pointer is on top of. A style
/// that has come off for a page is put back for the next document, which is not one — so the
/// refusal is the default and the allowance is the exception made and taken away again.
fn ex_style_for(style: isize, runs: bool) -> isize {
    let noactivate = WS_EX_NOACTIVATE.0 as isize;

    if runs {
        style & !noactivate
    } else {
        style | noactivate
    }
}

/// Retrieve and dispatch whatever is waiting, without blocking.
fn pump_messages() {
    let mut message = MSG::default();

    unsafe {
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

/// The folder the engine keeps its own state in — under the local profile, because a
/// browser profile is not something to synchronize between machines.
///
/// One folder is one browser at a time, and that is the whole reason this is a folder
/// *per run* rather than one folder for the app. An environment pointed at a folder
/// another browser is already holding is answered with `ERROR_INVALID_STATE` when it
/// asks for its controller, and the browser that holds it may be one left behind by a
/// run that ended badly — a zombie whose app is gone and which holds the folder for as
/// long as it lives, which can be indefinitely. A run that keeps its state in a folder
/// of its own cannot be blocked by any of that: the worst a leftover can do is take up
/// space, and the next start clears it away.
///
/// `RHP_WEBVIEW_PROFILE` names a folder for one run of the probes, so that a probe can
/// run beside a running app without either of them noticing the other.
fn user_data_folder() -> PathBuf {
    if let Some(folder) = std::env::var_os("RHP_WEBVIEW_PROFILE") {
        return PathBuf::from(folder);
    }

    profile_root().join(std::process::id().to_string())
}

/// The runs that left a profile folder behind.
///
/// Every folder under the browser's profile root is named for the run that made it,
/// so a folder that is not this run's names a run whose browser may still be holding
/// it — which is what reaches a browser left by a version of this app that wrote no
/// record of it. Read at startup, before this run has a folder of its own.
pub fn stale_profile_pids() -> Vec<u32> {
    engine_processes::stale_run_pids(&profile_root())
}

/// The folder the per-run folders live in.
fn profile_root() -> PathBuf {
    directories::BaseDirs::new()
        .map(|dirs| {
            dirs.data_local_dir()
                .join("rust-hover-preview")
                .join("webview")
        })
        .unwrap_or_else(std::env::temp_dir)
}

/// Clear away the folders earlier runs left, which nothing is using any more.
///
/// A folder a browser is still holding is left where it is: it cannot be removed while
/// it is open, and it is not this run's to insist on. What is left is taken by the next
/// start, once whatever held it is gone. Called once, before anything can create a
/// folder of this run's own.
pub fn clear_stale_profiles() {
    let root = profile_root();
    let ours = std::process::id().to_string();

    let Ok(entries) = std::fs::read_dir(&root) else {
        return;
    };

    for entry in entries.flatten() {
        if entry.file_name() == std::ffi::OsStr::new(&ours) {
            continue;
        }

        if entry.path().is_dir() {
            let _ = std::fs::remove_dir_all(entry.path());
        }
    }
}

fn create_environment(
    user_data_folder: &Path,
    deadline: Instant,
) -> Option<ICoreWebView2Environment> {
    let folder = wide(&user_data_folder.to_string_lossy());
    let (sender, receiver) = mpsc::channel();

    // What a document names is never fetched: the engine is given a resolver rule that
    // answers for no host at all, so an `<image href="http://…">` is a picture that
    // does not arrive rather than a request this app made. That is the promise the rest
    // of it keeps — a hover reads the file under the pointer and nothing else — and a
    // browser would otherwise break it without a line of anything being written.
    let options = CoreWebView2EnvironmentOptions::default();
    unsafe {
        options.set_additional_browser_arguments(BROWSER_ARGUMENTS.to_string());
    }
    let options: ICoreWebView2EnvironmentOptions = options.into();

    CreateCoreWebView2EnvironmentCompletedHandler::wait_for_async_operation(
        Box::new(move |handler| unsafe {
            CreateCoreWebView2EnvironmentWithOptions(
                PCWSTR::null(),
                PCWSTR(folder.as_ptr()),
                Some(&options),
                &handler,
            )
            .map_err(webview2_com::Error::WindowsError)
        }),
        Box::new(move |error_code, environment| {
            error_code?;
            sender
                .send(environment.ok_or_else(|| windows::core::Error::from(E_POINTER)))
                .expect("the waiting thread is gone");
            Ok(())
        }),
    )
    .ok()?;

    receiver.recv_timeout(remaining(deadline)).ok()?.ok()
}

/// How long the engine may take to be had at all — the environment and the controller
/// together — before the attempt is read as one that will not answer.
///
/// Both calls are asynchronous and both of them were waited on without a bound: the
/// completion handlers this thread parks on are the runtime's to fire, and one that never
/// fires leaves this thread waiting for the rest of the run — with every document after it
/// queued behind a wait nobody ends, and no failure notice to take a hover's spinner down
/// with, because the thread that would post one is the thread that is waiting. Twenty
/// seconds is far past what having an engine costs (a quarter of a second measured warm,
/// and the retries a held profile folder asks for are inside it), so what is past it is
/// silence rather than work.
const HOST_CREATION_TIMEOUT: Duration = Duration::from_secs(20);

/// What is left of a deadline, as a wait: `recv_timeout` given nothing waits nothing and
/// answers `Timeout` at once, which is the answer a wait past its deadline wants.
fn remaining(deadline: Instant) -> Duration {
    deadline.saturating_duration_since(Instant::now())
}

fn create_controller(
    environment: ICoreWebView2Environment,
    hwnd: HWND,
    deadline: Instant,
) -> Option<ICoreWebView2Controller> {
    let (sender, receiver) = mpsc::channel();

    CreateCoreWebView2ControllerCompletedHandler::wait_for_async_operation(
        Box::new(move |handler| unsafe {
            environment
                .CreateCoreWebView2Controller(hwnd, &handler)
                .map_err(webview2_com::Error::WindowsError)
        }),
        Box::new(move |error_code, controller| {
            trace(&format!(
                "controller callback: error={error_code:?} controller={}",
                controller.is_some()
            ));
            error_code?;
            sender
                .send(controller.ok_or_else(|| windows::core::Error::from(E_POINTER)))
                .expect("the waiting thread is gone");
            Ok(())
        }),
    )
    .ok()?;

    receiver.recv_timeout(remaining(deadline)).ok()?.ok()
}

/// The browser's own furniture, all of it taken away: the context menu, the tools, the zoom,
/// the status bar, the web messages a page could reach this app's threads through, and the
/// accelerator keys a browser answers for itself — a find bar, a print dialog — which belong
/// to no preview.
///
/// None of these says anything about the document. Whether it runs is not decided here, and
/// that is the one setting of the browser's that is not the same for every document this
/// engine is given (see `set_scripts`).
fn configure(webview: &ICoreWebView2) {
    unsafe {
        if let Ok(settings) = webview.Settings() {
            let _ = settings.SetAreDefaultContextMenusEnabled(false);
            let _ = settings.SetAreDevToolsEnabled(false);
            let _ = settings.SetIsZoomControlEnabled(false);
            let _ = settings.SetIsStatusBarEnabled(false);
            let _ = settings.SetIsWebMessageEnabled(false);
        }

        // The keys a browser answers for itself — a find bar, a print dialog — belong
        // to no preview.
        if let Ok(settings3) = webview.Settings().and_then(|settings| {
            settings.cast::<webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Settings3>()
        }) {
            let _ = settings3.SetAreBrowserAcceleratorKeysEnabled(false);
        }
    }
}

/// Whether the engine is to run the code in the document it is pointed at: the one setting of
/// the browser's that is a property of the document in it rather than of the engine, and so is
/// made as the document is pointed at (see `Host::navigate`).
fn set_scripts(webview: &ICoreWebView2, on: bool) {
    if let Ok(settings) = unsafe { webview.Settings() } {
        let _ = unsafe { settings.SetIsScriptEnabled(on) };
    }
}

/// A file's version: when it was last written, in milliseconds since the epoch, and
/// nothing at all when that cannot be read.
///
/// It goes into the URL the document is drawn by, so an edited file is a different URL
/// and the browser's cache cannot answer a hover with the document as it was. A version
/// that cannot be read is a URL a hover shares with the previous one, which is no worse
/// than not asking: what comes back is the cache's copy of a document that has not been
/// changed as far as this app can tell.
fn file_version(path: &Path) -> u64 {
    std::fs::metadata(path)
        .and_then(|meta| meta.modified())
        .ok()
        .and_then(|modified| modified.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|since| since.as_millis() as u64)
        .unwrap_or(0)
}

/// The engine's background, which is what the preview's own backdrop setting means to a
/// window that composites for itself.
fn background_color(background: TransparentBackground) -> COREWEBVIEW2_COLOR {
    match background {
        TransparentBackground::Transparent => COREWEBVIEW2_COLOR {
            A: 0,
            R: 0,
            G: 0,
            B: 0,
        },
        TransparentBackground::Black => COREWEBVIEW2_COLOR {
            A: 255,
            R: 0,
            G: 0,
            B: 0,
        },
        TransparentBackground::White => COREWEBVIEW2_COLOR {
            A: 255,
            R: 255,
            G: 255,
            B: 255,
        },
        // A checkerboard is the one backdrop the engine cannot be given: it is drawn by
        // whatever composites the frame, and this window composites its own. The page
        // paints it instead — see `frame_page` — so what this colour is is what stands
        // behind the page until it has been drawn: mid grey, which is what the squares
        // average to.
        TransparentBackground::Checkerboard => COREWEBVIEW2_COLOR {
            A: 255,
            R: 184,
            G: 184,
            B: 184,
        },
    }
}

/// A local file as a URL, which is what the engine is pointed at.
///
/// A verbatim path is not something a URL may contain: a browser pointed at one fails at
/// once, silently, and what a hover shows is nothing. So the Shell's prefix comes off
/// first (`crate::paths::plain_path`), and a share keeps its server: `\\?\UNC\server\share`
/// becomes `file://server/share`, a drive becomes `file:///C:/…`. The characters that
/// would end the path early — a space, a hash, a question mark, a percent — are escaped;
/// anything else is left as it is written.
fn file_url(path: &Path) -> Option<String> {
    if !path.is_absolute() {
        return None;
    }

    let local = plain_path(path);

    let (mut url, rest) = match local.strip_prefix(r"\\") {
        Some(share) => (String::from("file://"), share.to_string()),
        None => (String::from("file:///"), local),
    };

    for character in rest.chars() {
        match character {
            '\\' => url.push('/'),
            ' ' => url.push_str("%20"),
            '#' => url.push_str("%23"),
            '?' => url.push_str("%3F"),
            '%' => url.push_str("%25"),
            other => url.push(other),
        }
    }

    Some(url)
}

fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(test)]
mod tests;

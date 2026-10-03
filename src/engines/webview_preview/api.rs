use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use webview2_com::Microsoft::Web::WebView2::Win32::GetAvailableCoreWebView2BrowserVersionString;
use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{LPARAM, WPARAM};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::UI::WindowsAndMessaging::{PostThreadMessageW, WM_APP};

use super::engine::Engine;

use crate::config::config::TransparentBackground;
use crate::formats::font_formats;
use crate::formats::text_formats;
use crate::readers::svg_preview;
use crate::CONFIG;

/// The window class the engine's window is made from. It exists to refuse activation, and
/// the refusal is a default rather than a fixed answer: a document and a specimen never
/// take the keyboard away from what the pointer is over, and a page that runs is the one
/// thing a browser is brought here for — so a window of this class asks whether it is
/// holding a page before it says no (see `mouse_activate_answers`).
pub(super) const WEBVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWebView");

/// The browser arguments this app passes. The first answers for no host at all, so
/// nothing a document links to is fetched from anywhere (see `create_environment`); the
/// second keeps a scrollbar out of a preview if a document is ever a pixel larger than
/// the window it was given. Both are quoted because the browser's command line is a
/// command line: unquoted, an argument with a space in it arrives as several.
pub(super) const BROWSER_ARGUMENTS: &str =
    "--host-resolver-rules=\"MAP * ~NOTFOUND\" --hide-scrollbars";

/// What one engine cost to begin and to point at a document, in milliseconds. It is
/// kept for the same reason the other probes exist: "the preview is slow" is answered
/// by a number, and this is the module the number belongs to.
#[derive(Clone, Copy, Debug, Default)]
pub(crate) struct Timings {
    pub(crate) environment_ms: u64,
    pub(crate) controller_ms: u64,
    pub(crate) navigate_ms: u64,
}

pub(super) static LAST_TIMINGS: Lazy<Mutex<Timings>> = Lazy::new(|| Mutex::new(Timings::default()));

/// Whether the runtime is on this machine, answered once: the check reads the version
/// of the installed runtime, and that does not change while the app runs.
static RUNTIME: Lazy<Option<String>> = Lazy::new(runtime_version);

/// Whether the engine's window is on screen. The preview loop reads this to know when
/// to take its own window down, so it is an atomic rather than a message.
pub(super) static SHOWING: AtomicBool = AtomicBool::new(false);

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
pub(super) static RUNNING_DOCUMENT: AtomicBool = AtomicBool::new(false);

/// Where the engine's window is while it is on screen, in screen coordinates: the box
/// the document was last handed over in. Kept as a rectangle rather than as the window
/// handle, because a rectangle is what the preview loop asks about a preview.
static SHOWING_RECT: Lazy<Mutex<Option<ScreenRect>>> = Lazy::new(|| Mutex::new(None));

/// A rectangle in screen coordinates, as the preview loop asks about one.
type ScreenRect = (i32, i32, i32, i32);

/// The engine's thread, once one has been started.
pub(super) static ENGINE: Lazy<Mutex<Option<Engine>>> = Lazy::new(|| Mutex::new(None));

/// Where the engine is asked to put its window, in screen coordinates.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) struct Area {
    pub(crate) x: i32,
    pub(crate) y: i32,
    pub(crate) width: i32,
    pub(crate) height: i32,
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
pub(super) static WANTED: Lazy<Mutex<Option<Wanted>>> = Lazy::new(|| Mutex::new(None));

/// The generation of the newest want. Every request carries the generation it was made
/// under, and one that is not this is not carried out — which is the question the engine's
/// thread asks of a navigation that is still running, in the middle of pumping a browser's
/// messages and unable to take a lock the loop may be holding.
pub(super) static WANTED_GENERATION: AtomicU64 = AtomicU64::new(0);

/// The box a document the engine is *holding* is to be moved to, published under one cell
/// rather than a queue (see `WANTED`).
///
/// A drag is many placements and only the place the hand let go at is worth putting up. This
/// is what a placement is made of, and it is deliberately not a want: a want names a document
/// and is replaced only by another document, while this is the same document in another box.
/// Publishing a new one here therefore *replaces* what was there, which is the whole of the
/// coalescing — kept apart from `WANTED` because a document that has landed keeps its want
/// and a box is asked for far more often than a document is (see `place`, `ask_place`).
pub(super) static PLACED: Lazy<Mutex<Option<Placement>>> = Lazy::new(|| Mutex::new(None));

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
pub(super) static PLACE_ASKED: AtomicBool = AtomicBool::new(false);

/// A box a document the engine is holding is to be moved to, with the two facts it is
/// published under: the file it is for, so a box left behind by a document the preview has
/// moved on from is not carried out, and the backdrop, which is the page's own as well as the
/// controller's and so can make a placement a page to write again (see `Host::needs_page`).
#[derive(Clone)]
pub(super) struct Placement {
    pub(super) path: PathBuf,
    pub(super) background: TransparentBackground,
    pub(super) area: Area,
}

/// The newest placement asked for, taken back for the engine's thread to carry out, and the
/// flag that says one is owed released — the two are taken together because a placement left
/// owed with nothing behind it is a window that would never move again.
pub(super) fn take_placement() -> Option<Placement> {
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
pub(super) fn drop_placement() {
    PLACE_ASKED.store(false, Ordering::Release);

    if let Ok(mut placed) = PLACED.lock() {
        *placed = None;
    }
}

/// The engine thread's id, once it is running: what a want published while that thread is
/// parked on `GetMessage` is posted to, so that the wait a navigation is in the middle of
/// ends the moment what is wanted is not what it is waiting for.
pub(super) static ENGINE_THREAD: AtomicU32 = AtomicU32::new(0);

/// What a wakeup to the engine's thread is: a message the window procedure ignores, whose
/// whole job is to bring `GetMessage` back so a wait can look at what is wanted now.
const ENGINE_WAKE: u32 = WM_APP;

/// Wake the engine's thread if it is waiting on something.
///
/// A want published while that thread is parked on `GetMessage` would otherwise be noticed
/// only when the browser happened to say something — which, for a navigation that never
/// lands, is never. A post is not lost by the thread being between two messages: it waits
/// in the queue and brings the next `GetMessage` back at once.
pub(super) fn wake_engine_thread() {
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
pub(super) static HOST_HWND: AtomicIsize = AtomicIsize::new(0);

/// The document the loop wants drawn, as the engine's thread takes it up: the want and the
/// generation it was made under, or nothing when nothing is wanted.
#[derive(Clone)]
pub(super) struct Wanted {
    pub(super) generation: u64,
    pub(super) path: PathBuf,
    pub(super) background: TransparentBackground,
    pub(super) area: Area,
}

/// What is wanted this moment, for the engine's thread to take up.
pub(super) fn wanted() -> Option<Wanted> {
    WANTED.lock().ok().and_then(|wanted| wanted.clone())
}

/// Whether what was asked for under `generation` is still what is wanted.
pub(super) fn is_wanted(generation: u64) -> bool {
    WANTED_GENERATION.load(Ordering::Acquire) == generation
}

/// The box the newest want for `generation` asks for, when that is still the newest want:
/// where a navigation that has just landed puts its window, since a pointer that moved
/// while the document was being drawn has moved the preview with it.
pub(super) fn wanted_area(generation: u64) -> Option<Area> {
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
pub(crate) fn runtime_version() -> Option<String> {
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
pub(crate) fn is_available() -> bool {
    RUNTIME.is_some()
}

/// Whether the engine has a window on screen.
pub(crate) fn is_showing() -> bool {
    SHOWING.load(Ordering::Acquire)
}

/// Where the engine's window is, when it has one on screen: the box the document was
/// last handed over in, in screen coordinates.
///
/// The preview loop reads it for the reason it reads its own window's rectangle: a
/// preview that is under the parked pointer is one a keyboard hover placed there, and a
/// document the engine draws is a preview of this app's even though the window is not.
pub(crate) fn screen_rect() -> Option<ScreenRect> {
    SHOWING_RECT.lock().ok().and_then(|rect| *rect)
}

/// The document the engine's window is showing, when it is showing one: the file the page
/// that landed was written for.
///
/// The preview loop asks this to tell its own document from another hover's: a page that
/// arrives is what a wait ends on, and a wait is for one file — one that lands for a hover
/// the loop has left is not a wait that has ended (see `SHOWN`).
pub(crate) fn showing_path() -> Option<PathBuf> {
    SHOWN.lock().ok().and_then(|shown| shown.clone())
}

/// The window the engine draws in, when it has one: the handle a hit test compares against
/// the window under the pointer, so that a document the engine draws is touched the way
/// this app's own preview is — the window is the engine's, and the preview is this app's.
pub(crate) fn showing_hwnd() -> isize {
    HOST_HWND.load(Ordering::Acquire)
}

/// Say what the engine's window is showing and where it is, or that it is showing nothing
/// and is nowhere: the three things the preview loop asks a preview about, published together
/// because they are one fact and are read together (see `is_showing`, `showing_path`,
/// `screen_rect`).
///
/// Nothing is published as the window comes down, which is what `hide` is for.
pub(super) fn publish(shown: Option<(PathBuf, Area)>) {
    if let Ok(mut path) = SHOWN.lock() {
        *path = shown.as_ref().map(|(path, _)| path.clone());
    }

    SHOWING.store(shown.is_some(), Ordering::Release);

    publish_rect(shown.map(|(_, area)| area));
}

/// Say where the engine's window is — or that it is nowhere — for `screen_rect` to answer
/// with. The one half of `publish` a window being *moved* is also owed.
pub(super) fn publish_rect(area: Option<Area>) {
    if let Ok(mut rect) = SHOWING_RECT.lock() {
        *rect = area.map(|area| (area.x, area.y, area.x + area.width, area.y + area.height));
    }
}

/// What the engine cost last time it was asked for a document. Read by the probe: it
/// is the number that says whether the engine is worth keeping warm at all.
#[cfg(test)]
pub(crate) fn last_timings() -> Timings {
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
pub(crate) fn draws(path: &Path) -> bool {
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
pub(crate) fn page_runs(path: &Path) -> bool {
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
pub(super) fn page_runs_under(renders_html: bool, path: &Path) -> bool {
    renders_html && text_formats::is_html_extension(path)
}

/// Whether the configuration asks for a page of HTML to be drawn rather than its markup.
pub(super) fn renders_html() -> bool {
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
pub(super) fn owed(path: &Path) -> bool {
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
pub(crate) fn is_behind(path: &Path) -> bool {
    !is_failing() && (showing_hwnd() != 0 || owed(path))
}

/// Whether the engine can draw anything at all.
///
/// The runtime is what draws a document or a specimen, so a machine without it — or one the
/// engine has stood down on — has neither: this is the question the layout asks before it
/// measures one, and answering no is what keeps a hover from opening a box that nothing
/// would be drawn into.
pub(crate) fn can_draw() -> bool {
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
pub(crate) fn take_failure_notice() -> bool {
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
pub(super) fn note_failure() {
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
pub(super) fn note_document_failed() {
    FAILURE_NOTICE.store(true, Ordering::Release);
}

pub(super) fn note_engine_up() {
    if let Ok(mut failed) = ENGINE_FAILED_AT.lock() {
        *failed = None;
    }
}

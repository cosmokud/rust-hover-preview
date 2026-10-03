use crate::app::engine_processes;
use crate::config::config::{OfficeEngine, PreviewType};
use crate::engines::document_cache::{self, Page, PageKind};
use crate::formats::office_formats::{app_for, container_kind, OfficeApp};
use crate::shell::cloud_files;
use crate::ui::preview_window;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU32, AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant, SystemTime};

use windows::Win32::Foundation::{LPARAM, WPARAM};
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, MsgWaitForMultipleObjectsEx, PeekMessageW, PostThreadMessageW,
    TranslateMessage, MSG, MWMO_INPUTAVAILABLE, PM_REMOVE, QS_ALLINPUT, WM_APP,
};

use super::engines::{wait_for_exit, Engine, Engines};
use super::renderers::{render_target, slide_export_width, take_render, PreparedSource};

/// The message a request is announced with on the worker's own thread queue.
const WM_OFFICE_RENDER: u32 = WM_APP + 3;

/// How long a file that refused a page is left alone. Office refused it for a
/// reason — a password, a repair dialog, a document in Protected View — and the
/// answer will not be different a moment later, so the wait is the user's. It is
/// short enough that a refusal that was really the machine's — an Office that would
/// not start, a license that had to be sorted out — is tried again before long.
const FAILURE_BACKOFF: Duration = Duration::from_secs(120);
/// How long a queued request is still worth running. A render can take seconds,
/// and by the time a long one is done the pointer has moved on.
const REQUEST_STALE_SECS: u64 = 30;

/// How long the worker may be inside one piece of work before it is given up on.
///
/// The work that can block is not only the render: quitting an engine is another
/// COM call, and a dialog inside Office holds any of them for as long as it is up.
/// An Office start, an export and the dialogs in between are seconds at worst, so a
/// worker that has been inside one of them this long is not coming back. It is
/// minutes rather than seconds because a very large document — half a gigabyte of
/// Word, say — is slow to open without being stuck at all, and abandoning it would
/// throw away the page it was about to cache.
const WORKER_GIVE_UP: Duration = Duration::from_secs(120);
/// Files remembered as having refused a page. The map is cleared wholesale when it
/// is full, the way the app's other memories are: what a refusal costs is a wait,
/// so a few of them are worth remembering and a list of them is not.
const FAILURE_CACHE_MAX_ENTRIES: usize = 64;

pub(super) struct RenderRequest {
    pub(super) source: PathBuf,
    pub(super) width: u32,
    pub(super) height: u32,
    pub(super) generation: u64,
    pub(super) requested: Instant,
}

/// What failed, the version of the file that failed, and when. Held in memory
/// only: a failure is not a property of the file worth keeping across runs. What
/// Office said about it is kept beside this, in the diagnostic below.
struct Failure {
    modified: Option<SystemTime>,
    len: u64,
    at: Instant,
}

static REQUEST: Lazy<Mutex<Option<RenderRequest>>> = Lazy::new(|| Mutex::new(None));
static FAILURES: Lazy<Mutex<HashMap<PathBuf, Failure>>> = Lazy::new(|| Mutex::new(HashMap::new()));
static WORKER_HANDLE: Lazy<Mutex<Option<std::thread::JoinHandle<()>>>> =
    Lazy::new(|| Mutex::new(None));
static WORKER_STARTING: Lazy<Mutex<()>> = Lazy::new(|| Mutex::new(()));
/// Which worker is the current one.
///
/// A worker that has been given up on keeps running — nothing can stop a COM call
/// — and this is what keeps it from being mistaken for the worker that replaced
/// it, including when it finally returns and cleans up.
pub(super) static WORKER_GENERATION: AtomicU64 = AtomicU64::new(0);
static WORKER_THREAD: AtomicU32 = AtomicU32::new(0);
/// The work a worker is inside, and which generation of worker is inside it. Every
/// step that can block is done inside this marker — the render, and the engine
/// being ended — so "the worker is stuck" is a question about the whole thread
/// rather than about one call in it.
static WORKER_BUSY: Lazy<Mutex<Option<(u64, Instant)>>> = Lazy::new(|| Mutex::new(None));

/// Whether the render tier may run at all: the `Document` gate in the tray's
/// `Preview Types` submenu.
///
/// A cache budget of nothing is not a switch. An Office document has no other
/// source for its preview, so a page still has to be rendered to be shown — it is
/// simply not kept once the hover that asked for it is over.
pub(crate) fn enabled() -> bool {
    PreviewType::Document.enabled()
}

/// The page held for this document, if one has been drawn for it.
///
/// What the page *is* is not asked here: whether a page can actually be read out of the file it
/// is kept as is a question for the side that draws it (`office_preview`), whose threads may
/// talk to the PDF engine — this one is apartment-threaded for Office, and a WinRT call waited
/// on from here would deadlock.
pub(crate) fn held_page(source: &Path) -> Option<Page> {
    document_cache::page(source, OfficeEngine::MicrosoftOffice.as_str())
}

/// Whether the page held for this document is narrower than one rendered for a box `width` wide
/// would be.
///
/// It is a question about a raster a family exported *for the box*: a slide is exported at the
/// width it is asked for, so a deck that was first previewed on a smaller display holds a page
/// that a larger one would draw softer than it could — worth replacing, which is what this
/// answers — and one already exported at the cap is not, which is what keeps asking for it from
/// being a loop. The one other raster Office draws for this app is the picture a workbook is
/// answered with where no page can be exported, and it is not exported for the box at all: it is
/// the used range at the range's own size, bounded by the picture's own limits, so a second
/// render of it would write the same picture again — asking for one would be a render paid on
/// every hover (see `page_is_workbook_picture`). Every other family's page is a PDF the size its
/// document makes it, whatever box the render was asked for. What the held page's own width is,
/// is read from the page itself (see `document_cache::size`).
pub(crate) fn page_is_narrower_than(source: &Path, page: &Page, width: u32) -> bool {
    page.kind == PageKind::Png
        && !page_is_workbook_picture(source, page)
        && document_cache::size(source, OfficeEngine::MicrosoftOffice.as_str())
            .is_some_and(|(export_width, _)| export_width < slide_export_width(width))
}

/// Whether the page held for this document is the picture a workbook is answered with where no
/// page can be exported, rather than a page drawn for the document.
///
/// It is the one thing about a page that the kind it is kept under no longer says: a workbook's
/// picture is a PNG now, exactly as a slide's export is (see `copy_used_range_picture`), and the
/// two are told apart by the family the document belongs to — which is the file's own answer,
/// read here the same way the render tier reads it before it draws anything. What the question is
/// asked for: a picture is only as good as the pixels it holds, so it is placed as a bitmap
/// rather than enlarged to the display (`office_preview::SourceKind`), and a display larger than
/// it is not a reason to render it again. A `.bmp` answers as a picture whatever family it is
/// kept for, since that is the only thing that ever wrote one.
pub(crate) fn page_is_workbook_picture(source: &Path, page: &Page) -> bool {
    matches!(page.kind, PageKind::Png | PageKind::Bmp)
        && (page.kind == PageKind::Bmp || app_for(source) == Some(OfficeApp::Excel))
}

/// The hover a page was drawn for is over: it is no longer being waited on, so at a budget of
/// nothing it is given up now rather than lingering until the next render happens to make room.
pub(crate) fn hover_ended(source: &Path) {
    document_cache::hover_ended(source);
}

/// Ask for a page to be rendered, unless one is already waiting or the file has
/// just refused to give one.
pub(crate) fn request(source: &Path, width: u32, height: u32, generation: u64) {
    if !enabled() {
        return;
    }

    if failed_recently(source) {
        // Nothing to wait for: the preview is told at once so the spinner it may
        // be showing comes down rather than timing out.
        preview_window::notify_office_render(source, generation, false);
        return;
    }

    // A worker that has been inside one piece of work for longer than any of them
    // takes is not coming back — a dialog inside Office holds a COM call for good
    // — so it is left where it is, the Office process it started is ended with it,
    // and the request goes to a worker that can answer it. This is what keeps one
    // stuck document, or one stuck quit, from costing every document after it.
    if work_is_stuck() {
        abandon_worker();
    }

    if let Ok(mut slot) = REQUEST.lock() {
        *slot = Some(RenderRequest {
            source: source.to_path_buf(),
            width,
            height,
            generation,
            requested: Instant::now(),
        });
    }

    let thread = WORKER_THREAD.load(Ordering::Acquire);
    if thread != 0 {
        unsafe {
            let _ = PostThreadMessageW(thread, WM_OFFICE_RENDER, WPARAM(0), LPARAM(0));
        }
        return;
    }

    start_worker();
}

/// Remember that this version of the file produced no page, so the next hover
/// does not ask again.
pub(crate) fn remember_failure(source: &Path) {
    let metadata = std::fs::metadata(source).ok();
    let failure = Failure {
        modified: metadata
            .as_ref()
            .and_then(|metadata| metadata.modified().ok()),
        len: metadata.map(|metadata| metadata.len()).unwrap_or(0),
        at: Instant::now(),
    };

    if let Ok(mut failures) = FAILURES.lock() {
        if failures.len() >= FAILURE_CACHE_MAX_ENTRIES && !failures.contains_key(source) {
            failures.clear();
        }
        failures.insert(source.to_path_buf(), failure);
    }
}

/// Stop the worker, joining it only when it is not inside anything that can block:
/// a COM call into Office cannot be cancelled, and waiting on one would hold the
/// app's exit for as long as Office takes.
///
/// A worker that cannot be joined is one whose engines are never let go of itself —
/// the thread ends when the process does, and nothing of it runs again — so what it
/// started is ended here instead. Those are this app's own Office processes and
/// nobody else's: an instance the user is working in is not one this app created,
/// and is never one of these.
pub(crate) fn shutdown() {
    if work_in_flight() {
        engine_processes::terminate_all_owned();
        return;
    }

    if let Ok(mut handle) = WORKER_HANDLE.lock() {
        if let Some(handle) = handle.take() {
            let _ = handle.join();
        }
    }
}

/// Let go of every engine at once, for a preview kind that has been switched off.
///
/// The processes are ended from here rather than waited for on the worker: what the
/// worker is asked for is the tidy half — its slots cleared, the COM objects released
/// on the apartment that made them — and a worker that is inside a call it cannot cut
/// short would hold a process nothing can be previewed from for good. So it is told
/// to look at the gate now, and the processes go whether it looks or not.
///
/// Called from the tray, which may not wait on anything: every call here is a request.
pub(crate) fn stop_engines() {
    engine_processes::terminate_all_owned();

    let thread = WORKER_THREAD.load(Ordering::Acquire);
    if thread != 0 {
        unsafe {
            let _ = PostThreadMessageW(thread, WM_OFFICE_RENDER, WPARAM(0), LPARAM(0));
        }
    }
}

fn work_in_flight() -> bool {
    WORKER_BUSY
        .lock()
        .ok()
        .map(|busy| busy.is_some())
        .unwrap_or(false)
}

/// Whether the worker has been inside one piece of work for longer than any of
/// them takes.
fn work_is_stuck() -> bool {
    WORKER_BUSY
        .lock()
        .ok()
        .and_then(|busy| *busy)
        .map(|(_, since)| since.elapsed() >= WORKER_GIVE_UP)
        .unwrap_or(false)
}

/// Stop counting on the worker that is inside a piece of work, and end the Office
/// process it started.
///
/// The thread itself is left where it is: it is inside a COM call nothing can
/// cancel, and what it holds is one document's render, which is not worth the
/// tier. Its generation is moved on so that everything it does on the way out —
/// clearing the thread id, ending an engine — is conditional on it still being the
/// current worker, which it no longer is.
fn abandon_worker() {
    engine_processes::terminate_all_owned();
    WORKER_GENERATION.fetch_add(1, Ordering::AcqRel);
    WORKER_THREAD.store(0, Ordering::Release);
    if let Ok(mut busy) = WORKER_BUSY.lock() {
        *busy = None;
    }
}

fn failed_recently(source: &Path) -> bool {
    let metadata = std::fs::metadata(source).ok();
    let modified = metadata
        .as_ref()
        .and_then(|metadata| metadata.modified().ok());
    let len = metadata.map(|metadata| metadata.len()).unwrap_or(0);

    let Ok(failures) = FAILURES.lock() else {
        return false;
    };
    let Some(failure) = failures.get(source) else {
        return false;
    };

    failure.modified == modified && failure.len == len && failure.at.elapsed() < FAILURE_BACKOFF
}

// --------------------------------------------------------------- the worker

/// What one request came to.
///
/// The difference between the last two matters: a document that refused a page is
/// worth leaving alone for a while, while an engine that would not start is the
/// machine's problem and says nothing about the document — so only a refusal is
/// remembered against the file.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum RenderOutcome {
    /// The page is in the cache.
    Rendered,
    /// Office was asked and gave no page.
    Refused,
    /// There was no engine to ask: Office would not start, or the family has none.
    NoEngine,
}

fn start_worker() {
    let Ok(_starting) = WORKER_STARTING.lock() else {
        return;
    };
    if WORKER_THREAD.load(Ordering::Acquire) != 0 {
        return;
    }

    let handle = std::thread::spawn(worker_main);
    if let Ok(mut slot) = WORKER_HANDLE.lock() {
        *slot = Some(handle);
    }
}

/// The engine thread: an apartment-threaded COM thread with a message pump.
///
/// Office automation is OLE automation, and OLE automation is apartment-bound:
/// the calls marshal back into this thread, which has to be pumping messages for
/// them to be delivered. That is why this thread is apartment-threaded and
/// pumped — the opposite of every other worker in this app, which initializes a
/// multithreaded apartment for WinRT.
fn worker_main() {
    // Which worker this is. One that has been given up on is still running, and
    // everything it does on the way out is conditional on this still being the
    // current generation, so it cannot take the place of its replacement.
    let generation = WORKER_GENERATION.fetch_add(1, Ordering::AcqRel) + 1;

    // The thread makes itself known before anything expensive happens in it. What
    // the id is for is answering whether there is a worker to post to, and a worker
    // that is still initializing is one: publishing it after the apartment had been
    // initialized left a window — the whole of `CoInitializeEx`, which loads DLLs —
    // in which this thread existed but reported itself absent, so a request arriving
    // in it started a second worker beside this one. Two workers are not a lost
    // render either way, because the request is in the slot and whichever worker is
    // there takes it on its next wake; what they do cost is the race to store their
    // own ids, where the loser can leave a dead thread's id in the slot for every
    // request after it.
    unsafe {
        WORKER_THREAD.store(GetCurrentThreadId(), Ordering::Release);

        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        // The thread's message queue has to exist before another thread can post
        // to it, and it only exists once something has asked for it. A post that
        // arrives before this is refused by the system and costs nothing: the
        // request it announced is already in the slot, waiting to be taken.
        let mut message = MSG::default();
        let _ = PeekMessageW(&mut message, None, 0, 0, PM_REMOVE);
    }

    // The thread id is handed back however this thread ends, a panic included:
    // losing it would leave every later request posted to a thread that is gone,
    // which is a render tier that never answers again.
    let _registration = WorkerRegistration { generation };

    let mut engines = Engines::new();

    while crate::RUNNING.load(Ordering::Acquire) && is_current_worker(generation) {
        pump_messages();

        if let Some(request) = take_request() {
            if request.requested.elapsed() > Duration::from_secs(REQUEST_STALE_SECS) {
                continue;
            }

            begin_work(generation);
            // A panic inside one document's render is that document's failure, not
            // the tier's: the thread goes on to the next hover.
            let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                render_request(&mut engines, &request, generation)
            }))
            .unwrap_or(RenderOutcome::Refused);
            end_work(generation);

            match outcome {
                // The page is already held: storing it is what trimmed the cache.
                RenderOutcome::Rendered => {}
                // A refusal is the document's, and is remembered against it so the
                // next hover does not ask again straight away. An engine that would
                // not start is not: that is the machine's business, and the file is
                // worth asking about again — and neither is a refusal met while the
                // tier was switched off, where it was the engines being ended under
                // the render that refused it, which would otherwise leave the file
                // unprompted for minutes over a switch.
                RenderOutcome::Refused => {
                    if enabled() {
                        remember_failure(&request.source);
                    }
                }
                RenderOutcome::NoEngine => {}
            }

            preview_window::notify_office_render(
                &request.source,
                request.generation,
                outcome == RenderOutcome::Rendered,
            );
            continue;
        }

        // Nothing to do. An engine that has gone long enough without a page is let
        // go, each on its own clock — so the family asked for most recently is the
        // one that outlives the others — and this thread ends with the last of
        // them: an app left alone should end up with no Office process and no
        // polling thread. Letting an engine go is a COM call of its own, so it
        // happens inside the marker like the rest.
        //
        // Every engine goes at once when the tier is switched off, which is what
        // clears these slots: a process ended from outside — which is what the tray
        // does, for a worker too stuck to look — leaves this side holding COM objects
        // over a process that is gone, and the hover that would have been answered
        // from one is never even asked for, since the gate is what asks.
        let switched_off = !enabled();
        if engines.has_idle() || switched_off {
            begin_work(generation);
            if switched_off {
                engines.drop_all();
            } else {
                engines.drop_idle();
            }
            end_work(generation);
        }
        if engines.is_empty() {
            break;
        }

        wait_for_message(500);
    }

    // Whatever engines are left are let go the same way, inside the marker.
    if !engines.is_empty() {
        begin_work(generation);
        engines.drop_all();
        end_work(generation);
    }

    // A request that arrived while this thread was shutting down would otherwise
    // sit in the slot with nobody to run it — and only the worker that is still
    // the current one may hand it on, since an abandoned thread starting a worker
    // would leave two of them running.
    if is_current_worker(generation) {
        WORKER_THREAD.store(0, Ordering::Release);
        if take_request().is_some() && crate::RUNNING.load(Ordering::Acquire) {
            start_worker();
        }
    }

    unsafe {
        windows::Win32::System::Com::CoUninitialize();
    }
}

fn is_current_worker(generation: u64) -> bool {
    WORKER_GENERATION.load(Ordering::Acquire) == generation
}

/// Hands the worker's thread id back when the thread stops being the current one,
/// whether it ends normally or unwinds out of something unexpected.
struct WorkerRegistration {
    generation: u64,
}

impl Drop for WorkerRegistration {
    fn drop(&mut self) {
        if is_current_worker(self.generation) {
            WORKER_THREAD.store(0, Ordering::Release);
        }
    }
}

fn begin_work(generation: u64) {
    if let Ok(mut busy) = WORKER_BUSY.lock() {
        *busy = Some((generation, Instant::now()));
    }
}

fn end_work(generation: u64) {
    if let Ok(mut busy) = WORKER_BUSY.lock() {
        if busy.map(|(busy_generation, _)| busy_generation) == Some(generation) {
            *busy = None;
        }
    }
}

fn take_request() -> Option<RenderRequest> {
    REQUEST.lock().ok().and_then(|mut slot| slot.take())
}

/// Take every message off this thread's queue. There is no window to dispatch
/// to, so a message is consumed by being retrieved; what the pump is for is the
/// COM calls that Office makes back into this thread while it waits.
pub(super) fn pump_messages() {
    unsafe {
        let mut message = MSG::default();
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

pub(super) fn wait_for_message(milliseconds: u32) {
    unsafe {
        let _ = MsgWaitForMultipleObjectsEx(None, milliseconds, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
    }
}

pub(super) fn render_request(
    engines: &mut Engines,
    request: &RenderRequest,
    generation: u64,
) -> RenderOutcome {
    let Some(app_kind) = app_for(&request.source) else {
        return RenderOutcome::NoEngine;
    };

    // The gate every loader asks before it reads: a document whose content is
    // still in the cloud would be downloaded by the open, and the engine is not an
    // exception to that.
    if cloud_files::needs_download(&request.source) {
        return RenderOutcome::Refused;
    }

    // What the file is *called* is what got it here; what it *is* is its own
    // header's answer, and a file that is not a document at all is not worth an
    // Office start.
    if container_kind(&request.source).is_none() {
        return RenderOutcome::Refused;
    }

    // Another hover may have produced this page while this request waited.
    if held_page(&request.source).is_some() {
        return RenderOutcome::Rendered;
    }

    let Some(target) = render_target() else {
        return RenderOutcome::Refused;
    };

    // An instance that gives no page is given up on and the render is tried once
    // more on a new one, whether it was made for this request or kept warm from the
    // last: an Office that is still starting, or one whose license has run out —
    // which answers for a while and then refuses the copy with "the license to use
    // this application has expired", an answer about the instance rather than about
    // the document — refuses work for reasons that have nothing to do with the
    // file, and a fresh instance answers where the old one would not. The failed
    // instance is dropped rather than asked again, since a warm one is exactly
    // where a license runs out. Once, so that a document which will never render is
    // not asked for twice on every hover.
    let mut retried = false;
    loop {
        // A worker that has been given up on starts nothing. Its own engines have
        // already been ended, and a process started here would be a second engine
        // for this family beside the one the worker that replaced it is starting —
        // which is the one thing the tier does not do. What is given up for it is
        // one document's render, which is not worth a duplicate Office.
        if !is_current_worker(generation) {
            return RenderOutcome::NoEngine;
        }

        if engines.get(app_kind).is_none() {
            // One engine per family, and never a second beside one that is still
            // going: whatever this app started for this family and has not seen end
            // is ended here and waited out. A process that will not end is a preview
            // that does not render this time, which is nothing next to two Office
            // processes on the machine.
            if !engine_processes::end_recorded(app_kind.image_name()) {
                return RenderOutcome::NoEngine;
            }

            let Some(created) = Engine::create(app_kind) else {
                return RenderOutcome::NoEngine;
            };
            *engines.slot(app_kind) = Some(created);
        }

        let rendered = {
            let engine = engines
                .slot(app_kind)
                .as_mut()
                .expect("an engine was just created");
            let source = PreparedSource::new(&request.source);
            let width = request.width.max(1);
            let height = request.height.max(1);
            let rendered = engine.render(&source.path, &target, width, height, retried);
            // The engine that was just asked is the one that was just used, so it
            // is that engine's own clock that starts again.
            engine.idle_since = Instant::now();
            source.cleanup();
            rendered
        };

        // What the renderer wrote is the answer, whatever it chose to write: the
        // file it left behind is read here and deleted, and the page is kept where
        // both engines' documents are drawn from (see `document_cache`). Reading it
        // whatever the render reported is also what keeps the scratch folder empty —
        // a file a failed export left behind is not one the next attempt is allowed
        // to find.
        let page = take_render(&target);
        if let Some((kind, bytes)) = page.filter(|_| rendered) {
            document_cache::store(
                &request.source,
                OfficeEngine::MicrosoftOffice.as_str(),
                kind,
                &bytes,
            );
            return RenderOutcome::Rendered;
        }
        if retried {
            return RenderOutcome::Refused;
        }

        // The instance that failed is let go — ended rather than asked to quit,
        // which is another call it may not answer, made worse by the wait that
        // follows an unanswered quit — and the next pass asks a new one.
        if let Some(engine) = engines.slot(app_kind).take() {
            let (pid, owned) = (engine.owned_pid, !engine.attached);
            engine.abandon();

            // What is started next is not started beside it: the process is waited
            // out, and one that outlasts the wait is not replaced at all.
            if owned && !wait_for_exit(pid) {
                return RenderOutcome::NoEngine;
            }
        }
        retried = true;
    }
}

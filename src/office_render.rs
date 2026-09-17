//! The Office render tier: Word, Excel or PowerPoint draws a document's first
//! page, once, into a cache file the preview then draws from.
//!
//! Nothing here is ever on the hover path. A page is asked for only after the
//! pointer has rested on a file, it is drawn on a thread of its own, and the
//! preview shows a spinner in the meantime — so what a document costs is bounded
//! by the render tier even when it is an Office start, an export and a dialogs
//! worth of waiting.
//!
//! Three rules shape the rest:
//!
//! * **The user's Office is never disturbed.** An automation instance may attach
//!   to a running Word or Excel — the applications are registered for multiple
//!   use — so nothing is hidden that is already visible, the settings that are
//!   changed are restored, and only an instance this app created is quit or
//!   ended.
//! * **A stuck engine must not wedge the app.** The calls below are COM calls
//!   into another process and cannot be cancelled or bounded; the thread that
//!   makes them may be lost to a modal dialog inside Office. What that costs is
//!   one render, never the tier: a worker that has been inside one piece of work
//!   for too long is given up on, the process it started is ended with it, and
//!   the next hover is answered by a fresh worker. Nothing waits on this thread.
//! * **Nothing is held in memory.** What a hover shows comes off the disk, and
//!   the only things kept here are the request slot, the files that have refused a
//!   page, and which worker is current.

use crate::cloud_files;
use crate::config::sanitize_office_cache_mb;
use crate::office_formats::{app_for, container_kind, OfficeApp};
use crate::preview_window;
use crate::CONFIG;
use directories::BaseDirs;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU32, AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use windows::core::{GUID, PCWSTR, PWSTR, VARIANT};
use windows::Win32::Foundation::{CloseHandle, HGLOBAL, LPARAM, WPARAM};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, IDispatch, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
};
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, GetClipboardData, OpenClipboard,
};
use windows::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W, TH32CS_SNAPPROCESS,
};
use windows::Win32::System::Memory::{GlobalLock, GlobalSize, GlobalUnlock};
use windows::Win32::System::Ole::CF_DIB;
use windows::Win32::System::Threading::{
    GetCurrentThreadId, GetExitCodeProcess, OpenProcess, QueryFullProcessImageNameW,
    TerminateProcess, PROCESS_NAME_WIN32, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, MsgWaitForMultipleObjectsEx, PeekMessageW, PostThreadMessageW,
    TranslateMessage, MSG, MWMO_INPUTAVAILABLE, PM_REMOVE, QS_ALLINPUT, WM_APP,
};

/// The message a request is announced with on the worker's own thread queue.
const WM_OFFICE_RENDER: u32 = WM_APP + 3;

/// How long the engine is kept alive after its last render, so a folder of
/// documents costs one Office start rather than one per file.
const ENGINE_IDLE_SECS: u64 = 60;
/// How long a file that refused a page is left alone. Office refused it for a
/// reason — a password, a repair dialog, a document in Protected View — and the
/// answer will not be different a moment later, so the wait is the user's. It is
/// short enough that a refusal that was really the machine's — an Office that would
/// not start, a license that had to be sorted out — is tried again before long.
const FAILURE_BACKOFF: Duration = Duration::from_secs(120);
/// How long a queued request is still worth running. A render can take seconds,
/// and by the time a long one is done the pointer has moved on.
const REQUEST_STALE_SECS: u64 = 30;
/// The width a slide is exported at, bounded: enough for any preview box, and
/// never a poster.
const MIN_SLIDE_EXPORT_WIDTH: u32 = 640;
const MAX_SLIDE_EXPORT_WIDTH: u32 = 1920;
/// Paths at least this long are handed to Office as a copy in the temp folder.
/// Office is not a long-path consumer of the plain form the app converts to.
const MAX_OFFICE_PATH: usize = 240;
/// How much of a worksheet's used range is copied out when the machine has no
/// printer to export a page with: the corner a person sees first, not every row
/// the sheet holds.
const PICTURE_MAX_ROWS: i32 = 40;
const PICTURE_MAX_COLUMNS: i32 = 14;
/// The most a worksheet's picture may be, in pixels.
///
/// What a preview can show is bounded by the display, and a report whose rows are
/// tall — wrapped headings, merged cells — or whose columns are wide would
/// otherwise be copied out at several million pixels: slow to copy, slow to write
/// and slow to draw, all to show a corner no preview can hold. The window the
/// picture is taken from is cut down until what it would draw fits these — and they
/// are deliberately smaller than a display, because a picture is shown at its own
/// size (the configured scale) rather than fitted to the screen the way a page is,
/// and a corner of a sheet filling the screen is not a preview of anything.
const PICTURE_MAX_PIXELS_WIDTH: f64 = 900.0;
const PICTURE_MAX_PIXELS_HEIGHT: f64 = 700.0;
/// The least of a sheet still worth copying: less than this shows too little to be
/// a preview of anything.
const PICTURE_MIN_ROWS: i32 = 4;
const PICTURE_MIN_COLUMNS: i32 = 3;
/// How many times the window may be cut down before it is copied as it stands.
const PICTURE_FIT_ATTEMPTS: usize = 6;
/// Points to pixels, as Excel lays a sheet out at 96 DPI: a point is a 72nd of an
/// inch.
const PICTURE_PIXELS_PER_POINT: f64 = 96.0 / 72.0;
/// The clipboard is shared with every other process: a look that cannot open it,
/// or finds nothing in it, is retried this many times.
const CLIPBOARD_ATTEMPTS: usize = 5;
/// How many times a picture is asked for, and how long between the asks.
///
/// The first copy of a workbook Excel has just opened is regularly refused — an
/// exception saying the `CopyPicture` property of the range could not be got, which
/// is Excel still laying the sheet out rather than anything about the document — and
/// the next attempt a moment later is answered. Asking only once is what left whole
/// workbooks with no preview while their neighbours in the same folder had one.
const PICTURE_COPY_ATTEMPTS: usize = 3;
const PICTURE_COPY_RETRY_MS: u64 = 120;
/// `xlScreen` and `xlBitmap`: the appearance and format `CopyPicture` is asked for.
const XL_SCREEN: i32 = 1;
const XL_BITMAP: i32 = 2;
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
/// The longest image path `QueryFullProcessImageNameW` is given room for.
const MAX_IMAGE_PATH: usize = 260;
/// What `GetExitCodeProcess` reports for a process that is still running.
const STILL_ACTIVE: u32 = 259;
/// The value that switches macro execution off entirely.
const MSO_AUTOMATION_SECURITY_FORCE_DISABLE: i32 = 3;
/// What a property put is identified by, in the parameter block that carries it.
const DISPID_PROPERTYPUT: i32 = -3;

/// Which of the files a render can produce.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum RenderedKind {
    /// A one-page PDF, which the existing PDF renderer draws.
    Pdf,
    /// A PNG of the first slide.
    Png,
    /// A bitmap of a workbook's used range, for a machine whose Excel cannot
    /// export a page at all — see `render_excel`.
    Bmp,
}

impl RenderedKind {
    fn extension(self) -> &'static str {
        match self {
            Self::Pdf => "pdf",
            Self::Png => "png",
            Self::Bmp => "bmp",
        }
    }
}

/// A page that has been rendered for a document and is waiting in the cache.
pub(crate) struct CachedRender {
    pub(crate) path: PathBuf,
    pub(crate) kind: RenderedKind,
}

struct RenderRequest {
    source: PathBuf,
    width: u32,
    height: u32,
    generation: u64,
    requested: Instant,
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
static WORKER_GENERATION: AtomicU64 = AtomicU64::new(0);
static WORKER_THREAD: AtomicU32 = AtomicU32::new(0);
/// The work a worker is inside, and which generation of worker is inside it. Every
/// step that can block is done inside this marker — the render, and the engine
/// being ended — so "the worker is stuck" is a question about the whole thread
/// rather than about one call in it.
static WORKER_BUSY: Lazy<Mutex<Option<(u64, Instant)>>> = Lazy::new(|| Mutex::new(None));
/// The Office process this app started, with the family it belongs to.
///
/// An instance the user already had open is never recorded here, and never ended:
/// what a stuck render costs is the render, not the user's work.
static OWNED_ENGINE: Lazy<Mutex<Option<(OfficeApp, u32)>>> = Lazy::new(|| Mutex::new(None));

/// Whether the render tier may run at all: the tray's `Render With Office`
/// setting, and a cache that can hold what it produces — a page that cannot be
/// kept is not worth an Office start.
pub(crate) fn enabled() -> bool {
    CONFIG
        .lock()
        .map(|config| {
            config.office_render_enabled && sanitize_office_cache_mb(config.office_cache_mb) > 0
        })
        .unwrap_or(false)
}

fn cache_limit_bytes() -> u64 {
    let megabytes = CONFIG
        .lock()
        .map(|config| sanitize_office_cache_mb(config.office_cache_mb))
        .unwrap_or(0);

    megabytes as u64 * 1024 * 1024
}

/// The page rendered for this version of the file, if a file for one is there.
///
/// What the file *holds* is not asked here: whether a page can actually be read out
/// of it is a question for the side that draws it (`office_preview`), whose threads
/// may talk to the PDF engine — this one is apartment-threaded for Office, and a
/// WinRT call waited on from here would deadlock.
pub(crate) fn cached_render(source: &Path) -> Option<CachedRender> {
    let folder = cache_folder()?;
    let key = cache_key(source);

    for kind in [RenderedKind::Pdf, RenderedKind::Png, RenderedKind::Bmp] {
        let path = folder.join(format!("{key}.{}", kind.extension()));
        if let Ok(metadata) = std::fs::metadata(&path) {
            if metadata.len() > 0 {
                return Some(CachedRender { path, kind });
            }
        }
    }

    None
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
pub(crate) fn shutdown() {
    if work_in_flight() {
        return;
    }

    if let Ok(mut handle) = WORKER_HANDLE.lock() {
        if let Some(handle) = handle.take() {
            let _ = handle.join();
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
    terminate_owned_engine();
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
enum RenderOutcome {
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

    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        // The thread's message queue has to exist before another thread can post
        // to it, and it only exists once something has asked for it.
        let mut message = MSG::default();
        let _ = PeekMessageW(&mut message, None, 0, 0, PM_REMOVE);

        WORKER_THREAD.store(GetCurrentThreadId(), Ordering::Release);
    }

    // The thread id is handed back however this thread ends, a panic included:
    // losing it would leave every later request posted to a thread that is gone,
    // which is a render tier that never answers again.
    let _registration = WorkerRegistration { generation };

    let mut engine: Option<Engine> = None;
    let mut idle_since = Instant::now();

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
                render_request(&mut engine, &request)
            }))
            .unwrap_or(RenderOutcome::Refused);
            end_work(generation);
            idle_since = Instant::now();

            match outcome {
                RenderOutcome::Rendered => trim_cache(),
                // A refusal is the document's, and is remembered against it so the
                // next hover does not ask again straight away. An engine that would
                // not start is not: that is the machine's business, and the file is
                // worth asking about again.
                RenderOutcome::Refused => remember_failure(&request.source),
                RenderOutcome::NoEngine => {}
            }

            preview_window::notify_office_render(
                &request.source,
                request.generation,
                outcome == RenderOutcome::Rendered,
            );
            continue;
        }

        // Nothing to do. The engine is kept warm for a minute after the last
        // render and then let go, which is also when this thread ends: an idle app
        // should have no Office process and no polling thread. Letting it go is a
        // COM call of its own, so it happens inside the marker like the rest.
        if engine.is_some() && idle_since.elapsed() >= Duration::from_secs(ENGINE_IDLE_SECS) {
            begin_work(generation);
            engine = None;
            end_work(generation);
        }
        if engine.is_none() {
            break;
        }

        wait_for_message(500);
    }

    // Whatever engine is left is let go the same way, inside the marker.
    if engine.is_some() {
        begin_work(generation);
        drop(engine.take());
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
fn pump_messages() {
    unsafe {
        let mut message = MSG::default();
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

fn wait_for_message(milliseconds: u32) {
    unsafe {
        let _ = MsgWaitForMultipleObjectsEx(None, milliseconds, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
    }
}

fn render_request(engine: &mut Option<Engine>, request: &RenderRequest) -> RenderOutcome {
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
    if cached_render(&request.source).is_some() {
        return RenderOutcome::Rendered;
    }

    let Some(target) = cache_target(&request.source) else {
        return RenderOutcome::Refused;
    };

    // A different family's engine is dropped first, which quits it: an engine is
    // kept for the family it was created for and nothing else.
    //
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
        if engine.as_ref().map(|engine| engine.app_kind) != Some(app_kind) {
            *engine = None;
        }
        if engine.is_none() {
            let Some(created) = Engine::create(app_kind) else {
                return RenderOutcome::NoEngine;
            };
            *engine = Some(created);
        }

        let rendered = {
            let engine = engine.as_ref().expect("an engine was just created");
            let source = PreparedSource::new(&request.source);
            let width = request.width.max(1);
            let height = request.height.max(1);
            let rendered = engine.render(&source.path, &target, width, height);
            source.cleanup();
            rendered
        };

        // What the renderer wrote is the answer, whatever it chose to write: the
        // cache says which file it was.
        if rendered && cached_render(&request.source).is_some() {
            return RenderOutcome::Rendered;
        }
        if retried {
            return RenderOutcome::Refused;
        }

        // The instance that failed is let go, and the next pass asks a new one.
        *engine = None;
        retried = true;
    }
}

// ------------------------------------------------------------------ engines

/// An automation instance, and what it takes to leave it as it was found.
struct Engine {
    app_kind: OfficeApp,
    app: Object,
    /// Whether the instance was already running for the user when it was reached.
    /// Such an instance is never hidden, never quit, and never ended.
    attached: bool,
    /// The process this app started, when it started one: what may be ended if the
    /// engine stops answering, and what must never be ended otherwise.
    owned_pid: u32,
    previous_alerts: Option<VARIANT>,
    previous_security: Option<VARIANT>,
}

impl Engine {
    fn create(app_kind: OfficeApp) -> Option<Self> {
        // What is running before the instance is created, so the process this app
        // starts can be told from one the user already had open.
        let before = processes_named(app_kind.image_name());

        let app = Object::create(app_kind.prog_id())?;

        // An instance this app created is hidden already; the user's is visible,
        // and hiding it would take their windows away.
        let attached = app
            .value("Visible")
            .and_then(|value| bool::try_from(&value).ok())
            .unwrap_or(false);

        let owned_pid = if attached {
            0
        } else {
            processes_named(app_kind.image_name())
                .into_iter()
                .find(|pid| !before.contains(pid))
                .unwrap_or(0)
        };
        if owned_pid != 0 {
            remember_owned_engine(app_kind, owned_pid);
        }

        let previous_alerts = app.value("DisplayAlerts");
        let previous_security = app.value("AutomationSecurity");

        // Macros are switched off wholesale and dialogs are suppressed, so a
        // document cannot put a window in front of the user — an attached
        // instance's settings are put back when the engine is dropped.
        let _ = app.set(
            "AutomationSecurity",
            VARIANT::from(MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
        );
        let _ = app.set("DisplayAlerts", alerts_off(app_kind));

        if !attached {
            let _ = app.set("Visible", VARIANT::from(false));
        }

        Some(Self {
            app_kind,
            app,
            attached,
            owned_pid,
            previous_alerts,
            previous_security,
        })
    }

    fn render(&self, source: &Path, target: &CacheTarget, width: u32, height: u32) -> bool {
        match self.app_kind {
            OfficeApp::Word => render_word(&self.app, source, target),
            OfficeApp::Excel => render_excel(&self.app, source, target),
            OfficeApp::PowerPoint => render_powerpoint(&self.app, source, target, width, height),
        }
    }
}

impl Drop for Engine {
    fn drop(&mut self) {
        if let Some(alerts) = self.previous_alerts.take() {
            let _ = self.app.set("DisplayAlerts", alerts);
        }
        if let Some(security) = self.previous_security.take() {
            let _ = self.app.set("AutomationSecurity", security);
        }

        if !self.attached {
            let _ = self.app.call("Quit", &[]);
        }

        if self.owned_pid != 0 {
            // An Office application finishes quitting while its automation client
            // is pumping: one that is told to quit and then abandoned can sit there
            // for good, which is what an Excel instance left behind every time
            // looks like. So the quit is given a moment of pumping, and a process
            // that is still there after it — one this app started, and verified as
            // still being that application — is ended rather than left running.
            if !wait_for_exit(self.owned_pid) {
                terminate_process(self.owned_pid, self.app_kind.image_name());
            }
            forget_owned_engine(self.owned_pid);
        }
    }
}

/// Give a process a moment to exit, pumping messages while it does — which is what
/// an Office application that has been asked to quit may be waiting for. `true`
/// when it is gone.
fn wait_for_exit(pid: u32) -> bool {
    let deadline = Instant::now() + Duration::from_secs(2);

    while Instant::now() < deadline {
        if !process_is_running(pid) {
            return true;
        }
        wait_for_message(50);
    }

    !process_is_running(pid)
}

/// Whether a process is still running, by id. A process that cannot be opened is
/// one that is gone.
fn process_is_running(pid: u32) -> bool {
    unsafe {
        let Ok(handle) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) else {
            return false;
        };

        let mut code = 0u32;
        let running = GetExitCodeProcess(handle, &mut code).is_ok() && code == STILL_ACTIVE;
        let _ = CloseHandle(handle);
        running
    }
}

/// Record the Office process this app started, so a render that never returns can
/// end it — and only it.
fn remember_owned_engine(app_kind: OfficeApp, pid: u32) {
    if let Ok(mut owned) = OWNED_ENGINE.lock() {
        *owned = Some((app_kind, pid));
    }
}

fn forget_owned_engine(pid: u32) {
    if let Ok(mut owned) = OWNED_ENGINE.lock() {
        if owned.map(|(_, owned_pid)| owned_pid) == Some(pid) {
            *owned = None;
        }
    }
}

/// End the Office process this app started, if it is still there and still that
/// application.
///
/// Requested and never waited on, the way a stuck `ffplay` is: a process inside
/// kernel I/O cannot be ended by anyone in user mode, and what it costs is the
/// render in flight rather than the app. Only an instance this app started is
/// ended — the user's own Office is never touched — and only after the record of
/// it has been taken, so two callers cannot both end it.
fn terminate_owned_engine() {
    let Some((app_kind, pid)) = OWNED_ENGINE.lock().ok().and_then(|mut owned| owned.take()) else {
        return;
    };

    terminate_process(pid, app_kind.image_name());
}

fn terminate_process(pid: u32, image_name: &str) {
    unsafe {
        let Ok(handle) = OpenProcess(
            PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            pid,
        ) else {
            return;
        };

        // Only a process that is still the application it was: a recycled id must
        // never hit something else.
        let mut name = [0u16; MAX_IMAGE_PATH];
        let mut length = name.len() as u32;
        let named = QueryFullProcessImageNameW(
            handle,
            PROCESS_NAME_WIN32,
            PWSTR(name.as_mut_ptr()),
            &mut length,
        )
        .is_ok()
            && String::from_utf16_lossy(&name[..length as usize])
                .to_ascii_uppercase()
                .ends_with(image_name);

        if named {
            let _ = TerminateProcess(handle, 1);
        }
        let _ = CloseHandle(handle);
    }
}

/// The processes running an image of this name, by executable name.
fn processes_named(image_name: &str) -> Vec<u32> {
    let mut found = Vec::new();

    unsafe {
        let Ok(snapshot) = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) else {
            return found;
        };

        let mut entry = PROCESSENTRY32W {
            dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
            ..Default::default()
        };

        if Process32FirstW(snapshot, &mut entry).is_ok() {
            loop {
                let name = String::from_utf16_lossy(&entry.szExeFile);
                if name.trim_end_matches('\0').eq_ignore_ascii_case(image_name) {
                    found.push(entry.th32ProcessID);
                }
                if Process32NextW(snapshot, &mut entry).is_err() {
                    break;
                }
            }
        }

        let _ = CloseHandle(snapshot);
    }

    found
}

/// What suppresses a dialog in each application: Word and PowerPoint take a
/// level, Excel a boolean.
fn alerts_off(app_kind: OfficeApp) -> VARIANT {
    match app_kind {
        OfficeApp::Word => VARIANT::from(0i32),       // wdAlertsNone
        OfficeApp::Excel => VARIANT::from(false),     // DisplayAlerts is a boolean
        OfficeApp::PowerPoint => VARIANT::from(1i32), // ppAlertsNone
    }
}

// --------------------------------------------------------------- the render

fn render_word(app: &Object, source: &Path, target: &CacheTarget) -> bool {
    let Some(documents) = app.member("Documents") else {
        return false;
    };
    let opened = documents.call(
        "Open",
        &[
            ("FileName", path_variant(source)),
            ("ReadOnly", VARIANT::from(true)),
            ("AddToRecentFiles", VARIANT::from(false)),
            ("ConfirmConversions", VARIANT::from(false)),
            // The document opens in a hidden window even when the application
            // itself is the user's and visible.
            ("Visible", VARIANT::from(false)),
        ],
    );
    let Some(document) = opened.and_then(Object::from_variant) else {
        return false;
    };

    let rendered = document
        .call(
            "ExportAsFixedFormat",
            &[
                ("OutputFileName", path_variant(&target.file("pdf"))),
                ("ExportFormat", VARIANT::from(17i32)), // wdExportFormatPDF
                ("OpenAfterExport", VARIANT::from(false)),
                ("Range", VARIANT::from(3i32)), // wdExportFromTo
                ("From", VARIANT::from(1i32)),
                ("To", VARIANT::from(1i32)),
            ],
        )
        .is_some();

    let _ = document.call("Close", &[("SaveChanges", VARIANT::from(false))]);
    rendered
}

fn render_excel(app: &Object, source: &Path, target: &CacheTarget) -> bool {
    let Some(workbooks) = app.member("Workbooks") else {
        return false;
    };
    let opened = workbooks.call(
        "Open",
        &[
            ("FileName", path_variant(source)),
            ("UpdateLinks", VARIANT::from(0i32)),
            ("ReadOnly", VARIANT::from(true)),
            ("AddToMru", VARIANT::from(false)),
            ("IgnoreReadOnlyRecommended", VARIANT::from(true)),
        ],
    );
    let Some(workbook) = opened.and_then(Object::from_variant) else {
        return false;
    };

    // A workbook's first page is what it prints, and exporting one goes through
    // the print pipeline: Excel needs a printer on the machine for it, and a
    // machine with none — no printer at all, which is not the same as a sheet
    // without a print area — cannot export a page however it is asked.
    let printed = printer_installed(app) && export_first_page(&workbook, target);
    let rendered = printed || copy_used_range_picture(&workbook, target);

    let _ = workbook.call("Close", &[("SaveChanges", VARIANT::from(false))]);
    rendered
}

/// The first worksheet's first printed page, as a PDF.
fn export_first_page(workbook: &Object, target: &CacheTarget) -> bool {
    let Some(sheet) = workbook
        .member("Worksheets")
        .and_then(|sheets| sheets.item(1))
    else {
        return false;
    };

    let output = target.file("pdf");
    let exported = sheet
        .call(
            "ExportAsFixedFormat",
            &[
                ("Type", VARIANT::from(0i32)), // xlTypePDF
                ("Filename", path_variant(&output)),
                ("From", VARIANT::from(1i32)),
                ("To", VARIANT::from(1i32)),
                ("OpenAfterPublish", VARIANT::from(false)),
            ],
        )
        .is_some();

    exported && output.exists()
}

/// Whether the machine has a printer, which is what an export to a page needs.
///
/// Excel answers `ActivePrinter` with a name when there is one, and with the
/// sentence "unknown printer (check your Control Panel)" when there is not — so
/// the question is asked before the export rather than answered by its failure.
fn printer_installed(app: &Object) -> bool {
    app.value("ActivePrinter")
        .map(|value| value.to_string())
        .map(|name| !name.trim().is_empty() && !name.contains("unknown printer"))
        .unwrap_or(false)
}

/// The used range's top-left, copied out of Excel as a picture.
///
/// This is what a machine with no printer gets instead of a page: the range is
/// copied the way a person copies it — `CopyPicture` — and the bitmap Excel puts
/// on the clipboard is written to the cache as a BMP, the one image format that
/// is exactly the bytes the clipboard holds. What it shows is the corner of the
/// sheet a person would see first rather than the sheet's printed layout, which
/// is the most such a machine can produce.
fn copy_used_range_picture(workbook: &Object, target: &CacheTarget) -> bool {
    let Some(sheet) = workbook
        .member("Worksheets")
        .and_then(|sheets| sheets.item(1))
    else {
        return false;
    };
    let Some(used) = sheet.member("UsedRange") else {
        return false;
    };
    let (Some(rows), Some(columns)) = (
        collection_count(used.member("Rows")),
        collection_count(used.member("Columns")),
    ) else {
        return false;
    };

    let Some(range) = picture_range(&used, rows, columns) else {
        return false;
    };

    for attempt in 0..PICTURE_COPY_ATTEMPTS {
        let copied = range
            .call_args(
                "CopyPicture",
                &[VARIANT::from(XL_SCREEN), VARIANT::from(XL_BITMAP)],
            )
            .is_some();
        if copied {
            if let Some(dib) = clipboard_dib() {
                return write_bmp(&target.file("bmp"), &dib).is_ok();
            }
        }

        if attempt + 1 < PICTURE_COPY_ATTEMPTS {
            std::thread::sleep(Duration::from_millis(PICTURE_COPY_RETRY_MS));
        }
    }

    false
}

/// The window of the used range the picture is taken from: its top-left corner, cut
/// down until what it would draw fits the box a preview could ever show.
///
/// Only the range's own measurements are asked for — never its cells — so a sheet
/// of a million rows costs the same as a small one, and what comes back is the
/// corner a person would see first rather than a page of it. Each cut is in
/// proportion to how far over the budget the range is, so a sheet whose rows are
/// tall keeps as many of them as the budget allows instead of being halved until a
/// sliver is left.
fn picture_range(used: &Object, rows: i32, columns: i32) -> Option<Object> {
    let mut rows = rows.clamp(1, PICTURE_MAX_ROWS);
    let mut columns = columns.clamp(1, PICTURE_MAX_COLUMNS);
    let mut range = resize_range(used, rows, columns);

    for _ in 0..PICTURE_FIT_ATTEMPTS {
        let Some(current) = range.as_ref() else {
            break;
        };
        let (Some(width), Some(height)) =
            (point_size(current, "Width"), point_size(current, "Height"))
        else {
            break;
        };

        let width_px = width * PICTURE_PIXELS_PER_POINT;
        let height_px = height * PICTURE_PIXELS_PER_POINT;
        if width_px <= PICTURE_MAX_PIXELS_WIDTH && height_px <= PICTURE_MAX_PIXELS_HEIGHT {
            break;
        }

        let fitted_rows = if height_px > PICTURE_MAX_PIXELS_HEIGHT {
            (rows as f64 * PICTURE_MAX_PIXELS_HEIGHT / height_px).floor() as i32
        } else {
            rows
        };
        let fitted_columns = if width_px > PICTURE_MAX_PIXELS_WIDTH {
            (columns as f64 * PICTURE_MAX_PIXELS_WIDTH / width_px).floor() as i32
        } else {
            columns
        };

        let fitted_rows = fitted_rows.clamp(PICTURE_MIN_ROWS, rows);
        let fitted_columns = fitted_columns.clamp(PICTURE_MIN_COLUMNS, columns);
        // Whatever is left is smaller than a cell: copy it as it stands.
        if fitted_rows == rows && fitted_columns == columns {
            break;
        }

        rows = fitted_rows;
        columns = fitted_columns;
        range = resize_range(used, rows, columns);
    }

    range
}

/// The top-left window of a range, `rows` by `columns` cells of it.
fn resize_range(range: &Object, rows: i32, columns: i32) -> Option<Object> {
    range
        .call_args("Resize", &[VARIANT::from(rows), VARIANT::from(columns)])
        .and_then(Object::from_variant)
}

/// One of a range's own measurements, in points.
fn point_size(range: &Object, property: &str) -> Option<f64> {
    range
        .value(property)
        .and_then(|value| f64::try_from(&value).ok())
}

/// How many items a collection holds.
fn collection_count(collection: Option<Object>) -> Option<i32> {
    collection
        .and_then(|collection| collection.value("Count"))
        .and_then(|count| i32::try_from(&count).ok())
}

/// The bitmap Excel has just put on the clipboard.
fn clipboard_dib() -> Option<Vec<u8>> {
    clipboard_dib_inner(true)
}

/// `empty` says whether the clipboard is cleared once the picture has been taken.
/// It is, in the app: leaving Excel's copy on the clipboard is what makes Excel ask,
/// on its way out, whether a large amount of information should stay there.
fn clipboard_dib_inner(empty: bool) -> Option<Vec<u8>> {
    for attempt in 0..CLIPBOARD_ATTEMPTS {
        if let Some(dib) = read_clipboard_dib(empty) {
            return Some(dib);
        }
        if attempt + 1 < CLIPBOARD_ATTEMPTS {
            std::thread::sleep(Duration::from_millis(30));
        }
    }

    None
}

fn read_clipboard_dib(empty: bool) -> Option<Vec<u8>> {
    unsafe {
        if OpenClipboard(None).is_err() {
            return None;
        }

        let mut dib = None;
        if let Ok(handle) = GetClipboardData(CF_DIB.0 as u32) {
            let handle = HGLOBAL(handle.0);
            let size = GlobalSize(handle);
            let pointer = GlobalLock(handle) as *const u8;
            if !pointer.is_null() {
                if size > 0 {
                    dib = Some(std::slice::from_raw_parts(pointer, size).to_vec());
                }
                let _ = GlobalUnlock(handle);
            }
        }

        // The picture is taken rather than borrowed. Leaving it on the clipboard is
        // what makes Excel ask, on its way out, whether a large amount of
        // information should stay there — a dialog no one is present to answer,
        // which holds the quit and with it the worker.
        if dib.is_some() && empty {
            let _ = EmptyClipboard();
        }

        let _ = CloseClipboard();
        dib
    }
}

/// What the clipboard held, written as a BMP file: a DIB is a `BITMAPINFO` and
/// its pixels, and a BMP file is those bytes with a fourteen-byte header in front
/// of them.
///
/// It is written beside its name and moved into place, so a reader never sees half
/// a picture: a hover can land on a file whose render is being written.
fn write_bmp(path: &Path, dib: &[u8]) -> std::io::Result<()> {
    let offset = dib_pixel_offset(dib).unwrap_or(54);
    let mut file = Vec::with_capacity(dib.len() + 14);
    file.extend_from_slice(b"BM");
    file.extend_from_slice(&((dib.len() + 14) as u32).to_le_bytes());
    file.extend_from_slice(&0u16.to_le_bytes()); // reserved
    file.extend_from_slice(&0u16.to_le_bytes()); // reserved
    file.extend_from_slice(&offset.to_le_bytes());
    file.extend_from_slice(dib);

    let writing = path.with_extension("bmp.writing");
    std::fs::write(&writing, file)?;
    std::fs::rename(&writing, path)
}

/// Where a DIB's pixels start, past its header and its colour table.
fn dib_pixel_offset(dib: &[u8]) -> Option<u32> {
    let header = u32::from_le_bytes(dib.get(0..4)?.try_into().ok()?) as usize;
    if header < 40 || header > dib.len() {
        return None;
    }

    let bit_count = u16::from_le_bytes(dib.get(14..16)?.try_into().ok()?) as u32;
    let used_colors = u32::from_le_bytes(dib.get(32..36)?.try_into().ok()?);
    let palette_entries = if bit_count <= 8 {
        if used_colors != 0 {
            used_colors
        } else {
            1u32 << bit_count
        }
    } else {
        0
    };

    Some((14 + header) as u32 + palette_entries * 4)
}

fn render_powerpoint(
    app: &Object,
    source: &Path,
    target: &CacheTarget,
    width: u32,
    height: u32,
) -> bool {
    let Some(presentations) = app.member("Presentations") else {
        return false;
    };
    let opened = presentations.call(
        "Open",
        &[
            ("FileName", path_variant(source)),
            ("ReadOnly", VARIANT::from(-1i32)), // msoTrue
            ("Untitled", VARIANT::from(0i32)),  // msoFalse
            ("WithWindow", VARIANT::from(0i32)),
        ],
    );
    let Some(presentation) = opened.and_then(Object::from_variant) else {
        return false;
    };

    // A slide is exported as an image rather than the deck as a PDF: the export
    // writes slide 1 alone instead of every slide the deck holds. It is reached
    // through the collection's item, which PowerPoint exposes as a method —
    // asking `Slides` itself for one is answered with "member not found".
    let rendered = presentation
        .member("Slides")
        .and_then(|slides| slides.item(1))
        .map(|slide| {
            let (export_width, export_height) = slide_export_size(&presentation, width, height);
            slide
                .call(
                    "Export",
                    &[
                        ("FileName", path_variant(&target.file("png"))),
                        ("FilterName", VARIANT::from("PNG")),
                        ("ScaleWidth", VARIANT::from(export_width)),
                        ("ScaleHeight", VARIANT::from(export_height)),
                    ],
                )
                .is_some()
        })
        .unwrap_or(false);

    let _ = presentation.call("Close", &[]);
    rendered
}

/// A single-precision property. PowerPoint records a slide's size as one, and
/// automation may hand a number back as either width.
fn single_of(value: &VARIANT) -> Option<f32> {
    f64::try_from(value).ok().map(|value| value as f32)
}

/// The size slide 1 is exported at: the width the preview asked for, bounded,
/// and the height that keeps the slide's own aspect ratio.
fn slide_export_size(presentation: &Object, width: u32, height: u32) -> (i32, i32) {
    let export_width = width.clamp(MIN_SLIDE_EXPORT_WIDTH, MAX_SLIDE_EXPORT_WIDTH) as i32;

    let slide_width = presentation
        .member("PageSetup")
        .and_then(|setup| setup.value("SlideWidth"))
        .and_then(|value| single_of(&value))
        .filter(|value| *value > 1.0);
    let slide_height = presentation
        .member("PageSetup")
        .and_then(|setup| setup.value("SlideHeight"))
        .and_then(|value| single_of(&value))
        .filter(|value| *value > 1.0);

    let ratio = match (slide_width, slide_height) {
        (Some(slide_width), Some(slide_height)) => slide_height / slide_width,
        _ => height.max(1) as f32 / width.max(1) as f32,
    };

    (
        export_width,
        ((export_width as f32 * ratio).round() as i32).max(1),
    )
}

// -------------------------------------------------------------- the sources

/// The document handed to the engine, and whether it is a copy made for it.
///
/// Office is not a consumer of the verbatim `\\?\` paths the rest of this app
/// canonicalizes to, a path longer than the plain limit is one it cannot open at
/// all, and a document carrying a zone identifier is one Word and Excel open in
/// Protected View — where the export is refused. Any of those is answered with a
/// copy in the temp folder, which also keeps the engine from locking a file the
/// user may be working in.
struct PreparedSource {
    path: PathBuf,
    copy: bool,
}

impl PreparedSource {
    fn new(source: &Path) -> Self {
        let Some(plain) = plain_path(source) else {
            return Self {
                path: source.to_path_buf(),
                copy: false,
            };
        };

        if plain.chars().count() < MAX_OFFICE_PATH && !has_zone_identifier(&plain) {
            return Self {
                path: PathBuf::from(plain),
                copy: false,
            };
        }

        match copy_to_temp(&plain) {
            Some(path) => Self { path, copy: true },
            None => Self {
                path: PathBuf::from(plain),
                copy: false,
            },
        }
    }

    fn cleanup(&self) {
        if self.copy {
            let _ = std::fs::remove_file(&self.path);
        }
    }
}

/// The path without the verbatim prefix the hook canonicalizes to.
fn plain_path(path: &Path) -> Option<String> {
    let text = path.to_string_lossy();

    if let Some(rest) = text.strip_prefix(r"\\?\UNC\") {
        return Some(format!(r"\\{rest}"));
    }
    if let Some(rest) = text.strip_prefix(r"\\?\") {
        return Some(rest.to_string());
    }

    Some(text.to_string())
}

fn has_zone_identifier(path: &str) -> bool {
    std::fs::metadata(format!("{path}:Zone.Identifier")).is_ok()
}

fn copy_to_temp(source: &str) -> Option<PathBuf> {
    static COUNTER: AtomicU32 = AtomicU32::new(0);

    let folder = std::env::temp_dir().join("rust-hover-preview");
    std::fs::create_dir_all(&folder).ok()?;

    let extension = Path::new(source)
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or("bin");
    let name = format!(
        "render-{}-{}.{extension}",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    );
    let target = folder.join(name);

    std::fs::copy(source, &target).ok()?;
    Some(target)
}

// ------------------------------------------------------------ the disk cache

fn cache_folder() -> Option<PathBuf> {
    BaseDirs::new().map(|dirs| dirs.cache_dir().join("rust-hover-preview").join("office"))
}

/// What a cached render is keyed by: the file, and the version of it that was
/// rendered.
fn cache_key(source: &Path) -> String {
    let metadata = std::fs::metadata(source).ok();
    let mut hasher = std::collections::hash_map::DefaultHasher::new();

    source.to_string_lossy().to_lowercase().hash(&mut hasher);
    if let Some(metadata) = metadata {
        metadata.len().hash(&mut hasher);
        if let Ok(modified) = metadata.modified() {
            if let Ok(since_epoch) = modified.duration_since(UNIX_EPOCH) {
                since_epoch.as_nanos().hash(&mut hasher);
            }
        }
    }

    format!("{:016x}", hasher.finish())
}

/// Where a render for this version of the file is written. Which file it is —
/// which extension — is the renderer's to choose, since what a document can be
/// drawn from is not known until it has been asked.
struct CacheTarget {
    folder: PathBuf,
    key: String,
}

impl CacheTarget {
    fn file(&self, extension: &str) -> PathBuf {
        self.folder.join(format!("{}.{extension}", self.key))
    }
}

fn cache_target(source: &Path) -> Option<CacheTarget> {
    let folder = cache_folder()?;
    std::fs::create_dir_all(&folder).ok()?;

    Some(CacheTarget {
        folder,
        key: cache_key(source),
    })
}

/// Drop the oldest renders until the cache fits its budget. A cache that cannot
/// be trimmed is not an error: what it costs is disk, not correctness.
fn trim_cache() {
    let Some(folder) = cache_folder() else {
        return;
    };
    let limit = cache_limit_bytes();
    let Ok(entries) = std::fs::read_dir(&folder) else {
        return;
    };

    let mut files: Vec<(SystemTime, u64, PathBuf)> = Vec::new();
    let mut total: u64 = 0;

    for entry in entries.flatten() {
        let Ok(metadata) = entry.metadata() else {
            continue;
        };
        if !metadata.is_file() {
            continue;
        }

        total += metadata.len();
        files.push((
            metadata.modified().unwrap_or(UNIX_EPOCH),
            metadata.len(),
            entry.path(),
        ));
    }

    if total <= limit {
        return;
    }

    files.sort_by_key(|(modified, _, _)| *modified);

    for (_, len, path) in files {
        if total <= limit {
            break;
        }
        if std::fs::remove_file(&path).is_ok() {
            total = total.saturating_sub(len);
        }
    }
}

// ----------------------------------------------------------- late binding

/// A late-bound automation object: an `IDispatch`, called by name.
///
/// Every Office application is available through `IDispatch`, so nothing here
/// depends on a type library, a build step or an interface this app would have to
/// keep in step with the installed Office. The price is that a parameter is named
/// as a string — and a name this Office does not know is dropped rather than
/// failing the call, since every parameter wanted here is optional and the
/// document's own name is the one that is not.
struct Object(IDispatch);

impl Object {
    fn create(prog_id: &str) -> Option<Self> {
        let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
        let class = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr())).ok()? };
        let dispatch: IDispatch = unsafe {
            CoCreateInstance(
                &class,
                None::<&windows::core::IUnknown>,
                CLSCTX_LOCAL_SERVER,
            )
            .ok()?
        };

        Some(Self(dispatch))
    }

    fn from_variant(value: VARIANT) -> Option<Self> {
        Some(Self(IDispatch::try_from(&value).ok()?))
    }

    fn dispatch_id(&self, name: &str) -> Option<i32> {
        let wide = wide_string(name);
        let mut id = 0i32;

        unsafe {
            self.0
                .GetIDsOfNames(&GUID::zeroed(), &PCWSTR(wide.as_ptr()), 1, 0, &mut id)
                .ok()?;
        }

        Some(id)
    }

    /// A parameter's dispatch ID, asked for with its member: the names go in as
    /// the member followed by the parameter, and the answer for the parameter is
    /// the second one.
    fn parameter_dispatch_id(&self, member: &str, parameter: &str) -> Option<i32> {
        let member = wide_string(member);
        let parameter = wide_string(parameter);
        let names = [PCWSTR(member.as_ptr()), PCWSTR(parameter.as_ptr())];
        let mut ids = [0i32; 2];

        unsafe {
            self.0
                .GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 2, 0, ids.as_mut_ptr())
                .ok()?;
        }

        Some(ids[1])
    }

    fn invoke(
        &self,
        name: &str,
        member: i32,
        flags: DISPATCH_FLAGS,
        values: &mut [VARIANT],
        ids: &mut [i32],
    ) -> Option<VARIANT> {
        let params = DISPPARAMS {
            rgvarg: if values.is_empty() {
                std::ptr::null_mut()
            } else {
                values.as_mut_ptr()
            },
            rgdispidNamedArgs: if ids.is_empty() {
                std::ptr::null_mut()
            } else {
                ids.as_mut_ptr()
            },
            cArgs: values.len() as u32,
            cNamedArgs: ids.len() as u32,
        };
        let mut result = VARIANT::new();
        // An automation failure carries its reason in the exception, not in the
        // HRESULT, so it is asked for rather than left behind.
        let mut exception = EXCEPINFO::default();

        let outcome = unsafe {
            self.0.Invoke(
                member,
                &GUID::zeroed(),
                0,
                flags,
                &params,
                Some(&mut result),
                Some(&mut exception),
                None,
            )
        };

        // A failed call is remembered with its reason: what it costs is one string
        // per failure, and what it buys is a document that can be looked at instead
        // of guessed about.
        if let Err(error) = &outcome {
            record_failure(name, error, &exception);
        }

        outcome.ok()?;
        Some(result)
    }

    /// A property's value.
    fn value(&self, name: &str) -> Option<VARIANT> {
        let member = self.dispatch_id(name)?;
        self.invoke(name, member, DISPATCH_PROPERTYGET, &mut [], &mut [])
    }

    /// An item out of a collection, reached the way VBA reaches it when it writes
    /// `Slides(1)`.
    ///
    /// Some collections expose `Item` as a property and others as a method —
    /// PowerPoint's `Slides` is the second kind, and asking it as a property is
    /// answered with "member not found" — so the invoke says it may be either,
    /// which is what both kinds answer to.
    fn item(&self, index: i32) -> Option<Self> {
        let member = self.dispatch_id("Item")?;
        let mut values = [VARIANT::from(index)];
        let value = self.invoke(
            "Item",
            member,
            DISPATCH_PROPERTYGET | DISPATCH_METHOD,
            &mut values,
            &mut [],
        )?;

        Self::from_variant(value)
    }

    /// A call with positional arguments, in the order they are written here: the
    /// parameter block carries them reversed, which is what the server expects.
    ///
    /// It is the way to reach the members whose parameter names do not resolve —
    /// `Range("A1:B2")` is one — and it needs no names to be right.
    fn call_args(&self, name: &str, args: &[VARIANT]) -> Option<VARIANT> {
        let member = self.dispatch_id(name)?;
        let mut values: Vec<VARIANT> = args.iter().rev().cloned().collect();

        self.invoke(
            name,
            member,
            DISPATCH_PROPERTYGET | DISPATCH_METHOD,
            &mut values,
            &mut [],
        )
    }

    /// An object property.
    fn member(&self, name: &str) -> Option<Self> {
        Self::from_variant(self.value(name)?)
    }

    fn set(&self, name: &str, value: VARIANT) -> Option<()> {
        let member = self.dispatch_id(name)?;
        let mut values = [value];
        let mut ids = [DISPID_PROPERTYPUT];

        self.invoke(name, member, DISPATCH_PROPERTYPUT, &mut values, &mut ids)?;
        Some(())
    }

    /// A method call, with its arguments passed by name.
    fn call(&self, name: &str, args: &[(&str, VARIANT)]) -> Option<VARIANT> {
        let member = self.dispatch_id(name)?;

        let mut values: Vec<VARIANT> = Vec::with_capacity(args.len());
        let mut ids: Vec<i32> = Vec::with_capacity(args.len());
        for (parameter, value) in args {
            let Some(id) = self.parameter_dispatch_id(name, parameter) else {
                continue;
            };
            ids.push(id);
            values.push(value.clone());
        }

        self.invoke(name, member, DISPATCH_METHOD, &mut values, &mut ids)
    }
}

fn wide_string(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}

fn path_variant(path: &Path) -> VARIANT {
    VARIANT::from(path.to_string_lossy().as_ref())
}

thread_local! {
    /// What the last automation call failed with. It is kept because a render that
    /// produces nothing is otherwise silent — the file is simply left alone for a
    /// while — and this is what says which call refused and why: the diagnostic
    /// below prints it, and a failure remembered for a file carries it.
    static LAST_FAILURE: std::cell::RefCell<Option<String>> =
        const { std::cell::RefCell::new(None) };
}

fn record_failure(name: &str, error: &windows::core::Error, exception: &EXCEPINFO) {
    // An automation failure's reason is in the exception rather than in the
    // HRESULT, which for a failed call is only DISP_E_EXCEPTION.
    let description = exception.bstrDescription.to_string();
    let detail = if description.is_empty() {
        error.message()
    } else {
        format!("{} — {description}", error.message())
    };

    LAST_FAILURE.with(|slot| {
        *slot.borrow_mut() = Some(format!("{name}: 0x{:08X} {detail}", error.code().0));
    });
}

/// What the last automation call failed with, if one did. Read by the diagnostic
/// below: a render that produces nothing is otherwise silent.
#[cfg(test)]
pub(crate) fn last_failure() -> Option<String> {
    LAST_FAILURE.with(|slot| slot.borrow().clone())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_the_plain_form_of_a_verbatim_path() {
        assert_eq!(
            plain_path(Path::new(r"\\?\C:\docs\report.docx")).as_deref(),
            Some(r"C:\docs\report.docx")
        );
        assert_eq!(
            plain_path(Path::new(r"\\?\UNC\server\share\report.docx")).as_deref(),
            Some(r"\\server\share\report.docx")
        );
        assert_eq!(
            plain_path(Path::new(r"C:\docs\report.docx")).as_deref(),
            Some(r"C:\docs\report.docx")
        );
    }

    #[test]
    fn keys_a_render_by_the_file_and_its_version() {
        // A folder of this module's own: the tests run beside each other, and one
        // of them clearing its fixtures must not take another's with it.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("render");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("keyed.docx");
        std::fs::write(&path, b"one").expect("a written document");

        let first = cache_key(&path);
        assert_eq!(first, cache_key(&path));

        std::fs::write(&path, b"a longer document").expect("a rewritten document");
        assert_ne!(first, cache_key(&path));

        let _ = std::fs::remove_file(&path);
    }

    /// Whether one Excel instance can copy more than one picture, and what the
    /// clipboard has to do with it: reads 1 and 2 leave the clipboard alone, read 3
    /// clears it, read 4 shows whether the next copy still works.
    ///
    /// `$env:RHP_OFFICE_PROBE = "C:\docs\one.xlsx;C:\docs\two.xls"`
    /// `cargo test -- --ignored --nocapture copy_picture_repeat_probe`
    #[test]
    #[ignore = "starts the installed Excel"]
    fn copy_picture_repeat_probe() {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }

        let Ok(list) = std::env::var("RHP_OFFICE_PROBE") else {
            println!("set RHP_OFFICE_PROBE to one or more paths, separated by ';'");
            return;
        };

        let Some(engine) = Engine::create(OfficeApp::Excel) else {
            println!("no Excel");
            return;
        };

        for (index, path) in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
            .enumerate()
        {
            println!("\n--- {path} ---");
            let source = PreparedSource::new(Path::new(path));
            let Some(workbook) = engine
                .app
                .member("Workbooks")
                .and_then(|workbooks| {
                    workbooks.call(
                        "Open",
                        &[
                            ("FileName", path_variant(&source.path)),
                            ("UpdateLinks", VARIANT::from(0i32)),
                            ("ReadOnly", VARIANT::from(true)),
                            ("AddToMru", VARIANT::from(false)),
                            ("IgnoreReadOnlyRecommended", VARIANT::from(true)),
                        ],
                    )
                })
                .and_then(Object::from_variant)
            else {
                println!("open failed");
                source.cleanup();
                continue;
            };

            let range = workbook
                .member("Worksheets")
                .and_then(|sheets| sheets.item(1))
                .and_then(|sheet| sheet.member("UsedRange"))
                .and_then(|used| resize_range(&used, 10, 8));

            for attempt in 1..=4 {
                let Some(range) = range.as_ref() else {
                    break;
                };
                let copied = range
                    .call_args(
                        "CopyPicture",
                        &[VARIANT::from(XL_SCREEN), VARIANT::from(XL_BITMAP)],
                    )
                    .is_some();
                // The third read is the one the app's own path does — it clears the
                // clipboard; the others leave it as Excel left it.
                let dib = clipboard_dib_inner(attempt == 3);
                println!(
                    "  attempt {attempt} (file {}): copied={copied} picture={:?} failure={:?}",
                    index + 1,
                    dib.as_deref().and_then(dib_dimensions),
                    last_failure()
                );
            }

            let _ = workbook.call("Close", &[("SaveChanges", VARIANT::from(false))]);
            source.cleanup();
        }

        drop(engine);
    }

    /// What Excel reports for a range's own size, next to what the range's first
    /// rows and columns report — the two do not agree, which is why the picture's
    /// window is fitted the way it is.
    ///
    /// `$env:RHP_OFFICE_PROBE = "C:\docs\one.xlsx"`
    /// `cargo test -- --ignored --nocapture worksheet_measurement_probe`
    #[test]
    #[ignore = "starts the installed Excel"]
    fn worksheet_measurement_probe() {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }

        let Ok(list) = std::env::var("RHP_OFFICE_PROBE") else {
            println!("set RHP_OFFICE_PROBE to one or more paths, separated by ';'");
            return;
        };

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
        {
            println!("\n--- {path} ---");
            let source = PreparedSource::new(Path::new(path));

            // A fresh instance for every file, so a refusal can be told from one the
            // instance was already in.
            let Some(engine) = Engine::create(OfficeApp::Excel) else {
                println!("no Excel");
                source.cleanup();
                continue;
            };
            let Some(workbooks) = engine.app.member("Workbooks") else {
                source.cleanup();
                continue;
            };
            let Some(workbook) = workbooks
                .call(
                    "Open",
                    &[
                        ("FileName", path_variant(&source.path)),
                        ("UpdateLinks", VARIANT::from(0i32)),
                        ("ReadOnly", VARIANT::from(true)),
                        ("AddToMru", VARIANT::from(false)),
                        ("IgnoreReadOnlyRecommended", VARIANT::from(true)),
                    ],
                )
                .and_then(Object::from_variant)
            else {
                println!("open failed");
                source.cleanup();
                continue;
            };

            if let Some(used) = workbook
                .member("Worksheets")
                .and_then(|sheets| sheets.item(1))
                .and_then(|sheet| sheet.member("UsedRange"))
            {
                println!(
                    "used: {}x{} cells, {}x{} points",
                    collection_count(used.member("Rows")).unwrap_or(0),
                    collection_count(used.member("Columns")).unwrap_or(0),
                    point_size(&used, "Width").unwrap_or(f64::NAN),
                    point_size(&used, "Height").unwrap_or(f64::NAN),
                );
                for rows in [40, 20, 10, 5, 4] {
                    match resize_range(&used, rows, 14) {
                        Some(range) => println!(
                            "  resize({rows}, 14): {}x{} points",
                            point_size(&range, "Width").unwrap_or(f64::NAN),
                            point_size(&range, "Height").unwrap_or(f64::NAN),
                        ),
                        None => println!("  resize({rows}, 14): none"),
                    }

                    // What that window actually comes out as, which is the only
                    // answer that matters: the copy is what the picture is.
                    if let Some(range) = resize_range(&used, rows, 14) {
                        let copied = range
                            .call_args(
                                "CopyPicture",
                                &[VARIANT::from(XL_SCREEN), VARIANT::from(XL_BITMAP)],
                            )
                            .is_some();
                        match copied.then(clipboard_dib).flatten() {
                            Some(dib) => println!("    copy: {:?} pixels", dib_dimensions(&dib)),
                            None => println!(
                                "    copy: nothing — {}",
                                last_failure().unwrap_or_else(|| "no failure recorded".to_string())
                            ),
                        }
                    }
                }
                for rows in [1, 4, 9, 40] {
                    let height = used
                        .call_args("Rows", &[VARIANT::from(rows)])
                        .and_then(Object::from_variant)
                        .and_then(|range| point_size(&range, "Height"));
                    println!("  rows({rows}).Height: {height:?}");
                }
                for columns in [1, 3, 14] {
                    let width = used
                        .call_args("Columns", &[VARIANT::from(columns)])
                        .and_then(Object::from_variant)
                        .and_then(|range| point_size(&range, "Width"));
                    println!("  columns({columns}).Width: {width:?}");
                }
            }

            let _ = workbook.call("Close", &[("SaveChanges", VARIANT::from(false))]);
            source.cleanup();
            drop(engine);
        }
    }

    /// The engine's own lifecycle: one is started for each family, let go, and
    /// nothing is left behind.
    #[test]
    #[ignore = "starts the installed Office"]
    fn engine_lifecycle_probe() {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }

        for app_kind in [OfficeApp::Word, OfficeApp::Excel, OfficeApp::PowerPoint] {
            println!("\n--- {app_kind:?} ---");
            println!("before: {:?}", processes_named(app_kind.image_name()));

            let Some(engine) = Engine::create(app_kind) else {
                println!("created: no");
                continue;
            };
            println!(
                "created: attached={} owned_pid={}",
                engine.attached, engine.owned_pid
            );

            drop(engine);
            std::thread::sleep(Duration::from_secs(3));
            println!("after: {:?}", processes_named(app_kind.image_name()));
        }
    }

    /// A cache file that cannot be read is not a page: the side that draws a preview
    /// drops it, so the document is rendered again rather than answered with a
    /// preview that blinks away every time it is hovered.
    #[test]
    fn forgets_a_cache_file_it_cannot_read() {
        use crate::office_preview::{measure, source_kind, SourceKind};

        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("cache");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let source = folder.join("cached.docx");
        std::fs::write(&source, b"a document").expect("a written document");

        let cache = cache_folder().expect("a cache folder");
        std::fs::create_dir_all(&cache).expect("a cache folder");
        let rendered = cache.join(format!("{}.bmp", cache_key(&source)));

        // A picture that can be read is a page, and its own size is what the layout
        // places the preview by.
        std::fs::write(&rendered, bmp_bytes(2, 2, [10, 20, 30, 255])).expect("a written page");
        assert_eq!(measure(&source), Some((2, 2)), "a page that can be read");
        assert_eq!(source_kind(&source), SourceKind::Raster);

        // One that cannot is dropped, and the answer is that nothing is rendered yet.
        std::fs::write(&rendered, b"not a picture at all").expect("a written broken page");
        assert_eq!(source_kind(&source), SourceKind::None, "nothing to draw");
        assert!(!rendered.exists(), "the broken file is gone");

        let _ = std::fs::remove_file(&source);
    }

    /// A DIB's own size, as a picture of it would be.
    fn dib_dimensions(dib: &[u8]) -> Option<(i32, i32)> {
        let width = i32::from_le_bytes(dib.get(4..8)?.try_into().ok()?);
        let height = i32::from_le_bytes(dib.get(8..12)?.try_into().ok()?);
        Some((width, height))
    }

    /// A BMP holding one colour, written the way a workbook's picture is.
    fn bmp_bytes(width: u32, height: u32, color: [u8; 4]) -> Vec<u8> {
        let pixel_bytes = (width * height * 4) as u32;
        let mut dib = Vec::new();
        dib.extend_from_slice(&40u32.to_le_bytes());
        dib.extend_from_slice(&(width as i32).to_le_bytes());
        dib.extend_from_slice(&(height as i32).to_le_bytes());
        dib.extend_from_slice(&1u16.to_le_bytes());
        dib.extend_from_slice(&32u16.to_le_bytes());
        dib.extend_from_slice(&0u32.to_le_bytes());
        dib.extend_from_slice(&pixel_bytes.to_le_bytes());
        for _ in 0..4 {
            dib.extend_from_slice(&0i32.to_le_bytes());
        }
        for _ in 0..(width * height) {
            dib.extend_from_slice(&color);
        }

        let mut file = Vec::new();
        file.extend_from_slice(b"BM");
        file.extend_from_slice(&((dib.len() + 14) as u32).to_le_bytes());
        file.extend_from_slice(&0u16.to_le_bytes());
        file.extend_from_slice(&0u16.to_le_bytes());
        file.extend_from_slice(&54u32.to_le_bytes());
        file.extend_from_slice(&dib);
        file
    }

    /// A file that was downloaded carries a zone identifier, which is what puts
    /// Word and Excel into Protected View — where a page cannot be exported — so it
    /// is rendered from a copy made without it. A file that was never downloaded
    /// must not be copied: a large document is expensive to copy for nothing.
    #[test]
    fn sees_a_zone_identifier() {
        // A folder of this module's own: the tests run beside each other, and one
        // of them clearing its fixtures must not take another's with it.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("zones");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let path = folder.join("marked.docx");
        std::fs::write(&path, b"a document").expect("a written document");
        let plain = path.to_string_lossy().to_string();
        assert!(!has_zone_identifier(&plain), "nothing marks it yet");

        // The stream Windows writes beside a downloaded file, written the way any
        // process could write it.
        std::fs::write(
            format!("{plain}:Zone.Identifier"),
            b"[ZoneTransfer]\r\nZoneId=3\r\n",
        )
        .expect("a written stream");
        assert!(has_zone_identifier(&plain), "the stream is seen");

        let _ = std::fs::remove_file(&path);
    }

    /// Renders documents that are already on disk, named in `RHP_OFFICE_PROBE`
    /// (separated by `;`), through the real code path and reports what happened.
    ///
    /// Ignored like the smoke test, and the way to look at a document that will not
    /// preview:
    /// `$env:RHP_OFFICE_PROBE = "C:\docs\one.xlsx;C:\docs\two.xls"`
    /// `cargo test -- --ignored --nocapture office_render_probe`
    #[test]
    #[ignore = "starts the installed Office"]
    fn office_render_probe() {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }

        let Ok(list) = std::env::var("RHP_OFFICE_PROBE") else {
            println!("set RHP_OFFICE_PROBE to one or more paths, separated by ';'");
            return;
        };

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
        {
            let path = PathBuf::from(path);
            println!("\n--- {} ---", path.display());
            println!(
                "exists: {} container: {:?} engine: {:?}",
                path.exists(),
                container_kind(&path),
                app_for(&path)
            );

            let mut engine = None;
            let request = RenderRequest {
                source: path.clone(),
                width: 1280,
                height: 800,
                generation: 1,
                requested: Instant::now(),
            };

            let started = Instant::now();
            let outcome = render_request(&mut engine, &request);
            println!(
                "rendered: {} in {:?}",
                outcome == RenderOutcome::Rendered,
                started.elapsed()
            );
            match engine.as_ref() {
                Some(engine) => println!(
                    "engine: attached={} owned_pid={}",
                    engine.attached, engine.owned_pid
                ),
                None => println!("engine: none"),
            }
            println!(
                "last failure: {}",
                last_failure().unwrap_or_else(|| "none recorded".to_string())
            );
            match cached_render(&path) {
                Some(cached) => println!(
                    "cache: {} ({} bytes)",
                    cached.path.display(),
                    std::fs::metadata(&cached.path)
                        .map(|m| m.len())
                        .unwrap_or(0)
                ),
                None => println!("cache: none"),
            }

            // The half a hover does after the render: measure it, then draw it —
            // on a thread of its own, in a multithreaded apartment, which is where
            // the app does both.
            let drawing = path.clone();
            let (measured, drawn) = std::thread::spawn(move || {
                crate::pdf_preview::initialize_apartment();
                let measured = crate::office_preview::measure(&drawing);
                let drawn = crate::office_preview::render(&drawing, 1200, 900, None);
                (measured, drawn)
            })
            .join()
            .expect("the drawing thread");

            println!("measured: {measured:?}");
            match drawn {
                Some((pixels, width, height)) => {
                    println!("drawn: {width}x{height} ({} pixels)", pixels.len() / 4)
                }
                None => println!("drawn: nothing"),
            }

            drop(engine);
        }
    }

    /// A document of the given family, written by the application itself, with a
    /// little content in it so that a page has something to show.
    fn write_sample(app_kind: OfficeApp, path: &Path) -> Option<()> {
        let app = Object::create(app_kind.prog_id())?;
        // The hygiene the engine applies, which a fixture that leaves it out hangs
        // on: an alert is a dialog no one is present to answer.
        let _ = app.set("DisplayAlerts", alerts_off(app_kind));

        let collection = app.member(match app_kind {
            OfficeApp::Word => "Documents",
            OfficeApp::Excel => "Workbooks",
            OfficeApp::PowerPoint => "Presentations",
        })?;
        let document = Object::from_variant(collection.call("Add", &[])?)?;

        match app_kind {
            OfficeApp::Word => {}
            OfficeApp::Excel => {
                if let Some(sheet) = document
                    .member("Worksheets")
                    .and_then(|sheets| sheets.item(1))
                {
                    for (cell, value) in [
                        ("A1", "Item"),
                        ("B1", "Qty"),
                        ("A2", "Widget"),
                        ("B2", "3"),
                        ("A3", "Gadget"),
                        ("B3", "7"),
                    ] {
                        if let Some(range) = sheet
                            .call_args("Range", &[VARIANT::from(cell)])
                            .and_then(Object::from_variant)
                        {
                            let _ = range.set("Value2", VARIANT::from(value));
                        }
                    }
                }
            }
            OfficeApp::PowerPoint => {
                if let Some(slides) = document.member("Slides") {
                    let _ = slides.call_args(
                        "Add",
                        &[VARIANT::from(1i32), VARIANT::from(1i32)], // index, layout
                    );
                }
            }
        }

        let saved = match app_kind {
            OfficeApp::Word => document.call(
                "SaveAs2",
                &[
                    ("FileName", path_variant(path)),
                    ("FileFormat", VARIANT::from(12i32)), // wdFormatXMLDocument
                ],
            ),
            OfficeApp::Excel => document.call(
                "SaveAs",
                &[
                    ("Filename", path_variant(path)),
                    ("FileFormat", VARIANT::from(51i32)), // xlOpenXMLWorkbook
                ],
            ),
            OfficeApp::PowerPoint => document.call(
                "SaveAs",
                &[
                    ("FileName", path_variant(path)),
                    ("Format", VARIANT::from(24i32)), // ppSaveAsOpenXMLPresentation
                ],
            ),
        };

        let _ = document.call("Close", &[("SaveChanges", VARIANT::from(false))]);
        let _ = app.call("Quit", &[]);

        saved.map(|_| ())
    }

    /// One document of each family, written and then rendered through the real
    /// code path, reporting what every step did.
    ///
    /// Ignored because it starts the installed Office and writes sample documents
    /// into the scratchpad. Run it when a document produces no page:
    /// `cargo test -- --ignored --nocapture office_render_smoke_test`.
    #[test]
    #[ignore = "starts the installed Office and writes sample documents"]
    fn office_render_smoke_test() {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }

        let folder = std::env::var_os("COMMANDCODE_SCRATCHPAD")
            .map(PathBuf::from)
            .unwrap_or_else(std::env::temp_dir)
            .join("office-smoke-test");
        std::fs::create_dir_all(&folder).expect("a scratch folder");

        for (app_kind, name) in [
            (OfficeApp::Word, "sample.docx"),
            (OfficeApp::Excel, "sample.xlsx"),
            (OfficeApp::PowerPoint, "sample.pptx"),
        ] {
            println!("\n--- {name} ---");
            let path = folder.join(name);
            let _ = std::fs::remove_file(&path);

            let started = Instant::now();
            match write_sample(app_kind, &path) {
                Some(()) => println!("written: {} in {:?}", path.display(), started.elapsed()),
                None => {
                    println!(
                        "written: no ({})",
                        last_failure().unwrap_or_else(|| "no failure recorded".to_string())
                    );
                    continue;
                }
            }

            // What the hover would show before anything is rendered: nothing, so the
            // spinner's own box is what the layout places, and the page is asked for
            // in the box that family's pages have.
            println!(
                "before a render: source {:?}, spinner box {}, render box {:?}",
                crate::office_preview::source_kind(&path),
                crate::office_preview::WAITING_BOX,
                crate::office_formats::default_page_size(&path)
            );

            let mut engine = None;
            let request = RenderRequest {
                source: path.clone(),
                width: 1280,
                height: 800,
                generation: 1,
                requested: Instant::now(),
            };
            let started = Instant::now();
            let outcome = render_request(&mut engine, &request);
            println!(
                "rendered: {} in {:?}",
                outcome == RenderOutcome::Rendered,
                started.elapsed()
            );

            // What the engine made of the instance: an attached one is the user's
            // and is never hidden, quit or ended; one this app started is.
            match engine.as_ref() {
                Some(engine) => println!(
                    "engine: attached={} owned_pid={}",
                    engine.attached, engine.owned_pid
                ),
                None => println!("engine: none"),
            }

            if outcome != RenderOutcome::Rendered {
                println!(
                    "last failure: {}",
                    last_failure().unwrap_or_else(|| "none recorded".to_string())
                );
            }
            match cached_render(&path) {
                Some(cached) => println!(
                    "cache file: {} ({} bytes, {:?})",
                    cached.path.display(),
                    std::fs::metadata(&cached.path)
                        .map(|m| m.len())
                        .unwrap_or(0),
                    cached.kind
                ),
                None => println!("cache file: none"),
            }

            // Drawing happens on a thread of its own, in a multithreaded
            // apartment, which is where the app draws too: this thread is
            // apartment-threaded for the automation above, and a WinRT call waited
            // on from a single-threaded apartment deadlocks without a pump.
            let drawing = path.clone();
            let started = Instant::now();
            let drawn = std::thread::spawn(move || {
                crate::pdf_preview::initialize_apartment();
                crate::office_preview::render(&drawing, 800, 600, None)
            })
            .join()
            .ok()
            .flatten();

            match drawn {
                Some((pixels, width, height)) => println!(
                    "drawn: {width}x{height}, {} pixels, in {:?}",
                    pixels.len() / 4,
                    started.elapsed()
                ),
                None => println!("drawn: nothing, in {:?}", started.elapsed()),
            }

            drop(engine);
        }
    }
}

//! The Office render tier: Word, Excel or PowerPoint draws a document's first
//! page, once, into a cache file the preview then draws from.
//!
//! The saved thumbnail is the fast path (see `office_thumbnail`), but not every
//! document has one — Excel only writes one when "Save Thumbnails" is on, a macro
//! that saves a workbook writes none, and anything produced outside Office may
//! carry no picture at all. This is what covers those documents, and it is also
//! what replaces a small thumbnail with a real page for a hover that rests.
//!
//! Nothing here is ever on the hover path. A render is asked for only after the
//! pointer has rested on a file, it runs on a thread of its own, and the preview
//! shows the thumbnail (or a spinner) until the page is there.
//!
//! Two rules shape the rest:
//!
//! * **The user's Office is never disturbed.** An automation instance may attach
//!   to a running Word or Excel — the applications are registered for multiple
//!   use — so nothing is hidden that is already visible, the settings that are
//!   changed are restored, and only an instance this app created is quit.
//! * **A stuck engine must not wedge the app.** The calls below are COM calls
//!   into another process and cannot be cancelled or bounded; the thread that
//!   makes them may be lost to a modal dialog inside Office, and what that costs
//!   is the render tier, never the preview window. Nothing waits on this thread.

use crate::cloud_files;
use crate::config::sanitize_office_cache_mb;
use crate::office_formats::{app_for, OfficeApp};
use crate::preview_window;
use crate::CONFIG;
use directories::BaseDirs;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use windows::core::{GUID, PCWSTR, VARIANT};
use windows::Win32::Foundation::{HGLOBAL, LPARAM, WPARAM};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, IDispatch, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
};
use windows::Win32::System::DataExchange::{CloseClipboard, GetClipboardData, OpenClipboard};
use windows::Win32::System::Memory::{GlobalLock, GlobalSize, GlobalUnlock};
use windows::Win32::System::Ole::CF_DIB;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, MsgWaitForMultipleObjectsEx, PeekMessageW, PostThreadMessageW,
    TranslateMessage, MSG, MWMO_INPUTAVAILABLE, PM_REMOVE, QS_ALLINPUT, WM_APP,
};

/// The message a request is announced with on the worker's own thread queue.
const WM_OFFICE_RENDER: u32 = WM_APP + 3;

/// How long the engine is kept alive after its last render, so a folder of
/// documents costs one Office start rather than one per file.
const ENGINE_IDLE_SECS: u64 = 60;
/// How long a file that failed to render is left alone. Office refused it for a
/// reason — a password, a repair dialog, a document in Protected View — and the
/// answer will not be different a moment later, so the wait is the user's.
const FAILURE_BACKOFF: Duration = Duration::from_secs(600);
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
/// The clipboard is shared with every other process: a look that cannot open it,
/// or finds nothing in it, is retried this many times.
const CLIPBOARD_ATTEMPTS: usize = 5;
/// `xlScreen` and `xlBitmap`: the appearance and format `CopyPicture` is asked for.
const XL_SCREEN: i32 = 1;
const XL_BITMAP: i32 = 2;
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
static WORKER_THREAD: AtomicU32 = AtomicU32::new(0);
static WORKER_BUSY: AtomicBool = AtomicBool::new(false);

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

/// The page rendered for this version of the file, if one is waiting.
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
        failures.insert(source.to_path_buf(), failure);
    }
}

/// Stop the worker, joining it only when it is not inside a render: a COM call
/// into Office cannot be cancelled, and waiting on one would hold the app's exit
/// for as long as Office takes.
pub(crate) fn shutdown() {
    if WORKER_BUSY.load(Ordering::Acquire) {
        return;
    }

    if let Ok(mut handle) = WORKER_HANDLE.lock() {
        if let Some(handle) = handle.take() {
            let _ = handle.join();
        }
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
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        // The thread's message queue has to exist before another thread can post
        // to it, and it only exists once something has asked for it.
        let mut message = MSG::default();
        let _ = PeekMessageW(&mut message, None, 0, 0, PM_REMOVE);

        WORKER_THREAD.store(GetCurrentThreadId(), Ordering::Release);
    }

    let mut engine: Option<Engine> = None;
    let mut idle_since = Instant::now();

    while crate::RUNNING.load(Ordering::Acquire) {
        pump_messages();

        if let Some(request) = take_request() {
            if request.requested.elapsed() > Duration::from_secs(REQUEST_STALE_SECS) {
                continue;
            }

            WORKER_BUSY.store(true, Ordering::Release);
            let rendered = render_request(&mut engine, &request);
            WORKER_BUSY.store(false, Ordering::Release);
            idle_since = Instant::now();

            if rendered {
                trim_cache();
            } else {
                remember_failure(&request.source);
            }

            preview_window::notify_office_render(&request.source, request.generation, rendered);
            continue;
        }

        // Nothing to do. The engine is kept warm for a minute after the last
        // render and then let go, which is also when this thread ends: an idle
        // app should have no Office process and no polling thread.
        if engine.is_some() && idle_since.elapsed() >= Duration::from_secs(ENGINE_IDLE_SECS) {
            engine = None;
        }
        if engine.is_none() {
            break;
        }

        wait_for_message(500);
    }

    // A request that arrived while this thread was shutting down would otherwise
    // sit in the slot with nobody to run it.
    if take_request().is_some() && crate::RUNNING.load(Ordering::Acquire) {
        start_worker();
    }

    unsafe {
        windows::Win32::System::Com::CoUninitialize();
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

fn render_request(engine: &mut Option<Engine>, request: &RenderRequest) -> bool {
    let Some(app_kind) = app_for(&request.source) else {
        return false;
    };

    // The gate every loader asks before it reads: a document whose content is
    // still in the cloud would be downloaded by the open, and the engine is not
    // an exception to that.
    if cloud_files::needs_download(&request.source) {
        return false;
    }

    // Another hover may have produced this page while this request waited.
    if cached_render(&request.source).is_some() {
        return true;
    }

    let Some(target) = cache_target(&request.source) else {
        return false;
    };

    // A different family's engine is dropped first, which quits it: an engine is
    // kept for the family it was created for and nothing else.
    match engine {
        Some(engine) if engine.app_kind == app_kind => {}
        _ => {
            *engine = None;
            let Some(created) = Engine::create(app_kind) else {
                return false;
            };
            *engine = Some(created);
        }
    }

    let Some(engine) = engine.as_ref() else {
        return false;
    };

    let source = PreparedSource::new(&request.source);
    let width = request.width.max(1);
    let height = request.height.max(1);
    let rendered = engine.render(&source.path, &target, width, height);
    source.cleanup();

    // What the renderer wrote is the answer, whatever it chose to write: the
    // cache says which file it was.
    rendered && cached_render(&request.source).is_some()
}

// ------------------------------------------------------------------ engines

/// An automation instance, and what it takes to leave it as it was found.
struct Engine {
    app_kind: OfficeApp,
    app: Object,
    /// Whether the instance was already running for the user when it was
    /// reached. Such an instance is never hidden and never quit.
    attached: bool,
    previous_alerts: Option<VARIANT>,
    previous_security: Option<VARIANT>,
}

impl Engine {
    fn create(app_kind: OfficeApp) -> Option<Self> {
        let app = Object::create(app_kind.prog_id())?;

        // An instance this app created is hidden already; the user's is visible,
        // and hiding it would take their windows away.
        let attached = app
            .value("Visible")
            .and_then(|value| bool::try_from(&value).ok())
            .unwrap_or(false);

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
    }
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

    let rows = rows.clamp(1, PICTURE_MAX_ROWS);
    let columns = columns.clamp(1, PICTURE_MAX_COLUMNS);
    let Some(range) = used
        .call_args("Resize", &[VARIANT::from(rows), VARIANT::from(columns)])
        .and_then(Object::from_variant)
    else {
        return false;
    };

    let copied = range
        .call_args(
            "CopyPicture",
            &[VARIANT::from(XL_SCREEN), VARIANT::from(XL_BITMAP)],
        )
        .is_some();
    if !copied {
        return false;
    }

    let Some(dib) = clipboard_dib() else {
        return false;
    };

    write_bmp(&target.file("bmp"), &dib).is_ok()
}

/// How many items a collection holds.
fn collection_count(collection: Option<Object>) -> Option<i32> {
    collection
        .and_then(|collection| collection.value("Count"))
        .and_then(|count| i32::try_from(&count).ok())
}

/// The bitmap Excel has just put on the clipboard.
fn clipboard_dib() -> Option<Vec<u8>> {
    for attempt in 0..CLIPBOARD_ATTEMPTS {
        if let Some(dib) = read_clipboard_dib() {
            return Some(dib);
        }
        if attempt + 1 < CLIPBOARD_ATTEMPTS {
            std::thread::sleep(Duration::from_millis(30));
        }
    }

    None
}

fn read_clipboard_dib() -> Option<Vec<u8>> {
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

        let _ = CloseClipboard();
        dib
    }
}

/// What the clipboard held, written as a BMP file: a DIB is a `BITMAPINFO` and
/// its pixels, and a BMP file is those bytes with a fourteen-byte header in front
/// of them.
fn write_bmp(path: &Path, dib: &[u8]) -> std::io::Result<()> {
    let offset = dib_pixel_offset(dib).unwrap_or(54);
    let mut file = Vec::with_capacity(dib.len() + 14);
    file.extend_from_slice(b"BM");
    file.extend_from_slice(&((dib.len() + 14) as u32).to_le_bytes());
    file.extend_from_slice(&0u16.to_le_bytes()); // reserved
    file.extend_from_slice(&0u16.to_le_bytes()); // reserved
    file.extend_from_slice(&offset.to_le_bytes());
    file.extend_from_slice(dib);

    std::fs::write(path, file)
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

    /// A document of the given family, written by the application itself, with a
    /// little content in it so that a page has something to show.
    fn write_sample(app_kind: OfficeApp, path: &Path) -> Option<()> {
        let app = Object::create(app_kind.prog_id())?;
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

        let _ = document.call("Close", &[]);
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

            let started = Instant::now();
            match crate::office_thumbnail::thumbnail_for(&path, None) {
                Some(thumbnail) => println!(
                    "saved picture: {:?} {}x{} read in {:?}",
                    thumbnail.kind,
                    thumbnail.width,
                    thumbnail.height,
                    started.elapsed()
                ),
                None => println!("saved picture: none, looked for in {:?}", started.elapsed()),
            }

            let started = Instant::now();
            match crate::office_preview::render(&path, 800, 600, None) {
                Some((pixels, width, height)) => println!(
                    "drawn: {width}x{height}, {} pixels, in {:?}",
                    pixels.len() / 4,
                    started.elapsed()
                ),
                None => println!("drawn: nothing, in {:?}", started.elapsed()),
            }

            let mut engine = None;
            let request = RenderRequest {
                source: path.clone(),
                width: 1280,
                height: 800,
                generation: 1,
                requested: Instant::now(),
            };
            let started = Instant::now();
            let rendered = render_request(&mut engine, &request);
            println!("rendered: {rendered} in {:?}", started.elapsed());

            if !rendered {
                println!(
                    "last failure: {}",
                    last_failure().unwrap_or_else(|| "none recorded".to_string())
                );
            }
            match cached_render(&path) {
                Some(cached) => println!(
                    "cache file: {} ({} bytes)",
                    cached.path.display(),
                    std::fs::metadata(&cached.path)
                        .map(|m| m.len())
                        .unwrap_or(0)
                ),
                None => println!("cache file: none"),
            }

            drop(engine);
        }
    }
}

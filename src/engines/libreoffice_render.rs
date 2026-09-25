//! Design documents rendered by an installed LibreOffice.
//!
//! A CorelDRAW document is a vector drawing, and what this app reads of one by itself is
//! the picture the application saved for a file manager — small, and soft when the
//! preview is enlarged. The drawing itself is in a proprietary object format three
//! generations deep, which is why this app has no reader for it. LibreOffice has one: its
//! import filters ship the Document Liberation Project's libraries, `libcdr` among them,
//! and those parse the drawing rather than its thumbnail. So where LibreOffice is
//! installed it is asked first, and what it hands back is a PDF — a page of vector data,
//! which the PDF path already draws at whatever size a preview is shown at. Where it is
//! not installed, nothing changes: the readers of the pictures the files carry are the
//! fallback, and they are also what answers a document the engine cannot read.
//!
//! Nothing is bundled with this app and nothing is linked against: the engine is the
//! user's own installation, looked for where it installs — and beside `config.ini` for a
//! portable copy — and run as the user runs it. What that costs is a launch, which is
//! seconds, so a conversion happens once per document: the PDF it wrote is kept as a page,
//! named for the document, the version of it that was converted and this engine, and every
//! hover after the first is a read of that file (see `document_cache`). One conversion runs at
//! a time, because one LibreOffice at a time is what its own profile allows.
//!
//! The engine is asked about the names its filters read and no others — the formats of
//! the libraries above, CorelDRAW's and the rest — so a Photoshop document, a Krita
//! project or a Procreate file never pays for a launch that could not answer.
//!
//! One document of another kind is asked about as well, and it is the one case the list is
//! not consulted about: an Office document whose own application is not installed. There is
//! no Word, Excel or PowerPoint on the machine to draw a page for one, and this engine's
//! filters read the format, so the page is this engine's and the hover shows it as the
//! Office document the file is, at that kind's own scale (see `libre_formats`). Which files
//! those are is answered in one place, `libre_formats::engine_page_kind`, and it is asked by
//! the caller and by [`request`] alike: a page asked for by one side and refused by the
//! other would be a hover waiting for a conversion nothing was ever asked to make.
//!
//! A conversion is seconds, and it is not one a preview can wait on: the caller is the
//! preview loop, and a loop held inside a launch is a hover that does not come up, a tray
//! that does not answer and a pointer that cannot leave the file it is on. So the engine
//! runs on a thread of its own. What the loop asks is [`request`], which returns at once,
//! and what it waits on is the page appearing under [`rendered_page`] — the same wait an
//! Office document has, in the same box, with the hover replayed when the page lands.
//! Nothing on the drawing side converts a document either: the loader asks for a page that
//! is there and reads it, and a page that is not there yet is the wait above. The engine
//! draws one document at a time, because one profile is one seat, and the document waiting
//! behind it is the newest one asked for.
//!
//! An engine that has stopped answering is the other half of that. The filters of a
//! document the engine cannot really read do not always fail: they can spin — measured on
//! a Flash file, one core at a hundred percent, no page, past every bound — and a
//! conversion like that holds the seat, burns a core and would cost every document after
//! it. A conversion that has run longer than any conversion takes is therefore ended by
//! name and id, from both sides of it: the engine thread gives up on the process it is
//! waiting on, and a document asked for in the meantime ends that same process rather than
//! queueing behind it. The document it was on is remembered as one the engine will not
//! draw, so the launch is paid for once and never again, and the preview that asked for it
//! is answered rather than left spinning. What is ended is the engine itself rather than
//! the process that asked it for a page: a conversion handed to an engine that is being
//! kept runs *inside* it (see below), so a run that has stopped answering is ended by
//! letting the engine go, which is also what tells the next document it has a fresh one.
//!
//! What a launch costs is an application start — filter chains, fonts, the libraries above
//! — and it is paid again for every document, which is what the idle setting is about. An
//! engine started for one document can be kept for the next: started headless with a
//! document of this app's own held open, it is a running instance, and a `--convert-to`
//! issued beside it is handed to *that* instance rather than starting a second one — the
//! page it writes lands in the folder this side reads, and the process that was spawned to
//! ask for it does no work at all and exits. Measured on one document, 1.2 s cold against
//! 0.6 s handed to a kept engine; the work itself is the same either way, and what is
//! saved is the start. The document it holds is a stub of this app's own — a one-paragraph
//! flat OpenDocument text file written beside the profile — because an instance holding
//! nothing is not a running instance at all: started with `--accept` (a UNO socket) or
//! with `--invisible` and no document, it either takes no work or is evicted by the next
//! launch, and both were measured. Holding one costs what any LibreOffice costs to keep
//! open — a few hundred megabytes — which is what the setting is for: `LibreOffice TTL`
//! names how long past its last conversion the engine is kept, and `0 seconds` is the
//! setting switched off, which is an engine per document, exactly as it was.

use crate::config::config::{AppConfig, EngineIdle, OfficeEngine, DEFAULT_LIBREOFFICE_IDLE_SECS};
use crate::engines::document_cache::{self, Page, PageKind};
use crate::CONFIG;
use directories::BaseDirs;
use once_cell::sync::Lazy;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Condvar, Mutex};
use std::time::{Duration, Instant};

/// How long a conversion is given, and the point past which the engine is a stopped one
/// rather than a busy one: the conversion is ended where it stands, and the document it
/// was on is remembered as one the engine will not draw.
///
/// A conversion of an ordinary document is one to three seconds on the machine this was
/// measured on, the profile's first one included, so this is an order of magnitude past
/// anything a document the engine can draw costs — and a document it cannot draw does not
/// finish at all, whether it fails or spins. One number rather than two, because a
/// document waiting behind a hung engine is waiting on the same question, and the answer
/// is not a longer wait: it is that engine being ended. Waiting the longer of two bounds
/// would mean a hover that came after the stuck one paying the difference for nothing.
const CONVERSION_GIVE_UP: Duration = Duration::from_secs(30);
/// How often the wait above looks.
const CONVERSION_POLL: Duration = Duration::from_millis(100);

/// The engine's own program, by the name a record carries: what a conversion that is given
/// up on is ended by, and what the next run's reaper looks for.
const ENGINE_IMAGE: &str = "soffice.exe";
/// The engine behind it: the application itself runs as a child of that small launcher, and
/// it is the half that holds a document open (see `let_go` and `engine_child`).
const ENGINE_BIN_IMAGE: &str = "soffice.bin";

/// The document the kept engine holds: the name it is written under, beside the profile, and
/// the document itself — one paragraph of flat OpenDocument text, which is everything an
/// instance needs to be a running one. What it says does not matter and nothing reads it;
/// what matters is that the engine has a document open, which is what makes a launch beside
/// it hand the work over rather than start a second engine (see the module docs).
const HOLDER_NAME: &str = "holder.fodt";
const HOLDER_DOCUMENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">"#,
    r#"<office:body><office:text><text:p>rust-hover-preview</text:p></office:text></office:body>"#,
    r#"</office:document>"#,
);

/// How long a request waits for a kept engine to have its document open.
///
/// A request made before it does starts a second engine beside the first — measured, and the
/// one thing keeping an engine is not for — so the wait is what a conversion owes the engine
/// that is starting; what it costs is bounded by the engine start it replaces. An engine
/// whose process is gone is not waited for at all, so a broken installation costs this once
/// and never again.
const READY_WAIT: Duration = Duration::from_secs(3);
/// How often the wait above looks.
const READY_POLL: Duration = Duration::from_millis(50);
/// How often the engine thread wakes to look at the idle setting, which is what lets a kept
/// engine go while nothing is being asked of it. Nothing else wakes it between documents, and
/// one look a second is nothing beside the process it is looking at.
const IDLE_TICK: Duration = Duration::from_secs(1);

/// Where LibreOffice keeps its program, for the two places it installs and for a portable
/// copy a user may have put beside `config.ini`.
fn soffice() -> Option<PathBuf> {
    static FOUND: Lazy<Option<PathBuf>> = Lazy::new(|| {
        for candidate in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ] {
            let path = Path::new(candidate);
            if path.is_file() {
                return Some(path.to_path_buf());
            }
        }

        let portable = AppConfig::config_path()?
            .parent()?
            .join("libreoffice")
            .join("program")
            .join("soffice.exe");

        portable.is_file().then_some(portable)
    });

    FOUND.clone()
}

/// The folder the engine works in: the profile it runs under, the stub document a kept instance
/// holds open, and the folder a conversion is staged in.
///
/// It is this app's own folder under `%LOCALAPPDATA%`, where the other engine state lives —
/// the profiles a run of the browser is kept in, and the records of the processes an earlier run
/// started — rather than beside the pages: none of what is here is a page, and where pages are
/// kept is the cache's own business (see `document_cache`). Nothing of the user's stays in it
/// longer than a conversion takes: the stage folder is emptied before every conversion, and the
/// document a kept engine holds is a stub of this app's own.
fn workspace() -> Option<PathBuf> {
    Some(
        BaseDirs::new()?
            .data_local_dir()
            .join("rust-hover-preview")
            .join("libreoffice"),
    )
}

/// Whether an engine is installed to render these documents with.
pub fn available() -> bool {
    soffice().is_some()
}

/// The page already drawn for this version of the document, if there is one: a PDF the PDF path
/// reads the way it reads any other.
///
/// Nothing is started and nothing is waited on here. This is the question "has the engine
/// answered yet?" asked of the document cache — a preview is laid out from what it finds, and a
/// hover that is waiting for one keeps asking it until the answer is there (see [`request`]).
pub fn rendered_page(path: &Path) -> Option<PathBuf> {
    let page = document_cache::page(path, OfficeEngine::LibreOffice)?;

    usable(&page).then_some(page.path)
}

/// Whether the engine has already turned this version of the document down: the mark a
/// conversion that wrote no page leaves where a page would have been.
///
/// It is what a hover waiting on that document reads as "nothing is coming" — the spinner
/// comes down rather than running out its wait — and what keeps the launch from being paid for
/// a second time. A mark ages out, so a document that was undrawable for a while — locked, or
/// half-copied, or read while a filter was still being installed — is asked about again (see
/// `document_cache`).
pub fn refused(path: &Path) -> bool {
    document_cache::refused(path, OfficeEngine::LibreOffice)
}

/// Ask the engine for a page for `path`.
///
/// Nothing is waited on and nothing is answered: the conversion runs on the engine thread
/// below, and what a caller watches for is the page appearing in the folder it is kept in,
/// or the mark that says it is not coming. A hover that asked for one is replayed when it
/// is there.
pub fn request(path: &Path) {
    if !imports(path) || !available() {
        return;
    }
    if rendered_page(path).is_some() || refused(path) {
        return;
    }

    // An engine that has been inside one conversion for longer than any of them takes has
    // stopped answering, so it is ended here rather than queued behind: what that frees is
    // the seat this request needs, and the core the stuck one was holding.
    end_hung_engine();

    // A request for the document already being drawn is that request. One that arrives
    // while another waits replaces it, the way the loader's slot does: the newest hover is
    // the one the pointer is on, and a document whose hover has gone is one nobody is
    // waiting for.
    if running_source().as_deref() == Some(path) {
        return;
    }

    let (slot, ready) = &*REQUESTED;
    if let Ok(mut requested) = slot.lock() {
        *requested = Some(path.to_path_buf());
    }
    ready.notify_all();

    start_engine();
}

/// The document the engine is drawing now, if it is drawing one.
fn running_source() -> Option<PathBuf> {
    RUNNING
        .lock()
        .ok()?
        .as_ref()
        .map(|running| running.source.clone())
}

/// End a conversion that has outrun the engine's give-up.
///
/// Ending the process a conversion is waiting on is what ends the wait: the engine thread
/// reads it as a conversion that wrote no page, writes the mark that says so beside where
/// the page would have been, and releases the seat for the document behind it. What the
/// caller here is left with is a hung engine that costs nothing — no core, no seat — and
/// the answer it would have reached anyway.
///
/// The engine thread holds the same bound over the conversion it is inside, so this is the
/// same rule reached from the other side: a document asked for behind a hung engine ends it
/// at the moment it is asked for rather than waiting for the thread to notice.
fn end_hung_engine() {
    let hung = RUNNING.lock().ok().and_then(|running| {
        running
            .as_ref()
            .filter(|running| is_hung(running))
            .map(|running| running.pid)
    });

    if let Some(pid) = hung {
        // What a conversion that has outrun its bound is ended by is the engine rather than
        // the process that asked it for a page: a page handed to an engine that is being
        // kept is drawn inside it, so ending the launch alone would leave the engine
        // spinning on the same document with the seat still held. It is let go of here, and
        // the document behind it is answered by the engine that starts in its place.
        let_go();

        // Verified by name and start time before anything is ended, like every other
        // process this app holds a record of.
        crate::app::engine_processes::terminate_owned(pid);
    }
}

/// Whether a conversion in flight has had its chance: a document the engine has been drawing
/// for longer than any document takes is one it is not going to finish.
fn is_hung(running: &Running) -> bool {
    running.started.elapsed() >= CONVERSION_GIVE_UP
}

/// The engine this app is keeping, and when it last drew a page: the process holding it,
/// which is what it is let go by, and the moment the idle setting counts from.
struct Kept {
    pid: u32,
    idle_since: Instant,
}

static KEPT: Lazy<Mutex<Option<Kept>>> = Lazy::new(|| Mutex::new(None));

/// The profile the engine runs under: one of this app's own, beside the document it holds, so
/// that a LibreOffice the user has open is untouched by a conversion and its documents by this
/// app's engine. One string, made once and used by everything that runs the engine: a launch
/// only finds the instance to hand its work to when the profile it names is the one that
/// instance was started with, character for character.
fn profile_url(folder: &Path) -> String {
    let profile = folder.join("profile");
    format!("file:///{}", profile.to_string_lossy().replace('\\', "/"))
}

/// The document the kept engine holds, written when it is wanted: a stub of this app's own,
/// beside the profile, which is what makes the instance a running one rather than one that
/// yields its profile to the next launch (see the module docs).
fn holder_document() -> Option<PathBuf> {
    let folder = workspace()?;
    std::fs::create_dir_all(&folder).ok()?;

    // It is not written over an engine that has it open: what the engine holds is its
    // business while it holds it, and a file changed under a document is a document that
    // offers to reload itself.
    let holder = folder.join(HOLDER_NAME);
    if !holder.is_file() {
        write_holder(&holder)?;
    }

    Some(holder)
}

/// Write the stub the engine holds, over whatever is there.
fn write_holder(holder: &Path) -> Option<()> {
    std::fs::write(holder, HOLDER_DOCUMENT).ok()
}

/// The lock file LibreOffice keeps beside a document it has open, which is what says the
/// engine has hold of its own: `.~lock.<name>#`, written where the document is.
fn lock_file(document: &Path) -> PathBuf {
    let name = document
        .file_name()
        .map(|name| name.to_string_lossy().to_string())
        .unwrap_or_default();

    document.with_file_name(format!(".~lock.{name}#"))
}

/// How long past its last page the engine is kept, or `None` for a setting that keeps none
/// at all — the `0 seconds` of the tray's `LibreOffice TTL`, which is an engine per
/// document, exactly as every document was drawn before there was a setting.
///
/// It is the setting an engine marked `Persistent` is kept by and nothing else: one that is
/// not persistent is let go by the AFK timer instead, and this is not asked about it
/// (see `let_go_if_expired`).
fn kept_idle() -> Option<EngineIdle> {
    let idle = CONFIG
        .lock()
        .map(|config| config.libreoffice_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_LIBREOFFICE_IDLE_SECS));

    (idle != EngineIdle::Seconds(0)).then_some(idle)
}

/// Whether the engine is kept whatever the user is doing, which is the `Persistent` toggle
/// at the top of the same submenu.
///
/// Read rather than captured, for the reason the idle time is: the engine thread asks it
/// once a second while it waits for documents, so a click applies to an engine that is
/// already up.
fn engine_persistent() -> bool {
    CONFIG
        .lock()
        .map(|config| config.libreoffice_persistent)
        .unwrap_or(false)
}

/// The process holding the engine being kept, where there is one and it is still running.
///
/// An engine that is gone — ended by a user, crashed, or given up on — is not kept: what is
/// held is struck off and the next document starts another, so a machine that ends the
/// engine some other way costs a launch and never a preview.
fn kept_pid() -> Option<u32> {
    let pid = KEPT
        .lock()
        .ok()
        .and_then(|kept| kept.as_ref().map(|kept| kept.pid))?;

    if crate::app::engine_processes::is_running(pid) {
        return Some(pid);
    }

    let_go();
    None
}

/// Make sure an engine is being kept, and answer the process holding it.
///
/// It is called before a conversion, so that the page is asked of the engine being kept
/// rather than of a second one started beside it, and what it answers with is the process a
/// caller waits on before asking (see `wait_until_ready`). A setting that keeps no engine
/// answers with nothing, and the conversion is the launch it has always been.
fn keep_engine(program: &Path) -> Option<u32> {
    kept_idle()?;

    if let Some(pid) = kept_pid() {
        return Some(pid);
    }

    start_kept_engine(program)
}

/// Start the engine the setting keeps, holding this app's own stub document.
///
/// The process is recorded the way every engine this app starts is recorded: it is put in
/// the job that ends with the app, and written down so that a run which never got to end it
/// is answered for by the next one. It is also left running rather than waited on — what
/// holds it is the document it has open.
fn start_kept_engine(program: &Path) -> Option<u32> {
    let folder = workspace()?;
    let holder = holder_document()?;

    // One engine at a time is what one profile allows: an engine this app still holds — one
    // being let go of, or one a run never got to end — is ended and waited for here, so that
    // the instance about to hold the document is not starting beside another one. It is the
    // same rule the Office tier keeps a family to, and the same call.
    crate::app::engine_processes::end_recorded(ENGINE_BIN_IMAGE);
    crate::app::engine_processes::end_recorded(ENGINE_IMAGE);

    // A lock file left by an engine that was let go of is not this one's, and what a wait
    // reads is whether *this* engine has the document open: the mark goes before it starts,
    // so that it means something when it comes back.
    std::fs::remove_file(lock_file(&holder)).ok();

    let child = Command::new(program)
        .arg("--headless")
        .arg("--norestore")
        .arg(format!("-env:UserInstallation={}", profile_url(&folder)))
        .arg(&holder)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    crate::app::engine_processes::record(ENGINE_IMAGE, child.id());

    if let Ok(mut kept) = KEPT.lock() {
        *kept = Some(Kept {
            pid: child.id(),
            idle_since: Instant::now(),
        });
    }

    Some(child.id())
}

/// Whether the engine has its own document open yet, waiting for it while it starts.
///
/// A conversion asked for before then starts a second engine beside the first rather than
/// being handed to it, so the wait is what keeps the setting's promise; it is bounded, and
/// an engine whose process is already gone is not waited for at all. A wait that ends
/// either way answers with whether the engine is ready, which is what a caller decides
/// nothing on: the conversion is asked for regardless, and what differs is whether it is
/// the engine's or a fresh one's.
fn wait_until_ready(pid: u32) -> bool {
    let Some(holder) = holder_document() else {
        return false;
    };

    let ready = wait_until_ready_for(pid, &holder, READY_WAIT);
    // Whatever the wait answered, the engine behind the launcher is on the machine now or
    // soon will be, and what a run that never gets to end it leaves is that one.
    track_engine_child(pid);

    ready
}

fn wait_until_ready_for(pid: u32, holder: &Path, bound: Duration) -> bool {
    let lock = lock_file(holder);
    let deadline = Instant::now() + bound;

    loop {
        if lock.is_file() {
            return true;
        }
        if !crate::app::engine_processes::is_running(pid) || Instant::now() >= deadline {
            return false;
        }

        std::thread::sleep(READY_POLL);
    }
}

/// Say that the engine has just drawn a page, which is what the idle setting counts from.
fn touch_engine() {
    if let Ok(mut kept) = KEPT.lock() {
        if let Some(kept) = kept.as_mut() {
            kept.idle_since = Instant::now();
        }
    }
}

/// Let the engine go: end it, and stop calling it kept.
///
/// Both halves of the instance are ended: the launcher this app started, which is the process
/// that is recorded and that the job and the next run's reaper answer for, and the engine
/// behind it, which is what holds the document and what a conversion is handed to. Ending the
/// launcher alone was measured to leave an engine running with a document open and nothing
/// holding it, which is a few hundred megabytes kept for a setting that said otherwise.
///
/// The engine being let go of is one this app started — an instance the user opened is not
/// recorded here and is never one of these.
fn let_go() {
    let kept = KEPT.lock().ok().and_then(|mut kept| kept.take());

    if let Some(kept) = kept {
        // The engine behind the launcher goes first, because it is the half that holds the
        // document and the half the setting is about; the launcher is what carries the record
        // and what is left of the instance once its engine is gone.
        if let Some(engine) = engine_child(kept.pid) {
            crate::app::engine_processes::terminate_verified(engine, ENGINE_BIN_IMAGE, None);
        }

        crate::app::engine_processes::terminate_owned(kept.pid);
    }
}

/// The engine behind the launcher this app started, where it is there: LibreOffice runs the
/// application itself as a child of the small process that is started, and the child is what
/// holds the document — the instance a conversion is handed to.
fn engine_child(launcher: u32) -> Option<u32> {
    crate::app::engine_processes::processes_named_by_parent(ENGINE_BIN_IMAGE, launcher)
        .first()
        .copied()
}

/// Take the engine behind the launcher into the same record the launcher is in.
///
/// The record is what answers for a process a run never got to end — a crash, a kill — and
/// the launcher alone does not answer for the engine: what holds the document is the child,
/// and a run that is gone leaves that one running until it is reached. It is written down
/// beside its launcher as soon as it is there, which is what the wait below is watching for.
fn track_engine_child(launcher: u32) {
    if let Some(pid) = engine_child(launcher) {
        crate::app::engine_processes::record(ENGINE_BIN_IMAGE, pid);
    }
}

/// Let the engine go once it is time, which is one of two questions rather than one.
///
/// It is the engine thread that looks, once a second while it waits for documents: nothing
/// else runs between hovers, and an engine that is never let go of is a process left on the
/// machine for a setting that said otherwise.
///
/// Which question is asked is the `Persistent` toggle at the top of the same submenu. Not
/// persistent — the way this starts — the engine is let go once no Explorer window has been
/// reachable for the AFK timer, and the idle time is not consulted at all: what is being
/// bounded is the few hundred megabytes an engine holds while the user is somewhere else.
/// Marked persistent it is the idle time that is asked, exactly as it was before the toggle
/// existed, and a setting of `0 seconds` — the switch the tray offers as the bottom of the
/// list — lets go of an engine that is already running the moment it is looked at.
fn let_go_if_expired() {
    if !engine_persistent() {
        if crate::app::afk::expired() {
            let_go();
        }
        return;
    }

    let Some(idle) = kept_idle() else {
        let_go();
        return;
    };

    let expired = KEPT
        .lock()
        .ok()
        .and_then(|kept| {
            kept.as_ref()
                .map(|kept| idle.has_expired(kept.idle_since.elapsed()))
        })
        .unwrap_or(false);

    if expired {
        let_go();
    }
}

/// Start the thread conversions run on, once.
///
/// It is one of the app's threads rather than one per document: what it does between
/// conversions is wait on its own slot, which costs nothing, and what it is asked for is
/// one document at a time because the engine's profile is one seat.
fn start_engine() {
    if ENGINE_STARTED.swap(true, Ordering::AcqRel) {
        return;
    }

    std::thread::spawn(|| {
        while let Some(source) = next_request() {
            // Whatever the engine answers — a page, or a conversion that wrote none — it is
            // answered in the folder the pages are kept in, which is what the hover waiting
            // on this document is watching. A panic is contained here for the same reason
            // the loader contains one: one document's failure is that document's, and the
            // thread goes on to the next hover — a thread that died on one document would
            // take every document after it with it, in silence.
            let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| rendered(&source)));
        }
    });
}

/// The next document to draw, waiting for one.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made; a lock
/// poisoned by a panic on another thread is still the same slot, and a conversion queue is
/// not worth standing down over.
fn next_request() -> Option<PathBuf> {
    let (slot, ready) = &*REQUESTED;
    let mut requested = match slot.lock() {
        Ok(requested) => requested,
        Err(poisoned) => poisoned.into_inner(),
    };

    loop {
        if let Some(source) = requested.take() {
            return Some(source);
        }

        // The wait is bounded rather than endless, and what a bound that runs out is for is
        // the engine: one that is kept is let go of by this thread and no other, and the
        // idle setting is what says when (see `let_go_if_expired`).
        requested = match ready.wait_timeout(requested, IDLE_TICK) {
            Ok((requested, _)) => requested,
            Err(poisoned) => poisoned.into_inner().0,
        };

        let_go_if_expired();
    }
}

/// Whether the engine is the one that draws this file: a document of its own lists, named
/// by the configured list or recognized by its own bytes, or an Office document whose own
/// application is not installed. The question is asked where it is answered for every
/// caller — see `libre_formats::engine_page_kind` — so that a page one side asks for is a
/// page the other side will draw.
fn imports(path: &Path) -> bool {
    crate::formats::libre_formats::engine_page_kind(path).is_some()
}

/// Convert a document whose page is not kept yet, and keep what the engine drew.
///
/// The answer is the file the page is kept as; a document the engine would not draw is answered
/// with nothing, and that is written down so the launch is not paid for twice.
fn rendered(path: &Path) -> Option<PathBuf> {
    let program = soffice()?;
    if let Some(page) = kept_page(path) {
        return Some(page);
    }
    if document_cache::refused(path, OfficeEngine::LibreOffice) {
        return None;
    }

    // One engine at a time, and the seat is held across the conversion: LibreOffice's
    // profile is a single seat, so a second run beside the first would wait on it anyway.
    // A conversion in flight has had its chance by the give-up, so the seat is waited for
    // with the conversion ended rather than for as long as the bound below allows.
    end_hung_engine();

    let _seat = CONVERTING.lock().ok()?;
    if let Some(page) = kept_page(path) {
        return Some(page);
    }
    if document_cache::refused(path, OfficeEngine::LibreOffice) {
        return None;
    }

    let Some(page) = convert(&program, path) else {
        // An engine that would not draw this document is not asked again for a while: what it
        // answered is written down where the page it did not write would have been.
        document_cache::refuse(path, OfficeEngine::LibreOffice);
        return None;
    };

    Some(page.path)
}

/// The page kept for this document, where the engine has drawn one for this version of it and
/// it can be read.
///
/// The header is the check rather than a re-conversion: what is read back is a file this app
/// wrote, and a file that is not a PDF is not a page.
fn kept_page(path: &Path) -> Option<PathBuf> {
    let page = document_cache::page(path, OfficeEngine::LibreOffice)?;

    usable(&page).then_some(page.path)
}

/// Whether a kept page is there and is a PDF.
fn usable(page: &Page) -> bool {
    let Ok(mut file) = std::fs::File::open(&page.path) else {
        return false;
    };
    let mut header = [0u8; 5];
    std::io::Read::read_exact(&mut file, &mut header).is_ok() && &header == b"%PDF-"
}

/// Convert `source` into a page, by running the engine the way a user would: headless, with a
/// profile of this app's own so that a LibreOffice the user has open is untouched, and with a
/// wait that ends rather than holding a hover for good.
fn convert(program: &Path, source: &Path) -> Option<Page> {
    let folder = workspace()?;
    let stage = folder.join("stage");
    // Whatever a run before this one left behind is not read: the engine names what it
    // writes after what it was given, and only the file of this conversion is looked for.
    std::fs::remove_dir_all(&stage).ok();
    std::fs::create_dir_all(&stage).ok()?;

    let profile_url = profile_url(&folder);

    // The engine the setting keeps holds a document of this app's own, and a page asked for
    // beside it is drawn by it rather than by an engine started for this one document. A
    // request made before it has that document open starts a second engine beside the first
    // — the one thing keeping an engine is not for — so what the wait is for is the engine
    // that is starting. The page is asked for either way: what differs is whose engine draws
    // it.
    if let Some(pid) = keep_engine(program) {
        wait_until_ready(pid);
    }

    let mut child = Command::new(program)
        .arg("--headless")
        .arg("--norestore")
        .arg(format!("-env:UserInstallation={profile_url}"))
        .arg("--convert-to")
        .arg("pdf:draw_pdf_Export")
        .arg("--outdir")
        .arg(&stage)
        .arg(source)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    // A process this app started is put in the job every other engine is put in, and held
    // in a record beside it: one that outlives the app does not outlive it by much, one
    // that is given up on is ended by name and id wherever it is asked about from, and one
    // a run never got to end is answered for by the next run's reaper.
    crate::app::engine_processes::record(ENGINE_IMAGE, child.id());

    // What the engine is drawing, published for the threads that may decide it has stopped
    // answering while this one waits (see `end_hung_engine`).
    publish_running(Some(Running {
        source: source.to_path_buf(),
        pid: child.id(),
        started: Instant::now(),
    }));

    // The conversion's own bound, which is the same give-up every other thread reads: a
    // document still being drawn and an engine that has stopped drawing look the same from
    // the outside, and time is what tells them apart.
    let drawn = wait(&mut child, CONVERSION_GIVE_UP);

    publish_running(None);
    crate::app::engine_processes::forget(child.id());

    // Whatever the engine made of the document, it has drawn what it was going to draw: the
    // idle time the setting names counts from here.
    touch_engine();

    if !drawn {
        return None;
    }

    // The engine names what it wrote after what it was given, in the folder it was told —
    // and the name is the document's own, dots and all, which is why it is appended to
    // rather than replaced: a file called `drawing.v2.cdr` is written as
    // `drawing.v2.pdf`, not as `drawing.pdf`.
    let stem = source.file_stem()?;
    let written = stage.join(format!("{}.pdf", stem.to_string_lossy()));
    let page = std::fs::read(&written).ok()?;
    if !page.starts_with(b"%PDF-") {
        return None;
    }

    std::fs::remove_file(&written).ok();

    document_cache::store(source, OfficeEngine::LibreOffice, PageKind::Pdf, &page)
}

/// Wait for a process, ending it rather than waiting past `limit`.
fn wait(child: &mut Child, limit: Duration) -> bool {
    let deadline = Instant::now() + limit;
    loop {
        match child.try_wait() {
            Ok(Some(_)) => return true,
            Ok(None) if Instant::now() < deadline => std::thread::sleep(CONVERSION_POLL),
            _ => {
                let _ = child.kill();
                let _ = child.wait();
                return false;
            }
        }
    }
}

/// Say what the engine is drawing now, or that it has stopped.
fn publish_running(running: Option<Running>) {
    if let Ok(mut published) = RUNNING.lock() {
        *published = running;
    }
}

/// The one conversion running at a time, for the reason the engine has one profile.
static CONVERTING: Lazy<Mutex<()>> = Lazy::new(|| Mutex::new(()));

/// The conversion in flight: the document being drawn, the process drawing it, and since
/// when. It is what tells a busy engine from one that has stopped answering, and it is the
/// id a conversion that has to be ended is ended by.
struct Running {
    source: PathBuf,
    pid: u32,
    started: Instant,
}

static RUNNING: Lazy<Mutex<Option<Running>>> = Lazy::new(|| Mutex::new(None));

/// The document waiting to be drawn, and the signal that one is there: a queue of one, for
/// the reason there is one seat.
static REQUESTED: Lazy<(Mutex<Option<PathBuf>>, Condvar)> =
    Lazy::new(|| (Mutex::new(None), Condvar::new()));

/// Whether the engine thread has been started. It is one of the app's threads rather than
/// one per document, so it is started once and waits on its slot for the rest of the run.
static ENGINE_STARTED: AtomicBool = AtomicBool::new(false);

#[cfg(test)]
mod tests {
    use super::*;

    /// The engine is never started for a name its own filters do not read. A Photoshop
    /// document, a Krita project, a Sketch file and a Procreate document are this app's
    /// own readers' business, and asking an office suite about one would cost a launch and
    /// answer nothing. A Flash animation is the case that made the rule worth stating: the
    /// engine spins on one rather than failing, so a name like that must not reach it at
    /// all (see `libre_formats`).
    #[test]
    fn asks_the_engine_only_about_the_names_it_reads() {
        for name in [
            "poster.psd",
            "painting.kra",
            "design.sketch",
            "art.procreate",
            "icon.svg",
            "animation.swf",
        ] {
            assert!(
                !imports(Path::new(name)),
                "`{name}` is not one of its formats"
            );
        }
    }

    /// And a name it does not read is not queued either: nothing about a document is asked
    /// of the engine that its list has not claimed.
    #[test]
    fn queues_nothing_for_a_name_it_does_not_read() {
        let (slot, _) = &*REQUESTED;
        if let Ok(mut requested) = slot.lock() {
            *requested = None;
        }

        request(Path::new("animation.swf"));

        assert!(
            slot.lock()
                .map(|requested| requested.is_none())
                .unwrap_or(false),
            "the engine is not asked about a name no list of its own holds"
        );
    }

    /// And the give-up is not a decision on its own: what it names is ended, which is what
    /// frees the seat and the core the engine was holding and lets the document behind it be
    /// drawn. A process that stays up stands in for the engine — a test is not going to make
    /// LibreOffice spin on a file — recorded the way the engine is, by image name, which is
    /// the check that keeps an id from being acted on by itself.
    #[test]
    fn ends_the_engine_only_once_its_conversion_has_outrun_the_give_up() {
        let mut engine = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for the engine");
        let pid = engine.id();
        let running = |started: Instant| Running {
            source: PathBuf::from("drawing.cdr"),
            pid,
            started,
        };

        assert!(
            crate::app::engine_processes::processes_named("ping.exe").contains(&pid),
            "the stand-in runs the image the record below names it by"
        );
        crate::app::engine_processes::record("ping.exe", pid);

        // A conversion that has just started is a document being drawn, and is left to it.
        publish_running(Some(running(Instant::now())));
        end_hung_engine();
        assert!(
            crate::app::engine_processes::is_running(pid),
            "a conversion that has just started is not an engine to end"
        );

        // One that has outrun the give-up is an engine that has stopped answering.
        publish_running(Some(running(Instant::now() - CONVERSION_GIVE_UP)));
        end_hung_engine();
        assert!(
            !crate::app::engine_processes::is_running(pid),
            "the engine a conversion has outrun is ended"
        );

        publish_running(None);
        let _ = engine.wait();
    }

    /// And the bound the engine thread holds over its own conversion: a process that ends
    /// inside it is answered as it always was, and one that outlasts it is ended there
    /// rather than waited on — which is what makes the give-up a bound on the engine rather
    /// than only on a document that happens to be asked for after it.
    #[test]
    fn ends_a_conversion_rather_than_waiting_past_its_bound() {
        let mut quick = std::process::Command::new("ping")
            .args(["-n", "1", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process that ends on its own");
        assert!(
            wait(&mut quick, CONVERSION_GIVE_UP),
            "a conversion that finishes inside its bound is answered as it always was"
        );

        let mut engine = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for the engine");
        let pid = engine.id();
        let started = Instant::now();

        assert!(
            !wait(&mut engine, Duration::from_millis(300)),
            "a conversion that has outrun its bound is ended, not waited on"
        );
        assert!(
            started.elapsed() < Duration::from_secs(5),
            "and the wait ends at the bound it was given rather than when the process would have"
        );
        assert!(
            !crate::app::engine_processes::is_running(pid),
            "the engine is gone with it"
        );
    }

    /// What the give-up is decided from: a conversion that has just started is not hung —
    /// a document the engine can draw is seconds, and the profile's first one is a few of
    /// them — and one that has run past the give-up is.
    #[test]
    fn a_conversion_is_hung_only_once_it_has_outrun_the_give_up() {
        let running = |elapsed: Duration| Running {
            source: PathBuf::from("drawing.cdr"),
            pid: std::process::id(),
            started: Instant::now() - elapsed,
        };

        assert!(!is_hung(&running(Duration::from_secs(0))));
        assert!(!is_hung(&running(
            CONVERSION_GIVE_UP - Duration::from_secs(1)
        )));
        assert!(is_hung(&running(CONVERSION_GIVE_UP)));
        assert!(is_hung(&running(
            CONVERSION_GIVE_UP + Duration::from_secs(30)
        )));
    }

    /// And the names it does read are the CorelDRAW family and the formats of the same
    /// libraries, whatever case they are written in.
    #[test]
    fn reads_coreldraw_and_the_formats_beside_it() {
        for name in [
            "logo.cdr",
            "drawing.CDR",
            "artwork.cmx",
            "poster.pub",
            "plan.vsd",
        ] {
            assert!(imports(Path::new(name)), "`{name}` is one of its formats");
        }
    }

    /// What a kept engine is waited for is the document it holds being *open* — the lock file
    /// LibreOffice writes beside it — and a wait answers as soon as that is there.
    #[test]
    fn waits_for_a_kept_engine_until_its_document_is_open() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-libre-tests")
            .join("ready");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let holder = folder.join(HOLDER_NAME);
        write_holder(&holder).expect("a written stub");
        assert_eq!(
            std::fs::read_to_string(&holder).expect("the stub"),
            HOLDER_DOCUMENT,
            "and what it holds is the document the engine is given"
        );

        // The lock file is what says the engine has its document open: with one there the
        // wait is answered at once rather than at its bound. The wait is on this process,
        // which is running, so what is measured is the lock file and nothing else.
        std::fs::write(lock_file(&holder), b"").expect("a written lock file");
        let started = Instant::now();
        assert!(wait_until_ready_for(
            std::process::id(),
            &holder,
            Duration::from_secs(30)
        ));
        assert!(
            started.elapsed() < Duration::from_secs(1),
            "the wait ends when the engine is ready rather than at the bound"
        );

        // Without it, the wait is the bound and no more — an engine that has not opened its
        // document yet is waited for, and one that never does is not waited for forever.
        std::fs::remove_file(lock_file(&holder)).expect("the lock file removed");
        let started = Instant::now();
        assert!(!wait_until_ready_for(
            std::process::id(),
            &holder,
            Duration::from_millis(200)
        ));
        assert!(started.elapsed() >= Duration::from_millis(200));

        // And a process that is gone is not waited for at all: what that costs is a launch
        // for the document at hand, which is what a machine without an engine pays anyway.
        let mut gone = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for an engine that is gone");
        let pid = gone.id();
        let _ = gone.kill();
        let _ = gone.wait();

        let started = Instant::now();
        assert!(!wait_until_ready_for(pid, &holder, Duration::from_secs(30)));
        assert!(
            started.elapsed() < Duration::from_secs(2),
            "an engine whose process is gone is not waited for, and the bound is not paid"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// The idle setting is what says whether an engine is kept at all, and `0 seconds` — the
    /// bottom of the tray's list — is the setting that keeps none: every document is the
    /// launch it has always been. It is read through the app's own configuration, so what is
    /// asserted here is the shape of the answer rather than the value a machine holds.
    #[test]
    fn a_setting_of_no_seconds_keeps_no_engine() {
        let keep = kept_idle().is_some();
        let configured = CONFIG
            .lock()
            .map(|config| config.libreoffice_idle)
            .expect("the configuration");

        assert_eq!(
            keep,
            configured != EngineIdle::Seconds(0),
            "an engine is kept exactly where the idle setting is not `0 seconds`"
        );
    }

    /// What the idle setting buys, measured: the same document converted with nothing kept,
    /// and then again with the engine the first one started. Ignored because it starts the
    /// installed LibreOffice, and run when that setting is being looked at:
    /// `cargo test -- --ignored --nocapture engine_warmth_probe`.
    #[test]
    #[ignore = "starts the installed LibreOffice"]
    fn engine_warmth_probe() {
        let Some(program) = soffice() else {
            println!("no LibreOffice installed: nothing to measure");
            return;
        };
        let Some(holder) = holder_document() else {
            println!("no folder for the stub document");
            return;
        };
        let Some(folder) = workspace() else {
            println!("no folder for the engine to work in");
            return;
        };
        std::fs::create_dir_all(&folder).ok();

        let source = folder.join("probe-source.fodt");
        std::fs::write(&source, HOLDER_DOCUMENT).ok();

        // Nothing is kept to begin with, so the first row is the launch every document paid
        // for on its own before there was a setting.
        document_cache::forget(&source, OfficeEngine::LibreOffice);
        let_go();
        let started = Instant::now();
        let first = convert(&program, &source);
        println!(
            "first document: {} in {:?} — the engine started and handed the page",
            first.is_some(),
            started.elapsed()
        );

        // And the same document again, with the engine the first one started.
        document_cache::forget(&source, OfficeEngine::LibreOffice);
        let started = Instant::now();
        let next = convert(&program, &source);
        println!(
            "next document: {} in {:?} — the engine kept, the page handed to it",
            next.is_some(),
            started.elapsed()
        );

        let launcher = kept_pid();
        let child = launcher.and_then(engine_child);
        println!("kept launcher {launcher:?}, engine behind it {child:?}");

        let_go();

        // A process that has been terminated is still in the process table for a moment, so
        // what is reported is the check that matters: whether it is still running.
        let gone = |pid: u32| !crate::app::engine_processes::is_running(pid);
        println!(
            "left after letting go: launcher gone {}, engine gone {}",
            launcher.map(gone).unwrap_or(true),
            child.map(gone).unwrap_or(true),
        );

        // And the bottom row of the setting: `0 seconds` keeps no engine, so one that is
        // running when it is chosen is let go of at the next look — which is this call, made
        // from the engine thread once a second while it waits for documents.
        let Some(pid) = keep_engine(&program) else {
            println!("no engine started for the last check");
            return;
        };
        wait_until_ready(pid);
        if let Ok(mut config) = CONFIG.lock() {
            config.libreoffice_idle = EngineIdle::Seconds(0);
        }
        let_go_if_expired();
        println!("0 seconds: engine gone {}", gone(pid));
        if let Ok(mut config) = CONFIG.lock() {
            config.libreoffice_idle = EngineIdle::Seconds(DEFAULT_LIBREOFFICE_IDLE_SECS);
        }

        document_cache::forget(&source, OfficeEngine::LibreOffice);
        std::fs::remove_file(&source).ok();
        std::fs::remove_file(lock_file(&holder)).ok();
    }
}

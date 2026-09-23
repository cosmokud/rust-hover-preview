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
//! seconds, so a conversion happens once per document: the PDF it wrote is kept under
//! [`AppConfig::rendered_dir`], named for the document's path and the version of it that
//! was converted, and every hover after the first is a read of that file. One conversion
//! runs at a time, because one LibreOffice at a time is what its own profile allows.
//!
//! The engine is asked about the names its filters read and no others — the formats of
//! the libraries above, CorelDRAW's and the rest — so a Photoshop document, a Krita
//! project or a Procreate file never pays for a launch that could not answer.
//!
//! A conversion is seconds, and it is not one a preview can wait on: the caller is the
//! preview loop, and a loop held inside a launch is a hover that does not come up, a tray
//! that does not answer and a pointer that cannot leave the file it is on. So the engine
//! runs on a thread of its own. What the loop asks is [`request`], which returns at once,
//! and what it waits on is the page appearing under [`rendered_page`] — the same wait an
//! Office document has, in the same box, with the hover replayed when the page lands.
//! The engine draws one document at a time, because one profile is one seat, and the
//! document waiting behind it is the newest one asked for.
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
//! is answered rather than left spinning.

use crate::config::AppConfig;
use once_cell::sync::Lazy;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
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
/// What a name is remembered as when the engine would not draw it: the conversion is not
/// tried again for that version of the document, because a name this app was wrong about —
/// one in the list the engine has no filter for — would otherwise start an engine on every
/// hover to reach the same answer.
const REFUSED_SUFFIX: &str = "none";

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

        let portable = AppConfig::rendered_dir()?
            .parent()?
            .join("libreoffice")
            .join("program")
            .join("soffice.exe");

        portable.is_file().then_some(portable)
    });

    FOUND.clone()
}

/// Whether an engine is installed to render these documents with.
pub fn available() -> bool {
    soffice().is_some()
}

/// The page already drawn for this version of the document, if there is one: a PDF under
/// the app's own folder, which the PDF path reads the way it reads any other.
///
/// Nothing is started and nothing is waited on here. This is the question "has the engine
/// answered yet?" asked of the folder the answers are kept in — a preview is laid out from
/// what it finds, and a hover that is waiting for one keeps asking it until the answer is
/// there (see [`request`]).
pub fn rendered_page(path: &Path) -> Option<PathBuf> {
    let page = rendered_path(path)?;

    usable(&page).then_some(page)
}

/// Whether the engine has already turned this version of the document down: the mark a
/// conversion that wrote no page leaves where a page would have been.
///
/// It is what a hover waiting on that document reads as "nothing is coming" — the spinner
/// comes down rather than running out its wait — and what keeps the launch from being paid
/// for a second time.
pub fn refused(path: &Path) -> bool {
    rendered_path(path).is_some_and(|page| refused_marker(&page).is_some())
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

/// The same as [`rendered_page`] for a document the `[libre]` list does not hold — one of
/// the Office kind, asked for here only where the Office engine is not installed. The
/// caller decides that: what a name means is the lists' business, and this is the engine
/// that draws whatever it is given.
///
/// The conversion is this call's own, which is the one place a page is still drawn on a
/// caller's thread: the Office fallback is asked for by the loader, whose wait is a hover's
/// wait like any other, and whose page is measured from the file the moment it exists.
pub fn pdf_for_office(path: &Path) -> Option<PathBuf> {
    rendered(path)
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
        // Verified by name and start time before anything is ended, like every other
        // process this app holds a record of.
        crate::engine_processes::terminate_owned(pid);
    }
}

/// Whether a conversion in flight has had its chance: a document the engine has been drawing
/// for longer than any document takes is one it is not going to finish.
fn is_hung(running: &Running) -> bool {
    running.started.elapsed() >= CONVERSION_GIVE_UP
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

        requested = match ready.wait(requested) {
            Ok(requested) => requested,
            Err(poisoned) => poisoned.into_inner(),
        };
    }
}

/// Whether the list this app keeps says the engine is the one to draw `path`.
fn imports(path: &Path) -> bool {
    crate::libre_formats::is_libre_file(path)
}

/// The rendered page of a document, converting it if it has not been converted before.
fn rendered(path: &Path) -> Option<PathBuf> {
    let program = soffice()?;
    let page = rendered_path(path)?;
    if usable(&page) {
        return Some(page);
    }
    if refused_marker(&page).is_some() {
        return None;
    }

    // One engine at a time, and the seat is held across the conversion: LibreOffice's
    // profile is a single seat, so a second run beside the first would wait on it anyway.
    // A conversion in flight has had its chance by the give-up, so the seat is waited for
    // with the conversion ended rather than for as long as the bound below allows.
    end_hung_engine();

    let _seat = CONVERTING.lock().ok()?;
    if usable(&page) {
        return Some(page);
    }
    if refused_marker(&page).is_some() {
        return None;
    }

    if convert(&program, path, &page).is_none() {
        // An engine that would not draw this document is not asked again: what it answered
        // is written down beside the page it did not write.
        std::fs::write(refused_path(&page), b"").ok();
        return None;
    }

    Some(page)
}

/// Where the rendered page of `path` is kept: named for the document — its path, the
/// version of it that was converted, and nothing else — so a document saved again is
/// rendered again and a document that has not been is a read.
fn rendered_path(path: &Path) -> Option<PathBuf> {
    let folder = AppConfig::rendered_dir()?;
    std::fs::create_dir_all(&folder).ok()?;

    let mut hasher = DefaultHasher::new();
    path.to_string_lossy().to_lowercase().hash(&mut hasher);
    let metadata = std::fs::metadata(path).ok();
    metadata.as_ref().map(|metadata| metadata.len()).hash(&mut hasher);
    metadata
        .and_then(|metadata| metadata.modified().ok())
        .and_then(|modified| modified.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|since| since.as_secs())
        .hash(&mut hasher);

    Some(folder.join(format!("{:016x}.pdf", hasher.finish())))
}

/// Whether a rendered page is there and is a PDF: what is read back is a file this app
/// wrote, so the header is the check and not a re-render.
fn usable(rendered: &Path) -> bool {
    let Ok(mut file) = std::fs::File::open(rendered) else {
        return false;
    };
    let mut header = [0u8; 5];
    std::io::Read::read_exact(&mut file, &mut header).is_ok() && &header == b"%PDF-"
}

/// The mark left beside a page that was not written, for the same document and the same
/// version of it, or nothing when the engine has not been asked about it yet.
fn refused_marker(page: &Path) -> Option<PathBuf> {
    let refused = refused_path(page);

    refused.is_file().then_some(refused)
}

fn refused_path(page: &Path) -> PathBuf {
    page.with_extension(REFUSED_SUFFIX)
}

/// Convert `source` into `rendered`, by running the engine the way a user would: headless,
/// with a profile of this app's own so that a LibreOffice the user has open is untouched,
/// and with a wait that ends rather than holding a hover for good.
fn convert(program: &Path, source: &Path, rendered: &Path) -> Option<()> {
    let folder = rendered.parent()?;
    let stage = folder.join("stage");
    // Whatever a run before this one left behind is not read: the engine names what it
    // writes after what it was given, and only the file of this conversion is looked for.
    std::fs::remove_dir_all(&stage).ok();
    std::fs::create_dir_all(&stage).ok()?;

    let profile = folder.join("profile");
    let profile_url = format!("file:///{}", profile.to_string_lossy().replace('\\', "/"));

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
    crate::engine_processes::record(ENGINE_IMAGE, child.id());

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
    crate::engine_processes::forget(child.id());

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

    std::fs::write(rendered, &page).ok()?;
    std::fs::remove_file(&written).ok();
    prune(folder);

    Some(())
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

/// Keep the folder within the size the `Performance → Cache → Libre` setting names, oldest
/// first: a rendered page is built again from its document whenever it is wanted, so nothing
/// here is worth growing a folder for. A budget of nothing drops every page, which is what
/// the setting means — nothing is kept between hovers.
fn prune(folder: &Path) {
    let budget = crate::CONFIG
        .lock()
        .map(|config| config.libre_cache_mb as u64 * 1024 * 1024)
        .unwrap_or(0);
    let Ok(entries) = std::fs::read_dir(folder) else {
        return;
    };

    let mut pages: Vec<(std::time::SystemTime, u64, PathBuf)> = entries
        .flatten()
        .filter(|entry| {
            entry
                .path()
                .extension()
                .is_some_and(|extension| extension == "pdf" || extension == REFUSED_SUFFIX)
        })
        .filter_map(|entry| {
            let metadata = entry.metadata().ok()?;
            Some((metadata.modified().ok()?, metadata.len(), entry.path()))
        })
        .collect();

    let mut total: u64 = pages.iter().map(|(_, size, _)| size).sum();
    if total <= budget {
        return;
    }

    pages.sort_by_key(|(modified, _, _)| *modified);
    for (_, size, path) in pages {
        if total <= budget {
            break;
        }
        if std::fs::remove_file(&path).is_ok() {
            total = total.saturating_sub(size);
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
            assert!(!imports(Path::new(name)), "`{name}` is not one of its formats");
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
            crate::engine_processes::processes_named("ping.exe").contains(&pid),
            "the stand-in runs the image the record below names it by"
        );
        crate::engine_processes::record("ping.exe", pid);

        // A conversion that has just started is a document being drawn, and is left to it.
        publish_running(Some(running(Instant::now())));
        end_hung_engine();
        assert!(
            crate::engine_processes::is_running(pid),
            "a conversion that has just started is not an engine to end"
        );

        // One that has outrun the give-up is an engine that has stopped answering.
        publish_running(Some(running(Instant::now() - CONVERSION_GIVE_UP)));
        end_hung_engine();
        assert!(
            !crate::engine_processes::is_running(pid),
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
            !crate::engine_processes::is_running(pid),
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
        for name in ["logo.cdr", "drawing.CDR", "artwork.cmx", "poster.pub", "plan.vsd"] {
            assert!(imports(Path::new(name)), "`{name}` is one of its formats");
        }
    }
}

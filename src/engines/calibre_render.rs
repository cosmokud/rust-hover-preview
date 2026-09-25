//! Ebook pages produced by an installed Calibre.
//!
//! This app draws a PDF from its own reader and nothing else as a book. Everything a person
//! actually buys, borrows or scans is a format that reader does not open — the Kindle and
//! Mobipocket families, the open EPUB, the FictionBook, the Sony container, the Palm databases the
//! first ebooks arrived as, and the scanned book — and a hover onto one of them shows nothing at
//! all today. Calibre is the tool that opens them, and what it hands back is a PDF of what it read,
//! which is the one thing this app already knows how to draw.
//!
//! What is asked of it is the engine's own converter, `ebook-convert`, run the way a user runs it:
//! a book in, a PDF out, nothing else on the command line. That PDF is what the preview is made
//! from — the page at the `Ebook` kind's scale, over the `Ebook` kind's backdrop, under the
//! `Ebook` kind's switch — and it is kept as a page in the document cache, named for the book, the
//! version of it and this engine, so every hover after the first is a read of that file rather than
//! another conversion (see `document_cache`).
//!
//! The engine is asked about the names its own input plugins declare and no others — see
//! `calibre_formats`, which is the list and the reasons each name is or is not in it — so a picture
//! or a document never pays for a conversion that could not answer. A name the engine cannot read
//! is answered with no preview once and then remembered, so a name put in the list by mistake
//! costs one conversion and never another.
//!
//! A conversion is a whole book rather than a page, and it is seconds at the least: the engine is a
//! Python program that boots itself and its whole framework before it reads a byte, and what it
//! then does is walk every page of the book and write them out. So it is not something a preview
//! can wait on — the caller is the preview loop, and a loop held inside a launch is a hover that
//! does not come up, a tray that does not answer and a pointer that cannot leave the file it is on.
//! The conversion therefore runs on a thread of its own. What the loop asks is [`request`], which
//! returns at once, and what it waits for is the page appearing under [`rendered_page`] — the same
//! wait a document an installed render engine draws has, in the same box, with the hover replayed
//! when the page lands. One conversion runs at a time, and the book waiting behind it is the newest
//! one asked for.
//!
//! An engine that has stopped answering is the other half of that, and here it is the ordinary
//! case rather than the exception: a book the engine cannot read — one whose DRM it will not open,
//! one whose scan it cannot decode — can take minutes to say so, and some do not say so at all. A
//! conversion that has outrun [`CONVERSION_GIVE_UP`] is therefore ended where it stands, and the
//! book it was on is remembered as one the engine will not convert. What is bounded is also the
//! *first* hover of a large book: a hover stops waiting after the app's own cap whatever the engine
//! is doing, and the conversion runs on, so the book is drawn for the next hover rather than for
//! the one that asked.
//!
//! **What this module is not is a process the app keeps, and that is why there is no `Calibre TTL`
//! row in the `Engine` submenu.** It is the shape `imagemagick_render` and `peazip_render` have
//! rather than the shape `office_render`, `libreoffice_render` and the browser have, and the reason
//! is the engine's own: `ebook-convert` is a converter — it is handed a book, writes the book back
//! out as a PDF and exits — so there is nothing to hold open between books and nothing for an idle
//! time to bound. The three engines that have a TTL are the ones that are *applications*: an Office
//! instance driven through automation, a LibreOffice kept running with a document of this app's own
//! open beside it, and a browser whose whole point is to be pointed at the next page. None of that
//! exists here. What a user who wants their books to open faster changes is not a TTL but the page
//! cache: a second hover of a book whose page is still held costs no conversion at all — which is
//! the same answer `peazip_render` gives, and the same reason.

use crate::config::config::{read_within_budget, AppConfig};
use crate::engines::document_cache::{self, PageKind};
use once_cell::sync::Lazy;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Condvar, Mutex};
use std::time::{Duration, Instant};

/// The name this engine's pages are kept under, which is what tells one of its pages from the page
/// an Office application or an installed render engine drew for the same file (see
/// `document_cache::key`).
pub const ENGINE_NAME: &str = "calibre";

/// The program inside an installation that reads a book: the engine's converter, which is the one
/// entry point every one of these formats is opened through.
const ENGINE_IMAGE: &str = "ebook-convert.exe";

/// How long a conversion is given, and the point past which the engine is a stopped one rather
/// than a busy one: the conversion is ended where it stands, and the book it was on is remembered
/// as one the engine will not convert.
///
/// Twice the render engine's bound, and for a reason that is the format's rather than the
/// engine's: what a document is converted *from* is one document, while what a book is converted
/// from is every page of a book — a novel of five hundred pages, a scanned one of a thousand, a
/// comic of a hundred plates. An ordinary book is a few seconds, most of which is the engine
/// booting itself; the bound is an order of magnitude past that, and a book past it is one the
/// engine is not going to finish. One number rather than two, for the reason every other engine's
/// give-up is one: a hover waiting behind a hung conversion is waiting on the same question, and
/// the answer is that conversion being ended rather than a longer wait.
const CONVERSION_GIVE_UP: Duration = Duration::from_secs(60);

/// How often the wait above looks.
const CONVERSION_POLL: Duration = Duration::from_millis(100);

/// How long the engine thread waits on its slot before looking again.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made and this is only
/// a ceiling on how long anything else — a conversion asked to be given up on, a run that is
/// ending — goes unnoticed.
const IDLE_TICK: Duration = Duration::from_secs(1);

/// Where Calibre keeps its programs, for the two places the installer puts them and for a portable
/// copy a user may have put beside `config.ini`.
///
/// The folder is the one the Windows installer writes by default — `Calibre2`, which is not a
/// version but the name the application has installed itself under since its second major version
/// — and the portable copies are the two places every other engine of this app's is looked for
/// beside: the folder `config.ini` lives in, and the folder the app runs from. Nothing is guessed
/// at from the `PATH`: the engine this app runs is the one inside a Calibre installation, and
/// nothing else on the machine is asked as one.
fn converter() -> Option<PathBuf> {
    static FOUND: Lazy<Option<PathBuf>> = Lazy::new(|| {
        for candidate in [
            r"C:\Program Files\Calibre2",
            r"C:\Program Files (x86)\Calibre2",
        ] {
            let path = Path::new(candidate).join(ENGINE_IMAGE);
            if path.is_file() {
                return Some(path);
            }
        }

        let mut portable: Vec<PathBuf> = Vec::new();

        if let Some(beside_config) =
            AppConfig::config_path().and_then(|path| path.parent().map(Path::to_path_buf))
        {
            portable.push(beside_config.join("calibre").join(ENGINE_IMAGE));
        }
        if let Some(beside_app) = std::env::current_exe()
            .ok()
            .and_then(|exe| exe.parent().map(Path::to_path_buf))
        {
            portable.push(beside_app.join("calibre").join(ENGINE_IMAGE));
        }

        portable.into_iter().find(|path| path.is_file())
    });

    FOUND.clone()
}

/// Whether an engine is installed to convert these books with.
pub fn available() -> bool {
    converter().is_some()
}

/// The page already converted for this version of the book, if there is one: a PDF the PDF path
/// reads the way it reads any other.
///
/// Nothing is started and nothing is waited on here. This is the question "has the engine answered
/// yet?" asked of the document cache — a preview is laid out from what it finds, and a hover that
/// is waiting for one keeps asking it until the answer is there (see [`request`]).
pub fn rendered_page(path: &Path) -> Option<PathBuf> {
    document_cache::page(path, ENGINE_NAME).map(|page| page.path)
}

/// Whether the engine has already turned this version of the book down: the mark a conversion that
/// wrote no page leaves where a page would have been.
///
/// It is what a hover waiting on that book reads as "nothing is coming" — the spinner comes down
/// rather than running out its wait — and what keeps the conversion from being paid for a second
/// time. A mark ages out, so a book that was unreadable for a while — locked, or half-copied, or
/// opened while the engine was still installing itself — is asked about again (see
/// `document_cache`).
pub fn refused(path: &Path) -> bool {
    document_cache::refused(path, ENGINE_NAME)
}

/// Ask the engine for a page for `path`.
///
/// Nothing is waited on and nothing is answered: the conversion runs on the engine thread below,
/// and what a caller watches for is the page appearing in the folder it is kept in, or the mark
/// that says it is not coming. A hover that asked for one is replayed when it is there.
pub fn request(path: &Path) {
    if !imports(path) || !available() {
        return;
    }
    if rendered_page(path).is_some() || refused(path) {
        return;
    }

    // A conversion that has been inside one book for longer than any of them takes has stopped
    // answering, so it is ended here rather than queued behind: what that frees is the thread it
    // was holding and the book that is waiting for it.
    end_hung_conversion();

    // A request for the book already being converted is that request. One that arrives while
    // another waits replaces it, the way the loader's slot does: the newest hover is the one the
    // pointer is on, and a book whose hover has gone is one nobody is waiting for.
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

/// Whether the engine is the one that reads this file: a name of its own list, or the bytes of a
/// book it reads under a name no list holds — a `.mobi` renamed to `.dat`, say.
///
/// The question is asked where it is answered for every caller — see
/// `calibre_formats::is_engine_ebook` — so that a book one side asks about is a book the other
/// side will convert. Nothing is asked of a file that is what it is called but is not one of the
/// engine's formats, and nothing is asked of one whose bytes are another kind's.
fn imports(path: &Path) -> bool {
    crate::formats::calibre_formats::is_engine_ebook(path)
}

/// The book the engine is converting now, if it is converting one.
fn running_source() -> Option<PathBuf> {
    RUNNING
        .lock()
        .ok()?
        .as_ref()
        .map(|running| running.source.clone())
}

/// End a conversion that has outrun the engine's give-up.
///
/// Ending the process a conversion is waiting on is what ends the wait: the engine thread reads it
/// as a run that wrote nothing, remembers the book as one the engine will not convert, and takes
/// up the book behind it. What the caller here is left with is an engine that costs nothing and the
/// answer it would have reached anyway.
fn end_hung_conversion() {
    let hung = RUNNING.lock().ok().and_then(|running| {
        running
            .as_ref()
            .filter(|running| is_hung(running))
            .map(|running| running.pid)
    });

    if let Some(pid) = hung {
        // Verified by name and start time before anything is ended, like every other process this
        // app holds a record of.
        crate::app::engine_processes::terminate_owned(pid);
    }
}

/// Whether a conversion in flight has had its chance: a book the engine has been reading for longer
/// than any of them takes is one it is not going to finish.
fn is_hung(running: &Running) -> bool {
    running.started.elapsed() >= CONVERSION_GIVE_UP
}

/// Start the thread conversions run on, once.
///
/// It is one of the app's threads rather than one per book: what it does between conversions is
/// wait on its own slot, which costs nothing, and what it is asked for is one book at a time
/// because what it is holding is one process.
fn start_engine() {
    if ENGINE_STARTED.swap(true, Ordering::AcqRel) {
        return;
    }

    std::thread::spawn(|| {
        while let Some(source) = next_request() {
            // Whatever the engine makes of the book is written where a page is kept, or left
            // unwritten as a mark: there is nothing to answer with here, and nothing to send. The
            // side that asked reads the folder the page lands in (see `request`). A panic is
            // contained for the reason the loader contains one: one book's failure is that book's,
            // and the thread goes on to the next hover.
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| convert(&source))).ok();
        }
    });
}

/// The next book to convert, waiting for one.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made; a lock poisoned
/// by a panic on another thread is still the same slot, and a queue of one is not worth standing
/// down over.
fn next_request() -> Option<PathBuf> {
    let (slot, ready) = &*REQUESTED;
    let mut requested = match slot.lock() {
        Ok(requested) => requested,
        Err(poisoned) => poisoned.into_inner(),
    };

    loop {
        if let Some(requested) = requested.take() {
            return Some(requested);
        }

        // The wait is bounded rather than endless, and what a bound that runs out is for is a
        // request that may have been made in the meantime.
        requested = match ready.wait_timeout(requested, IDLE_TICK) {
            Ok((requested, _)) => requested,
            Err(poisoned) => poisoned.into_inner().0,
        };
    }
}

/// Convert `source` into a page, by running the engine the way a user would: one book in, one PDF
/// out, and a wait that ends rather than holding a hover for good.
///
/// What the engine wrote is kept as the book's page. A book it would not convert is answered with
/// nothing, and that is written down where the page would have been, so the conversion is not paid
/// for twice.
fn convert(source: &Path) {
    let Some(program) = converter() else {
        return;
    };
    let Some(stage) = stage_folder() else {
        document_cache::refuse(source, ENGINE_NAME);
        return;
    };

    // Whatever a run before this one left behind is not read: the conversion is given the name of
    // the file it is to write, and only that file is looked for.
    std::fs::remove_dir_all(&stage).ok();
    if std::fs::create_dir_all(&stage).is_err() {
        document_cache::refuse(source, ENGINE_NAME);
        return;
    }

    let written = stage.join("page.pdf");
    let Some(page) = run(&program, source, &written) else {
        let _ = std::fs::remove_dir_all(&stage);
        document_cache::refuse(source, ENGINE_NAME);
        return;
    };

    document_cache::store(source, ENGINE_NAME, PageKind::Pdf, &page);
    let _ = std::fs::remove_dir_all(&stage);
}

/// Run the engine over one book, and answer the bytes of the PDF it wrote — or nothing where it
/// wrote none, ended badly, or was given up on.
///
/// The book is named to the engine by its whole path, and the page by a name of this app's own
/// rather than by the book's: a conversion reads what it is given and writes what it is told, so
/// nothing about the file's own name can move a file of this app's or be read as one of the
/// engine's own switches.
fn run(program: &Path, source: &Path, written: &Path) -> Option<Vec<u8>> {
    let mut child = Command::new(program)
        .arg(source)
        .arg(written)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(crate::app::engine_processes::CREATE_NO_WINDOW)
        .spawn()
        .ok()?;

    // A process this app started is put in the job every other engine is put in, and held in a
    // record beside it: one that outlives the app does not outlive it by much, one that is given
    // up on is ended by name and id wherever it is asked about from, and one a run never got to
    // end is answered for by the next run's reaper (see `engine_processes`).
    crate::app::engine_processes::record(ENGINE_IMAGE, child.id());

    // What the engine is converting, published for the threads that may decide it has stopped
    // answering while this one waits (see `end_hung_conversion`).
    publish_running(Some(Running {
        source: source.to_path_buf(),
        pid: child.id(),
        started: Instant::now(),
    }));

    // The conversion's own bound, which is the same give-up every other thread reads: a book still
    // being converted and an engine that has stopped converting look the same from the outside,
    // and time is what tells them apart.
    let converted = wait(&mut child, CONVERSION_GIVE_UP);

    publish_running(None);
    crate::app::engine_processes::forget(child.id());

    if !converted {
        return None;
    }

    // Nothing of a conversion is believed but the file it wrote: an engine that ended well and
    // wrote nothing, and one that wrote something that is not a PDF, are both a book it would not
    // convert. What is read is bounded like every other read of this app's, by the ceiling one
    // hover may decode for — a book of a thousand scanned pages is a PDF of gigabytes, and one
    // past the budget is one no preview could draw.
    let page = read_within_budget(written)?;
    page.starts_with(b"%PDF-").then_some(page)
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

/// The folder a conversion is staged in.
///
/// Nothing of the user's stays in it longer than a conversion takes — it is emptied before every
/// one and taken away after it — and where it is is the app's own scratch folder under the temp
/// folder, which is where every other scratch file an engine can only answer with is written (see
/// `document_cache::temp_folder`).
fn stage_folder() -> Option<PathBuf> {
    Some(document_cache::temp_folder().join("calibre"))
}

/// Say what the engine is converting now, or that it has stopped.
fn publish_running(running: Option<Running>) {
    if let Ok(mut published) = RUNNING.lock() {
        *published = running;
    }
}

/// The conversion in flight: the book being converted, the process converting it, and since when.
/// It is what tells a busy engine from one that has stopped answering, and it is the id a
/// conversion that has to be ended is ended by.
struct Running {
    source: PathBuf,
    pid: u32,
    started: Instant,
}

static RUNNING: Lazy<Mutex<Option<Running>>> = Lazy::new(|| Mutex::new(None));

/// The book waiting to be converted, and the signal that one is there: a queue of one, for the
/// reason there is one engine at a time.
static REQUESTED: Lazy<(Mutex<Option<PathBuf>>, Condvar)> =
    Lazy::new(|| (Mutex::new(None), Condvar::new()));

/// Whether the engine thread has been begun.
static ENGINE_STARTED: AtomicBool = AtomicBool::new(false);

#[cfg(test)]
mod tests {
    use super::*;

    /// The list of names this engine is asked about is the format list's, and the answer is the
    /// same whichever side asks it: what the request side asks the engine about is what the box
    /// and the loader ask it about, so a hover is never waiting for a conversion nothing was asked
    /// to make.
    #[test]
    fn asks_about_the_books_its_own_list_names() {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.calibre_extensions =
                crate::formats::calibre_formats::sanitize_calibre_extensions(
                    crate::formats::calibre_formats::DEFAULT_CALIBRE_EXTENSIONS,
                );
        }

        for name in [
            "book.epub",
            "book.mobi",
            "book.azw3",
            "book.fb2",
            "book.djvu",
        ] {
            assert!(
                imports(Path::new(name)),
                "`{name}` is one of the engine's books"
            );
        }

        for name in ["photo.png", "report.docx", "notes.txt", "archive.zip"] {
            assert!(
                !imports(Path::new(name)),
                "`{name}` is another kind's, and the engine is not asked about it"
            );
        }
    }

    /// What a page is kept under is a name of this engine's own, so that a book and a document
    /// that is the same file cannot be handed each other's page (see `document_cache::key`).
    #[test]
    fn keeps_its_pages_under_its_own_name() {
        assert_eq!(ENGINE_NAME, "calibre");
        assert_ne!(
            ENGINE_NAME,
            crate::config::config::OfficeEngine::MicrosoftOffice.as_str()
        );
        assert_ne!(
            ENGINE_NAME,
            crate::config::config::OfficeEngine::LibreOffice.as_str()
        );
    }
}

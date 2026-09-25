//! Archive listings produced by an installed PeaZip.
//!
//! This app reads the archives people meet — a zip, a 7z, a rar, a tar — from their own tables of
//! contents, and shows a page of what each holds. Everything else is a format no reader here has:
//! a cabinet file out of a driver package, the arj of an old backup, an `lzh`, an iso, a wim, a
//! vhd, the `.deb` and `.rpm` of a Linux download, the single-stream `.gz`, `.bz2`, `.xz` and
//! `.zst` a file is put through before it is sent. PeaZip is the tool that opens them, and what it
//! is asked for here is the one thing that page needs: what is in the file.
//!
//! What is asked, and of which of the tools PeaZip ships, is settled by the name a file carries —
//! see `peazip_formats::Backend`, which answers it for every name this kind is gated by. Most
//! names are the console archiver's: the 7-Zip console PeaZip carries in `res\bin\7z`, which reads
//! more formats than anything else it ships and can be asked for a report meant to be read by
//! something other than a person. Three belong to the backends beside it, each of which reads a
//! format that one cannot open at all — FreeArc's archiver for an `.arc`, zpaq for a `.zpaq`, and
//! Zstandard's own tool for a `.zst`, which the archiver reads too but without saying how large the
//! stream was before it was compressed, which is the whole of what a hover on one is for.
//!
//! And three have no listing to ask for at all. A `.br`, a `.bcm` and an `.lpaq8` are single
//! streams: the tools that write them put one file into one file and have no command that prints
//! what is inside, because there is nothing inside but the bytes. What this module answers for one
//! of those is the member an extraction would write — its name, taken from the file's own name the
//! way every single-stream name here is, and no size, since the only way to learn that one is to
//! decompress — and no process is started for it: the page is this app's own answer, and what says
//! the name may be previewed at all is the tool being there.
//!
//! What a tool hands back is the answer it prints for an archive — the archive's table of contents,
//! entry by entry, as text — and that text is what this module carries back to the rest of the app.
//! It is read into the same listing every other archive's is read into (see `archive_listing`),
//! held in the same cache under the same key, and drawn by the same page: a `.cab` is previewed
//! exactly like a `.zip`, because that is what it is, a list of what the file holds. Nothing of a
//! tool reaches the screen and nothing of it is kept: what is kept between hovers is the listing,
//! and a second hover of the same archive is a lookup that starts nothing at all.
//!
//! Nothing is bundled with this app and nothing is linked against: the tools are the user's own
//! installation of PeaZip, looked for where it installs and beside `config.ini` for a portable
//! copy, and run as the user runs them. A name is answered only where the tool that reads it is
//! there, so a portable copy carrying a few of them shows the names those few read and no others.
//! What a listing costs is a launch — a fraction of a second for an ordinary archive — and it is
//! not one a preview can wait on: the caller is the preview loop, and a loop held inside a launch
//! is a hover that does not come up, a tray that does not answer and a pointer that cannot leave
//! the file it is on. So a tool runs on a thread of its own. What the loop asks is [`request`],
//! which returns at once, and what it waits for is the answer that thread sends back through the
//! preview channel — the same wait a page an engine drew has, in the same box, with the hover
//! replayed when the answer lands. One listing runs at a time, and the file waiting behind it is
//! the newest one asked for.
//!
//! An engine that has stopped answering is the other half of that. A file one of these tools cannot
//! open does not always fail quickly — a damaged archive, or one whose format it reads through a
//! delegate this machine was never given, can leave a listing running past any bound a hover is
//! owed — so a listing that has outrun [`LISTING_GIVE_UP`] is ended where it stands, and the file
//! it was on is remembered as one the engine will not list: the launch is paid for once and never
//! again, and the hover that asked for it is answered rather than left spinning. What is
//! remembered lives for the run rather than on disk, like the answer itself. One of these tools was
//! measured doing exactly that rather than failing, and it is worth naming: the zpaq PeaZip 10.9.0
//! ships — zpaqfranz — does not come back at all from a file that is not a zpaq, so a `.zpaq` whose
//! bytes are something else is a hover the give-up answers.
//!
//! What this module is *not* is a process the app keeps. It is the shape `imagemagick_render` has
//! rather than the shape `libreoffice_render` has, and for the same reason: each of these tools is
//! a converter — it is handed an archive, prints what is inside it and exits — so there is nothing
//! to hold open between files and nothing for an idle time to bound. What a user who wants their
//! archives to open faster changes is not a TTL but the listing cache: a second hover of an archive
//! whose listing is still held costs no tool at all. See the `Engine` submenu, where the engines
//! that *are* kept have their `… TTL` rows and this one has none.

use crate::formats::peazip_formats::Backend;
use once_cell::sync::Lazy;
use std::io::Read;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, ExitStatus, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{mpsc, Condvar, Mutex};
use std::time::{Duration, Instant, SystemTime};

/// How long a listing is given, and the point past which the engine is a stopped one rather than
/// a busy one: the run is ended where it stands, and the file it was on is remembered as one the
/// engine will not list.
///
/// Reading a table of contents is not the extraction that could take minutes — it is a walk over
/// the archive's own headers, and every archive this was measured on answers inside a second —
/// so this is an order of magnitude past anything an archive the engine can read costs, while a
/// file it cannot read does not finish at all, whether it fails slowly or spins. One number
/// rather than two, for the reason the image converter's give-up is one: a hover waiting behind a
/// hung listing is waiting on the same question, and the answer is that engine being ended rather
/// than a longer wait.
const LISTING_GIVE_UP: Duration = Duration::from_secs(30);

/// How often the wait above looks.
const LISTING_POLL: Duration = Duration::from_millis(50);

/// How long the report the engine wrote is waited for once the process that wrote it has ended.
///
/// The report is read from the engine's own output as it is written, on a thread of its own, so
/// that a directory of twenty thousand entries does not deadlock against the pipe between the two
/// processes: what is waited for here is the last of those bytes, which arrive as the process
/// closes its end. A process that has ended has closed it, so the wait is a formality — it is
/// bounded rather than endless for the one way it cannot be: a tool that leaves something else
/// holding that end.
const OUTPUT_WAIT: Duration = Duration::from_secs(5);

/// How long the engine thread waits on its slot before looking again.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made and this is only
/// a ceiling on how long anything else — a file the engine thread is asked to give up on, a run
/// that is ending — goes unnoticed.
const IDLE_TICK: Duration = Duration::from_secs(1);

/// What the engine's answers about a run are held for: which files it listed, and which it would
/// not. One hover of one file asks for one answer, and a folder is swept a file at a time, so the
/// list is a session's worth of files rather than a scan's.
const ANSWERS_MAX_ENTRIES: usize = 512;

/// The folder PeaZip installs into, under each program directory.
const ENGINE_FOLDER: &str = "PeaZip";

/// The tools of this machine's PeaZip, looked for under the folders it installs into.
///
/// It is not PeaZip's own executable that is looked for. `peazip.exe` is the windowed frontend, and
/// what it does when it is pointed at an archive is open a window — its `-ext2list` switch has it
/// *show* a listing rather than print one, in a window of its own, which is neither a preview nor
/// something a hover may do, and nothing of it would reach a pipe anyway. The work is done by the
/// archivers PeaZip carries beside it, and which of them a file is handed to is settled by the name
/// it carries — see `Backend`.
///
/// Each tool is looked for separately, so an installation missing one of them — a portable copy
/// carrying a few — still answers for the names the others read, and a name whose own tool is not
/// there shows nothing rather than a page of something else.
struct Backends {
    seven_zip: Option<PathBuf>,
    arc: Option<PathBuf>,
    zpaq: Option<PathBuf>,
    zstd: Option<PathBuf>,
    brotli: Option<PathBuf>,
    bcm: Option<PathBuf>,
    lpaq: Option<PathBuf>,
}

impl Backends {
    /// The tool this machine has for a file, or nothing where the installation carries none.
    fn program(&self, backend: Backend) -> Option<&Path> {
        match backend {
            Backend::SevenZip => self.seven_zip.as_deref(),
            Backend::Arc => self.arc.as_deref(),
            Backend::Zpaq => self.zpaq.as_deref(),
            Backend::Zstd => self.zstd.as_deref(),
            Backend::Brotli => self.brotli.as_deref(),
            Backend::Bcm => self.bcm.as_deref(),
            Backend::Lpaq => self.lpaq.as_deref(),
        }
    }
}

/// The folders a PeaZip installation is looked for under: the two the installer writes to, and the
/// two a portable copy sits in.
///
/// The two program directories are where the installer puts it — `C:\Program Files\PeaZip` and its
/// 32-bit sibling — and the portable copies are the two places every other engine of this app's is
/// looked for beside: the folder `config.ini` lives in, and the folder the app runs from. Nothing
/// is guessed at from the `PATH`: a tool this app runs is the one inside a PeaZip installation, and
/// a `7z.exe` found on the `PATH` is another program's copy of the same archiver — one with no
/// PeaZip codecs beside it and no PeaZip the user installed.
fn installation_folders() -> Vec<PathBuf> {
    let mut roots: Vec<PathBuf> = vec![
        Path::new(r"C:\Program Files").join(ENGINE_FOLDER),
        Path::new(r"C:\Program Files (x86)").join(ENGINE_FOLDER),
    ];

    if let Some(beside_config) = config_folder() {
        roots.push(beside_config.join("peazip"));
    }
    if let Some(beside_app) = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.parent().map(Path::to_path_buf))
    {
        roots.push(beside_app.join("peazip"));
    }

    roots
}

/// One tool of the installation, looked for under each folder in turn.
fn find_program(backend: Backend) -> Option<PathBuf> {
    installation_folders()
        .into_iter()
        .map(|folder| folder.join(backend.program()))
        .find(|program| program.is_file())
}

/// The folder `config.ini` lives in, which is where a portable engine is looked for beside.
fn config_folder() -> Option<PathBuf> {
    crate::config::config::AppConfig::config_path()
        .and_then(|path| path.parent().map(Path::to_path_buf))
}

/// Every tool this machine has, looked for once: the answer is a fact about an installation rather
/// than a question about a hover.
static BACKENDS: Lazy<Backends> = Lazy::new(|| Backends {
    seven_zip: find_program(Backend::SevenZip),
    arc: find_program(Backend::Arc),
    zpaq: find_program(Backend::Zpaq),
    zstd: find_program(Backend::Zstd),
    brotli: find_program(Backend::Brotli),
    bcm: find_program(Backend::Bcm),
    lpaq: find_program(Backend::Lpaq),
});

/// Whether an engine is installed to list these archives with: the console archiver, which is what
/// the list is mostly made of and what the tray's own row is about. A name one of the other tools
/// reads is asked about where it is routed, by `available_for`.
pub fn available() -> bool {
    BACKENDS.seven_zip.is_some()
}

/// Whether the tool that reads this file is installed, which is what says a hover onto it may be
/// answered at all. A machine without PeaZip has none of them, and a portable copy carrying a few
/// answers for the names those few read.
pub fn available_for(path: &Path) -> bool {
    BACKENDS.program(Backend::of(path)).is_some()
}

/// Whether the engine is the one that lists this file: a name of its own list, or the bytes of an
/// archive it reads under a name no list holds — a `.cab` renamed to `.dat`, say.
///
/// The question is asked where it is answered for every caller — see
/// `peazip_formats::is_engine_archive` — so that an archive one side asks about is an archive the
/// other side will list. Nothing is asked of a file that is what it is called but is not one of
/// the engine's formats, and nothing is asked of one whose bytes are another kind's.
fn imports(path: &Path) -> bool {
    crate::formats::peazip_formats::is_engine_archive(path)
}

/// A file and the version of it an answer is about: its path, what it weighed and when it was last
/// written, which is what says it is a file to list again.
#[derive(Clone, PartialEq, Eq, Hash)]
struct Key {
    path: PathBuf,
    len: u64,
    modified: Option<SystemTime>,
}

impl Key {
    fn of(path: &Path) -> Self {
        let metadata = std::fs::metadata(path).ok();

        Self {
            path: path.to_path_buf(),
            len: metadata
                .as_ref()
                .map(|metadata| metadata.len())
                .unwrap_or(0),
            modified: metadata.and_then(|metadata| metadata.modified().ok()),
        }
    }
}

/// The files the engine would not list, by file and version.
///
/// It is what keeps a name the list was wrong about — one whose format the engine reads through a
/// delegate this machine was not given, an archive whose table of contents is encrypted, a file
/// that is not an archive at all — from starting an engine on every hover to reach the same
/// answer.
static REFUSED: Lazy<Mutex<Vec<Key>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Whether the engine has already turned this file down, for this version of it.
pub fn refused(path: &Path) -> bool {
    let key = Key::of(path);

    REFUSED
        .lock()
        .map(|refused| refused.contains(&key))
        .unwrap_or(false)
}

/// Whether a listing for this file is already in hand, wherever it came from.
///
/// It is the question the request side asks before it asks the engine for one — an archive whose
/// listing is held is an archive there is nothing to ask about — and the answer is the listing
/// cache's own (see `archive_listing::is_listed`), so a listing this app's own readers produced
/// answers it exactly as one the engine did.
pub fn listed(path: &Path) -> bool {
    crate::readers::archive_listing::is_listed(path)
}

/// Ask the engine for a listing of `path`.
///
/// Nothing is waited on and nothing is answered: the run happens on the engine thread below, and
/// what a caller watches for is the message it sends when it is done — or the mark that says the
/// listing is not coming. A hover that asked for one is replayed when the answer lands.
pub fn request(path: &Path, generation: u64) {
    if !imports(path) || !available_for(path) || refused(path) || listed(path) {
        return;
    }

    // A listing that has been inside one file for longer than any of them takes has stopped
    // answering, so it is ended here rather than queued behind: what that frees is the thread it
    // was holding and the file that is waiting for it.
    end_hung_listing();

    // A request for the file already being listed is that request. One that arrives while another
    // waits replaces it, the way the loader's slot does: the newest hover is the one the pointer
    // is on, and a file whose hover has gone is one nobody is waiting for.
    if running_source().as_deref() == Some(path) {
        return;
    }

    let (slot, ready) = &*REQUESTED;
    if let Ok(mut requested) = slot.lock() {
        *requested = Some(Requested {
            path: path.to_path_buf(),
            generation,
        });
    }
    ready.notify_all();

    start_engine();
}

/// The file the engine is listing now, if it is listing one.
fn running_source() -> Option<PathBuf> {
    RUNNING
        .lock()
        .ok()?
        .as_ref()
        .map(|running| running.source.clone())
}

/// End a listing that has outrun the engine's give-up.
///
/// Ending the process a listing is waiting on is what ends the wait: the engine thread reads it as
/// a run that reported nothing, remembers the file as one the engine will not list, and takes up
/// the file behind it. What the caller here is left with is an engine that costs nothing and the
/// answer it would have reached anyway.
fn end_hung_listing() {
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

/// Whether a listing in flight has had its chance: a file the engine has been reading for longer
/// than any of them takes is one it is not going to finish.
fn is_hung(running: &Running) -> bool {
    running.started.elapsed() >= LISTING_GIVE_UP
}

/// Start the thread listings run on, once.
///
/// It is one of the app's threads rather than one per file: what it does between listings is wait
/// on its own slot, which costs nothing, and what it is asked for is one file at a time because
/// what it is holding is one process.
fn start_engine() {
    if ENGINE_STARTED.swap(true, Ordering::AcqRel) {
        return;
    }

    std::thread::spawn(|| {
        while let Some(requested) = next_request() {
            // Whatever the engine answers — a listing, or a run that reported none — the hover
            // waiting on this file is told either way. A panic is contained here for the reason
            // the loader contains one: one file's failure is that file's, and the thread goes on
            // to the next hover.
            let listing =
                std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| list(&requested.path)))
                    .unwrap_or(None);

            let ok = listing.is_some();
            if let Some(listing) = listing {
                crate::readers::archive_listing::remember_engine_listing(&requested.path, listing);
            } else {
                refuse(&Key::of(&requested.path));
            }

            crate::ui::preview_window::notify_peazip_ready(
                &requested.path,
                requested.generation,
                ok,
            );
        }
    });
}

/// The next file to list, waiting for one.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made; a lock poisoned
/// by a panic on another thread is still the same slot, and a queue of one is not worth standing
/// down over.
fn next_request() -> Option<Requested> {
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

/// List a file, by asking the tool of the installation that reads it the way a user would.
///
/// What comes back is the archive's table of contents read into the shape the rest of the app reads
/// a listing in, or nothing where the tool reported none — which is the answer for an archive it
/// cannot open, one whose own table of contents is encrypted, a run that was given up on, and a
/// process that could not be started at all. And for a single-stream file — a `.br`, a `.bcm`, an
/// `.lpaq8` — there is no tool to ask and nothing to ask it: what comes back is the member an
/// extraction would write, derived from the file's own name (see
/// `archive_listing::stream_listing`).
fn list(path: &Path) -> Option<crate::readers::archive_listing::Listing> {
    let backend = Backend::of(path);
    let file_size = std::fs::metadata(path).ok()?.len();

    if !backend.has_a_listing() {
        return crate::readers::archive_listing::stream_listing(path, file_size);
    }

    let program = BACKENDS.program(backend)?;
    let (report, status) = contents(backend, program, path)?;

    let mut listing = match backend {
        Backend::SevenZip => {
            crate::readers::archive_listing::engine_listing(&report, path, file_size)
        }
        Backend::Arc => crate::readers::archive_listing::arc_listing(&report, file_size),
        Backend::Zpaq => crate::readers::archive_listing::zpaq_listing(&report, file_size),
        Backend::Zstd => crate::readers::archive_listing::zstd_listing(&report, path, file_size),
        // A tool with no listing is never run, so there is never a report of its to read: the arm
        // is here because a name has to be routed to something.
        Backend::Brotli | Backend::Bcm | Backend::Lpaq => None,
    }?;

    // A run that reported a table of contents and then ended badly — an archive it could read part
    // of and not the rest — is a listing that stops where the tool stopped, which is the caveat the
    // page states in its own words.
    listing.read_truncated = !status.success() || listing.read_truncated;

    Some(listing)
}

/// Ask a tool what is in `source`, and answer the report it printed and how it ended.
///
/// What each is asked is a decision of its own, measured against PeaZip 10.9.0:
///
/// * The console archiver is asked for a technical listing — `-slt`, the form meant to be read by
///   something other than a person — and FreeArc for its verbose one, `v`, which is the only one of
///   its three listings that carries a folder's attributes. Both are given `--` to end the
///   switches, so a file whose own name begins with a dash is still a file rather than an argument.
/// * zpaq is asked to list (`l`) with no `--` after it, because it takes none: given one it prints
///   its usage instead, which this side reads as no report at all and remembers as a file that
///   could not be listed.
/// * And zstd is asked for `-l`, its frame table, which is the whole reason `.zst` is routed to it
///   rather than to the archiver: the archiver lists a `.zst` too, and leaves the sizes blank.
///
/// What is shared is what every one of them needs:
///
/// * The report is read from the tool's own output as it is written, on a thread of its own, so that
///   an archive holding twenty thousand entries does not deadlock against the pipe between the two
///   processes. It is bounded like every other read of this app, by the ceiling one hover may
///   decode for, so a tool answering with more text than any listing is costs the budget rather
///   than the machine.
/// * Its input is closed (`Stdio::null()`) and that is not a detail: an archive whose table of
///   contents is encrypted makes one of these ask for a password on its own input, and a process
///   asking a question nobody can answer is a preview that never comes up. Reading from nothing, it
///   is answered with an end of input and gives up on the file, which is the answer this side wants
///   and the reason it is safe to ask at all.
/// * And it is started without a console window of its own (`CREATE_NO_WINDOW`): these are console
///   programs and this app is not, so Windows would otherwise give every listing a window of its
///   own — a black rectangle over whatever the pointer was on, for as long as the launch lasted.
///   See `engine_processes` for the flag and for what it is said of.
fn contents(backend: Backend, program: &Path, source: &Path) -> Option<(String, ExitStatus)> {
    let mut command = Command::new(program);

    match backend {
        Backend::SevenZip => {
            command.args(["l", "-slt", "--"]);
        }
        Backend::Arc => {
            command.args(["v", "--"]);
        }
        Backend::Zpaq => {
            command.arg("l");
        }
        Backend::Zstd => {
            command.arg("-l");
        }
        Backend::Brotli | Backend::Bcm | Backend::Lpaq => return None,
    }

    let mut child = command
        .arg(source)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .creation_flags(crate::app::engine_processes::CREATE_NO_WINDOW)
        .spawn()
        .ok()?;

    // A process this app started is put in the job every other engine is put in, so that whatever
    // ends the app ends it. It is adopted rather than recorded, the way a video's probes are: a
    // listing lives a fraction of a second and writing it down would be a record made and struck
    // off again inside the hover that started it, and the job is what answers for one left by a
    // run that ended badly (see `engine_processes`).
    crate::app::engine_processes::adopt(child.id());

    let output = child.stdout.take()?;
    let (written, written_rx) = mpsc::channel();
    std::thread::spawn(move || {
        let mut report = Vec::new();
        let read = output.take(decode_budget()).read_to_end(&mut report);
        let _ = written.send(read.ok().map(|_| report));
    });

    // What the engine is reading, published for the threads that may decide it has stopped
    // answering while this one waits (see `end_hung_listing`).
    publish_running(Some(Running {
        source: source.to_path_buf(),
        pid: child.id(),
        started: Instant::now(),
    }));

    let status = wait(&mut child, LISTING_GIVE_UP);

    publish_running(None);

    let report = written_rx.recv_timeout(OUTPUT_WAIT).ok().flatten();

    let (report, status) = (report?, status?);

    // What the engine read is text, and a report that is not text at all is not a report: the run
    // is answered as one that reported nothing rather than as one that reported nonsense.
    Some((String::from_utf8_lossy(&report).into_owned(), status))
}

/// The ceiling one hover may read for, in bytes, read from the configuration each time so that an
/// edit applies without a restart.
fn decode_budget() -> u64 {
    crate::config::config::decode_budget_bytes()
}

/// Wait for a process, ending it rather than waiting past `limit`.
///
/// The status comes back where the process ended inside the bound and nothing where it was ended
/// here — the same distinction an image converter's wait makes, and for the same reason: a run
/// this side gave up on is not a run whose own answer may be believed.
fn wait(child: &mut Child, limit: Duration) -> Option<ExitStatus> {
    let deadline = Instant::now() + limit;
    loop {
        match child.try_wait() {
            Ok(Some(status)) => return Some(status),
            Ok(None) if Instant::now() < deadline => std::thread::sleep(LISTING_POLL),
            _ => {
                let _ = child.kill();
                let _ = child.wait();
                return None;
            }
        }
    }
}

/// Remember that the engine will not list this file, so that it is not asked twice.
fn refuse(key: &Key) {
    let Ok(mut refused) = REFUSED.lock() else {
        return;
    };

    if refused.len() >= ANSWERS_MAX_ENTRIES {
        refused.clear();
    }

    if !refused.contains(key) {
        refused.push(key.clone());
    }
}

/// Say what the engine is listing now, or that it has stopped.
fn publish_running(running: Option<Running>) {
    if let Ok(mut published) = RUNNING.lock() {
        *published = running;
    }
}

/// The listing in flight: the file being listed, the process listing it, and since when. It is
/// what tells a busy engine from one that has stopped answering, and it is the id a run that has
/// to be ended is ended by.
struct Running {
    source: PathBuf,
    pid: u32,
    started: Instant,
}

static RUNNING: Lazy<Mutex<Option<Running>>> = Lazy::new(|| Mutex::new(None));

/// A file to list, and the hover that asked for it.
struct Requested {
    path: PathBuf,
    generation: u64,
}

/// The file waiting to be listed, and the signal that one is there: a queue of one, for the reason
/// there is one engine at a time.
static REQUESTED: Lazy<(Mutex<Option<Requested>>, Condvar)> =
    Lazy::new(|| (Mutex::new(None), Condvar::new()));

/// Whether the engine thread has been started. It is one of the app's threads rather than one per
/// file, so it is started once and waits on its slot for the rest of the run.
static ENGINE_STARTED: AtomicBool = AtomicBool::new(false);

#[cfg(test)]
mod tests {
    use super::*;

    /// The engine is never started for a name its own list does not hold: an archive this app
    /// reads itself, a picture, a document and a video are all somebody else's, and asking an
    /// archiver about one would cost a launch and answer nothing.
    #[test]
    fn asks_the_engine_only_about_the_names_it_reads() {
        for name in [
            "archive.zip",
            "archive.7z",
            "archive.rar",
            "sources.tar.gz",
            "photo.png",
            "clip.mp4",
            "drawing.svg",
            "report.pdf",
            "notes.txt",
            "font.ttf",
            "document.pmd",
            "animation.swf",
        ] {
            assert!(
                !imports(Path::new(name)),
                "`{name}` is not one of its formats"
            );
        }

        for name in [
            "backup.cab",
            "image.iso",
            "package.msi",
            "readme.bz2",
            "readme.zst",
        ] {
            assert!(imports(Path::new(name)), "`{name}` is one of its formats");
        }
    }

    /// And a name it does not read is not queued either: nothing about a file is asked of the
    /// engine that its list has not claimed.
    #[test]
    fn queues_nothing_for_a_name_it_does_not_read() {
        let (slot, _) = &*REQUESTED;
        if let Ok(mut requested) = slot.lock() {
            *requested = None;
        }

        request(Path::new("photo.png"), 7);

        assert!(
            slot.lock()
                .map(|requested| requested.is_none())
                .unwrap_or(false),
            "the engine is not asked about a name no list of its own holds"
        );
    }

    /// What a file's refusals are remembered by: the file and the version of it that was read. A
    /// file saved again is a file to list again — and one the engine would not list is asked about
    /// once, not once per hover.
    #[test]
    fn remembers_a_refusal_for_the_version_of_the_file_it_was_read_from() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-peazip-tests")
            .join("remembering");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let archive = folder.join("backup.cab");
        std::fs::write(&archive, b"a cabinet, of a sort").expect("a written file");

        refuse(&Key::of(&archive));
        assert!(refused(&archive));
        assert!(!refused(&folder.join("other.cab")));

        // Saved again: what was known about the file it was is not what it is now.
        std::fs::write(&archive, b"a cabinet, edited and then some").expect("a written file");
        assert!(!refused(&archive));

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// What the engine ends is only what it started, and only once that run has had its chance. A
    /// process that stays up stands in for the engine — a test is not going to make a real one
    /// spin on an archive — recorded the way the engine is, by image name, which is the check that
    /// keeps an id from being acted on by itself.
    #[test]
    fn ends_a_listing_only_once_it_has_outrun_the_give_up() {
        let mut engine = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for the engine");
        let pid = engine.id();
        let running = |started: Instant| Running {
            source: PathBuf::from("backup.cab"),
            pid,
            started,
        };

        assert!(
            crate::app::engine_processes::processes_named("ping.exe").contains(&pid),
            "the stand-in runs the image the record below names it by"
        );
        crate::app::engine_processes::record("ping.exe", pid);

        // A listing that has just started is a file being read, and is left to it.
        publish_running(Some(running(Instant::now())));
        end_hung_listing();
        assert!(
            crate::app::engine_processes::is_running(pid),
            "a listing that has just started is not an engine to end"
        );

        // One that has outrun the give-up is an engine that has stopped answering.
        publish_running(Some(running(Instant::now() - LISTING_GIVE_UP)));
        end_hung_listing();
        assert!(
            !crate::app::engine_processes::is_running(pid),
            "the engine a listing has outrun is ended"
        );

        publish_running(None);
        let _ = engine.wait();
    }

    /// And the bound the engine thread holds over its own run: a process that ends inside it comes
    /// back with its own status, and one that outlasts it is ended there rather than waited on.
    #[test]
    fn ends_a_listing_rather_than_waiting_past_its_bound() {
        let mut quick = std::process::Command::new("ping")
            .args(["-n", "1", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process that ends on its own");
        let status = wait(&mut quick, LISTING_GIVE_UP).expect("a status");
        assert!(
            status.success(),
            "a run that finishes inside its bound is answered as it always was"
        );

        let mut engine = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for the engine");
        let pid = engine.id();
        let started = Instant::now();

        assert!(
            wait(&mut engine, Duration::from_millis(300)).is_none(),
            "a run that has outrun its bound is ended, not waited on"
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

    /// What the give-up is decided from: a run that has just started is not hung — a listing is
    /// under a second — and one that has run past the give-up is.
    #[test]
    fn a_listing_is_hung_only_once_it_has_outrun_the_give_up() {
        let running = |elapsed: Duration| Running {
            source: PathBuf::from("backup.cab"),
            pid: std::process::id(),
            started: Instant::now() - elapsed,
        };

        assert!(!is_hung(&running(Duration::from_secs(0))));
        assert!(!is_hung(&running(LISTING_GIVE_UP - Duration::from_secs(1))));
        assert!(is_hung(&running(LISTING_GIVE_UP)));
        assert!(is_hung(&running(LISTING_GIVE_UP + Duration::from_secs(30))));
    }

    /// Every tool of the installation is found inside a PeaZip folder and nowhere else: what the app
    /// runs are the archivers PeaZip carries, and a `7z.exe` another program installed is not one
    /// of them. Whatever this machine has is what the expectations are written from, and the rule
    /// is what is asserted.
    #[test]
    fn finds_the_tools_inside_an_installation_and_never_one_of_its_own() {
        for backend in [
            Backend::SevenZip,
            Backend::Arc,
            Backend::Zpaq,
            Backend::Zstd,
            Backend::Brotli,
            Backend::Bcm,
            Backend::Lpaq,
        ] {
            let program = match backend {
                Backend::SevenZip => &BACKENDS.seven_zip,
                Backend::Arc => &BACKENDS.arc,
                Backend::Zpaq => &BACKENDS.zpaq,
                Backend::Zstd => &BACKENDS.zstd,
                Backend::Brotli => &BACKENDS.brotli,
                Backend::Bcm => &BACKENDS.bcm,
                Backend::Lpaq => &BACKENDS.lpaq,
            };

            // A tool this installation does not carry is nothing to have expectations about: the
            // names it reads are answered with no preview, which the routing test in
            // `peazip_formats` is the one that says.
            let Some(program) = program else {
                continue;
            };

            assert!(
                program.is_file(),
                "a tool that was found is a program that is there"
            );
            assert_eq!(
                program.file_name().and_then(|name| name.to_str()),
                backend
                    .program()
                    .rsplit('\\')
                    .next()
                    .map(str::to_string)
                    .as_deref(),
                "{backend:?} runs under the name its own path inside the installation ends with"
            );

            // And it is inside a folder called what the installation root is called, which is what
            // a copy on the `PATH` would not be.
            let inside = program.to_string_lossy().replace('/', "\\");
            assert!(
                inside
                    .to_lowercase()
                    .contains(&format!("\\{}\\", ENGINE_FOLDER.to_lowercase())),
                "{backend:?} is inside a PeaZip folder: {inside}"
            );
        }

        assert_eq!(
            available(),
            BACKENDS.seven_zip.is_some(),
            "and whether an engine is installed is the answer the tray's own row goes by"
        );

        // A name is asked about where it is routed rather than of the installation as a whole, so
        // a machine carrying the archiver and not the backends still answers for the archiver's
        // names and shows nothing for the others.
        assert_eq!(
            available_for(Path::new("backup.zpaq")),
            BACKENDS.zpaq.is_some(),
            "a `.zpaq` is answered where zpaq is installed"
        );
        assert_eq!(
            available_for(Path::new("backup.cab")),
            BACKENDS.seven_zip.is_some(),
            "and a `.cab` where the console archiver is"
        );
    }

    /// What a listing is asked for, measured against the installed engine: an archive no reader
    /// here has is answered with its table of contents, a file that is not an archive is answered
    /// with nothing at all, and neither leaves anything on the disk. Ignored because it starts the
    /// installed PeaZip, and run when the command is being looked at:
    /// `cargo test -- --ignored --nocapture engine_listing_probe`.
    #[test]
    #[ignore = "starts the installed PeaZip"]
    fn engine_listing_probe() {
        let Some(program) = BACKENDS.program(Backend::SevenZip) else {
            println!("no PeaZip installed: nothing to measure");
            return;
        };
        println!("engine: {}", program.display());

        let folder = std::env::temp_dir().join("rust-hover-preview-peazip-probe");
        std::fs::create_dir_all(&folder).expect("a probe folder");

        // An archive to list, made with the same engine: what is being measured is the listing
        // rather than the format it is asked about, and every format is asked about the same way.
        // What is made is an xz of a text file — a format no reader of this app's own has, and a
        // single-stream one, which is the shape whose member carries no name inside it to read
        // (see `archive_listing::stream_member_name`).
        let source = folder.join("probe-source.txt");
        std::fs::write(&source, b"a text file, to be compressed\n").expect("a written file");

        let sample = folder.join("probe-source.txt.xz");
        let made = Command::new(program)
            .args(["a", "-txz", "-y"])
            .arg(&sample)
            .arg(&source)
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .status();
        println!("an archive to list: {made:?}");

        let started = Instant::now();
        let listed = list(&sample);
        println!(
            "listing it: {:?} entries in {:?}",
            listed
                .as_ref()
                .map(|listing| listing.entries.len())
                .unwrap_or(0),
            started.elapsed()
        );
        if let Some(listing) = &listed {
            for entry in listing.entries.iter().take(5) {
                println!(
                    "  {} ({} bytes, packed {:?}, dir {})",
                    entry.name, entry.size, entry.packed, entry.is_dir
                );
            }
        }

        // And a file that is not an archive at all: nothing is reported and nothing is held.
        let broken = folder.join("probe.cab");
        std::fs::write(&broken, b"not an archive at all").expect("a written file");
        println!("a file it cannot read: {:?}", list(&broken));

        let _ = std::fs::remove_dir_all(&folder);
    }
}

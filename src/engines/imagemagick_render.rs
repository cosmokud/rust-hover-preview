//! Pictures converted by an installed ImageMagick.
//!
//! A camera raw is a sensor reading rather than a picture: a `.nef`, a `.cr3`, an `.arw`,
//! a `.raf` or a `.dng` holds what the sensor saw, a picture wrapped around it and the
//! recipe for turning one into the other, and every application that shows one does the
//! same two things — demosaic the reading and develop it. This app has no reader for any
//! of it, and Windows has no codec for one either, so a hover onto a raw shows nothing at
//! all. ImageMagick is the tool that has the reader: the raw decoder it is built with,
//! LibRaw, is the library the whole world develops raws with, and its coders cover the
//! rest of the formats beside them that nothing else on the machine opens.
//!
//! So where ImageMagick is installed it is asked first, and what it hands back is a PNG —
//! the picture, developed, turned the right way up and resized to the box the preview is
//! shown at. That picture is read the way any picture of this app's is read: decoded into a
//! frame, composited over `image_background`, held in the image cache under the file, its
//! version and the pixel size it was made for. Nothing of the engine reaches the screen:
//! what is drawn is a picture of this app's, which is why a `.nef` follows the picture scale
//! and the picture backdrop rather than a scale and a backdrop of its own, and it is also
//! the whole point of the kind — a developed raw shown at the size it hovers as a photograph
//! rather than as a thumbnail.
//!
//! Nothing of it reaches the disk either. The engine is a converter rather than an
//! application: `magick.exe` reads a file, writes one and exits, so there is nothing to hold
//! open between files, and what it writes is written to its own standard output rather than
//! to a path — read on this side as the bytes of a picture, decoded from memory, and gone
//! when the frame is built. What is kept between hovers is what every other picture is kept
//! in: the frame, in the image cache, bounded by `image_cache_mb` and dropped least recently
//! used first. A raw whose frame is still there is a hover that costs nothing at all — no
//! engine, no decode — and one whose frame has been given up is developed again, which is
//! the price of holding no file of our own rather than a setting nobody has to make; see
//! Magick Previews for what that costs in memory.
//!
//! Nothing is bundled with this app and nothing is linked against: the engine is the user's
//! own installation, looked for where it installs and beside `config.ini` for a portable
//! copy, and run as the user runs it. What a conversion costs is a launch — a fraction of a
//! second for an ordinary raw, a second for a large one — and it is not one a preview can
//! wait on: the caller is the preview loop, and a loop held inside a launch is a hover that
//! does not come up, a tray that does not answer and a pointer that cannot leave the file it
//! is on. So the engine runs on a thread of its own. What the loop asks is [`request`],
//! which returns at once, and what it waits for is the answer that thread sends back through
//! the preview channel — the same wait an Office page has, in the same box, with the hover
//! replayed when the answer lands. One conversion runs at a time, and the file waiting
//! behind it is the newest one asked for.
//!
//! An engine that has stopped answering is the other half of that. A file a coder cannot
//! decode does not always fail quickly — a damaged raw, or a name in the list whose format
//! the engine reads through a delegate this machine was never given, can leave a conversion
//! running past any bound a hover is owed — so a conversion that has outrun
//! [`CONVERSION_GIVE_UP`] is ended where it stands, and the file it was on is remembered as
//! one the engine will not draw: the launch is paid for once and never again, and the hover
//! that asked for it is answered rather than left spinning. What is remembered lives for the
//! run rather than on disk, like the answer itself.

use crate::config::config::decode_budget_bytes;
use once_cell::sync::Lazy;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{mpsc, Condvar, Mutex};
use std::time::{Duration, Instant, SystemTime};

/// How long a conversion is given, and the point past which the engine is a stopped one
/// rather than a busy one: the conversion is ended where it stands, and the picture it was
/// on is remembered as one the engine will not draw.
///
/// A conversion of an ordinary raw is well under a second on the machine this was measured
/// on — the launch included — so this is an order of magnitude past anything a picture the
/// engine can read costs, while a file it cannot read does not finish at all, whether it
/// fails slowly or spins. One number rather than two, for the reason the render engine's
/// give-up is one: a hover waiting behind a hung conversion is waiting on the same
/// question, and the answer is that engine being ended rather than a longer wait.
const CONVERSION_GIVE_UP: Duration = Duration::from_secs(30);
/// How often the wait above looks.
const CONVERSION_POLL: Duration = Duration::from_millis(50);
/// How long the bytes the engine wrote are waited for once the process that wrote them has
/// ended.
///
/// The picture is read from the engine's own output as it is written, on a thread of its
/// own, so that a conversion larger than the pipe between them does not deadlock: what is
/// waited for here is the last of those bytes, which arrive as the process closes its end.
/// A process that has ended has closed it, so the wait is a formality — it is bounded
/// rather than endless for the one way it cannot be: a coder that leaves something else
/// holding that end.
const OUTPUT_WAIT: Duration = Duration::from_secs(5);
/// How long the engine thread waits on its slot before looking again.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made and this
/// is only a ceiling on how long anything else — a file the engine thread is asked to give up
/// on, a run that is ending — goes unnoticed.
const IDLE_TICK: Duration = Duration::from_secs(1);

/// The engine's own program, by the name it installs under, and the name the older one is
/// installed as.
///
/// `magick.exe` is ImageMagick 7 and `convert.exe` is ImageMagick 6, and both are looked
/// for because both are still installed. The older name is the one that has to be looked
/// for carefully: Windows ships a `convert.exe` of its own in `System32` — the file system
/// converter — and it is on the `PATH` of every machine, so a search that took the first
/// file of that name would run a system tool with a photograph for an argument (see
/// [`find_engine`]).
const ENGINE_IMAGE: &str = "magick.exe";
const LEGACY_ENGINE_IMAGE: &str = "convert.exe";

/// How many files the answers about a conversion are held for: what each engine developed
/// for a file, and which files it would not draw.
///
/// One hover of one file asks for one answer, and a folder is swept a file at a time, so
/// the list is a session's worth of files rather than a scan's: what it bounds is a pointer
/// dragged across a large folder of raws, and what it costs when it is reached is everything
/// held, the way every other cache of this app's shape answers that.
const ANSWERS_MAX_ENTRIES: usize = 512;

/// Where ImageMagick keeps its program, for the three places it installs and for a
/// portable copy a user may have put beside `config.ini`.
///
/// The two versioned folders are ImageMagick's own arrangement — `ImageMagick-7.1.2-Q16`
/// and its kin, one folder per build — so what is looked for is the family rather than one
/// of its members: the newest build installed is as good as the only one, and which of
/// them is there is not something this app decides. A machine with more than one has them
/// in the order the folder listing gives, which is the newest first where the version is
/// being read and an arbitrary one otherwise, and either is a working engine.
///
/// The older `convert.exe` is only ever taken from one of those folders, never from the
/// `PATH`: what is on the `PATH` under that name is Windows' own file system converter,
/// and running it would be a system change in place of a preview (see [`ENGINE_IMAGE`]).
fn find_engine() -> Option<PathBuf> {
    let mut folders: Vec<PathBuf> = Vec::new();

    for root in [r"C:\Program Files", r"C:\Program Files (x86)"] {
        let Ok(entries) = std::fs::read_dir(root) else {
            continue;
        };

        let mut installed: Vec<PathBuf> = entries
            .flatten()
            .map(|entry| entry.path())
            .filter(|path| {
                path.is_dir()
                    && path
                        .file_name()
                        .and_then(|name| name.to_str())
                        .is_some_and(|name| name.starts_with("ImageMagick"))
            })
            .collect();
        // Newest first, by the folder's own name: `ImageMagick-7.1.2-31` after
        // `ImageMagick-7.1.2-Q16` is a question about text rather than about versions, so
        // what the order is really for is that the same machine answers the same way twice.
        installed.sort();
        installed.reverse();

        folders.extend(installed);
    }

    // A portable copy, beside `config.ini` the way the render engine's is looked for
    // beside it, and a copy beside the app itself for the archive a user unpacks.
    if let Some(beside_config) = config_folder() {
        folders.push(beside_config.join("imagemagick"));
    }
    if let Some(own) = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.parent().map(Path::to_path_buf))
        .map(|folder| folder.join("imagemagick"))
    {
        folders.push(own);
    }

    for folder in folders {
        for image in [ENGINE_IMAGE, LEGACY_ENGINE_IMAGE] {
            let program = folder.join(image);
            if program.is_file() {
                return Some(program);
            }
        }
    }

    // And the `PATH` last, for an installation that was put somewhere else entirely —
    // for `magick.exe` alone, because the other name is Windows' own tool more often than
    // it is ImageMagick's (see `ENGINE_IMAGE`).
    std::env::var_os("PATH")
        .map(|path| std::env::split_paths(&path).collect::<Vec<PathBuf>>())
        .unwrap_or_default()
        .into_iter()
        .map(|folder| folder.join(ENGINE_IMAGE))
        .find(|program| program.is_file())
}

/// The folder `config.ini` lives in, which is where a portable engine is looked for beside.
fn config_folder() -> Option<PathBuf> {
    crate::config::config::AppConfig::config_path()
        .and_then(|path| path.parent().map(Path::to_path_buf))
}

/// The engine this machine has, looked for once: the answer is a fact about an
/// installation rather than a question about a hover.
static ENGINE: Lazy<Option<PathBuf>> = Lazy::new(find_engine);

/// Whether an engine is installed to develop these pictures with.
pub fn available() -> bool {
    ENGINE.is_some()
}

/// Whether the engine is the one that develops this file: a name of its own list, or the
/// bytes of a picture it reads under a name no list holds — a `.sgi` renamed to `.bin`, say.
///
/// The question is asked where it is answered for every caller — see
/// `magick_formats::is_engine_picture` — so that a picture one side asks about is a picture
/// the other side will read. Nothing is asked of a file that is what it is called but is not
/// one of the engine's formats, and nothing is asked of one whose bytes are another kind's.
fn imports(path: &Path) -> bool {
    crate::formats::magick_formats::is_engine_picture(path)
}

/// A file and the version of it an answer is about: its path, what it weighed and when it was
/// last written, which is what says it is a file to develop again.
#[derive(Clone, PartialEq, Eq, Hash)]
struct Key {
    path: PathBuf,
    len: u64,
    modified: Option<SystemTime>,
}

fn key_of(path: &Path) -> Key {
    let metadata = std::fs::metadata(path).ok();

    Key {
        path: path.to_path_buf(),
        len: metadata.as_ref().map(|metadata| metadata.len()).unwrap_or(0),
        modified: metadata.and_then(|metadata| metadata.modified().ok()),
    }
}

/// A picture the engine developed, waiting to be drawn: the size the picture says it is, and
/// the bytes it wrote it as.
///
/// It is held in memory and nowhere else, and it is handed over rather than kept: what takes
/// it is the loader building the frame for the hover that asked, and what is kept after that
/// is the frame, in the image cache every other picture is kept in.
pub struct Developed {
    /// The size the picture says it is, read from its own header rather than from the box it
    /// was asked for: what the engine writes is a picture fitted into the room, which is the
    /// size it is shown at only when the file is larger than the display.
    pub width: u32,
    pub height: u32,
    pub png: Vec<u8>,
}

/// The picture the engine developed for each file, waiting to be drawn, or nothing.
static LAST: Lazy<Mutex<Option<Held>>> = Lazy::new(|| Mutex::new(None));

/// What each file's picture turned out to be, by file and version: the size a conversion
/// developed it at.
///
/// It is what the layout measures a preview of one of these from, and what the loader keys
/// the frame it builds by — and it is held for the run rather than on disk, since it is the
/// engine's answer about a file rather than the file's picture: a file saved again is a file
/// to develop again, and a run that has just started knows nothing about anything.
static DEVELOPED_SIZE: Lazy<Mutex<Vec<(Key, (u32, u32))>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// The files the engine would not draw, by file and version.
///
/// It is what keeps a name the list was wrong about — one whose format the engine reads
/// through a delegate this machine has not been given, or a file that is not the picture it
/// is called — from starting an engine on every hover to reach the same answer.
static REFUSED: Lazy<Mutex<Vec<Key>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// The size a picture the engine developed for this file turned out to be, if it has
/// developed one.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let key = key_of(path);

    DEVELOPED_SIZE
        .lock()
        .ok()?
        .iter()
        .find(|(known, _)| *known == key)
        .map(|(_, size)| *size)
}

/// Whether the engine has already turned this file down, for this version of it.
pub fn refused(path: &Path) -> bool {
    let key = key_of(path);

    REFUSED
        .lock()
        .map(|refused| refused.contains(&key))
        .unwrap_or(false)
}

/// Whether a picture the engine developed for this file is in hand and not yet drawn.
pub fn developed(path: &Path) -> bool {
    let key = key_of(path);

    LAST
        .lock()
        .map(|last| last.as_ref().is_some_and(|last| last.key == key))
        .unwrap_or(false)
}

/// Take the picture the engine developed for this file, if it is the one in hand.
///
/// What takes it is the loader building the frame for the hover that asked; a picture
/// developed for another file is left where it is, and one nobody takes is replaced by the
/// next conversion — what is held is one picture, and it is held for the milliseconds
/// between the engine answering and the hover being drawn rather than as a cache.
pub fn take_developed(path: &Path) -> Option<Developed> {
    let key = key_of(path);
    let mut last = LAST.lock().ok()?;

    // A picture developed for another file is not one this caller asked for, and it is not
    // this caller's to drop: what it is waiting for is its own hover, which is a replay away.
    if !last.as_ref().is_some_and(|held| held.key == key) {
        return None;
    }

    last.take().map(|held| Developed {
        width: held.width,
        height: held.height,
        png: held.png,
    })
}

/// The picture in hand, with the key it belongs to: which file and which version of it the
/// picture answers for, so that a hover of another file — or of this one since it was saved
/// again — is not handed a picture of something else.
struct Held {
    key: Key,
    width: u32,
    height: u32,
    png: Vec<u8>,
}

/// Ask the engine for a picture for `path`, at the room the preview may take.
///
/// Nothing is waited on and nothing is answered: the conversion runs on the engine thread
/// below, and what a caller watches for is the message it sends when it is done — or the
/// mark that says the picture is not coming. A hover that asked for one is replayed when
/// the answer lands.
pub fn request(path: &Path, room: (u32, u32), generation: u64) {
    if !imports(path) || !available() || refused(path) {
        return;
    }

    // A picture already in hand is one there is nothing to ask for: it is the hover that
    // asked, waiting to be drawn, and a second conversion beside it would be a launch spent
    // on a picture nothing would look at.
    if developed(path) {
        return;
    }

    // A conversion that has been inside one file for longer than any of them takes has
    // stopped answering, so it is ended here rather than queued behind: what that frees is
    // the thread it was holding and the file that is waiting for it.
    end_hung_conversion();

    // A request for the file already being converted is that request. One that arrives
    // while another waits replaces it, the way the loader's slot does: the newest hover is
    // the one the pointer is on, and a file whose hover has gone is one nobody is waiting
    // for.
    if running_source().as_deref() == Some(path) {
        return;
    }

    let (slot, ready) = &*REQUESTED;
    if let Ok(mut requested) = slot.lock() {
        *requested = Some(Requested {
            path: path.to_path_buf(),
            room,
            generation,
        });
    }
    ready.notify_all();

    start_engine();
}

/// The file the engine is converting now, if it is converting one.
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
/// reads it as a conversion that wrote no picture, remembers the file as one the engine will
/// not draw, and takes up the file behind it. What the caller here is left with is an engine
/// that costs nothing and the answer it would have reached anyway.
fn end_hung_conversion() {
    let hung = RUNNING.lock().ok().and_then(|running| {
        running
            .as_ref()
            .filter(|running| is_hung(running))
            .map(|running| running.pid)
    });

    if let Some(pid) = hung {
        // Verified by name and start time before anything is ended, like every other
        // process this app holds a record of.
        crate::app::engine_processes::terminate_owned(pid);
    }
}

/// Whether a conversion in flight has had its chance: a file the engine has been reading
/// for longer than any of them takes is one it is not going to finish.
fn is_hung(running: &Running) -> bool {
    running.started.elapsed() >= CONVERSION_GIVE_UP
}

/// Start the thread conversions run on, once.
///
/// It is one of the app's threads rather than one per file: what it does between conversions
/// is wait on its own slot, which costs nothing, and what it is asked for is one file at a
/// time because what it is holding is one process.
fn start_engine() {
    if ENGINE_STARTED.swap(true, Ordering::AcqRel) {
        return;
    }

    std::thread::spawn(|| {
        while let Some(requested) = next_request() {
            // Whatever the engine answers — a picture, or a conversion that wrote none — the
            // hover waiting on this file is told either way. A panic is contained here for
            // the reason the loader contains one: one file's failure is that file's, and the
            // thread goes on to the next hover — a thread that died on one picture would take
            // every picture after it with it, in silence.
            let picture = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                develop(&requested.path, requested.room)
            }))
            .unwrap_or(None);

            let ok = picture.is_some();
            if let Some(picture) = picture {
                remember(&picture.key, (picture.width, picture.height));
                hold(picture);
            } else {
                refuse(&key_of(&requested.path));
            }

            crate::ui::preview_window::notify_magick_ready(
                &requested.path,
                requested.generation,
                ok,
            );
        }
    });
}

/// The next file to convert, waiting for one.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made; a lock
/// poisoned by a panic on another thread is still the same slot, and a queue of one is not
/// worth standing down over.
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

        // The wait is bounded rather than endless, and what a bound that runs out is for is
        // a request that may have been made in the meantime.
        requested = match ready.wait_timeout(requested, IDLE_TICK) {
            Ok((requested, _)) => requested,
            Err(poisoned) => poisoned.into_inner().0,
        };
    }
}

/// Develop a file, by running the engine the way a user would.
///
/// What comes back is the picture it wrote to its own output, with the size its header says
/// it is, or nothing where it wrote none — which is the answer for a file it cannot read, a
/// conversion that was given up on, and a process that could not be started at all.
fn develop(path: &Path, room: (u32, u32)) -> Option<Held> {
    let program = ENGINE.as_ref()?;
    let key = key_of(path);

    let png = convert(program, path, room)?;
    let (width, height) = png_dimensions(&png)?;

    Some(Held {
        key,
        width,
        height,
        png,
    })
}

/// Convert `source` with the engine, and answer the bytes it wrote.
///
/// What is asked for is the picture the file holds, turned the way its own header says it
/// should be shown, resized to the box the preview is drawn in and written as an eight-bit
/// PNG on the engine's own output. Every one of those is a decision:
///
/// * What it is written to is `png:-`, its standard output, rather than a path under the
///   app's folder: what is read back is decoded from memory and never written down, so a
///   hover leaves nothing on the disk — no file to prune, no folder to watch, and nothing
///   left behind by a run that ended badly. The bytes are read as they are written, on a
///   thread of its own, because a picture larger than the pipe between the two processes
///   would otherwise be a conversion that blocks forever on a write nobody is reading.
/// * `-delete 1--1` keeps the first picture of a file that holds several, which is what makes
///   the rest of this work at all: an astronomer's `.fits`, a multi-page `.dcx` and an
///   animated `.mng` are more than one picture in one file, and what the engine writes for one
///   of those asked of a single name is a *set* of files — or, on an output that has no names
///   to number, several pictures one after another, which is a picture this side reads as one
///   that is not whole. What is dropped is what the preview would not show: the frames after
///   the first, on a file whose first frame is the only one a hover has anywhere to put. A
///   file that holds one picture is untouched by it, which is every other format here.
/// * `-auto-orient` asks the engine to apply the orientation a camera writes beside its
///   reading. It is the one place in this app where orientation is applied at all — a picture
///   of a phone's is drawn sideways here until that question is answered for the whole of the
///   picture path — and it is asked here because a raw is the format that always carries one:
///   a photograph developed without it is a photograph on its side, and there is no reader of
///   this app's behind this engine to disagree with it.
/// * `-resize {width}x{height}>` is the box the preview is shown in, and the `>` is what
///   makes it a ceiling rather than a size: a picture smaller than the box keeps the size it
///   has — which is the size the preview would be drawn at anyway — and one larger than it is
///   developed at the size it is shown rather than at the size of the sensor, so a
///   forty-megapixel raw costs the preview and not the file.
/// * `-depth 8` is the eight bits to the channel a frame is composed in, so the engine gives
///   up the sixteen an install is usually built for on this side rather than in the decoder,
///   and the bytes it writes are half of what they would otherwise be.
/// * `-strip` drops the metadata — the maker notes, the ICC profile, the thumbnail a raw
///   carries inside it — which nothing here reads and which is most of what a converted raw
///   would otherwise hold.
///
/// The output is named as a PNG explicitly rather than by what a file extension would say:
/// the coder is what is being asked for, and the format is this side's business.
fn convert(program: &Path, source: &Path, room: (u32, u32)) -> Option<Vec<u8>> {
    let geometry = format!("{}x{}>", room.0.max(1), room.1.max(1));
    let mut child = Command::new(program)
        .arg(source)
        .arg("-delete")
        .arg("1--1")
        .arg("-auto-orient")
        .arg("-resize")
        .arg(&geometry)
        .arg("-depth")
        .arg("8")
        .arg("-strip")
        .arg("png:-")
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    // A process this app started is put in the job every other engine is put in, so that
    // whatever ends the app ends it. It is adopted rather than recorded, the way a video's
    // probes are: a conversion lives a second and writing it down would be a record made
    // and struck off again inside the hover that started it, and the job is what answers
    // for one left by a run that ended badly (see `engine_processes`).
    crate::app::engine_processes::adopt(child.id());

    // The picture is read as it is written rather than after the process has ended: a
    // conversion that writes more than the pipe can hold would otherwise be waiting on a
    // reader that is waiting on it. What is read is bounded like every other read of this
    // app — by the ceiling one hover may decode for — so a coder that answers with a picture
    // nobody asked for the size of costs the budget rather than the machine.
    let output = child.stdout.take()?;
    let (written, written_rx) = mpsc::channel();
    std::thread::spawn(move || {
        let mut png = Vec::new();
        let read = output.take(decode_budget_bytes()).read_to_end(&mut png);
        let _ = written.send(read.ok().map(|_| png));
    });

    // What the engine is reading, published for the threads that may decide it has stopped
    // answering while this one waits (see `end_hung_conversion`).
    publish_running(Some(Running {
        source: source.to_path_buf(),
        pid: child.id(),
        started: Instant::now(),
    }));

    let finished = wait(&mut child, CONVERSION_GIVE_UP);

    publish_running(None);

    let png = written_rx.recv_timeout(OUTPUT_WAIT).ok().flatten();

    if !finished {
        return None;
    }

    // Whether the engine wrote a picture is the bytes' own answer rather than its exit code:
    // a coder that failed writes nothing and says so on its error stream, so what came back
    // empty is a conversion that wrote no picture — and what did come back is checked for the
    // header of the format it was asked for before anything is decoded (see `png_dimensions`).
    png.filter(|png| !png.is_empty())
}

/// The size a PNG says it is, read from the header every one of them opens with: the eight
/// bytes of its signature, the length of its first chunk and the name of that chunk, and the
/// width and height that follow — four bytes each, most significant first, as every number in
/// the format is written.
///
/// Nothing is decoded to ask this: the size is what the layout places a preview by, and it is
/// read here rather than by the decoder so that the bytes the engine wrote are decoded once,
/// on the side that builds the frame.
fn png_dimensions(png: &[u8]) -> Option<(u32, u32)> {
    const SIGNATURE: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
    const IHDR_AT: usize = 16;

    if png.get(..SIGNATURE.len())? != SIGNATURE {
        return None;
    }

    let width = u32::from_be_bytes(png.get(IHDR_AT..IHDR_AT + 4)?.try_into().ok()?);
    let height = u32::from_be_bytes(png.get(IHDR_AT + 4..IHDR_AT + 8)?.try_into().ok()?);

    (width > 0 && height > 0).then_some((width, height))
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

/// Remember what a file's picture turned out to be, so that the layout can measure a preview
/// of it without a second conversion.
fn remember(key: &Key, size: (u32, u32)) {
    let Ok(mut developed) = DEVELOPED_SIZE.lock() else {
        return;
    };

    if developed.len() >= ANSWERS_MAX_ENTRIES {
        developed.clear();
    }

    developed.retain(|(known, _)| known != key);
    developed.push((key.clone(), size));
}

/// Hold the picture the engine developed, for the hover that asked to draw.
fn hold(picture: Held) {
    if let Ok(mut last) = LAST.lock() {
        *last = Some(picture);
    }
}

/// Remember that the engine will not draw this file, so that it is not asked twice.
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

/// Say what the engine is converting now, or that it has stopped.
fn publish_running(running: Option<Running>) {
    if let Ok(mut published) = RUNNING.lock() {
        *published = running;
    }
}

/// The conversion in flight: the file being converted, the process converting it, and since
/// when. It is what tells a busy engine from one that has stopped answering, and it is the id
/// a conversion that has to be ended is ended by.
struct Running {
    source: PathBuf,
    pid: u32,
    started: Instant,
}

static RUNNING: Lazy<Mutex<Option<Running>>> = Lazy::new(|| Mutex::new(None));

/// A file to convert, the box to convert it into, and the hover that asked for it.
struct Requested {
    path: PathBuf,
    room: (u32, u32),
    generation: u64,
}

/// The file waiting to be converted, and the signal that one is there: a queue of one, for
/// the reason there is one engine at a time.
static REQUESTED: Lazy<(Mutex<Option<Requested>>, Condvar)> =
    Lazy::new(|| (Mutex::new(None), Condvar::new()));

/// Whether the engine thread has been started. It is one of the app's threads rather than
/// one per file, so it is started once and waits on its slot for the rest of the run.
static ENGINE_STARTED: AtomicBool = AtomicBool::new(false);

#[cfg(test)]
mod tests {
    use super::*;

    /// The engine is never started for a name its own list does not hold: a picture this
    /// app decodes itself, a document an engine of its own draws and a video are all
    /// somebody else's, and asking an image converter about one would cost a launch and
    /// answer nothing.
    #[test]
    fn asks_the_engine_only_about_the_names_it_reads() {
        for name in [
            "photo.png",
            "photo.jpg",
            "photo.heic",
            "texture.dds",
            "render.exr",
            "drawing.svg",
            "report.pdf",
            "notes.txt",
            "font.ttf",
            "clip.mp4",
            "drawing.cdr",
        ] {
            assert!(!imports(Path::new(name)), "`{name}` is not one of its formats");
        }

        for name in ["shot.nef", "shot.CR3", "shot.arw", "shot.dng", "scan.dcm"] {
            assert!(imports(Path::new(name)), "`{name}` is one of its formats");
        }
    }

    /// And a name it does not read is not queued either: nothing about a file is asked of
    /// the engine that its list has not claimed.
    #[test]
    fn queues_nothing_for_a_name_it_does_not_read() {
        let (slot, _) = &*REQUESTED;
        if let Ok(mut requested) = slot.lock() {
            *requested = None;
        }

        request(Path::new("photo.png"), (1920, 1080), 7);

        assert!(
            slot.lock()
                .map(|requested| requested.is_none())
                .unwrap_or(false),
            "the engine is not asked about a name no list of its own holds"
        );
    }

    /// What the engine wrote is read for the size in its own header and for nothing else:
    /// the signature every PNG opens with, and the width and height of the picture that
    /// follows it. What is not a whole picture is not one to place a preview by — a
    /// conversion that was cut short, an error message, a file of another format.
    #[test]
    fn reads_the_size_a_png_says_it_is() {
        let mut png = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
        png.extend_from_slice(&[0, 0, 0, 13]); // the length of the first chunk
        png.extend_from_slice(b"IHDR");
        png.extend_from_slice(&1600u32.to_be_bytes());
        png.extend_from_slice(&1200u32.to_be_bytes());

        assert_eq!(png_dimensions(&png), Some((1600, 1200)));

        // A conversion that wrote nothing, one that wrote something else, one that was cut
        // short inside its header, and a picture that claims no size at all.
        assert_eq!(png_dimensions(b""), None);
        assert_eq!(png_dimensions(b"%PDF-1.4"), None);
        assert_eq!(png_dimensions(&png[..18]), None);

        let mut empty = png.clone();
        empty[16..20].copy_from_slice(&0u32.to_be_bytes());
        assert_eq!(png_dimensions(&empty), None);

        // The size is what the picture says rather than what it was asked for: what the
        // engine writes is a picture fitted into the room, which is the size it is shown at
        // only when the file is larger than the display.
        let mut smaller = png.clone();
        smaller[16..20].copy_from_slice(&800u32.to_be_bytes());
        smaller[20..24].copy_from_slice(&600u32.to_be_bytes());
        assert_eq!(png_dimensions(&smaller), Some((800, 600)));
    }

    /// What the engine developed is held for the hover that asked and handed over once: it
    /// is a picture waiting to be drawn, not a cache, and the second caller — a hover that
    /// was replayed twice — is answered with nothing rather than with the same picture again.
    #[test]
    fn hands_the_picture_it_developed_over_once() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-magick-tests")
            .join("holding");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let raw = folder.join("shot.nef");
        let other = folder.join("other.nef");
        std::fs::write(&raw, b"a raw, of a sort").expect("a written file");
        std::fs::write(&other, b"another").expect("a written file");

        let held = || Held {
            key: key_of(&raw),
            width: 1600,
            height: 1200,
            png: vec![0x89, b'P', b'N', b'G'],
        };

        hold(held());
        assert!(developed(&raw), "the picture is in hand");
        assert!(!developed(&other), "and it is not another file's");

        let taken = take_developed(&raw).expect("the picture");
        assert_eq!((taken.width, taken.height), (1600, 1200));
        assert!(!developed(&raw), "and there is one of it");
        assert!(take_developed(&raw).is_none(), "which cannot be taken twice");

        // A picture developed for one file is left where it is by a caller asking about
        // another: what it is waiting for is its own hover.
        hold(held());
        assert!(take_developed(&other).is_none());
        assert!(developed(&raw), "so the picture it is holding is still there");
        take_developed(&raw);

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// What a file's size and its refusals are remembered by: the file and the version of it
    /// that was read. A file saved again is a file to develop again — and one the engine
    /// would not draw is asked about once, not once per hover.
    #[test]
    fn remembers_an_answer_for_the_version_of_the_file_it_was_read_from() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-magick-tests")
            .join("remembering");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let raw = folder.join("shot.nef");
        std::fs::write(&raw, b"a raw, of a sort").expect("a written file");

        remember(&key_of(&raw), (1600, 1200));
        refuse(&key_of(&raw));
        assert_eq!(dimensions(&raw), Some((1600, 1200)));
        assert!(refused(&raw));

        assert_eq!(dimensions(&folder.join("other.nef")), None);
        assert!(!refused(&folder.join("other.nef")));

        // Saved again: what was known about the file it was is not what it is now.
        std::fs::write(&raw, b"a raw, edited and then some").expect("a written file");
        assert_eq!(dimensions(&raw), None);
        assert!(!refused(&raw));

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// What the engine ends is only what it started, and only once that conversion has had
    /// its chance. A process that stays up stands in for the engine — a test is not going
    /// to make ImageMagick spin on a file — recorded the way the engine is, by image name,
    /// which is the check that keeps an id from being acted on by itself.
    #[test]
    fn ends_a_conversion_only_once_it_has_outrun_the_give_up() {
        let mut engine = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for the engine");
        let pid = engine.id();
        let running = |started: Instant| Running {
            source: PathBuf::from("shot.nef"),
            pid,
            started,
        };

        assert!(
            crate::app::engine_processes::processes_named("ping.exe").contains(&pid),
            "the stand-in runs the image the record below names it by"
        );
        crate::app::engine_processes::record("ping.exe", pid);

        // A conversion that has just started is a file being read, and is left to it.
        publish_running(Some(running(Instant::now())));
        end_hung_conversion();
        assert!(
            crate::app::engine_processes::is_running(pid),
            "a conversion that has just started is not an engine to end"
        );

        // One that has outrun the give-up is an engine that has stopped answering.
        publish_running(Some(running(Instant::now() - CONVERSION_GIVE_UP)));
        end_hung_conversion();
        assert!(
            !crate::app::engine_processes::is_running(pid),
            "the engine a conversion has outrun is ended"
        );

        publish_running(None);
        let _ = engine.wait();
    }

    /// And the bound the engine thread holds over its own conversion: a process that ends
    /// inside it is answered as it always was, and one that outlasts it is ended there
    /// rather than waited on.
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
    /// a picture the engine can read is under a second — and one that has run past the
    /// give-up is.
    #[test]
    fn a_conversion_is_hung_only_once_it_has_outrun_the_give_up() {
        let running = |elapsed: Duration| Running {
            source: PathBuf::from("shot.nef"),
            pid: std::process::id(),
            started: Instant::now() - elapsed,
        };

        assert!(!is_hung(&running(Duration::from_secs(0))));
        assert!(!is_hung(&running(CONVERSION_GIVE_UP - Duration::from_secs(1))));
        assert!(is_hung(&running(CONVERSION_GIVE_UP)));
        assert!(is_hung(&running(CONVERSION_GIVE_UP + Duration::from_secs(30))));
    }

    /// The engine's own program is found by name, and the older name only ever where
    /// ImageMagick is installed under it: a `convert.exe` on the `PATH` is Windows' own
    /// file system converter, and the one that is there on every machine must not be
    /// mistaken for the engine. Whatever this machine has is what the expectations are
    /// written from, and the rule is what is asserted.
    #[test]
    fn finds_the_engine_where_it_installs_and_never_a_system_tool() {
        let found = find_engine();

        if let Some(program) = &found {
            let name = program
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or_default();
            assert!(
                name.eq_ignore_ascii_case(ENGINE_IMAGE) || name.eq_ignore_ascii_case(LEGACY_ENGINE_IMAGE),
                "the engine runs under one of its own two names"
            );

            if name.eq_ignore_ascii_case(LEGACY_ENGINE_IMAGE) {
                let folder = program
                    .parent()
                    .and_then(|folder| folder.file_name())
                    .and_then(|folder| folder.to_str())
                    .unwrap_or_default();
                assert!(
                    folder.starts_with("ImageMagick"),
                    "the older name is only ever ImageMagick's own folder, not a system tool"
                );
            }
        }

        assert_eq!(
            available(),
            found.is_some(),
            "and whether an engine is installed is the answer this app goes by"
        );
    }

    /// What a conversion is asked for, measured against the installed engine: the picture
    /// comes back as bytes with the size its header says, a file the engine cannot read
    /// comes back as nothing at all, and neither leaves anything on the disk. Ignored because
    /// it starts the installed ImageMagick, and run when the command is being looked at:
    /// `cargo test -- --ignored --nocapture engine_conversion_probe`.
    #[test]
    #[ignore = "starts the installed ImageMagick"]
    fn engine_conversion_probe() {
        let Some(program) = ENGINE.as_ref() else {
            println!("no ImageMagick installed: nothing to measure");
            return;
        };
        println!("engine: {}", program.display());

        let folder = std::env::temp_dir().join("rust-hover-preview-magick-probe");
        std::fs::create_dir_all(&folder).expect("a probe folder");

        // A picture to convert, made with the same engine: what is being measured is the
        // conversion rather than the format it is asked about, and every format is asked
        // about the same way.
        let sample = folder.join("probe-source.png");
        let made = Command::new(program)
            .args(["-size", "1200x800", "gradient:red-blue"])
            .arg(format!("png:{}", sample.display()))
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .status();
        println!("a picture to convert: {made:?}");

        let started = Instant::now();
        let first = convert(program, &sample, (1920, 1080));
        println!(
            "first conversion: {:?} bytes in {:?}",
            first.as_ref().map(Vec::len),
            started.elapsed()
        );
        println!(
            "and its size, from its own header: {:?} — the picture fitted into the room",
            first.as_deref().and_then(png_dimensions)
        );

        let started = Instant::now();
        let again = convert(program, &sample, (600, 400));
        println!(
            "again, at a smaller room: {:?} bytes in {:?}, size {:?}",
            again.as_ref().map(Vec::len),
            started.elapsed(),
            again.as_deref().and_then(png_dimensions)
        );

        // And a file the engine cannot read: nothing is written and nothing is held.
        let broken = folder.join("probe.xcf");
        std::fs::write(&broken, b"not a picture at all").expect("a written file");
        println!(
            "a file it cannot read: {:?}",
            convert(program, &broken, (1920, 1080))
        );

        let _ = std::fs::remove_dir_all(&folder);
    }
}

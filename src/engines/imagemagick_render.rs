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
//! shown at. That file is read the way any picture of this app's own is read: decoded
//! into a frame, composited over `image_background`, held in the image cache under the
//! file, its version and the pixel size it was made for. Nothing of the engine reaches the
//! screen: what is drawn is a picture of this app's, which is why a `.nef` follows the
//! picture scale and the picture backdrop rather than a scale and a backdrop of its own,
//! and it is also the whole point of the kind — a developed raw shown at the size it is
//! hovers as a photograph rather than as a thumbnail.
//!
//! Nothing is bundled with this app and nothing is linked against: the engine is the
//! user's own installation, looked for where it installs and beside `config.ini` for a
//! portable copy, and run as the user runs it. What that costs is a launch — a conversion
//! of an ordinary raw is a fraction of a second to a second, and the first one of a
//! session pays for the process itself — so a conversion happens once per picture: the PNG
//! it wrote is kept under [`AppConfig::rendered_dir`], named for the document's path, the
//! version of it that was converted and the box it was converted for, and every hover
//! after the first is a read of that file.
//!
//! A conversion is not one a preview can wait on: the caller is the preview loop, and a
//! loop held inside a launch is a hover that does not come up, a tray that does not answer
//! and a pointer that cannot leave the file it is on. So the engine runs on a thread of
//! its own. What the loop asks is [`request`], which returns at once, and what it waits
//! for is the answer that thread sends back through the preview channel — the same wait an
//! Office page has, in the same box, with the hover replayed when the answer lands. One
//! conversion runs at a time, and the file waiting behind it is the newest one asked for.
//!
//! An engine that has stopped answering is the other half of that. A file a coder cannot
//! decode does not always fail quickly — a damaged raw, or a name in the list whose format
//! the engine reads through a delegate this machine was never given, can leave a
//! conversion running past any bound a hover is owed — so a conversion that has outrun
//! [`CONVERSION_GIVE_UP`] is ended where it stands, and the file it was on is remembered as
//! one the engine will not draw: the launch is paid for once and never again, and the
//! hover that asked for it is answered rather than left spinning.
//!
//! What an engine keeps between hovers is the one thing that is worth it — that is the
//! question every other engine's TTL answers — and ImageMagick is the one engine of them
//! all that has nothing to keep: `magick.exe` reads a file, writes a file and exits, the
//! way `ffprobe` does, and there is no instance to hold open and no seat to hand the next
//! conversion to the way LibreOffice's is handed one. So what `Engine → ImageMagick TTL`
//! bounds is what the engine *did* leave behind: the converted picture. A conversion that
//! has not been hovered for longer than the setting names is let go — the file is dropped,
//! and the next hover of that picture is a conversion — and `0 seconds`, the bottom of the
//! list, is a picture that is converted every time it is hovered. The alternative, a
//! conversion dropped the moment it was read, is the same thing with the setting held at
//! zero, and keeping them is what makes a pointer swept over a folder of raws cost one
//! launch per picture rather than one per hover.

use crate::config::config::{AppConfig, EngineIdle, DEFAULT_MAGICK_IDLE_SECS};
use crate::CONFIG;
use once_cell::sync::Lazy;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Condvar, Mutex};
use std::time::{Duration, Instant};

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

/// The engine's own program, by the name a record carries, and the name the older one is
/// installed as.
///
/// `magick.exe` is ImageMagick 7 and `convert.exe` is ImageMagick 6, and both are looked
/// for because both are still installed. The older name is the one that has to be looked
/// for carefully: Windows ships a `convert.exe` of its own in `System32` — the file
/// system converter — and it is on the `PATH` of every machine, so a search that took the
/// first file of that name would run a system tool with a photograph for an argument (see
/// [`find_engine`]).
const ENGINE_IMAGE: &str = "magick.exe";
const LEGACY_ENGINE_IMAGE: &str = "convert.exe";

/// What a name is remembered as when the engine would not read it: the conversion is not
/// tried again for that file, because a name this app was wrong about — one in the list
/// the engine has no coder for — would otherwise start one on every hover to reach the
/// same answer.
const REFUSED_SUFFIX: &str = "none";

/// How often the engine thread wakes to look at the idle setting, which is what lets a
/// converted picture go while nothing is being asked of it. Nothing else wakes it between
/// conversions, and one look a second is nothing beside the process it is looking at.
const IDLE_TICK: Duration = Duration::from_secs(1);

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
    if let Some(portable) = AppConfig::rendered_dir()
        .and_then(|folder| folder.parent().map(Path::to_path_buf))
        .map(|folder| folder.join("imagemagick"))
    {
        folders.push(portable);
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

/// The engine this machine has, looked for once: the answer is a fact about an
/// installation rather than a question about a hover.
static ENGINE: Lazy<Option<PathBuf>> = Lazy::new(find_engine);

/// Whether an engine is installed to develop these pictures with.
pub fn available() -> bool {
    ENGINE.is_some()
}

/// Whether the engine is the one that develops this file: a name of its own list, or the bytes
/// of a picture it reads under a name no list holds — a `.sgi` renamed to `.bin`, say.
///
/// The question is asked where it is answered for every caller — see
/// `magick_formats::is_engine_picture` — so that a picture one side asks about is a picture the
/// other side will read. Nothing is asked of a file that is what it is called but is not one
/// of the engine's formats, and nothing is asked of one whose bytes are another kind's.
fn imports(path: &Path) -> bool {
    crate::formats::magick_formats::is_engine_picture(path)
}

/// The picture already converted for this file, if there is one.
///
/// Nothing is started and nothing is waited on here. This is the question "has the engine
/// answered yet?" asked of the folder the answers are kept in — the layout measures a picture
/// from what it finds, and the loader draws it — and a hover that is waiting for one keeps
/// asking it until the answer is there (see [`request`]).
pub fn converted(path: &Path) -> Option<PathBuf> {
    let picture = converted_path(path)?;

    usable(&picture).then_some(picture)
}

/// Whether the engine has already turned this file down: the mark a conversion that wrote
/// no picture leaves where a picture would have been.
///
/// It is what a hover waiting on that file reads as "nothing is coming" — the spinner comes
/// down rather than running out its wait — and what keeps the launch from being paid for a
/// second time.
pub fn refused(path: &Path) -> bool {
    converted_path(path).is_some_and(|picture| refused_path(&picture).is_file())
}

/// Ask the engine for a picture for `path`, at the room the preview may take.
///
/// Nothing is waited on and nothing is answered: the conversion runs on the engine thread
/// below, and what a caller watches for is the message it sends when it is done — or the
/// mark that says the picture is not coming. A hover that asked for one is replayed when
/// the answer lands.
pub fn request(path: &Path, room: (u32, u32), generation: u64) {
    if !imports(path) || !available() {
        return;
    }
    if converted(path).is_some() || refused(path) {
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
/// reads it as a conversion that wrote no picture, writes the mark that says so beside
/// where the picture would have been, and takes up the file behind it. What the caller here
/// is left with is an engine that costs nothing and the answer it would have reached
/// anyway.
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

/// How long past its last hover a converted picture is kept, or `None` for a setting that
/// keeps none at all — the `0 seconds` of the tray's `ImageMagick TTL`, which converts
/// every picture every time it is hovered.
fn kept_idle() -> Option<EngineIdle> {
    let idle = CONFIG
        .lock()
        .map(|config| config.magick_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_MAGICK_IDLE_SECS));

    (idle != EngineIdle::Seconds(0)).then_some(idle)
}

/// Say that a converted picture has just been used, which is what the idle setting counts
/// from.
///
/// The moment is the file's own rather than a list in memory: what is kept between hovers
/// is the file, the engine thread that lets it go is not the thread that read it, and a
/// second run of the app is not entitled to drop what this one has just shown. A
/// conversion sets it by being written, and a read sets it here.
pub fn touch(picture: &Path) {
    // Opened for writing rather than for reading: what a moment is set with on Windows is the
    // right to write attributes, and a handle that may only read is not one that carries it.
    // Nothing is written and nothing is truncated — the handle is what the moment is set
    // through, and it is this app's own file.
    let _ = std::fs::File::options()
        .write(true)
        .open(picture)
        .and_then(|file| file.set_modified(std::time::SystemTime::now()));
}

/// Let go of the converted pictures whose last use is older than the setting names.
///
/// It is the engine thread that looks, once a second while it waits for files: nothing else
/// runs between hovers, and a picture that is never let go of is a file left on the disk
/// for a setting that said otherwise. A setting of `0 seconds` — the bottom of the tray's
/// list — lets go of every picture the moment it is looked at, which is a conversion per
/// hover rather than a picture per file. A machine that has never converted anything has no
/// folder to look at, and a look at one is nothing.
fn let_go_if_expired() {
    let Some(folder) = magick_dir() else {
        return;
    };

    let_go_of_expired(&folder, kept_idle());
}

/// What one look at the folder does, with the idle time it is looking against: the two names
/// a conversion writes, and nothing else, since a folder of this app's own that holds
/// something else is not a folder to empty.
///
/// The moment is the file's own rather than a list kept in memory — see [`touch`], which is
/// what a read moves it with — so a picture that is still being hovered is a picture that is
/// still fresh, and a second run of the app is not entitled to drop what this one has just
/// shown.
fn let_go_of_expired(folder: &Path, idle: Option<EngineIdle>) {
    let Ok(entries) = std::fs::read_dir(folder) else {
        return;
    };

    let now = std::time::SystemTime::now();
    for entry in entries.flatten() {
        let path = entry.path();
        if !is_converted_name(&path) {
            continue;
        }

        let keep = idle.is_some_and(|idle| {
            entry
                .metadata()
                .ok()
                .and_then(|metadata| metadata.modified().ok())
                .and_then(|modified| now.duration_since(modified).ok())
                .is_some_and(|age| !idle.has_expired(age))
        });

        if !keep {
            std::fs::remove_file(&path).ok();
        }
    }
}

/// Whether a file in the engine's folder is one of the two a conversion writes: the
/// picture, or the mark left where a picture was not.
fn is_converted_name(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension == "png" || extension == REFUSED_SUFFIX)
}

/// Start the thread conversions run on, once.
///
/// It is one of the app's threads rather than one per file: what it does between
/// conversions is wait on its own slot, which costs nothing, and what it is asked for is
/// one file at a time because what it is holding is one process.
fn start_engine() {
    if ENGINE_STARTED.swap(true, Ordering::AcqRel) {
        return;
    }

    std::thread::spawn(|| {
        while let Some(requested) = next_request() {
            // Whatever the engine answers — a picture, or a conversion that wrote none —
            // it is answered in the folder the pictures are kept in, and the hover waiting
            // on this file is told either way. A panic is contained here for the reason the
            // loader contains one: one file's failure is that file's, and the thread goes on
            // to the next hover — a thread that died on one picture would take every picture
            // after it with it, in silence.
            let ok = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                converted_now(&requested.path, requested.room).is_some()
            }))
            .unwrap_or(false);

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
        // the pictures: this thread and no other lets go of the ones that have gone idle
        // (see `let_go_if_expired`).
        requested = match ready.wait_timeout(requested, IDLE_TICK) {
            Ok((requested, _)) => requested,
            Err(poisoned) => poisoned.into_inner().0,
        };

        let_go_if_expired();
    }
}

/// The picture of a file, converting it if it has not been converted before.
fn converted_now(path: &Path, room: (u32, u32)) -> Option<PathBuf> {
    let program = ENGINE.as_ref()?;
    let picture = converted_path(path)?;
    if usable(&picture) {
        return Some(picture);
    }
    if refused_path(&picture).is_file() {
        return None;
    }

    // One conversion at a time, and what holds the turn is the file: the engine is started
    // per file rather than kept, so the cost of a second conversion beside the first is a
    // second process rather than a corrupt answer — but a folder swept over by a pointer is
    // a queue of conversions nobody is waiting for, and dropping the ones behind the newest
    // is what the slot above already does.
    let _turn = CONVERTING.lock().ok()?;
    if usable(&picture) {
        return Some(picture);
    }
    if refused_path(&picture).is_file() {
        return None;
    }

    if convert(program, path, &picture, room).is_none() {
        // A file the engine would not read is not asked about again: what it answered is
        // written down beside the picture it did not write.
        std::fs::write(refused_path(&picture), b"").ok();
        return None;
    }

    Some(picture)
}

/// Where the converted picture of `path` is kept: named for the file — its path and the
/// version of it that was converted — so a picture saved again is converted again and a
/// picture that has not been is a read.
///
/// The box the picture was written into is deliberately *not* part of the name, which is the
/// one thing about this that differs from the picture cache and from the pages a render engine
/// writes: what the engine is asked for is a ceiling rather than a size — a picture smaller
/// than the room keeps the size it has — so the same file converted for a smaller room is a
/// smaller picture of the same thing rather than a different picture, and a second one written
/// beside it would be kept and read as though it were. What is measured and drawn is the
/// picture's own size either way, so a display that has changed since a conversion costs the
/// size it is drawn at and never the right picture.
fn converted_path(path: &Path) -> Option<PathBuf> {
    let folder = magick_dir()?;
    std::fs::create_dir_all(&folder).ok()?;

    let mut hasher = DefaultHasher::new();
    path.to_string_lossy().to_lowercase().hash(&mut hasher);
    let metadata = std::fs::metadata(path).ok();
    metadata
        .as_ref()
        .map(|metadata| metadata.len())
        .hash(&mut hasher);
    metadata
        .and_then(|metadata| metadata.modified().ok())
        .and_then(|modified| modified.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|since| since.as_secs())
        .hash(&mut hasher);

    Some(folder.join(format!("{:016x}.png", hasher.finish())))
}

/// The folder the converted pictures are kept in, beside the pages the render engine drew.
///
/// It is a folder of its own inside that one rather than the folder itself: what a picture
/// here is named for is the file and the box, what a page there is named for is the
/// document, and the two are pruned by different settings.
fn magick_dir() -> Option<PathBuf> {
    AppConfig::rendered_dir().map(|folder| folder.join("magick"))
}

/// Whether a converted picture is there and is a picture: what is read back is a file this
/// app wrote, so the header is the check and not a second conversion.
///
/// The header is the eight bytes every PNG opens with rather than any picture's: what the
/// engine was asked for is a PNG, and a file that does not carry it is a conversion that
/// was cut short — which is a conversion that wrote no picture.
fn usable(picture: &Path) -> bool {
    let Ok(mut file) = std::fs::File::open(picture) else {
        return false;
    };
    let mut header = [0u8; 8];
    std::io::Read::read_exact(&mut file, &mut header).is_ok()
        && header == [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]
}

fn refused_path(picture: &Path) -> PathBuf {
    picture.with_extension(REFUSED_SUFFIX)
}

/// Convert `source` into `picture`, by running the engine the way a user would.
///
/// What is asked for is the picture the file holds, turned the way its own header says it
/// should be shown, resized to the box the preview is drawn in and written as an eight-bit
/// PNG. Every one of those is a decision:
///
/// * `-delete 1--1` keeps the first picture of a file that holds several, which is what makes
///   the rest of this work at all: an astronomist's `.fits`, a multi-page `.dcx` and an
///   animated `.mng` are more than one picture in one file, and what the engine writes for one
///   of those asked of a single name is a *set* of files — `picture-0.png`, `picture-1.png` —
///   with nothing at the name it was given, which is a conversion this side reads as one that
///   wrote no picture. What is deleted is what the preview would not show: the frames after the
///   first, on a file whose first frame is the only one a hover has anywhere to put. A file
///   that holds one picture is untouched by it, which is every other format here.
/// * `-auto-orient` asks the engine to apply the orientation a camera writes beside its
///   reading. It is the one place in this app where orientation is applied at all — a
///   picture of a phone's is drawn sideways here until that question is answered for the
///   whole of the picture path — and it is asked here because a raw is the format that
///   always carries one: a photograph developed without it is a photograph on its side, and
///   there is no reader of this app's behind this engine to disagree with it.
/// * `-resize {width}x{height}>` is the box the preview is shown in, and the `>` is what
///   makes it a ceiling rather than a size: a picture smaller than the box keeps the size
///   it has — which is the size the preview would be drawn at anyway — and one larger than
///   it is developed at the size it is shown rather than at the size of the sensor, so a
///   forty-megapixel raw costs the preview and not the file.
/// * `-depth 8` is the eight bits to the channel a frame is composed in, so the engine
///   gives up the sixteen an install is usually built for on this side rather than in the
///   decoder, and the file it writes is half the size in the folder and in the read.
/// * `-strip` drops the metadata — the maker notes, the ICC profile, the thumbnail a raw
///   carries inside it — which nothing here reads and which is most of what a converted
///   raw would otherwise hold.
///
/// The output is named as a PNG explicitly rather than left to the extension of a path
/// with a hash for a name: the coder is what is being asked for, and the name is this
/// side's business.
fn convert(program: &Path, source: &Path, picture: &Path, room: (u32, u32)) -> Option<()> {
    // A conversion that was cut short by a give-up leaves a file with no header, which is
    // read as no picture — and what is written over it is written whole, so the file it
    // leaves is the one this run made.
    std::fs::remove_file(picture).ok();

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
        .arg(format!("png:{}", picture.display()))
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    // A process this app started is put in the job every other engine is put in, so that
    // whatever ends the app ends it. It is adopted rather than recorded, the way a video's
    // probes are: a conversion lives a second and writing it down would be a record made
    // and struck off again inside the hover that started it, and the job is what answers
    // for one left by a run that ended badly (see `engine_processes`).
    crate::app::engine_processes::adopt(child.id());

    // What the engine is reading, published for the threads that may decide it has stopped
    // answering while this one waits (see `end_hung_conversion`).
    publish_running(Some(Running {
        source: source.to_path_buf(),
        pid: child.id(),
        started: Instant::now(),
    }));

    let converted = wait(&mut child, CONVERSION_GIVE_UP);

    publish_running(None);

    if !converted {
        std::fs::remove_file(picture).ok();
        return None;
    }

    // Whether the engine wrote a picture is the file's own answer rather than its exit
    // code: a coder that failed writes nothing and says so, and one that wrote a file this
    // app cannot read is the same answer here.
    usable(picture).then_some(())
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

/// Say what the engine is converting now, or that it has stopped.
fn publish_running(running: Option<Running>) {
    if let Ok(mut published) = RUNNING.lock() {
        *published = running;
    }
}

/// The one conversion running at a time.
static CONVERTING: Lazy<Mutex<()>> = Lazy::new(|| Mutex::new(()));

/// The conversion in flight: the file being converted, the process converting it, and
/// since when. It is what tells a busy engine from one that has stopped answering, and it
/// is the id a conversion that has to be ended is ended by.
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

    /// What a converted picture is named for: the file and the version of it that was read.
    /// A file saved again is converted again — which is the whole of what makes the folder a
    /// cache rather than a guess — and the box the picture was written into is deliberately
    /// not part of the name, because what the engine is asked for is a ceiling rather than a
    /// size.
    #[test]
    fn names_a_converted_picture_for_the_file_and_the_version_of_it() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-magick-tests")
            .join("naming");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let file = folder.join("shot.nef");
        std::fs::write(&file, b"a raw, of a sort").expect("a written file");

        let first = converted_path(&file).expect("a path");
        let same = converted_path(&file).expect("a path");
        let other = converted_path(&folder.join("other.nef")).expect("a path");

        assert_eq!(first, same, "the same file is one picture");
        assert_ne!(first, other, "two files are two pictures");
        assert_eq!(
            first.extension().and_then(|extension| extension.to_str()),
            Some("png"),
            "and what it is kept as is the format the engine was asked for"
        );

        std::fs::write(&file, b"a raw, edited").expect("a written file");
        assert_ne!(
            first,
            converted_path(&file).expect("a path"),
            "a file saved again is a file to convert again"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// What a conversion is read back by: the header every PNG opens with, and nothing
    /// else. A file with no header is a conversion that wrote no picture, whatever the
    /// engine's exit code said.
    #[test]
    fn a_converted_picture_is_read_by_its_header() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-magick-tests")
            .join("headers");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let picture = folder.join("picture.png");
        std::fs::write(&picture, [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]).expect("a written picture");
        assert!(usable(&picture), "a PNG is what the engine was asked for");

        let truncated = folder.join("truncated.png");
        std::fs::write(&truncated, b"").expect("a written file");
        assert!(!usable(&truncated), "a conversion cut short wrote no picture");

        let other = folder.join("other.png");
        std::fs::write(&other, b"%PDF-1.4").expect("a written file");
        assert!(!usable(&other), "and neither did one that wrote something else");

        assert!(!usable(&folder.join("absent.png")), "a file that is not there is not one either");

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

    /// The idle setting is what says whether a converted picture is kept at all, and
    /// `0 seconds` — the bottom of the tray's list — is the setting that keeps none: every
    /// hover is a conversion. It is read through the app's own configuration, so what is
    /// asserted here is the shape of the answer rather than the value a machine holds.
    #[test]
    fn a_setting_of_no_seconds_keeps_no_converted_picture() {
        let keep = kept_idle().is_some();
        let configured = CONFIG
            .lock()
            .map(|config| config.magick_idle)
            .expect("the configuration");

        assert_eq!(
            keep,
            configured != EngineIdle::Seconds(0),
            "a picture is kept exactly where the idle setting is not `0 seconds`"
        );
    }

    /// What the idle setting is read off: a conversion sets the moment by being written,
    /// and a read pushes it on, so a picture that is still being hovered is never the one
    /// that is let go.
    #[test]
    fn a_read_moves_the_moment_a_converted_picture_will_be_let_go_at() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-magick-tests")
            .join("idle");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let picture = folder.join("picture.png");
        std::fs::write(&picture, [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]).expect("a written picture");

        let old = std::time::SystemTime::now() - Duration::from_secs(3600);
        let file = std::fs::File::options()
            .write(true)
            .open(&picture)
            .expect("the picture");
        file.set_modified(old).expect("an older moment");

        let modified = || {
            std::fs::metadata(&picture)
                .and_then(|metadata| metadata.modified())
                .expect("the moment it was written")
        };
        assert!(modified() < std::time::SystemTime::now() - Duration::from_secs(1800));

        touch(&picture);

        assert!(
            modified() > std::time::SystemTime::now() - Duration::from_secs(60),
            "a picture that has just been read is a picture that has just been used"
        );

        // And what the sweep reads it as: a picture an hour old is older than ten minutes,
        // and one that has just been read is not.
        assert!(EngineIdle::Seconds(600).has_expired(Duration::from_secs(3600)));
        assert!(!EngineIdle::Seconds(600).has_expired(Duration::from_secs(1)));

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// And the sweep lets go of the two names a conversion writes, and of nothing else: a
    /// folder of this app's own that holds something else is not a folder to empty.
    #[test]
    fn sweeps_only_what_a_conversion_wrote() {
        assert!(is_converted_name(Path::new("6a5b.png")));
        assert!(is_converted_name(Path::new("6a5b.none")));
        assert!(!is_converted_name(Path::new("6a5b.pdf")));
        assert!(!is_converted_name(Path::new("notes.txt")));
        assert!(!is_converted_name(Path::new("magick")));
    }

    /// And what one look at the folder does, which is the whole of what the setting buys: a
    /// picture that nothing has hovered for longer than the setting names is dropped, one
    /// that has just been read is kept, `0 seconds` keeps nothing at all, and `indefinitely`
    /// keeps everything. What is not a conversion's own file is left alone whatever the
    /// setting says.
    #[test]
    fn lets_go_of_the_pictures_a_setting_gives_up_on() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-magick-tests")
            .join("sweep");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // A picture that has just been converted, one that was converted an hour ago, the
        // mark a refusal leaves, and a file this app did not make.
        let fresh = folder.join("fresh.png");
        let stale = folder.join("stale.png");
        let refused = folder.join("refused.none");
        let other = folder.join("notes.pdf");
        let an_hour_ago = std::time::SystemTime::now() - Duration::from_secs(3600);
        for path in [&fresh, &stale, &refused, &other] {
            std::fs::write(path, b"a file").expect("a written file");
        }
        for path in [&stale, &refused, &other] {
            let file = std::fs::File::options()
                .write(true)
                .open(path)
                .expect("the file");
            file.set_modified(an_hour_ago).expect("an older moment");
        }

        // Ten minutes: the hour-old ones go, the one just converted stays.
        let_go_of_expired(&folder, Some(EngineIdle::Seconds(600)));
        assert!(fresh.is_file(), "a picture that has just been converted is kept");
        assert!(!stale.is_file(), "one nothing has hovered for an hour is not");
        assert!(!refused.is_file(), "and neither is the mark beside where one would be");
        assert!(other.is_file(), "a file this app did not make is left alone");

        // Nothing at all: the setting the tray offers as `0 seconds`.
        let_go_of_expired(&folder, None);
        assert!(!fresh.is_file(), "`0 seconds` keeps no converted picture");
        assert!(other.is_file());

        // And indefinitely: what is there stays, whatever its age.
        std::fs::write(&stale, b"a file").expect("a written file");
        let file = std::fs::File::options()
            .write(true)
            .open(&stale)
            .expect("the file");
        file.set_modified(an_hour_ago).expect("an older moment");
        let_go_of_expired(&folder, Some(EngineIdle::Indefinite));
        assert!(stale.is_file(), "an engine kept for the run keeps its pictures");

        let _ = std::fs::remove_dir_all(&folder);
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

    /// What the engine is asked for, measured: the same file converted with nothing kept
    /// and then again with the picture the first one wrote. Ignored because it starts the
    /// installed ImageMagick, and run when the folder or the command is being looked at:
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

        let picture = folder.join("probe.png");
        std::fs::remove_file(&picture).ok();

        let started = Instant::now();
        let first = convert(program, &sample, &picture, (1920, 1080));
        println!(
            "first conversion: {first:?} in {:?} — the engine started and wrote the picture",
            started.elapsed()
        );

        let started = Instant::now();
        let second = convert(program, &sample, &picture, (1920, 1080));
        println!(
            "again: {second:?} in {:?} — the same picture, written over",
            started.elapsed()
        );

        // And a file the engine cannot read: nothing is written, and the mark beside the
        // picture is what says so.
        let broken = folder.join("probe.xcf");
        std::fs::write(&broken, b"not a picture at all").expect("a written file");
        std::fs::remove_file(&picture).ok();
        let absent = convert(program, &broken, &picture, (1920, 1080));
        println!(
            "a file it cannot read: {absent:?}, picture written: {}",
            picture.is_file()
        );

        let _ = std::fs::remove_dir_all(&folder);
    }
}

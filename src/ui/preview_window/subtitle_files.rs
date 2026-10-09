//! A film's subtitle tracks copied out of it once, into small files a hover reads instead of
//! the film.
//!
//! **The fault this exists to remove, measured on the user's own files before any of it was
//! written.** A hover of a film whose subtitles are *embedded* is slow because the `subtitles`
//! filter the player draws them with opens the film a second time and streams the whole
//! container before the first frame is drawn. On a 1.4 GB SubsPlease MKV on a SATA HDD: cold,
//! with the app's own filter arguments, the first frame took 14 904 ms and 1 423 MB was read;
//! the same film with no filter came up in 369 ms off 39 MB. Drawn from a 30 KB `.ass` this app
//! had extracted once, the same cold hover took 492 ms and read 41 MB — the extracted file plus
//! the film's own header. That is the whole of the difference: the filter opens a small file
//! rather than the film, which is also why a sidecar and an MP4 with SRT were already fast and
//! why the small test fixtures never showed the fault at all — they are warm and on an SSD.
//!
//! **One hover pays for the extraction; every hover after it reads the files.** The user chose
//! that the hover which starts the extraction draws *no* subtitles rather than the slow ones —
//! "fast, no subs once", no mid-hover upgrade — so `video_launch::subtitle_filter` answers
//! `None` while an extraction is coming, and `None` again where the one pass has failed: a
//! subtitle that is not there is a film played without subtitles, never a film played from
//! itself. The extraction itself runs on a thread of its own and is never waited on by a
//! preview: what starts it is the launch of the film, and what a later hover reads is the
//! geometry cache entry the thread updates when it is done (see `finish`).
//!
//! **One film is read at a time, and only while its window is showing.** Every extraction is a
//! whole read of a film, so two of them at once is two films read end to end for a preview that
//! can only be showing one of them — and a pass that outlives the window that asked for it is a
//! read nobody is waiting for. So the extraction is a single slot (see `claim`): a request for
//! another film drops the pass that is running, and the roads out of a preview — a hover that
//! ended, a pin that came down or was shown another file — drop it too (see
//! `keep_extraction_for`). A dropped pass kills its child, leaves no half-written folder behind
//! (see `discard`), and is not remembered as a failure, so the next time the film is shown the
//! copy is started again.
//!
//! **A folder is an answer only where the pass that wrote it said it finished.** What a probe
//! reads back is `sub<i>.<ext>` files under a folder carrying the pass's own `finished` mark,
//! and only where those files hold bytes. A pass cut short with its outputs already open, a
//! folder an older build left behind after a failure — those hold zero-byte or half-written
//! files, and naming one to the `subtitles` filter is a filtergraph that fails to build and
//! takes the film's picture down with it, on every showing and for good. Anything without the
//! mark is ignored rather than read, and the next showing deletes the folder and copies the
//! tracks again (see `resolve` and `discard`).
//!
//! **The files are a folder per film, under a key of the film and its version**, which is the
//! shape `document_cache` uses for the pages its engines draw and `image_cache` uses for
//! decoded pictures: `<key>/sub<i>.<ext>` for each subtitle track and `<key>/fonts/` for the
//! container's attached fonts. The key is the same recipe as `document_cache::key` without its
//! engine part — a film is extracted by one program and there is no second answer to tell
//! apart. Everything a film's extract wrote sits under its own key, so a film replaced in
//! place is extracted again rather than drawn with the subtitles of the file it used to be,
//! and a trim can give a whole film's files up at once.
//!
//! **The fonts are dumped in the same pass, by explicit name, and that is measured rather than
//! assumed.** The command the extraction runs is one `ffmpeg` pass over the film — one read of
//! a file that is the slow thing to read — and it writes both the subtitle tracks and the
//! attachments: FFmpeg's own spelling for "dump every attachment under the name it carries"
//! is `-dump_attachment:t ""`, and on FFmpeg 9.0.2 that is refused outright when the
//! attachment's `filename` tag is absolute — `Filename ... is unsafe`, exit code -22, and no
//! subtitle output written either, because the refusal happens while the input is opened. Every
//! fixture to hand has fonts tagged with absolute paths, so the spelling that works is one
//! dump per attached font, named relative to the folder the pass runs in:
//! `-dump_attachment:t:<n> font<n>.ttf`, with the fonts folder as the working directory. That
//! was measured to write both subtitle files and both fonts of a container with two
//! attachments in one pass, byte-identical to the same files dumped on their own.
//!
//! The `fontsdir` the filter is later given is measured the same way. On FFmpeg 9.0.2
//! (`libass` 0.17.5, "font provider directwrite (with GDI)"), a filter spelled
//! `subtitles='<file>':fontsdir='<dir>'` builds its graph and logs `Loading font file
//! '<dir>\font0.ttf'` — the file a container's font was dumped into, opened out of that
//! folder — while the same filter without `fontsdir` logs nothing of it and falls back to the
//! installed faces. Both values are spelled through `video_launch::escape_filter_path`, so the
//! same escaping the sidecar is drawn with protects the drive colon here (see `subtitle_filter`).
//!
//! **What the folder may hold is the `general_disk_cache_mb` budget, and the unit it gives up
//! is a film.** The entries are folders rather than files, so the walk sums each film's files
//! and drops the oldest folder first — by the modified time the folder itself carries, which is
//! what the `document_cache` table reads too. `0` means nothing is kept between hovers, and
//! since the extraction's whole cost is a read of the film, nothing is extracted at all at that
//! size: what such a size asks for is the fast hover without subtitles, which is also what a
//! hover gets before the first extraction has answered.

use super::*;
use std::hash::{DefaultHasher, Hash, Hasher};

/// The folder inside a film's own key that the container's attached fonts are dumped into, and
/// the folder the filter's `fontsdir` points at when it holds any (see `subtitle_filter`).
pub(super) const FONTS_FOLDER: &str = "fonts";

/// The prefix of every subtitle file under a film's key: `sub<i>.<ext>`, where `i` is the
/// subtitle-relative index — the index `-map 0:s:<i>` and the filter's own `si=` count in.
const SUBTITLE_PREFIX: &str = "sub";

/// The file the one pass writes after its last output, and the whole of what makes a folder an
/// answer: what a folder holds is read back only where the pass that wrote it left this beside
/// the files (see `resolve`).
const FINISHED_MARKER: &str = "finished";

/// How long one extraction is given before it is killed and the pass counted as failed.
///
/// The pass reads the whole film, so unlike a probe's ten seconds this is a read of a file of
/// any size on a disk of any speed: the measured 1.4 GB episode extracted in 9 659 ms cold, and
/// a film several times that on a slower volume is the case the bound has to leave room for. It
/// is a bound at all for the file FFmpeg never comes back from — a disk that has gone away, a
/// container it loops on — because what it guards is a thread left running for the rest of the
/// run. Five minutes is far past any extraction that is going to finish and short enough that a
/// stuck one is not a process left for the session.
pub(super) const SUBTITLE_EXTRACTION_TIMEOUT_SECS: u64 = 300;

/// The one extraction that may be running right now, and the two switches its thread reads.
///
/// It is a slot rather than a list of films because a machine must not read two films at once:
/// every extraction is a whole read of a film (see the note at the top of this module), and what
/// a window is showing is the one film worth reading (see `keep_extraction_for`). The `dropped`
/// switch is the whole of a drop: the thread that owns the pass looks at it while it waits,
/// kills the child, and gives up rather than finishing a read nobody is waiting for. The `ended`
/// switch is what a task that displaced this one waits on before starting its own pass, so that
/// at most one film is being read at any moment (see `spawn_subtitle_extraction`).
struct Extraction {
    /// The film whose tracks this extraction is copying.
    path: PathBuf,
    /// Set when this extraction is no longer the one wanted.
    dropped: AtomicBool,
    /// Set by this extraction's thread once it has ended: child reaped, folder settled.
    ended: AtomicBool,
}

/// The slot an extraction lives in, with the rules that hold one task at a time in one place.
///
/// The state machine is asked of a slot rather than of the one global one so that the tests can
/// hold a slot of their own (see `one_extraction_at_a_time_is_the_slot_the_newest_window_takes`):
/// the busy rules are the same either way, and a test that shared the running slot with the rest
/// of the binary would be a test racing every other test that shows or hides a preview.
#[derive(Default)]
struct Slot {
    running: Option<Arc<Extraction>>,
}

impl Slot {
    /// Claim the slot for `path`, answering the task to run and the task it displaced — or
    /// nothing where `path` is exactly what the slot already holds.
    ///
    /// The claim is the whole of "one at a time": a film already being copied is not copied
    /// again, and anything else in the slot is a task from a window that is no longer showing,
    /// which is dropped by the same claim that displaces it.
    fn claim(&mut self, path: &Path) -> Option<(Arc<Extraction>, Option<Arc<Extraction>>)> {
        if let Some(current) = self.running.as_ref() {
            if current.path == path && !current.dropped.load(Ordering::Acquire) {
                return None;
            }

            current.dropped.store(true, Ordering::Release);
        }

        let displaced = self.running.take();
        let extraction = Arc::new(Extraction {
            path: path.to_path_buf(),
            dropped: AtomicBool::new(false),
            ended: AtomicBool::new(false),
        });
        self.running = Some(Arc::clone(&extraction));

        Some((extraction, displaced))
    }

    /// Give up the slot, where it is still this task's.
    ///
    /// A task that was displaced left the slot to the task that displaced it, so what it gives
    /// up here is nothing: the check is what keeps an ending task from clearing a running task's
    /// claim.
    fn release(&mut self, extraction: &Arc<Extraction>) {
        if self
            .running
            .as_ref()
            .is_some_and(|current| Arc::ptr_eq(current, extraction))
        {
            self.running = None;
        }
    }

    /// Keep only the extraction of `keep` running: every other one is dropped.
    ///
    /// `None` is the answer for a preview that has gone — a hover ended, a pin came down — which
    /// asks for no extraction at all. What is dropped is not waited for here: the switch is set,
    /// and the thread that owns the pass kills its child on its next look (see
    /// `EXTRACTION_POLL_MS`).
    fn keep(&mut self, keep: Option<&Path>) {
        let Some(current) = self.running.as_ref() else {
            return;
        };

        if keep != Some(current.path.as_path()) {
            current.dropped.store(true, Ordering::Release);
        }
    }
}

/// The slot the extraction above lives in, and the only one: nothing running is the state no
/// task, one running is the state one film and no others.
static RUNNING: Lazy<Mutex<Slot>> = Lazy::new(|| Mutex::new(Slot::default()));

/// How often a running pass is looked in on while it is waited for: how long a dropped
/// extraction may go on reading its film before it is killed.
///
/// It is the only latency a drop has — the delivery of the switch the road that dropped it set —
/// and it is small because what a drop is *for* is the disk: the sooner the pass is killed, the
/// sooner the film that replaces it is read at full speed.
const EXTRACTION_POLL_MS: u64 = 25;

/// How long a task that displaced another waits for it to end before starting its own pass.
///
/// A displaced task notices within one poll and is never waited on for anything else, so this is
/// a bound on the impossible rather than the expected case: a machine where a killed child
/// cannot be reaped at all. Past it the new pass starts anyway, because a wait with no end is
/// worse than a brief overlap — and what the overlap would cost is nothing like what the pass
/// the switch was set for was going to cost.
const EXTRACTION_HANDOVER_WAIT_SECS: u64 = 5;

/// The folder every film's extracted files are kept in.
///
/// The app's own folder under the temp folder, one folder for everything this kind keeps,
/// exactly as the pages and the developed pictures each have (see `document_cache::temp_folder`).
/// What it holds is derived data — a film whose files have gone is extracted again the next time
/// it is hovered — so the temp folder is the place for it.
fn folder() -> PathBuf {
    // What a test writes is kept apart from what the app writes, and the folder is one folder
    // for the whole test process: a probe whose extraction runs on a thread of its own has to
    // see the files the thread that asked for them wrote. What a test gives up it gives up in a
    // folder of its own, named by the test rather than by the process (see `trim_folder`).
    #[cfg(test)]
    let root = std::env::temp_dir().join("rust-hover-preview-subtitle-tests");
    #[cfg(not(test))]
    let root = crate::engines::document_cache::temp_folder().join("general");

    root
}

/// The name this version of this film's extracted files are kept under.
///
/// The same recipe as `document_cache::key` — the lowercased path, its length and its
/// modification time in nanoseconds, hashed — without the engine part that one carries: a
/// film's subtitles are extracted by one program whatever plays the film afterwards, so there
/// is no second answer for a name to tell apart.
fn key(source: &Path) -> String {
    let mut hasher = DefaultHasher::new();
    source.to_string_lossy().to_lowercase().hash(&mut hasher);

    let metadata = std::fs::metadata(source).ok();
    metadata
        .as_ref()
        .map(|metadata| metadata.len())
        .hash(&mut hasher);
    metadata
        .and_then(|metadata| metadata.modified().ok())
        .and_then(|modified| modified.duration_since(SystemTime::UNIX_EPOCH).ok())
        .map(|since| since.as_nanos())
        .hash(&mut hasher);

    format!("{:016x}", hasher.finish())
}

/// The folder one film's extracted files are kept in: `<key>/sub<i>.<ext>` and `<key>/fonts/`.
pub(super) fn film_folder(path: &Path) -> PathBuf {
    folder().join(key(path))
}

/// The budget the folder is kept within, read from the configuration each time rather than
/// captured: the tray can change it at any moment, and the next extraction and the next trim
/// both ask for it again.
fn limit_bytes() -> u64 {
    CONFIG
        .lock()
        .map(|config| sanitize_general_disk_cache_mb(config.general_disk_cache_mb))
        .unwrap_or(DEFAULT_GENERAL_DISK_CACHE_MB) as u64
        * 1024
        * 1024
}

/// What the folder already holds for this film, read off the names its files carry.
///
/// A run before this one left the same files in the same place, so what a probe finds here is
/// an extraction that has already happened — including one that finished after the hover that
/// asked for it moved on, which is the whole point of keeping them: the next hover of the film
/// draws them rather than being the hover that starts a read of the whole film. `None` is the
/// answer that nothing exists yet, which is what a hover may spawn an extraction for.
///
/// `codecs` is the film's subtitle-relative codec names as the probe read them, and it is what
/// the slots are counted against: a track whose codec has no small form was never written, so
/// its slot stays `None` while its index is still the one `si=` counts to.
///
/// **A folder is an answer only where the pass that wrote it said it finished, and only where
/// the files it names hold bytes.** Both halves are the repair for the same fault, measured on a
/// machine whose cache the builds before this one had written: a pass that was killed, or that
/// failed after FFmpeg had opened its outputs, leaves files behind — twenty-three of the thirty
/// folders on that machine held zero-byte copies, a film with twenty-nine subtitle streams
/// holding twenty-nine empty `sub<i>.ass` — and naming one of those to the `subtitles` filter is
/// a filtergraph that fails to build, which takes the player, and with it the preview, down for
/// good. The marker is written after the pass's last file (see `spawn_subtitle_extraction`), so
/// a folder carrying it is a folder of whole files; anything else is ignored here rather than
/// read, and the next launch deletes it and copies the tracks again.
///
/// **And only the files the filter can draw are read**, whatever a folder happens to hold (see
/// `drawable_copy`): a folder an older build left a PGS copy in is not an answer, and naming one
/// of those to the filter is the same failed graph by another road.
pub(super) fn resolve(path: &Path, codecs: &[String]) -> Option<DerivedSubtitles> {
    // The gate, answered here for a geometry probed while the switch was on and for a
    // pass that outlived it (`finish` asks this after the switch may have moved): the
    // probe zeroes its own answers where the switch is off (see `probe_video_geometry`),
    // and this is the defence in depth beneath it — a folder nobody will draw a
    // subtitle from is not an answer, whatever the disk holds.
    if !subtitles_wanted() {
        return None;
    }

    let folder = film_folder(path);

    if !folder.join(FINISHED_MARKER).is_file() {
        return None;
    }

    let mut tracks: Vec<Option<PathBuf>> = vec![None; codecs.len()];
    let mut any = false;

    if let Ok(read) = std::fs::read_dir(&folder) {
        for entry in read.flatten() {
            if !entry.file_type().is_ok_and(|kind| kind.is_file()) {
                continue;
            }

            // A file with no bytes is not a copy, however it got there: FFmpeg's subtitle
            // demuxers cannot open one, so a filter handed it fails to build its graph and the
            // player goes with it (see the note above).
            if !entry.metadata().is_ok_and(|metadata| metadata.len() > 0) {
                continue;
            }

            let Some(name) = entry.file_name().to_str().map(str::to_owned) else {
                continue;
            };
            if !drawable_copy(&name) {
                continue;
            }

            let Some(index) = name
                .strip_prefix(SUBTITLE_PREFIX)
                .and_then(|rest| rest.split('.').next())
                .and_then(|index| index.parse::<usize>().ok())
            else {
                continue;
            };

            if let Some(slot) = tracks.get_mut(index) {
                *slot = Some(entry.path());
                any = true;
            }
        }
    }

    if !any {
        return None;
    }

    // The fonts folder is named only where it holds something: a container with no attached
    // fonts leaves it empty, and a `fontsdir` pointing at an empty folder is a filter argument
    // that says nothing.
    let fonts = folder.join(FONTS_FOLDER);
    Some(DerivedSubtitles {
        tracks,
        fonts: holds_files(&fonts).then_some(fonts),
    })
}

/// Whether a folder is there and holds at least one file.
fn holds_files(folder: &Path) -> bool {
    std::fs::read_dir(folder).is_ok_and(|read| {
        read.flatten()
            .any(|entry| entry.file_type().is_ok_and(|kind| kind.is_file()))
    })
}

/// Whether a name is one of the small files the `subtitles` filter draws: the three text forms
/// the copy writes out, and nothing else a film's folder may hold.
///
/// It is the filter's own condition read back (see `extension`): libass draws text subtitles and
/// refuses everything else, so a name that is not one of these is not an answer a probe may
/// trust whatever wrote it — a PGS copy an older build left behind above all.
fn drawable_copy(name: &str) -> bool {
    matches!(
        Path::new(name)
            .extension()
            .and_then(|extension| extension.to_str()),
        Some("ass" | "ssa" | "srt")
    )
}

/// Whether the one extraction pass is worth starting for what the probe just read.
///
/// The refusals are the whole of the rule. A sidecar beside the film is already a small file to
/// draw from, so there is nothing to gain. Files already extracted are the answer itself. A
/// film with no subtitle streams has nothing to copy. And a film none of whose tracks has a
/// small form — a container whose only subtitle is a codec this cannot write out — would be a
/// whole read producing nothing, which is the one thing this must never do.
///
/// Whether an extraction is *already running* is not asked here: what says that is the slot the
/// spawn itself claims (see `spawn_subtitle_extraction`), which is what makes two launches of
/// one film race safely.
pub(super) fn extraction_due(
    sidecar: Option<&Path>,
    derived: Option<&DerivedSubtitles>,
    subtitle_streams: usize,
    codecs: &[String],
) -> bool {
    sidecar.is_none()
        && derived.is_none()
        && subtitle_streams > 0
        && codecs.iter().any(|codec| extension(codec).is_some())
}

/// The extension a subtitle track of this codec is copied out as, or nothing for a codec with
/// no form this app draws from a small file.
///
/// The two are the text muxers FFmpeg's own build here has — `ass` and `srt` were each measured
/// to exist and to take a copied stream — and the codecs are the ones that reach them: `ass` and
/// `ssa` are the same muxer, and SubRip text with its siblings is the SubRip one. Anything else
/// is a track whose slot stays empty, which is the film drawn without it rather than the film
/// streamed for it.
///
/// **A codec the `subtitles` filter cannot draw is not copied, and PGS is the case that made
/// that a rule rather than a taste.** The filter draws text subtitles through libass and refuses
/// every other kind outright — its own check is `AV_CODEC_PROP_TEXT_SUB`, logged as "Only text
/// based subtitles are currently supported" — and a filter that refuses the file it was named
/// is a filtergraph that fails to build: the player exits without a window, which is a preview
/// that never appears however long anyone waits. A `sup` copy of a PGS track is exactly such a
/// file, and one cached under a film's key is a film that can never be previewed again: the
/// answer is read by every probe after it, in memory and on disk (see `resolve`).
fn extension(codec: &str) -> Option<&'static str> {
    match codec {
        "ass" | "ssa" => Some("ass"),
        "subrip" | "text" | "mov_text" => Some("srt"),
        _ => None,
    }
}

/// Whether an attachment of this codec is a font the filter could draw with.
///
/// FFprobe names the two font containers it knows — `ttf` and `otf` — and everything else a
/// container may attach (a cover image, a chapter file, a licence) is not something `fontsdir`
/// would ever read.
fn font_codec(codec: &str) -> bool {
    matches!(codec, "ttf" | "otf")
}

/// The argv of the one FFmpeg pass that writes a film's subtitle tracks and fonts out, or
/// nothing when no track of the film has a small form.
///
/// `folder` is the film's own folder under the cache (see `film_folder`), and the tracks are
/// written into it as `sub<i>.<ext>` — the names `resolve` reads back. The fonts are dumped
/// under *relative* names, which is what the command is run with the fonts folder as its
/// working directory to make true (see `spawn_subtitle_extraction`); the names only have to be
/// ones libass can find in the folder, not ones that resemble the fonts, because it reads what
/// is inside them (measured — see the note at the top of this module).
///
/// Pure and kept so: the shape of this command is what the tests assert, since running it in a
/// test would mean a film to run it on.
pub(super) fn extraction_args(
    folder: &Path,
    path: &Path,
    codecs: &[String],
    attachments: &[String],
) -> Option<Vec<String>> {
    let tracks: Vec<(usize, &'static str)> = codecs
        .iter()
        .enumerate()
        .filter_map(|(index, codec)| extension(codec).map(|extension| (index, extension)))
        .collect();

    if tracks.is_empty() {
        return None;
    }

    let mut args: Vec<String> = ["-y", "-v", "error", "-hide_banner"]
        .iter()
        .map(|argument| argument.to_string())
        .collect();

    // The attachments are written first, which is the arrangement FFmpeg's own command has
    // them in and the one the specifier counts in: `t:<n>` is the nth attachment stream of the
    // container, so a container whose first attachment is a cover image dumps its fonts at
    // their own numbers rather than at a number this app made up.
    for (index, codec) in attachments.iter().enumerate() {
        if font_codec(codec) {
            args.push(format!("-dump_attachment:t:{index}"));
            args.push(format!("font{index}.ttf"));
        }
    }

    args.push("-i".to_string());
    args.push(path.to_string_lossy().into_owned());

    for (index, extension) in tracks {
        args.push("-map".to_string());
        args.push(format!("0:s:{index}"));
        args.push("-c:s".to_string());
        args.push("copy".to_string());
        args.push(
            folder
                .join(format!("sub{index}.{extension}"))
                .to_string_lossy()
                .into_owned(),
        );
    }

    Some(args)
}

/// Copy a film's subtitle tracks out of it, on a thread of its own, once — the one extraction
/// this machine may be running (see `RUNNING`).
///
/// Nothing waits for it and nothing is held up by it: the window that asked is answered
/// immediately, without subtitles, and what this pass produces answers every launch after it
/// (see the note at the top of this module). What is asked here is the film a launch is about to
/// show (see `request_extraction`), so a film already being copied is not copied again, and any
/// *other* film's copy is dropped: the window that is on screen is the one whose subtitles are
/// worth reading a film for.
///
/// The displaced pass ends before this one's begins — its thread notices the switch within one
/// poll, kills its child, and sets its `ended` — so a machine never reads two films at once.
///
/// At a budget of nothing there is nothing to gain: the extraction's own product is the only
/// thing the budget holds, so at `0` the files would be written and given up inside the same
/// moment (see `sanitize_general_disk_cache_mb`). Nothing is extracted then, which is the fast
/// hover without subtitles drawn from a film that never pays the read.
pub(super) fn spawn_subtitle_extraction(path: &Path, codecs: &[String], attachments: &[String]) {
    if limit_bytes() == 0 {
        return;
    }

    let folder = film_folder(path);
    let Some(args) = extraction_args(&folder, path, codecs, attachments) else {
        return;
    };

    let Some((extraction, displaced)) = claim(path) else {
        return;
    };

    let path = path.to_path_buf();
    let codecs = codecs.to_vec();
    let fonts = folder.join(FONTS_FOLDER);
    let marker = folder.join(FINISHED_MARKER);

    std::thread::spawn(move || {
        // The task this one displaced ends first, so that the film it was reading is not being
        // read while this pass runs: one film at a time is the whole of what the slot is for.
        if let Some(displaced) = displaced {
            wait_for_end(
                &displaced,
                Duration::from_secs(EXTRACTION_HANDOVER_WAIT_SECS),
            );
        }

        // A task dropped while it was waiting for the one before it never reached the folder,
        // so there is nothing to kill, nothing to remove and nothing to write down.
        if !extraction.dropped.load(Ordering::Acquire) {
            // **A folder a pass before this one left is given up rather than written over.**
            // Nothing reads it (see `resolve`), and what it holds — the zero-byte outputs a pass
            // that failed with them open leaves, a copy of a shape this build no longer writes —
            // is what a name this pass does not happen to write again would keep forever. The
            // pass starts from nothing.
            discard(&path);

            // The fonts folder is created even where the container attached none, because it is
            // the working directory the pass runs in and a pass that writes nothing there leaves
            // it empty rather than absent — `resolve` names it only when it holds files.
            let ran =
                std::fs::create_dir_all(&fonts).is_ok() && copy_out(&args, &fonts, &extraction);
            let dropped = extraction.dropped.load(Ordering::Acquire);

            // **A pass that finished says so before it is read back.** The marker goes down after
            // the last file FFmpeg closed, so a folder that carries it is a folder of whole
            // files; a pass that was killed, or that failed with its outputs already open,
            // leaves files and no marker, and both are given up below.
            let finished = ran && std::fs::write(&marker, b"").is_ok();
            let ok = finished && resolve(&path, &codecs).is_some();

            // A pass that failed or came to nothing leaves no folder at all, marked or not: a
            // half-written answer is one a later probe would read were it not for the marker,
            // and the disk is better off without either (see `discard`).
            if !ok {
                discard(&path);
            }

            // A drop that produced nothing is not a failure and is not written down: the film is
            // asked for again the next time it is shown. A drop that raced a pass which did
            // finish keeps what it wrote — the marker and the answer with it (see `finish`).
            if !dropped || ok {
                finish(&path, &codecs, ok);
            }

            trim_now();
        }

        extraction.ended.store(true, Ordering::Release);
        release(&extraction);
    });
}

/// Ask for a film's subtitles to be copied out, where the probe's answer says there is anything
/// to copy and no copy has answered for it yet.
///
/// It is asked at the launch rather than at the probe, and that placement is the whole of the
/// "only while its window is showing" rule: a probe runs for a file the pointer may only have
/// passed over, while a launch is a film actually put on screen. Every road that shows a film
/// comes through it — a hover installed, a pin shown another file, a seek — so a film whose copy
/// was dropped when the user moved on is asked for again by the next launch that shows it.
pub(super) fn request_extraction(path: &Path, geometry: Option<&VideoGeometry>) {
    // The gate: this ask is the one that spawns the extraction pass — a whole read
    // of the film (see the note at the top of this module) — and it is placed at the
    // launch, so a film whose subtitles are not wanted is played without them rather
    // than paid for on every showing (see `video_subtitles`).
    if !subtitles_wanted() {
        return;
    }

    let Some(geometry) = geometry else {
        return;
    };

    // A pass that has already failed is not asked for again until the next run: the flag is the
    // memory of that failure (see `finish`), and a film whose copy cannot be made is a film
    // played without subtitles rather than one paid for on every launch.
    if geometry.subtitle_extraction_failed {
        return;
    }

    if !extraction_due(
        geometry.sidecar.as_deref(),
        geometry.derived.as_ref(),
        geometry.subtitles.count,
        &geometry.subtitle_codecs,
    ) {
        return;
    }

    spawn_subtitle_extraction(path, &geometry.subtitle_codecs, &geometry.attachment_codecs);
}

/// Keep only the extraction of `keep` running, asking the one slot to drop every other one (see
/// `Slot::keep` and `RUNNING`).
///
/// `pub(crate)` for the same reason `trim_now` is: the tray's `Volume → Video
/// Subtitles` switch reaches it through `preview_window`'s re-export, to drop an
/// extraction running for a film the switch now says no subtitle will be drawn
/// from (see `toggle_video_subtitles`).
pub(crate) fn keep_extraction_for(keep: Option<&Path>) {
    if let Ok(mut slot) = RUNNING.lock() {
        slot.keep(keep);
    }
}

/// Claim the one extraction slot for `path` (see `Slot::claim`).
fn claim(path: &Path) -> Option<(Arc<Extraction>, Option<Arc<Extraction>>)> {
    RUNNING.lock().ok()?.claim(path)
}

/// Give up the slot, where it is still this task's (see `Slot::release`).
fn release(extraction: &Arc<Extraction>) {
    if let Ok(mut slot) = RUNNING.lock() {
        slot.release(extraction);
    }
}

/// Wait for a task that was displaced to end, bounded by `timeout` (see
/// `EXTRACTION_HANDOVER_WAIT_SECS`).
fn wait_for_end(extraction: &Arc<Extraction>, timeout: Duration) {
    let deadline = Instant::now() + timeout;

    while !extraction.ended.load(Ordering::Acquire) {
        if Instant::now() >= deadline {
            return;
        }

        std::thread::sleep(Duration::from_millis(EXTRACTION_POLL_MS));
    }
}

/// Run the one pass, answering whether it finished successfully.
///
/// The process is hidden and adopted like every other one this app starts, and it is waited for
/// in slices rather than plainly so that the slot's switch reaches it: a pass that has been
/// dropped is killed and reaped where it stands, and the answer is that it did not finish. The
/// bound is the other thing the slices make — a pass that runs past
/// `SUBTITLE_EXTRACTION_TIMEOUT_SECS` is killed and reaped, and the answer is that it did not
/// finish, which is what a later launch is told as well (see `finish`).
fn copy_out(args: &[String], workdir: &Path, extraction: &Extraction) -> bool {
    let child = engine_processes::hidden_command("ffmpeg")
        .args(args)
        .current_dir(workdir)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn();

    let Ok(mut child) = child else {
        return false;
    };

    engine_processes::adopt(child.id());

    let handle = HANDLE(child.as_raw_handle());
    let deadline = Instant::now() + Duration::from_secs(SUBTITLE_EXTRACTION_TIMEOUT_SECS);
    let poll = Duration::from_millis(EXTRACTION_POLL_MS);

    loop {
        if extraction.dropped.load(Ordering::Acquire) || Instant::now() >= deadline {
            let _ = child.kill();
            let _ = child.wait();
            return false;
        }

        let slice = poll.min(deadline.saturating_duration_since(Instant::now()));
        let waited = unsafe { WaitForSingleObject(handle, slice.as_millis().max(1) as u32) };

        // Anything but the slice having run out is a process that has ended — or a handle the
        // wait could not read, which is the same answer a moment later — and what is left is
        // what it wrote and the status it finished with.
        if waited != WAIT_TIMEOUT {
            return child.wait().is_ok_and(|status| status.success());
        }
    }
}

/// The extraction of this film is done: hold what it wrote, or hold that it failed.
///
/// Both outcomes are remembered in the geometry cache entry for this file and version, which is
/// the whole of how a later launch hears about either without this thread being waited on — and
/// a success is also the one thing a pinned window is told, because a pin has its film on screen
/// already and begins its player again to draw what just landed (see
/// `reload_pinned_subtitles`). What a success leaves is the derived files themselves, so the
/// next launch draws them. What a failure leaves is the flag that says this film is not asked
/// again: the launch that reads it draws the film *without* subtitles and stays fast (see
/// `video_launch::subtitle_filter`), which is the whole of why a failure is not allowed to
/// matter past itself. A missing entry is not written here — a probe that has not answered yet
/// is answered by its own resolution a moment later, and an entry this thread wrote alone could
/// race the probe's own answer.
fn finish(path: &Path, codecs: &[String], ok: bool) {
    // Read off the disk before the cache lock is taken: the folder is somebody else's state at
    // this moment — the pass has just closed its last file — and the cache is this module's.
    let derived = ok.then(|| resolve(path, codecs)).flatten();

    let key = VideoGeometryKey {
        path: path.to_path_buf(),
        version: file_version(path),
    };
    let mut cache = video_geometry_cache();
    let Some(ProbedGeometry::Measured(geometry)) = cache.get_mut(&key) else {
        return;
    };

    let ready = match derived {
        Some(derived) => {
            geometry.derived = Some(derived);
            true
        }
        None => {
            geometry.subtitle_extraction_failed = true;
            false
        }
    };
    drop(cache);

    // **A pinned window is told, because it is the one thing that can use the answer now.** A
    // pin showing this film has a player that was begun before the copy existed, and the frame
    // it draws next is drawn without the subtitles the user is watching for; the message is
    // what begins that player again, at the second the bar is showing (see
    // `reload_pinned_subtitles`). A hover is told nothing — it was answered without subtitles
    // by the user's own choice, and the hover after it reads what this wrote (see
    // `video_launch::subtitle_filter`) — and a pass that *failed* tells nothing either: the
    // film's own track is not a route this app takes at all (see `subtitle_filter`).
    if ready {
        super::requests::notify_video_subtitles_ready(path);
    }
}

/// Give up a film's folder: what a pass that failed or was dropped leaves behind, and what the
/// next pass gives up before it writes.
///
/// A half-written folder is worse than no folder at all, because a later probe would read it as
/// an answer were it not for the pass's own mark (see `resolve`) — and the folder a build before
/// this one poisoned a film's preview with is exactly that, so it goes here rather than staying
/// on the disk. The removal is best-effort: a folder that cannot be given up is one the trim
/// will reach in the end, and nothing here is worth failing an extraction thread over.
fn discard(path: &Path) {
    let _ = std::fs::remove_dir_all(film_folder(path));
}

/// Trim the folder to the configured budget now, which is what the tray asks for when a smaller
/// size is chosen: what is over the new budget goes at the moment it is set rather than at the
/// next extraction that happens to pass through here.
pub(crate) fn trim_now() {
    trim_folder(&folder(), limit_bytes());
}

/// Drop films' folders, oldest first, until the folder fits inside `limit`.
///
/// The unit given up is one film's whole folder rather than one file of it. What a folder holds
/// is a handful of subtitle files and the fonts beside them, all written by the same pass, so a
/// film is the smallest thing that can be given up without splitting an answer across two
/// states; a file-grain trim is what `document_cache` does, because each of its files is a
/// whole page of a whole document, which one of these is not.
///
/// The order is the modified time each folder carries, which is when its files were last
/// written — the same reading of a folder `document_cache` trims on, and the one that costs a
/// stat per film rather than a clock kept anywhere. A folder that cannot be given up stays
/// counted, so the next trim tries it again.
pub(super) fn trim_folder(folder: &Path, limit: u64) {
    let Ok(read) = std::fs::read_dir(folder) else {
        return;
    };

    let mut films: Vec<(SystemTime, PathBuf, u64)> = Vec::new();
    let mut total = 0u64;

    for entry in read.flatten() {
        if !entry.file_type().is_ok_and(|kind| kind.is_dir()) {
            continue;
        }

        let bytes = folder_bytes(&entry.path());
        total += bytes;

        let left = entry
            .metadata()
            .and_then(|metadata| metadata.modified())
            .unwrap_or_else(|_| SystemTime::now());
        films.push((left, entry.path(), bytes));
    }

    if total <= limit {
        return;
    }

    films.sort_by_key(|(left, _, _)| *left);

    for (_, path, bytes) in films {
        if total <= limit {
            break;
        }

        if std::fs::remove_dir_all(&path).is_ok() {
            total = total.saturating_sub(bytes);
        }
    }
}

/// The bytes every file under a film's folder takes.
fn folder_bytes(folder: &Path) -> u64 {
    let Ok(read) = std::fs::read_dir(folder) else {
        return 0;
    };

    read.flatten()
        .map(|entry| match entry.file_type() {
            Ok(kind) if kind.is_dir() => folder_bytes(&entry.path()),
            Ok(kind) if kind.is_file() => {
                entry.metadata().map(|metadata| metadata.len()).unwrap_or(0)
            }
            _ => 0,
        })
        .sum()
}

/// The one lock every test that stands the `Subtitles` switch holds
/// for its whole run, and the guard that stands it — shared by every test
/// module whose answer the switch turns (this one's own, the `video_launch`
/// sidecar walks and the pin's adopted reload), because the switch is a
/// process-global every test runs beside the others on.
///
/// The tests need the switch at opposite ends — the tests of a folder a pass
/// already extracted need it on, and the gate tests need it off — so a guard
/// that only stood and restored would leave a neighbour mid-assertion
/// answering for the setting another test had just moved (the same reason the
/// slot test holds a `Slot` of its own rather than the `RUNNING` one, see
/// `one_extraction_at_a_time_is_the_slot_the_newest_window_takes`).
#[cfg(test)]
pub(super) static SUBTITLE_SWITCH_TESTS: Mutex<()> = Mutex::new(());

/// The `Volume → Video → Subtitles` switch, stood where a test wants it and put
/// back when the guard is dropped — the house pattern (`CardFontSettings` in
/// `tests::pin_cards`), because the configuration is process-global and these
/// tests run beside others that answer against the machine's own setting.
///
/// The guard holds `SUBTITLE_SWITCH_TESTS` for as long as it stands the switch,
/// which is the whole of the test: two tests that need the switch at different
/// answers never run at once (see the note above the lock).
#[cfg(test)]
pub(super) struct VideoSubtitlesSetting {
    was: bool,
    /// The lock held for the test's whole run, released after the switch is put
    /// back (fields drop after the `Drop` below has run).
    _tests: std::sync::MutexGuard<'static, ()>,
}

#[cfg(test)]
impl VideoSubtitlesSetting {
    /// Stands the switch where a test wants it, holding what the machine's own
    /// setting was to put back and the lock that keeps the switch stood for the
    /// whole of the test.
    pub(super) fn stood_at(wanted: bool) -> Self {
        let _tests = SUBTITLE_SWITCH_TESTS
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = config.video_subtitles;
        config.video_subtitles = wanted;

        Self { was, _tests }
    }
}

#[cfg(test)]
impl Drop for VideoSubtitlesSetting {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.video_subtitles = self.was;
        }
    }
}

#[cfg(test)]
mod tests;

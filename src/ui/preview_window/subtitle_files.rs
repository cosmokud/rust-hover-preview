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
//! `None` while an extraction is coming and only falls back to the film's own track once the
//! one pass has failed. The extraction itself runs on a thread of its own and is never waited
//! on: what starts it is the probe, and what a later hover reads is the geometry cache entry
//! the thread updates when it is done (see `finish`).
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

/// The films whose subtitles are being extracted right now, which is what keeps a film hovered
/// twice in the time one extraction takes from being read twice (see `begin` and `finish`).
static EXTRACTING: Lazy<Mutex<Vec<PathBuf>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Reads in flight this list holds before it is emptied, the same bound and the same reasoning
/// as the measure list's own (see `MEASURING_GAIN_MAX_ENTRIES`).
const EXTRACTING_MAX_ENTRIES: usize = 64;

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
pub(super) fn resolve(path: &Path, codecs: &[String]) -> Option<DerivedSubtitles> {
    let folder = film_folder(path);
    let mut tracks: Vec<Option<PathBuf>> = vec![None; codecs.len()];
    let mut any = false;

    if let Ok(read) = std::fs::read_dir(&folder) {
        for entry in read.flatten() {
            if !entry.file_type().is_ok_and(|kind| kind.is_file()) {
                continue;
            }

            let Some(name) = entry.file_name().to_str().map(str::to_owned) else {
                continue;
            };
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

/// Whether the one extraction pass is worth starting for what the probe just read.
///
/// The refusals are the whole of the rule. A sidecar beside the film is already a small file to
/// draw from, so there is nothing to gain. Files already extracted are the answer itself. A
/// film with no subtitle streams has nothing to copy. And a film none of whose tracks has a
/// small form — a container whose only subtitle is a codec this cannot write out — would be a
/// whole read producing nothing, which is the one thing this must never do.
///
/// Whether an extraction is *already running* is not asked here: what says that is the in-flight
/// guard the spawn itself takes (see `spawn_subtitle_extraction`), which is what makes two
/// probes of one film race safely.
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
/// The three are the muxers FFmpeg's own build here has — `ass`, `srt` and `sup` were each
/// measured to exist and to take a copied stream — and the codecs are the ones that reach
/// them: `ass` and `ssa` are the same muxer, SubRip text and its siblings are the SubRip one,
/// and PGS bitmaps are the raw `sup` stream. Anything else is a track whose slot stays empty,
/// which is the film drawn without it rather than the film streamed for it.
fn extension(codec: &str) -> Option<&'static str> {
    match codec {
        "ass" | "ssa" => Some("ass"),
        "subrip" | "text" | "mov_text" => Some("srt"),
        "hdmv_pgs_subtitle" => Some("sup"),
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

/// Copy a film's subtitle tracks out of it, on a thread of its own, once.
///
/// Nothing waits for it and nothing is held up by it: the hover that asked is answered
/// immediately, without subtitles, and what this pass produces answers every hover after it
/// (see the note at the top of this module). The guard is what keeps one film from being read
/// twice — a hover begun again while the first extraction runs is the ordinary case, since a
/// hover is a second or two and an extraction is more, and the second hover must not pay the
/// same whole read again.
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

    if !begin(path) {
        return;
    }

    let path = path.to_path_buf();
    let codecs = codecs.to_vec();
    let fonts = folder.join(FONTS_FOLDER);

    std::thread::spawn(move || {
        // The fonts folder is created even where the container attached none, because it is the
        // working directory the pass runs in and a pass that writes nothing there leaves it
        // empty rather than absent — `resolve` names it only when it holds files.
        let ran = std::fs::create_dir_all(&fonts).is_ok() && copy_out(&args, &fonts);
        let ok = ran && resolve(&path, &codecs).is_some();

        finish(&path, &codecs, ok);
        trim_now();
    });
}

/// Run the one pass, answering whether it finished successfully.
///
/// The process is hidden and adopted like every other one this app starts, and it is waited for
/// under a bound rather than plainly: a pass that runs past
/// `SUBTITLE_EXTRACTION_TIMEOUT_SECS` is killed and reaped, and the answer is that it did not
/// finish — which is what a later hover is told as well, so a film whose extraction is stuck is
/// not extracted again on every probe (see `finish`).
fn copy_out(args: &[String], workdir: &Path) -> bool {
    let child = engine_processes::hidden_command("ffmpeg")
        .args(args)
        .current_dir(workdir)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn();

    let Ok(child) = child else {
        return false;
    };

    engine_processes::adopt(child.id());

    wait_bounded(child, Duration::from_secs(SUBTITLE_EXTRACTION_TIMEOUT_SECS))
        .is_some_and(|output| output.status.success())
}

/// The extraction of this film is done: hold what it wrote, or hold that it failed.
///
/// Both outcomes are remembered in the geometry cache entry for this file and version, which is
/// the whole of how a later hover hears about either without this thread being waited on. What
/// a success leaves is the derived files themselves, so the next hover draws them. What a
/// failure leaves is the flag that opens the slow route: the film's own embedded track, which
/// draws a frame in fourteen seconds rather than not at all. A missing entry is not written
/// here — a probe that has not answered yet is answered by its own resolution a moment later,
/// and an entry this thread wrote alone could race the probe's own answer.
fn finish(path: &Path, codecs: &[String], ok: bool) {
    end(path);

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

    match derived {
        Some(derived) => geometry.derived = Some(derived),
        None => geometry.subtitle_extraction_failed = true,
    }
}

/// Say that this film's subtitles are being extracted, answering whether one already is.
fn begin(path: &Path) -> bool {
    let Ok(mut extracting) = EXTRACTING.lock() else {
        return false;
    };

    if extracting.iter().any(|running| running == path) {
        return false;
    }

    if extracting.len() >= EXTRACTING_MAX_ENTRIES {
        extracting.clear();
    }

    extracting.push(path.to_path_buf());

    true
}

/// Say that the extraction of this film's subtitles is done with.
fn end(path: &Path) {
    let Ok(mut extracting) = EXTRACTING.lock() else {
        return;
    };

    extracting.retain(|running| running != path);
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

#[cfg(test)]
mod tests;

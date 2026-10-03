//! What a file is: the read that holds everything the question needs, the question itself,
//! the table of names behind it, and what has been answered this run.
//!
//! The bytes are recognised by [`super::signatures`], and each question that recognition
//! is written with is in [`super::matchers`]. Both are asked through here, and both are
//! asked of the same four kilobytes.

use super::matchers::{names_the_container_of_a_raw, palm_ebook_or_program_database};
use super::signatures::SIGNATURES;
use crate::config::config::{AppConfig, PreviewType};
use infer::{MatcherType, Type};
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

/// How many answers are held between hovers. One hover of one file asks for one answer,
/// and a folder is swept a file at a time, so the list is a session's worth of files
/// rather than a scan's: what it bounds is a pointer dragged across a large folder, and
/// what it costs when it is reached is everything held, the way every other cache of this
/// app's shape answers that.
const ANSWERS_MAX_ENTRIES: usize = 512;

/// What a file is, where the answer is not what its name says it is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Content {
    /// The format the file holds — or, where no table named the bytes, the format its name
    /// is written for — is one of this app's kinds. Where the two disagree this is the kind
    /// that should have the file; where they do not, it is the kind the name already gave
    /// it.
    Kind(PreviewType),
    /// The content is a format no kind of this app previews — an executable, an audio
    /// file, a format this app has no reader for, or the other format a guarded name is
    /// written under (see [`GUARDS`]). There is nothing to show, and no engine of this
    /// app's is the engine for it.
    Foreign,
    /// No opinion: nothing was recognized, or what was recognized is what the name
    /// already said, or the name has no extension to disagree with. The name decides, as
    /// it always has.
    Unknown,
}

/// Everything one question about a file's own bytes needs read, and nothing else.
///
/// It is the half of this module that touches the disk, and it is separate from the half that
/// consults the configuration's lists because the two were once one function: a caller that
/// held the process-wide configuration lock to have the lists in hand was holding it across
/// the four kilobytes this reads, and a guard held across a read on a slow volume is a guard
/// every thread of the app — the one pumping this window's own messages included — waits on for
/// as long as the disk takes. Asking for the two halves in that order is the whole of the fix,
/// and it is a shape rather than a rule: [`read`] takes no configuration, so there is nothing
/// for a caller to be holding while it runs.
///
/// The front of the file is read to its front and no further, and the whole window is read here
/// only where the front settles nothing — which is exactly where the tables below would have
/// asked for it, so nothing is read that was not read before (see `head::PROBE_BYTES`). One
/// file's front and its whole window both being wanted by one hover is what made this worth a
/// type: a hover asks this question and five others about the same file, and every one of them
/// was reading the same directory entry for itself (see `crate::formats::head::Facts`).
pub struct Probe {
    path: PathBuf,
    facts: Option<crate::formats::head::Facts>,
    front: Option<std::sync::Arc<crate::formats::head::Head>>,
    /// The whole window, read here only where the front says nothing — the one case where the
    /// tables below cannot answer from the front alone.
    window: Option<std::sync::Arc<crate::formats::head::Head>>,
}

impl Probe {
    /// Everything the question below needs from the disk, read once.
    pub fn read(path: &Path) -> Self {
        #[cfg(test)]
        ENTRY_READS.with(|reads| reads.set(reads.get() + 1));

        // A file whose content is not on this machine is not opened at all, and that question
        // is answered out of the directory entry rather than by trying (see `cloud_files`).
        let Some(facts) = crate::formats::head::Facts::read(path) else {
            let front = crate::formats::head::of(path);
            let window = front.clone().filter(|head| needs_the_window(head));
            return Self {
                path: path.to_path_buf(),
                facts: None,
                front,
                window,
            };
        };

        let front = crate::formats::head::of_with_facts(path, &facts);
        let window = front
            .as_ref()
            .filter(|head| needs_the_window(head))
            .and_then(|_| crate::formats::head::full_with_facts(path, &facts));

        Self {
            path: path.to_path_buf(),
            facts: Some(facts),
            front,
            window,
        }
    }

    /// The same for a caller that has already read the file's own entry, which reads the head
    /// against that entry rather than reading the entry again.
    fn borrowed(path: &Path, facts: &crate::formats::head::Facts) -> Self {
        let front = crate::formats::head::of_with_facts(path, facts);
        let window = front
            .as_ref()
            .filter(|head| needs_the_window(head))
            .and_then(|_| crate::formats::head::full_with_facts(path, facts));

        Self {
            path: path.to_path_buf(),
            facts: None,
            front,
            window,
        }
    }

    /// The file's own directory entry, for a question that needs the version of the file to
    /// hold an answer under — the probe's verdict, the router's video claim — rather than to
    /// read anything.
    pub fn facts(&self) -> Option<&crate::formats::head::Facts> {
        self.facts.as_ref()
    }

    /// The file this probe is of.
    pub fn path(&self) -> &Path {
        self.path.as_path()
    }

    /// Whether reading the file would have to fetch its content first (see `cloud_files`).
    ///
    /// A caller that has read the entry gets the answer out of it; a file with no entry to
    /// read has nothing on this machine either way, and the question is asked of the path as
    /// it always has been.
    pub fn needs_download(&self) -> bool {
        self.facts
            .as_ref()
            .is_none_or(crate::formats::head::Facts::needs_download)
    }
}

// How many times a probe has read a file's directory entry, which is how many `fs::metadata`
// calls a run of hovers paid for reading "what is this file" against the disk.
//
// It is the count the whole of this module's split is for, and it is counted rather than
// argued about: six questions about one file used to make six of them, one per question,
// because nothing was handed down between them. It is per-thread because the tests that read
// it run beside each other, and a count of one thread's probes taken while another's are added
// to it says nothing about either. It is a comment rather than a doc because a `thread_local!`
// block carries no doc of its own.
#[cfg(test)]
thread_local! {
    static ENTRY_READS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// The entry reads this thread has made, and a way to start counting again from nothing.
#[cfg(test)]
pub(crate) fn entry_reads() -> usize {
    ENTRY_READS.with(std::cell::Cell::get)
}

#[cfg(test)]
pub(crate) fn count_entry_reads_from_now() {
    ENTRY_READS.with(|reads| reads.set(0));
}

/// Whether the tables below cannot answer from the front alone, and so the whole window is
/// what they will ask about: a front that names nothing, and a front that names the container
/// a camera raw is written in rather than a picture of its own.
///
/// The condition is the one `read` below uses to decide whether to go on to the window, written
/// once here so that the read above and the question below cannot disagree about it. It is
/// deliberately blind to the lists: a front that names a format no list claims is the one case
/// where the window is wanted and the front did not say so, and it is answered by reading the
/// window under whatever the caller happens to hold — which is `head`'s own cache for every
/// hover after the first.
fn needs_the_window(head: &crate::formats::head::Head) -> bool {
    head.complete()
        || head.front().is_none()
        || head
            .nature()
            .is_some_and(|nature| nature.container_of_a_raw)
}

/// The kind a file's content belongs to, where that is not what its name says.
///
/// Asked where a file's kind is decided — the hook's gate, the loader and the layout —
/// and answered from the cache where the same file has been asked about already in this
/// hover. `Content::Unknown`, where that is the answer, leaves the kind to the list the
/// name is written in.
///
/// It is [`Probe::read`] followed by [`answer`], which is the shape every caller that has more
/// than one question about a file wants: read once, answer as often as asked.
///
/// The configuration is passed in rather than taken here: the lists are what the names a
/// file's bytes answered with are turned into a kind by, and the caller either has them in
/// hand already or takes them once for this question. Nothing below takes that lock, which is
/// what makes a second acquisition on one thread impossible rather than merely unlikely: a
/// lock taken twice hangs the thread that asked, and this question is asked from the hook, the
/// loader, the layout and the engines' own request sides.
pub fn of(path: &Path, config: &AppConfig) -> Content {
    answer(&Probe::read(path), config)
}

/// The question [`of`] asks, with the configuration's lock taken here rather than handed in.
///
/// It is the four engine tiers' own form and it exists because those four are the only questions
/// in this layer with no configuration of their own to be handed one through: they are asked
/// before any hover is installed for the file — by a render tier deciding whether it is the
/// engine that draws it — so there is nothing in hand to hand them.
///
/// The order is the whole of it. The file's own entry is read first, with no guard held, and the
/// lists are taken under the guard only for the lookup that consults them: a guard held across a
/// `content_type::of` is a guard every other thread of the app waits on for as long as the disk
/// takes, and these are asked from the engines' own threads as well as the preview's. A
/// configuration that will not open is answered with the one answer that needs none of them, which
/// is what every caller that reads the lock itself already does (see `mod tests` below, which
/// still holds).
pub(crate) fn of_reaching_config(path: &Path) -> Content {
    // The file's own entry, read before the lock: it is the one thing this needs from the disk,
    // and reading it here rather than inside `of` is what keeps the guard off the file.
    let facts = crate::formats::head::Facts::read(path);

    let Ok(config) = crate::CONFIG.lock() else {
        return Content::Unknown;
    };

    match &facts {
        Some(facts) => of_with_facts(path, facts, &config),
        None => of(path, &config),
    }
}

/// The same question of a file whose own bytes have already been read into a [`Probe`].
///
/// This is the half that consults the lists, and it is a separate function from [`Probe::read`]
/// so that what it reads and what it consults cannot be done in the wrong order: the reading is
/// over before this is called, and a caller holding the configuration to have the lists in hand
/// is holding it over a cache lookup and a set of list comparisons rather than over a
/// `File::open`.
pub fn answer(probe: &Probe, config: &AppConfig) -> Content {
    let key = probe.facts.as_ref().map_or_else(
        || crate::formats::head::key(&probe.path),
        |facts| facts.key().clone(),
    );

    answered(key, probe, config)
}

/// The same question, for a caller that has already read the file's own entry and wants to hand
/// the reading on rather than keep it.
///
/// It is the form the hook asks it in — the entry is in hand there because the hook read it to
/// answer whether the file is a file at all — so nothing is read again and nothing is held that
/// the caller has not already chosen to hold (see `explorer_hook::normalize_media_path`).
pub fn of_with_facts(
    path: &Path,
    facts: &crate::formats::head::Facts,
    config: &AppConfig,
) -> Content {
    answered(facts.key().clone(), &Probe::borrowed(path, facts), config)
}

/// The kind a file's content belongs to, held under the key it was read at.
fn answered(key: crate::formats::head::Key, probe: &Probe, config: &AppConfig) -> Content {
    // The one answer that is not a table's, and it is asked before the cache rather than after
    // it: a container whose own streams a probe found a sound in and no picture is a sound
    // whatever its box is called, and the answer held below was read before that probe ran —
    // an `.mp4` that is really a song would be cached as the video its brand says it is and
    // stay one for the version of the file the probe had already corrected (see
    // `audio_formats::probed_audio_only`).
    //
    // Asked of the key rather than of the path, so that a caller who has already read the
    // directory entry does not have it read again to look this up in a table in memory.
    if crate::formats::audio_formats::probed_audio_only_in(&key) {
        return Content::Kind(PreviewType::Audio);
    }

    if let Ok(answers) = ANSWERS.lock() {
        if let Some(answer) = answers.get(&key) {
            return *answer;
        }
    }

    let answer = read(probe, config);

    if let Ok(mut answers) = ANSWERS.lock() {
        if answers.len() >= ANSWERS_MAX_ENTRIES {
            answers.clear();
        }
        answers.insert(key, answer);
    }

    answer
}

/// What the file holds, read once and not held. The lists the names it answered with are
/// turned into a kind by are the caller's, which is why nothing here takes the configuration
/// lock (see [`of`]).
///
/// The tables are asked in the order they can answer in: the bytes first, through the
/// formats every tool agrees on and then through the ones an engine here reads that no
/// such table carries, and the name the file is under last, for the formats whose own head
/// is nothing either table knows — see [`KIND_BY_NAME`] for what is in that one.
fn read(probe: &Probe, config: &AppConfig) -> Content {
    let path = probe.path.as_path();
    // The front of the file first, and the kind its own form settles by itself: the bytes
    // at the front of a picture are a picture's, whatever the file is called, and nothing
    // further is read of one. A front that settles nothing — a container a camera raw is
    // written in, a file whose signature is further in, a file whose form says nothing at
    // all — is the front the whole window is read for, and that read is `Probe`'s rather than
    // this function's: it is made before the caller had any configuration in hand, which is
    // the whole of what `Probe` is for.
    if let Some(head) = probe.front.as_deref() {
        if let Some(names) = head.front() {
            if !head
                .nature()
                .is_some_and(|nature| nature.container_of_a_raw)
            {
                if let Some(kind) = kind_claiming(names, config) {
                    return Content::Kind(kind);
                }
            }
        }
    }

    let Some(head) = window_of(probe) else {
        return Content::Unknown;
    };
    let probe = head.bytes();

    if let Some(names) = detected_names(probe) {
        return classify(path, names, config);
    }

    // Neither table named the front of the file, so the name it carries is what is left to
    // ask. A file with no name to ask about — one with no extension at all — has nothing
    // here to disagree with either, and is left to the lists.
    crate::formats::text_formats::lookup_extension(path)
        .and_then(|extension| kind_by_name(&extension, probe))
        .unwrap_or(Content::Unknown)
}

/// The window the tables are asked about, read into the probe where its front said the window
/// was what the question needed, and read here where it did not.
///
/// The second of those two is the one case where a file's whole four kilobytes are read while
/// the caller holds the configuration, and it is left there rather than read twice: a front that
/// names a format no list claims is the only way to reach it, the lists are what make that
/// answer "no", and reading the window for every front that names anything would be four
/// kilobytes read for every picture on the machine (see [`Probe::read`]).
fn window_of(probe: &Probe) -> Option<std::sync::Arc<crate::formats::head::Head>> {
    if let Some(window) = probe.window.as_ref() {
        return Some(std::sync::Arc::clone(window));
    }

    match probe.facts.as_ref() {
        Some(facts) => crate::formats::head::full_with_facts(&probe.path, facts),
        None => crate::formats::head::full(&probe.path),
    }
}

/// What the file's own bytes say it is: the names the format is known by to the lists, or
/// nothing where the front of the file is not a format either table names.
pub(super) fn detected_names(probe: &[u8]) -> Option<&'static [&'static str]> {
    // The formats every table agrees on, in the names this app's lists carry.
    if let Some(kind) = infer::get(probe) {
        if let Some(names) = names_of(&kind) {
            return Some(names);
        }
    }

    // And the formats no such table carries, which is what the table below is for.
    SIGNATURES
        .iter()
        .find(|signature| signature.matches(probe))
        .map(|signature| signature.names)
}

/// What one of the common formats means for this app, in the names its lists carry, or
/// nothing where the format is one the name should still decide.
///
/// A name is not the only way a format is written — a JPEG is a `jpg` and a motion JPEG is
/// an `mjpg`, an ISO base media file is an `mp4` and a `mov` — and every one of those
/// names belongs here, because the answer is not "what is this" but "is this what the file
/// is called": a name the file already carries is the two agreeing, and the reply is that
/// there is nothing to override.
///
/// What is *not* answered here is as deliberate as what is. A **box** — a zip, a 7z, a
/// tar — is not a kind, so nothing is answered for one and the name decides, which is what
/// keeps an OpenDocument, an iWork document and an Office package working. Ogg is the one
/// container that is sometimes a video, and which it is is a question for the table below
/// rather than for a guess. A format whose signature is its text — an SVG, an HTML page —
/// is left to the text lists for the same reason.
fn names_of(kind: &Type) -> Option<&'static [&'static str]> {
    match kind.extension() {
        "zip" | "7z" | "rar" | "tar" | "gz" | "tgz" | "bz2" | "xz" | "cab" | "iso" | "dmg"
        | "ogg" => None,

        // A disk image is a box like the ones above, and `qcow2` is the one of those that the
        // common table files under its own heading rather than among them: what a `.qcow2` holds
        // is a file system, and the answer this app has for one is the PeaZip engine's rather than
        // a program's — which is the answer the arm below would otherwise give it, and the reason
        // it is named here rather than left to it (see `peazip_formats`).
        "qcow2" => None,

        // Nothing to show at all: a program. No kind of this app previews one, so a
        // document's name that holds one starts no engine — which is the answer this whole
        // module exists to give for the files that were never documents.
        _ if matches!(kind.matcher_type(), MatcherType::App) => Some(&[]),

        // A sound, in the names the `[audio]` list carries: what the common table knows about
        // one is what every file manager agrees on, and what this app does with the file —
        // which of the two engines plays it, and what the card beside it says — is the list's
        // and the probe's business (see `audio_formats` and `audio_track`).
        //
        // A MIDI sequence is the one the common table names that this app has nothing for:
        // the GS Wavetable synthesizer is a MIDI *out* device and not a decoder, so neither
        // engine here plays a `.mid` under any name, and the answer is the no-preview answer
        // a program gets (see `TODO.md`).
        "midi" => Some(&[]),
        "mp3" => Some(&["mp3"]),
        "m4a" => Some(&["m4a", "m4b"]),
        "opus" => Some(&["opus", "ogg", "oga"]),
        "flac" => Some(&["flac"]),
        "wav" => Some(&["wav", "wave"]),
        "aac" => Some(&["aac"]),
        "aiff" => Some(&["aiff", "aif", "aifc"]),
        "amr" => Some(&["amr", "awb"]),
        "dsf" => Some(&["dsf"]),
        "ape" => Some(&["ape"]),

        "mp4" | "m4v" => Some(&["mp4", "m4v", "mov", "qt", "3gp", "3g2", "f4v"]),
        "mov" => Some(&["mov", "qt", "mp4"]),
        "mkv" => Some(&["mkv", "mk3d", "mka"]),
        "webm" => Some(&["webm", "mkv"]),
        "avi" => Some(&["avi", "divx"]),
        "flv" => Some(&["flv", "f4v"]),
        "wmv" | "asf" => Some(&["wmv", "asf", "dvr-ms"]),
        "mpg" => Some(&["mpg", "mpeg", "vob", "m2v", "m1v", "mpv"]),
        "swf" => Some(&["swf"]),

        "png" => Some(&["png", "apng"]),
        "jpg" => Some(&["jpg", "jpeg", "jpe", "jfif", "mjpg", "mjpeg"]),
        "gif" => Some(&["gif"]),
        "bmp" => Some(&["bmp"]),
        "tif" => Some(&["tif", "tiff"]),
        "webp" => Some(&["webp"]),
        "ico" => Some(&["ico"]),
        "psd" => Some(&["psd", "psb"]),
        "heic" => Some(&["heic", "heif"]),
        "avif" => Some(&["avif"]),
        // A JPEG XL picture is asked of the codec Windows has for it, and an OpenRaster
        // project is read out of the container it keeps its finished picture in: both are
        // names a list here carries, and the common table is what names them.
        "jxl" => Some(&["jxl"]),
        "ora" => Some(&["ora"]),

        "pdf" => Some(&["pdf"]),
        "ttf" => Some(&["ttf"]),
        "otf" => Some(&["otf"]),
        "woff" => Some(&["woff"]),
        "woff2" => Some(&["woff2"]),

        _ => None,
    }
}

/// One of FFmpeg's own formats: a name the video list carries that no signature above
/// names — a container whose header is not one of the common ones, a raw stream, a
/// capture, game or camera format.
const fn video(extension: &'static str) -> (&'static str, PreviewType) {
    (extension, PreviewType::Videos)
}

/// One of the documents the render engine draws: a name the `[libre]` list carries that
/// no signature above names.
const fn engine(extension: &'static str) -> (&'static str, PreviewType) {
    (extension, PreviewType::Libre)
}

/// One of the books the ebook engine reads: a name the `[calibre]` list carries that no
/// signature above names.
const fn ebook(extension: &'static str) -> (&'static str, PreviewType) {
    (extension, PreviewType::Calibre)
}

/// The names no table above answers for, and the kind each of them belongs to.
///
/// What is left once the two tables above have answered for everything they can is short,
/// and every entry in it is a name a head cannot settle:
///
/// * a document that is a container and says nothing about itself inside the probe: an
///   iWork document (`key`, `numbers`, `pages`), a StarOffice 5 one (`sda`, `sdc`, `sdd`,
///   `sdw`), a Visio one (`vsd`), a Publisher one, the Pocket Word one, a Zoner one, a
///   gzip-wrapped AbiWord or Gnumeric document, or a Word for the Macintosh file that is
///   one of the older shapes;
/// * a format whose signature is its own text, which the text lists answer for: a flat
///   OpenDocument, a Visio `.vdx`;
/// * a name without any probe of its own — `mvi` and `mxg`, which FFmpeg reads by name,
///   `psp` and `vw`, which no demuxer of its registers, and `cin`, whose demuxer is gone;
/// * and the two TiVo names, whose chunk headers sit a hundred and twenty-eight kilobytes
///   apart, which is not a question the head of a file can be asked.
///
/// A name here answers with the kind, because a name is the only thing left to answer with,
/// and it answers whether or not a list in `config.ini` still holds it: what a format *is*
/// does not change with a setting, and a user who wants none of these previewed has the
/// kind's own switch in the tray's `Preview Types`.
pub(super) static KIND_BY_NAME: &[(&str, PreviewType)] = &[
    // The video names nothing here can ask the bytes about.
    video("cin"),
    video("flm"),
    video("mvi"),
    video("mxg"),
    video("psp"),
    video("ty"),
    video("ty+"),
    video("vw"),
    // And the documents of the render engine's list that are containers, text, or both.
    engine("abw"),
    engine("fodg"),
    engine("fodp"),
    engine("fodt"),
    engine("gnm"),
    engine("gnumeric"),
    engine("key"),
    engine("mw"),
    engine("numbers"),
    engine("pages"),
    engine("pdb"),
    engine("psw"),
    engine("pub"),
    engine("sda"),
    engine("sdc"),
    engine("sdd"),
    engine("sdw"),
    engine("vdx"),
    engine("vsd"),
    engine("vsdm"),
    engine("vsdx"),
    engine("vstx"),
    engine("zabw"),
    engine("zmf"),
    // And the books of the ebook engine's list whose own head is nothing to ask about: the two
    // texts of the dedicated readers, the SQLite database a `.snb` is, and the zip an `.htmlz` is.
    // The books beside them are answered by their own bytes above, and every one of them is a name
    // this table would otherwise have no opinion about at all.
    ebook("htmlz"),
    ebook("pml"),
    ebook("snb"),
    ebook("tcr"),
];

/// The names in the table above that are more than one format's, and the question asked of
/// the bytes before any of them is answered.
///
/// A name is here where the format its engine reads shares its spelling with a format
/// nothing here previews, and where the two can be told apart from the front of the file.
/// What such a name answers is the format the engine reads, or nothing at all: a file that
/// is the other format is a file no kind of this app previews, and starting an engine that
/// can only turn it down is the one thing this is for.
pub(super) static GUARDS: &[(&str, Guard)] = &[("pdb", palm_ebook_or_program_database)];

/// The question a name in [`GUARDS`] is asked of the front of a file, answered in the same
/// three terms the tables answer in: the kind that previews what the bytes hold, nothing
/// where no kind of this app previews them, and no opinion where the bytes do not say.
type Guard = fn(&[u8]) -> Content;

/// What the name a file carries answers for it, where the tables above named nothing.
///
/// It is the last of the three questions this module asks of a file: the bytes first,
/// through the two signature tables, and the name after them, through the table above —
/// whose own exception, a name that is two formats, is asked of the bytes in between (see
/// [`GUARDS`]). `None` where the name is not one the table holds, which is where a file's
/// kind is left to the lists, as it has always been.
pub(super) fn kind_by_name(extension: &str, probe: &[u8]) -> Option<Content> {
    if let Some((_, guard)) = GUARDS
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(extension))
    {
        return Some(guard(probe));
    }

    KIND_BY_NAME
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(extension))
        .map(|(_, kind)| Content::Kind(*kind))
}

/// What the format the bytes named means for this app, for one file.
pub(super) fn classify(path: &Path, names: &[&str], config: &AppConfig) -> Content {
    // A file with no extension at all has no name for the content to disagree with, and
    // one whose own name is among the names the content answered with is the two agreeing:
    // there is nothing to override either way.
    let Some(own) = crate::formats::text_formats::lookup_extension(path) else {
        return Content::Unknown;
    };

    if names.iter().any(|name| name.eq_ignore_ascii_case(&own)) {
        return Content::Unknown;
    }

    // A name the ImageMagick engine is asked about is answered as that kind where the bytes
    // named the container its format is written in rather than a picture of its own: a
    // camera raw that is a TIFF is a `.nef`, a `.cr2`, an `.arw`, a `.dng` or a `.pef`, and
    // the name is the camera's own answer about what is inside the box (see
    // `names_the_container_of_a_raw`). Every other name is answered below, which is where a
    // `.tif` is taken for the picture it is.
    if names_the_container_of_a_raw(names) && crate::formats::lists::MAGICK.claims(path, config) {
        return Content::Kind(PreviewType::Magick);
    }

    match kind_claiming(names, config) {
        Some(kind) => Content::Kind(kind),
        None => Content::Foreign,
    }
}

/// The kind of this app that claims one of the names the content answered with, asked in the
/// one order every classification of a file is asked in.
///
/// Each name is asked of the same table the hook, the loader and the layout ask, by a name of
/// this module's own making (`content.<extension>`) that stands for the file and is never
/// opened: a list reads the extension and nothing else, so the file the caller holds is not
/// touched a second time. It is the name half of that question rather than the file half —
/// the two extensions a video list shares with the text lists are settled by content, and the
/// content is exactly what answered here, so there is nothing left to read (see
/// `routing::kind_of_name`). Nothing claimed by any list is a format this app does not
/// preview — see [`Content::Foreign`].
fn kind_claiming(names: &[&str], config: &AppConfig) -> Option<PreviewType> {
    names.iter().find_map(|name| {
        let named = PathBuf::from(format!("content.{name}"));

        crate::formats::routing::kind_of_name(&named, config)
    })
}

/// What has been asked and answered, for the hovers of this run. A file saved again is a
/// file to ask about again, and what says so is what says it everywhere else in this app:
/// when it was last written and what it weighed (see `head::key`).
static ANSWERS: Lazy<Mutex<HashMap<crate::formats::head::Key, Content>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

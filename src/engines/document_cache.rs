//! The page a document's engine drew, kept on disk.
//!
//! Three engines draw pages for this app — Microsoft Office, driven through automation, the
//! LibreOffice beside it, and the ebook engine that converts a book no reader here opens — and all
//! of them are asked the same question: has a page been drawn for this version of this document, by
//! this engine? The answer is the page itself, and it is kept here rather than in memory because
//! producing one costs far more than reading one back: an Office start and an export, a conversion
//! of one to three seconds, or an ebook converted at the length of the book, against a read of a
//! file a few hundred kilobytes long. What is kept is worth keeping across a restart for the same
//! reason, which is why it is a folder rather than a structure that goes when the process does.
//!
//! Everything about a page is in its name. `<key>.pdf`, `<key>.png` and `<key>.bmp` are the
//! pages the engines produce — a Word or Excel export and a conversion are PDFs, a slide is a
//! PNG, a workbook whose Excel cannot export a page at all is a picture of its used range, and a
//! book is a PDF too — and `<key>.none` is the mark an engine leaves for a document it would not
//! draw.
//!
//! The key is the document, the version of it, and the engine that drew it. The engine is
//! part of it because the two do not draw the same page: LibreOffice's rendering of a `.docx`
//! is not Word's, so a page one engine drew is never handed back as the other's work when the
//! choice between them changes (see `office_formats::page_engine`) — and neither of them is the
//! page the ebook engine writes for a book, which is a third name again.
//!
//! What is kept is bounded by `document_cache_mb`, least recently used first — and "used" is
//! the page file's own timestamp, set every time a page is read, so what is given up first is
//! what has not been looked at for longest rather than what was converted first.

use crate::config::config::{sanitize_document_cache_mb, AppConfig, DEFAULT_DOCUMENT_CACHE_MB};
use crate::readers::pdf_preview;
use crate::CONFIG;
use directories::BaseDirs;
use once_cell::sync::Lazy;
use std::collections::hash_map::DefaultHasher;
use std::collections::HashMap;
use std::ffi::OsStr;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{Duration, SystemTime};

/// Which of the files a page can be kept as.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum PageKind {
    /// A one-page PDF, which the PDF renderer draws: what Word and Excel export, and what the
    /// engine beside Office writes for every document it converts.
    Pdf,
    /// A PNG of the first slide.
    Png,
    /// A bitmap of a workbook's used range, for a machine whose Excel cannot export a page at
    /// all (see `office_render::render_excel`).
    Bmp,
}

impl PageKind {
    pub(crate) fn extension(self) -> &'static str {
        match self {
            Self::Pdf => "pdf",
            Self::Png => "png",
            Self::Bmp => "bmp",
        }
    }

    /// The kind a name names, for the sizes `prune` reads a folder by.
    fn from_extension(extension: &OsStr) -> Option<Self> {
        match extension.to_str()? {
            "pdf" => Some(Self::Pdf),
            "png" => Some(Self::Png),
            "bmp" => Some(Self::Bmp),
            _ => None,
        }
    }
}

/// The suffix the mark an engine leaves for a document it would not draw carries, where a page
/// would have been put.
const REFUSED_SUFFIX: &str = "none";

/// How long a document an engine would not draw is left alone.
///
/// A mark kept for good would be a lie about a document that is only undrawable for now — one
/// that was locked, or half-copied, or read while an engine was still installing a filter — so
/// what it holds is that the engine would not draw it *then*. It is the same two minutes the
/// render tier leaves a document it failed on (see `office_render`), and it is the mark's own
/// timestamp rather than a clock written down: a refusal is the one thing here that is not
/// read, so its file is the one timestamp a hit never touches.
const REFUSAL_TTL: Duration = Duration::from_secs(120);

/// How many pages' sizes are remembered at once. A size is a pair of numbers and nothing else,
/// so the number is high and what it costs is nothing; it is a ceiling so that a session spent
/// hovering every document in a folder cannot grow the table without end.
const SIZE_MEMO_MAX_ENTRIES: usize = 256;

/// A page that is there: the file it is kept as, and which of the files it is.
#[derive(Clone, Debug)]
pub(crate) struct Page {
    pub(crate) path: PathBuf,
    pub(crate) kind: PageKind,
}

/// A page's own size, remembered by the document it was drawn from.
static SIZES: Lazy<Mutex<HashMap<String, (u32, u32)>>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// The page a hover is waiting for: the one stored a moment ago and not drawn yet, and the one
/// page a trim never gives up (see `prune`). It is held with the document it was drawn from, which
/// is what lets a hover that ends say which page to release without knowing which engine drew it
/// (see `hover_ended`).
static HELD: Lazy<Mutex<Option<(String, PathBuf)>>> = Lazy::new(|| Mutex::new(None));

/// Let go of the page a hover was waiting for, where the caller's own question says it is that page.
///
/// One helper for the three callers that release it, because what is held is a pair of things and
/// what each of them knows differs: a page given up is asked about by its own name, while a hover
/// that has ended knows the document it was drawn from and not the engine that drew it (see
/// `hover_ended`).
fn release_held(is_the_page: impl Fn(&str, &Path) -> bool) {
    if let Ok(mut held) = HELD.lock() {
        if held
            .as_ref()
            .is_some_and(|(key, source)| is_the_page(key, source))
        {
            *held = None;
        }
    }
}

/// The app's own folder under the temp folder: where a render writes the file an engine can
/// only answer with, and where the pages are kept beside it.
///
/// A page is a cache and nothing else — a document whose page has gone is drawn again the next
/// time it is hovered — so what is kept is allowed to be *gone* without anything having to
/// notice, and the folder this app is given for files of that kind is the temp folder.
pub(crate) fn temp_folder() -> PathBuf {
    std::env::temp_dir().join("rust-hover-preview")
}

/// The folder the pages are kept in.
///
/// The one thing that must not clear it is this app's own start: what a run before this one
/// left behind in the temp folder is cleaned there, and the pages are not that (see
/// `discard_leftovers`).
fn folder() -> Option<PathBuf> {
    // What a test writes is kept apart from what the app writes, and the folder is one folder for
    // the whole test process: a probe whose render runs on a thread of its own has to see the page
    // the thread that asked for it kept. What a test gives up it gives up in a folder of its own,
    // named by the test rather than by the process (see `prune_folder`).
    #[cfg(test)]
    let root = std::env::temp_dir().join("rust-hover-preview-document-tests");
    #[cfg(not(test))]
    let root = temp_folder().join("document");

    Some(root)
}

/// The name the page drawn for this version of this document by this engine is kept under.
///
/// The engine is named by the name it writes its own setting under — `microsoft_office`,
/// `libreoffice`, or the ebook engine's own — because it is the engine's identity rather than its
/// kind: a name is what goes into the key, and two engines that draw the same document draw
/// different pages.
fn key(source: &Path, engine: &str) -> String {
    let mut hasher = DefaultHasher::new();
    source.to_string_lossy().to_lowercase().hash(&mut hasher);
    engine.hash(&mut hasher);

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

fn path_of(folder: &Path, key: &str, kind: PageKind) -> PathBuf {
    folder.join(format!("{key}.{}", kind.extension()))
}

fn refused_path(folder: &Path, key: &str) -> PathBuf {
    folder.join(format!("{key}.{REFUSED_SUFFIX}"))
}

/// The page kept for this version of this document, if there is one.
///
/// Reading a page is what using it is: its timestamp is set to now, which is how the budget
/// tells what has not been looked at from what has (see `prune`).
///
/// What the file *is* is not asked here. Whether a page can actually be read out of one is a
/// question for the side that draws it — a side whose threads are multithreaded apartments and
/// may talk to the PDF engine — and this is asked from threads that are not (see
/// `office_preview`).
pub(crate) fn page(source: &Path, engine: &str) -> Option<Page> {
    let folder = folder()?;
    let key = key(source, engine);

    for kind in [PageKind::Pdf, PageKind::Png, PageKind::Bmp] {
        let path = path_of(&folder, &key, kind);
        if std::fs::metadata(&path).is_ok() {
            touch(&path);
            return Some(Page { path, kind });
        }
    }

    None
}

/// Say that the page has just been read, which is what the order pages are given up in is made
/// of. A timestamp that cannot be written is not worth answering for: the page is still there,
/// and all that is lost is where it sits in that order.
fn touch(page: &Path) {
    let Ok(file) = std::fs::File::options().write(true).open(page) else {
        return;
    };

    let _ = file.set_modified(SystemTime::now());
}

/// Keep a page an engine has just drawn, replacing whatever was kept for the same version of
/// the same document.
pub(crate) fn store(source: &Path, engine: &str, kind: PageKind, bytes: &[u8]) -> Option<Page> {
    let folder = folder()?;
    std::fs::create_dir_all(&folder).ok()?;

    let key = key(source, engine);

    // Whatever was kept under this name is not the page any more: this render drew over it, and
    // a page of another kind left beside it would be found first (see `page`).
    for other in [PageKind::Pdf, PageKind::Png, PageKind::Bmp] {
        let _ = std::fs::remove_file(path_of(&folder, &key, other));
    }
    let _ = std::fs::remove_file(refused_path(&folder, &key));

    let path = path_of(&folder, &key, kind);
    write_whole(&path, bytes)?;

    // A size remembered for the page this one replaces is not this page's size: what the
    // document's own file has not changed about is the page, and a slide re-exported at another
    // width is exactly that.
    if let Ok(mut sizes) = SIZES.lock() {
        sizes.remove(&key);
    }

    // The page just drawn is the one a hover is waiting for, so it is held whatever the budget
    // says: giving it up here would leave the spinner with nothing to replace it. It is
    // released when that hover ends.
    if let Ok(mut held) = HELD.lock() {
        *held = Some((key, source.to_path_buf()));
    }

    prune();

    Some(Page { path, kind })
}

/// Write a file whole, under a name nothing reads until it is complete: a run ended inside the
/// write leaves half a page under a name of its own rather than a page a reader would open and
/// draw from.
fn write_whole(path: &Path, bytes: &[u8]) -> Option<()> {
    let writing = path.with_extension("writing");
    std::fs::write(&writing, bytes).ok()?;
    std::fs::rename(&writing, path).ok()?;

    Some(())
}

/// Give up the page kept for this version of this document: bytes that cannot be drawn — one a
/// render cut short, or one something else corrupted — are not a page, and what this gets is
/// the document drawn again rather than a preview that blinks away every time it is hovered.
pub(crate) fn forget(source: &Path, engine: &str) {
    let Some(folder) = folder() else {
        return;
    };

    let key = key(source, engine);

    for kind in [PageKind::Pdf, PageKind::Png, PageKind::Bmp] {
        let _ = std::fs::remove_file(path_of(&folder, &key, kind));
    }

    if let Ok(mut sizes) = SIZES.lock() {
        sizes.remove(&key);
    }
    release_held(|held, _| held == key);
}

/// Write down that an engine would not draw this version of this document, where a page would
/// have been put.
pub(crate) fn refuse(source: &Path, engine: &str) {
    let Some(folder) = folder() else {
        return;
    };
    if std::fs::create_dir_all(&folder).is_err() {
        return;
    }

    let key = key(source, engine);
    let _ = std::fs::write(refused_path(&folder, &key), b"");

    release_held(|held, _| held == key);

    prune();
}

/// Whether an engine has already turned this version of the document down.
///
/// A mark past its age is not an answer any more: what it says is that an engine would not draw
/// the document once, and a document that was locked, or half-copied, or read while a filter
/// was still being installed is one worth asking about again. The mark is dropped here rather
/// than left for a trim to find, because the ask itself is what says it is stale.
pub(crate) fn refused(source: &Path, engine: &str) -> bool {
    let Some(folder) = folder() else {
        return false;
    };

    let marker = refused_path(&folder, &key(source, engine));
    let Ok(metadata) = std::fs::metadata(&marker) else {
        return false;
    };

    let age = metadata
        .modified()
        .ok()
        .and_then(|refused| SystemTime::now().duration_since(refused).ok());
    if age.is_some_and(|age| age < REFUSAL_TTL) {
        return true;
    }

    let _ = std::fs::remove_file(&marker);

    false
}

/// The hover a page was drawn for is over: it is no longer being waited on, so at a budget of
/// nothing it goes now rather than lingering until the next render happens to make room.
///
/// What is released is the page of *this* document, whichever engine drew it, because the side
/// that ends a hover knows the document and not which of the engines was asked for a page.
pub(crate) fn hover_ended(source: &Path) {
    release_held(|_, held| held == source);

    prune();
}

/// Trim the cache to the configured size now, which is what the tray asks for when a smaller
/// size is chosen: what is over the new budget goes at the moment it is set rather than at the
/// next render that happens to pass through here.
pub(crate) fn trim_now() {
    prune();
}

/// The size of the page kept for this document: what a layout places a preview by, and what a
/// wider display compares a slide's export against.
///
/// It is read from the page the first time it is asked for and remembered from then on — the
/// same question is asked more than once per hover, and reading a size out of a PDF is a parse
/// of the document. What is remembered is keyed by the document, its version and its engine
/// rather than by the file the page is kept as: that file's timestamp says when the page was
/// last *used*, which is not a version to read a size against (see `touch`).
pub(crate) fn size(source: &Path, engine: &str) -> Option<(u32, u32)> {
    let key = key(source, engine);

    if let Ok(sizes) = SIZES.lock() {
        if let Some(size) = sizes.get(&key) {
            return Some(*size);
        }
    }

    let size = read_size(&page(source, engine)?)?;
    if let Ok(mut sizes) = SIZES.lock() {
        if sizes.len() >= SIZE_MEMO_MAX_ENTRIES {
            sizes.clear();
        }
        sizes.insert(key, size);
    }

    Some(size)
}

/// A page's own size, read from the file it is kept as.
fn read_size(page: &Page) -> Option<(u32, u32)> {
    match page.kind {
        PageKind::Pdf => pdf_preview::page_dimensions(&page.path),
        // A slide's PNG and a workbook's bitmap are both read the way any picture is: the
        // header answers with the size and the decode is not paid for.
        PageKind::Png | PageKind::Bmp => image::ImageReader::open(&page.path)
            .ok()?
            .with_guessed_format()
            .ok()?
            .into_dimensions()
            .ok(),
    }
}

/// Drop pages, least recently used first, until the folder fits inside the configured budget.
///
/// What a page costs is the size of its file, and a page that cannot be drawn is worth the same
/// as one that can until it is read: a mark left for a document an engine would not draw is
/// counted and given up with the rest, since it is the same folder's room either way.
///
/// The page a hover is waiting for is never one of them — it was stored a moment ago and has
/// not been drawn yet. At a budget of nothing it is the only page left, which is what that size
/// means: nothing kept *between* hovers rather than no page shown at all.
fn prune() {
    let limit = CONFIG
        .lock()
        .map(|config| sanitize_document_cache_mb(config.document_cache_mb))
        .unwrap_or(DEFAULT_DOCUMENT_CACHE_MB) as u64
        * 1024
        * 1024;

    let held = HELD.lock().ok().and_then(|held| held.clone());

    if let Some(folder) = folder() {
        prune_folder(&folder, limit, held.as_ref().map(|(key, _)| key.as_str()));
    }
}

/// The same, for a folder the caller names rather than the one pages are kept in — a test's own,
/// since what a trim gives up is given up for good.
fn prune_folder(folder: &Path, limit: u64, held: Option<&str>) {
    let Ok(entries) = std::fs::read_dir(folder) else {
        return;
    };

    let mut pages: Vec<(SystemTime, u64, PathBuf)> = entries
        .flatten()
        .filter(|entry| {
            entry
                .file_type()
                .map(|kind| kind.is_file())
                .unwrap_or(false)
        })
        .filter(|entry| {
            entry
                .path()
                .extension()
                .and_then(PageKind::from_extension)
                .is_some()
                || entry
                    .path()
                    .extension()
                    .is_some_and(|extension| extension == REFUSED_SUFFIX)
        })
        .filter(|entry| {
            let stem = entry
                .path()
                .file_stem()
                .and_then(OsStr::to_str)
                .map(str::to_string);
            stem.as_deref() != held
        })
        .filter_map(|entry| {
            let metadata = entry.metadata().ok()?;
            Some((metadata.modified().ok()?, metadata.len(), entry.path()))
        })
        .collect();

    let mut total: u64 = pages.iter().map(|(_, size, _)| *size).sum();
    if total <= limit {
        return;
    }

    pages.sort_by_key(|(modified, _, _)| *modified);
    for (_, size, path) in pages {
        if total <= limit {
            break;
        }
        if std::fs::remove_file(&path).is_ok() {
            total = total.saturating_sub(size);
        }
    }
}

/// Delete what earlier versions cached on disk, and whatever a render that was ended mid-flight
/// left behind in the temp folder.
///
/// The pages themselves are not that, and are left where they are — what is cleared there is
/// the files beside the folder they are kept in, since a scratch file a previous run never got
/// to read is not one the next attempt is allowed to find.
pub(crate) fn discard_leftovers() {
    // Only the `office` folder is removed — in the installed layout the app's own executable
    // and uninstaller live beside it in that same directory, and a page cache is not worth
    // risking them for.
    if let Some(dirs) = BaseDirs::new() {
        let _ = std::fs::remove_dir_all(dirs.cache_dir().join("rust-hover-preview").join("office"));
    }

    // The `rendered` folder beside `config.ini` is where a version before this one kept every
    // converted page of every document, and the profile its engine ran under. Nothing reads it
    // any more: what is drawn now is kept under the pages' own folder, and a folder that is no
    // longer written to is one to be rid of rather than left for a user to find on their disk.
    if let Some(folder) = AppConfig::config_path()
        .and_then(|path| path.parent().map(|folder| folder.join("rendered")))
    {
        let _ = std::fs::remove_dir_all(folder);
    }

    let Ok(entries) = std::fs::read_dir(temp_folder()) else {
        return;
    };
    for entry in entries.flatten() {
        if entry
            .file_type()
            .map(|kind| kind.is_file())
            .unwrap_or(false)
        {
            let _ = std::fs::remove_file(entry.path());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::config::OfficeEngine;

    /// The name the Office application's pages are kept under, and the render engine's beside it.
    /// They are the names the callers pass rather than the enum they come from, because the cache
    /// is keyed by the engine's own name — which is what lets an engine that is neither of the two
    /// keep pages here as well (see `key`).
    fn office() -> &'static str {
        OfficeEngine::MicrosoftOffice.as_str()
    }

    fn libre() -> &'static str {
        OfficeEngine::LibreOffice.as_str()
    }

    /// A document under the tests' own folder, named for the test that asks for it: the folder
    /// is one folder for the whole process (see `folder`), so the names are what keeps one
    /// test's document from being another's.
    fn document(name: &str) -> PathBuf {
        let folder = std::env::temp_dir().join("rust-hover-preview-document-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let path = folder.join(name);
        std::fs::write(&path, b"a document").expect("a written document");

        path
    }

    /// A page is named after the document, the version of it, and the engine that drew it.
    #[test]
    fn keys_a_page_by_the_document_its_version_and_its_engine() {
        let source = document("keyed.docx");

        let first = key(&source, office());
        assert_eq!(first, key(&source, office()));
        assert_ne!(
            first,
            key(&source, libre()),
            "the engine that drew it is part of the page's name"
        );

        std::fs::write(&source, b"a longer document").expect("a rewritten document");
        assert_ne!(
            first,
            key(&source, office()),
            "a document saved again is another page"
        );

        let _ = std::fs::remove_file(&source);
    }

    /// A page is read back by the name it was kept under, and one that has been given up — or
    /// one another engine drew — is not found.
    #[test]
    fn keeps_reads_and_forgets_a_page() {
        let source = document("held.docx");

        assert!(
            page(&source, office()).is_none(),
            "nothing is kept before anything is drawn"
        );

        let kept =
            store(&source, office(), PageKind::Pdf, b"%PDF-1.7 a page").expect("a kept page");
        assert_eq!(kept.kind, PageKind::Pdf);
        assert_eq!(
            std::fs::read(&kept.path).expect("a page to read"),
            b"%PDF-1.7 a page"
        );

        let found = page(&source, office()).expect("the page just kept");
        assert_eq!(found.path, kept.path);

        assert!(
            page(&source, libre()).is_none(),
            "the engine that did not draw it has no page for the document"
        );

        forget(&source, office());
        assert!(
            page(&source, office()).is_none(),
            "a page given up is not one to be found"
        );

        let _ = std::fs::remove_file(&source);
    }

    /// A mark left for a document an engine would not draw is read as one, and a mark past its
    /// age is dropped rather than believed for good.
    #[test]
    fn reads_a_refusal_until_it_ages_out() {
        let source = document("refused.cdr");

        assert!(
            !refused(&source, libre()),
            "nothing is marked before an engine has been asked"
        );

        refuse(&source, libre());
        assert!(refused(&source, libre()));
        assert!(
            !refused(&source, office()),
            "one engine's refusal is not another's"
        );

        // A mark older than the age it is kept for is not an answer any more, and reading it
        // is what drops it: the document is asked about again.
        let marker = folder()
            .expect("a folder")
            .join(format!("{}.{REFUSED_SUFFIX}", key(&source, libre())));
        std::fs::File::options()
            .write(true)
            .open(&marker)
            .expect("the mark")
            .set_modified(SystemTime::now() - REFUSAL_TTL - Duration::from_secs(1))
            .expect("an older mark");

        assert!(!refused(&source, libre()));
        assert!(!marker.is_file(), "and the mark is gone with the answer");

        let _ = std::fs::remove_file(&source);
    }

    /// What a budget gives up is the page that has not been read for longest, and the page a
    /// hover is waiting for is never one of them — which is what a budget of nothing means:
    /// nothing kept *between* hovers rather than no page shown at all.
    #[test]
    fn gives_up_the_page_that_was_read_longest_ago() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-document-tests")
            .join("pruned");
        let _ = std::fs::remove_dir_all(&folder);
        std::fs::create_dir_all(&folder).expect("a test folder");

        let older = folder.join("older.pdf");
        let newer = folder.join("newer.pdf");
        for (path, seconds) in [(&older, 60), (&newer, 30)] {
            write_whole(path, b"%PDF-1.7 a page").expect("a written page");
            std::fs::File::options()
                .write(true)
                .open(path)
                .expect("a page")
                .set_modified(SystemTime::now() - Duration::from_secs(seconds))
                .expect("a page read that long ago");
        }

        // Room for one page, and it is the one that was read most recently.
        let one_page = std::fs::metadata(&newer).expect("a page").len();
        prune_folder(&folder, one_page, None);
        assert!(!older.is_file(), "the page read longest ago goes");
        assert!(newer.is_file(), "and the one read most recently stays");

        // A budget of nothing gives that one up too — unless a hover is waiting for it, which
        // is the one page a trim never touches.
        prune_folder(&folder, 0, Some("newer"));
        assert!(
            newer.is_file(),
            "the page a hover is waiting for is not given up"
        );

        prune_folder(&folder, 0, None);
        assert!(!newer.is_file(), "nothing is kept between hovers");
    }

    /// The size a page is placed by is its own, read from the file it is kept as — and it is
    /// read once: the same question is asked more than once per hover.
    #[test]
    fn reads_and_remembers_a_pages_size() {
        let source = document("sized.png");

        assert_eq!(size(&source, office()), None);

        store(&source, office(), PageKind::Png, &png_bytes(64, 32)).expect("a kept page");

        assert_eq!(size(&source, office()), Some((64, 32)));
        assert_eq!(
            size(&source, libre()),
            None,
            "and by nothing about another engine"
        );

        let _ = std::fs::remove_file(&source);
    }

    /// A PNG of one colour, written the way a slide's export is.
    fn png_bytes(width: u32, height: u32) -> Vec<u8> {
        let mut written = std::io::Cursor::new(Vec::new());
        image::RgbaImage::from_pixel(width, height, image::Rgba([10, 20, 30, 255]))
            .write_to(&mut written, image::ImageFormat::Png)
            .expect("a written slide");

        written.into_inner()
    }
}

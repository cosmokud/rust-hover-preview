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
    // is what drops it: the document is asked about again. The age is the table's own —
    // what is on the disk is the mark, and when it was left is what this side remembers
    // having written (see `FolderIndex`).
    let refused_key = key(&source, libre());
    let marker = folder()
        .expect("a folder")
        .join(format!("{refused_key}.{REFUSED_SUFFIX}"));

    with_index(&folder().expect("a folder"), |index| {
        let Some(IndexEntry::Refused { left }) = index.entries.get_mut(&refused_key) else {
            panic!("the mark just left");
        };

        *left = SystemTime::now() - REFUSAL_TTL - Duration::from_secs(1);
    });

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

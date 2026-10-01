//! The names of the books this app previews as a page of one.
//!
//! Two things are what a person means by "a book" here, and this app already reads both: the PDF,
//! which it draws from page one of the file with a reader of its own, and the comic, which is a
//! container of pictures — a `.cbz` is a zip of plates, a `.cbr` a rar of them, and a `.cbc` is a
//! zip of the pages of several comics under a folder each, with a `comics.txt` naming them — whose
//! first page is a picture inside it. What the two have in common is the whole of why they are one
//! list: what a hover on either shows is a page, at the book kind's own scale, over the book kind's
//! backdrop, under the book kind's switch. See `pdf_preview` for the first and `comic_preview` for
//! the second, and `PreviewType::Ebook` for the kind they share.
//!
//! The list of names this answers for is a row of `crate::formats::lists` — the one table
//! every kind's list is a row of, and the one place a list is written down, the built-in
//! entries and the older lists this app shipped and then changed included.

use crate::config::config::PreviewType;
use crate::formats::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The book list itself, which is a row of `crate::formats::lists` like every other.
///
/// It is named here as well because the PDF reader's own tests still ask for it by this path, and
/// until they are moved across the table keeps the name they reach for.
#[cfg(test)]
pub use crate::formats::lists::DEFAULT_EBOOK_EXTENSIONS;

/// Whether the configured list claims `path`.
pub fn matches_ebook_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Whether the configured list claims `path`, without asking whether these previews are switched
/// on. The gate is asked beside it by the hook, the way every other kind's is.
pub fn is_ebook_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_ebook_list(path, &config.ebook_extensions))
        .unwrap_or(false)
}

/// Whether a preview may be shown for `path`: the file the configured list claims, and the `Ebook`
/// gate in the tray's `Preview Types` submenu.
///
/// Both halves ask it where a kind can be switched off under a preview that is already on screen:
/// a hover is not sent for a kind that is off, and the layout places nothing for a file whose kind
/// is off, which is how a preview of that kind comes down when the switch does. See
/// `PreviewType::enabled`.
pub fn is_ebook_preview(path: &Path) -> bool {
    is_ebook_file(path) && PreviewType::Ebook.enabled()
}

/// Whether the list holds this file's name *and* that name is one of the PDF's own spellings: the
/// names the PDF reader is the reader for.
///
/// It is the name half of the question `pdf_preview::is_pdf_file` asks — that function adds the one
/// content answer of its own, a drawing saved as a PDF — and it is asked of a list the caller
/// already holds, which is the form every caller that has the configuration in hand asks it in.
///
/// The hook resolves a hover with the configuration held, and every list it consults it consults
/// through the caller's own copy — `matches_ebook_list(path, &config.ebook_extensions)` and the
/// gates beside it — because a question that went and read the configuration again would wait on
/// a lock the same thread is already holding, and a lock taken twice on one thread is a deadlock
/// (see `explorer_hook::is_media_file`).
pub fn matches_page_name(path: &Path, extensions: &[String]) -> bool {
    page_spelling(path) && matches_ebook_list(path, extensions)
}

/// Whether the file is named with one of the three spellings the PDF reader draws: `pdf`, `pdfa`
/// and `epdf`. Whether such a name is still a book is the list's answer, asked beside this one.
fn page_spelling(path: &Path) -> bool {
    matches!(
        crate::formats::text_formats::lookup_extension(path).as_deref(),
        Some("pdf") | Some("pdfa") | Some("epdf")
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    /// What the list holds is what no other list of this app's names, and no name of it is a
    /// book the engine beside this app has to convert.
    #[test]
    fn holds_no_name_another_kind_already_reads() {
        let list = crate::formats::text_formats::sanitize_extension_list(
            crate::formats::lists::DEFAULT_EBOOK_EXTENSIONS,
        );
        let config = crate::config::config::AppConfig::default();

        let claimed_elsewhere = |path: &Path| {
            crate::formats::image_formats::matches_image_list(path, &config.image_extensions)
                || crate::formats::vector_formats::matches_vector_list(
                    path,
                    &config.vector_extensions,
                )
                || crate::formats::design_formats::matches_design_list(
                    path,
                    &config.design_extensions,
                )
                || crate::formats::font_formats::matches_font_list(path, &config.font_extensions)
                || crate::formats::libre_formats::matches_libre_list(path, &config.libre_extensions)
                || crate::formats::magick_formats::matches_magick_list(
                    path,
                    &config.magick_extensions,
                )
                || crate::formats::video_formats::matches_any_video_list(path, &config)
                || crate::formats::archive_formats::matches_archive_list(
                    path,
                    &config.archive_extensions,
                )
                || crate::formats::peazip_formats::matches_peazip_list(
                    path,
                    &config.peazip_extensions,
                )
                || crate::formats::calibre_formats::matches_calibre_list(
                    path,
                    &config.calibre_extensions,
                )
                || crate::formats::office_formats::matches_office_list(
                    path,
                    &config.office_extensions,
                )
                || crate::formats::text_formats::matches_text_lists(
                    path,
                    &config.text_extensions,
                    &config.text_names,
                )
        };

        // `cbz` is the name this list took off the archive list, so it is asked about twice: it
        // must be a comic here, and it must not be an archive there, or the same file would be two
        // answers again.
        for name in [
            "photo.png",
            "drawing.svg",
            "report.docx",
            "notes.txt",
            "bundle.zip",
        ] {
            let path = Path::new(name);
            assert!(
                claimed_elsewhere(path),
                "`{name}` is read by another kind, so this expectation is written the wrong way round"
            );
            assert!(
                !matches_ebook_list(path, &list),
                "`{name}` is read by another kind, so this list does not carry it"
            );
        }

        let comic = Path::new("chapter.cbz");
        assert!(
            matches_ebook_list(comic, &list),
            "a comic is this list's own name"
        );
        assert!(
            !crate::formats::archive_formats::matches_archive_list(
                comic,
                &config.archive_extensions
            ),
            "and it is not the archive list's any more: one name, one answer"
        );

        // And the two books that are not this app's at all: a Microsoft Reader book is a page the
        // ebook engine draws, and a compiled help file is the listing engine's, which answers
        // immediately — the reason it is not the engine's is the wait, not the format.
        for (name, claimed_by_calibre) in [("book.lit", true), ("help.chm", false)] {
            let path = Path::new(name);
            assert_eq!(
                crate::formats::calibre_formats::matches_calibre_list(
                    path,
                    &config.calibre_extensions
                ),
                claimed_by_calibre,
                "`{name}`: whether the ebook engine is the one asked about it"
            );
            assert!(
                crate::formats::peazip_formats::matches_peazip_list(
                    path,
                    &config.peazip_extensions
                ) != claimed_by_calibre,
                "`{name}` and the listing engine are the other way round"
            );
            assert!(
                !matches_ebook_list(path, &list),
                "`{name}` is not read here, so it is not a book of this kind either"
            );
        }
    }

    /// And what it does hold is the PDF's three spellings and the three comic containers.
    #[test]
    fn holds_the_pages_and_the_comics() {
        let list = crate::formats::text_formats::sanitize_extension_list(
            crate::formats::lists::DEFAULT_EBOOK_EXTENSIONS,
        );

        for name in ["book.pdf", "report.pdfa", "encapsulated.epdf"] {
            let path = Path::new(name);
            assert!(
                matches_ebook_list(path, &list),
                "`{name}` is a page the PDF reader draws"
            );
            assert!(matches_page_name(path, &list), "and it is read as one");
        }

        for name in ["chapter.cbz", "chapter.cbr", "collection.cbc"] {
            let path = Path::new(name);
            assert!(
                matches_ebook_list(path, &list),
                "`{name}` is a comic this app reads itself"
            );
            assert!(
                !matches_page_name(path, &list),
                "and the PDF reader is not asked about it: a comic is the other half of the list"
            );
        }

        // A name that is in neither half is neither: what the list does not hold is not a book.
        assert!(!matches_ebook_list(Path::new("book.epub"), &list));
        assert!(!matches_page_name(Path::new("book.epub"), &list));
    }

    /// The PDF's names are asked of the list a caller hands in — the form the hook asks them in,
    /// since it resolves a hover with the configuration held and a question that read the
    /// configuration again would be a lock taken twice on that thread (see `matches_page_name`).
    #[test]
    fn asks_for_the_pdf_names_of_the_list_it_is_given() {
        let list = crate::formats::text_formats::sanitize_extension_list(
            crate::formats::lists::DEFAULT_EBOOK_EXTENSIONS,
        );

        for name in ["book.pdf", "report.pdfa", "encapsulated.epdf"] {
            assert!(
                matches_page_name(Path::new(name), &list),
                "`{name}` is a page the PDF reader draws"
            );
        }

        assert!(
            !matches_page_name(Path::new("chapter.cbz"), &list),
            "a comic is the other half of this list and not a page of it"
        );
        assert!(
            !matches_page_name(Path::new("book.epub"), &list),
            "and a book the ebook engine converts is not one either"
        );
        assert!(
            !matches_page_name(
                Path::new("book.pdf"),
                &crate::formats::text_formats::sanitize_extension_list("cbz,cbr,cbc")
            ),
            "a name taken out of the list is a name this app stops drawing"
        );
    }

    /// A name a user adds is a comic if the file it names is a container of pictures, and the
    /// reader that decides that is the comic one — nothing here claims a name is a comic.
    #[test]
    fn a_name_added_by_hand_is_asked_of_the_reader_that_reads_it() {
        let added = crate::formats::text_formats::sanitize_extension_list(
            "cbc,cbr,cbz,myalbum,pdf,pdfa,epdf",
        );

        assert!(
            matches_ebook_list(Path::new("book.myalbum"), &added),
            "a name the list holds and the PDF reader has no spelling for is the comic reader's to answer"
        );
        assert!(
            !matches_page_name(Path::new("book.myalbum"), &added),
            "and the PDF reader is not asked about it"
        );
        assert!(matches_page_name(Path::new("book.pdf"), &added));
    }

    /// A name that is not a bare extension is dropped rather than matched against, the way every
    /// other list of this app's answers a hand-edited entry.
    #[test]
    fn reads_a_list_of_bare_extensions() {
        let extensions =
            crate::formats::text_formats::sanitize_extension_list(" .PDF , cbz,,cbr,pdf,pdfa");

        assert_eq!(extensions, vec!["pdf", "cbz", "cbr", "pdfa"]);
    }
}

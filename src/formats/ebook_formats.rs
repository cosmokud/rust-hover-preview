//! The names of the books this app previews as a page of one, and the three spellings of the one
//! half of them this side draws itself.
//!
//! Two things are what a person means by "a book" here, and this app already reads both: the PDF,
//! which it draws from page one of the file with a reader of its own, and the comic, which is a
//! container of pictures — a `.cbz` is a zip of plates, a `.cbr` a rar of them, and a `.cbc` is a
//! zip of the pages of several comics under a folder each, with a `comics.txt` naming them — whose
//! first page is a picture inside it. What the two have in common is the whole of why they are one
//! list: what a hover on either shows is a page, at the book kind's own scale, over the book
//! kind's backdrop, under the book kind's switch. See `pdf_preview` for the first and
//! `comic_preview` for the second, and `PreviewType::Ebook` for the kind they share.
//!
//! The list of names both halves are written in is the `[ebook]` row of `crate::formats::lists` —
//! the one table every kind's list is a row of, and the one place a list is written down, the
//! built-in entries and the older lists this app shipped and then changed included. What is left
//! here is the one question about those names the list cannot answer: which of them name a page
//! the PDF reader draws rather than a comic read for its first plate. Three spellings do — `pdf`,
//! `pdfa` and `epdf` — and an Illustrator document saved with PDF compatibility is a fourth that
//! no list holds, which is `pdf_preview`'s own answer of its own (see `is_pdf_file_in_of`).

use crate::formats::text_formats;
use std::path::Path;

/// Whether the file is named with one of the three spellings the PDF reader draws: `pdf`, `pdfa`
/// and `epdf`. Whether such a name is still a book is the list's answer, asked beside this one by
/// [`crate::readers::pdf_preview::is_pdf_file_in`], which is the question this is half of.
///
/// It is asked of the extension every list is matched by, through `lookup_extension`, so a dot
/// file is read the way every other name is and `PDF` is one name rather than two.
pub(crate) fn page_spelling(path: &Path) -> bool {
    matches!(
        text_formats::lookup_extension(path).as_deref(),
        Some("pdf") | Some("pdfa") | Some("epdf")
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::config::AppConfig;
    use crate::formats::lists;

    /// The three spellings are the whole of what a page is called, and they are read by the rule
    /// every list is read by: a leading dot is what a user types, the case it was typed in is not
    /// part of it, and a name with no extension or a longer one is a name rather than a spelling.
    #[test]
    fn a_page_is_named_by_one_of_three_spellings() {
        for name in [
            "book.pdf",
            "report.pdfa",
            "encapsulated.epdf",
            "REPORT.PDF",
            "book..pdf",
        ] {
            assert!(
                page_spelling(Path::new(name)),
                "`{name}` is one of the three spellings the PDF reader draws"
            );
        }

        for name in [
            "chapter.cbz",
            "book.epub",
            "notes.txt",
            "pdf",
            ".pdfx",
            "draw.pdf.gzip",
        ] {
            assert!(
                !page_spelling(Path::new(name)),
                "`{name}` is not one of the three: a comic is the other half of the kind, and a \
                 name with no extension or a longer one is not a spelling"
            );
        }
    }

    /// What the list holds is what no other kind's names: an archive this app reads itself, a
    /// document an engine of its own draws, and a picture and a text file are all answered
    /// elsewhere. A name in two lists is a file whose preview depends on the order rather than on
    /// the name, which is the one thing the table's own drift test exists to keep small — and
    /// this is where it is kept small for this row.
    #[test]
    fn holds_no_name_another_kind_already_reads() {
        let config = AppConfig::default();

        let claimed_elsewhere = |path: &Path| {
            lists::IMAGE.claims(path, &config)
                || lists::VECTOR.claims(path, &config)
                || lists::DESIGN.claims(path, &config)
                || lists::FONT.claims(path, &config)
                || lists::LIBRE.claims(path, &config)
                || lists::MAGICK.claims(path, &config)
                || crate::formats::video_formats::matches_any_video_list(path, &config)
                || lists::ARCHIVE.claims(path, &config)
                || lists::OFFICE.claims(path, &config)
                || lists::PEAZIP.claims(path, &config)
                || lists::CALIBRE.claims(path, &config)
                || lists::TEXT.claims(path, &config)
                || lists::NAMES.claims(path, &config)
        };

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
                !lists::EBOOK.claims(path, &config),
                "`{name}` is read by another kind, so this list does not carry it"
            );
        }

        // `cbz` is the name this list took off the archive list, so it is asked about twice: it
        // must be a comic here, and it must not be an archive there, or the same file would be two
        // answers again.
        let comic = Path::new("chapter.cbz");
        assert!(
            lists::EBOOK.claims(comic, &config),
            "a comic is this list's own name"
        );
        assert!(
            !lists::ARCHIVE.claims(comic, &config),
            "and it is not the archive list's any more: one name, one answer"
        );

        // And the two books that are not this app's at all: a Microsoft Reader book is a page the
        // ebook engine draws, and a compiled help file is the listing engine's, which answers
        // immediately — the reason it is not the engine's is the wait, not the format.
        for (name, claimed_by_calibre) in [("book.lit", true), ("help.chm", false)] {
            let path = Path::new(name);
            assert_eq!(
                lists::CALIBRE.claims(path, &config),
                claimed_by_calibre,
                "`{name}`: whether the ebook engine is the one asked about it"
            );
            assert!(
                lists::PEAZIP.claims(path, &config) != claimed_by_calibre,
                "`{name}` and the listing engine are the other way round"
            );
            assert!(
                !lists::EBOOK.claims(path, &config),
                "`{name}` is not read here, so it is not a book of this kind either"
            );
        }
    }

    /// And what it does hold is the PDF's three spellings and the three comic containers — the
    /// two halves of one kind, told apart by the spellings alone.
    #[test]
    fn holds_the_pages_and_the_comics() {
        let list = lists::EBOOK.built_in();
        let config = AppConfig::default();

        for name in ["book.pdf", "report.pdfa", "encapsulated.epdf"] {
            let path = Path::new(name);
            assert!(
                lists::EBOOK.claims(path, &config),
                "`{name}` is a page this list carries"
            );
            assert!(
                page_spelling(path) && text_formats::matches_configured_extension(path, &list),
                "and it is one of the three spellings the reader of a page draws it with"
            );
        }

        for name in ["chapter.cbz", "chapter.cbr", "collection.cbc"] {
            let path = Path::new(name);
            assert!(
                lists::EBOOK.claims(path, &config),
                "`{name}` is a comic this app reads itself"
            );
            assert!(
                !page_spelling(path),
                "and no spelling here asks the page reader for it: a comic is the other half"
            );
        }

        // A name that is in neither half is neither: what the list does not hold is not a book.
        assert!(!lists::EBOOK.claims(Path::new("book.epub"), &config));
    }

    /// The PDF's names are asked of the list a caller hands in — the form the hook asks them in,
    /// since it resolves a hover with the configuration held and a question that read the
    /// configuration again would be a lock taken twice on that thread (see
    /// `pdf_preview::is_pdf_file_in`).
    #[test]
    fn asks_for_the_pdf_names_of_the_list_it_is_given() {
        let list = lists::EBOOK.built_in();

        for name in ["book.pdf", "report.pdfa", "encapsulated.epdf"] {
            assert!(
                page_spelling(Path::new(name))
                    && text_formats::matches_configured_extension(Path::new(name), &list),
                "`{name}` is a page the PDF reader draws"
            );
        }

        assert!(
            !page_spelling(Path::new("chapter.cbz")),
            "a comic is the other half of this list and not a page of it"
        );
        assert!(
            !lists::EBOOK.built_in().iter().any(|entry| entry == "epub"),
            "and a book the ebook engine converts is not this list's at all"
        );
    }

    /// A name a user adds is a comic if the file it names is a container of pictures, and the
    /// reader that decides that is the comic one — nothing here claims a name is a comic.
    #[test]
    fn a_name_added_by_hand_is_asked_of_the_reader_that_reads_it() {
        let added = lists::sanitize_extension_list("cbc,cbr,cbz,myalbum,pdf,pdfa,epdf");

        assert!(
            text_formats::matches_configured_extension(Path::new("book.myalbum"), &added),
            "a name the list holds and the PDF reader has no spelling for is carried"
        );
        assert!(
            !page_spelling(Path::new("book.myalbum")),
            "and the PDF reader is not asked about it"
        );
        assert!(page_spelling(Path::new("book.pdf")));
    }
}

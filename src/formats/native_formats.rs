//! Which reader of this app's own draws a file: the native half of a route.
//!
//! A kind says what a file *is* and is answered by `routing`; a job says what of this app's
//! reads it. The two are separate because one kind can be drawn by several readers and a
//! reader's job is not the kind's business: a picture is an entry of the crate this app links
//! where the crate carries the format, the codec Windows has where it does not, and a texture
//! reader of this app's own behind both — and all three are the `Images` kind.
//!
//! Every job here is code, not a crate. The `zip` crate backs three of them — an archive a
//! user hovered, the plate of a comic, and the picture a design container keeps of itself —
//! and the `image` crate backs the pictures and, behind the animated readers, the frames of a
//! `.gif`. What a job names is the work, so the question "what reads a `.cbz`" has one answer
//! rather than one per crate it happens to reach.
//!
//! What is deliberately *not* here is equally worth saying. A job is chosen by the name and by
//! the bytes where the bytes settle it — a zip and a tarball are told apart by their magic
//! before their names are read, and a `.png` that turns out to hold more than one frame is the
//! animated reader's — but a job never decides *whether* a file is previewed, and no gate is
//! asked here. That is the kind's question, and it is asked before this one (see
//! `routing::kind_of`).
//!
//! The engine kinds have no job: a document, a book an engine converts, an archive an engine
//! lists and a picture an engine develops are not read by anything here, and the reader of
//! those is named by `routing::chain` rather than by this table.

use crate::config::config::{AppConfig, PreviewType};
use crate::readers::{
    archive_listing, metafile_image, pdf_preview, psd_image, svg_preview, wic_image,
};
use std::path::Path;

/// What of this app's own reads a file.
///
/// Most of these are a reader module; two are a transform every picture of a kind goes through
/// rather than a reader (`tone_map`, which is why it is not here); and two are a pair of steps
/// rather than a choice — see [`NativeJob::SvgDocument`] and [`NativeJob::FontSpecimen`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NativeJob {
    /// The `image` crate's decoders, at the one limit every picture is decoded under.
    Picture,
    /// The codec Windows has for a format the `image` crate does not carry — an AVIF, a HEIC,
    /// a JPEG XL, a still WebP — with this app's own readers behind it: the texture reader,
    /// which reads what the codec refuses, and libwebp, which is what a machine without the
    /// WebP codec falls back to (see `wic_image`, `dds_image`, `webp_image`).
    PictureCodec,
    /// An animated GIF, streamed frame by frame into the shared queue every animation is
    /// played through. A file of the name that does not move is the still picture behind it.
    AnimatedGif,
    /// The same for an animated PNG, which is a PNG whose frames the `image` crate reads.
    AnimatedApng,
    /// And for an animated WebP, which is libwebp's rather than the codec's.
    AnimatedWebp,
    /// A Photoshop document, read for the merged picture it keeps past its layers.
    Psd,
    /// A design container — a Krita or OpenRaster project, a Procreate document — read for the
    /// picture the format keeps of the whole document, with the encapsulated PostScript reader
    /// behind it for the drawings that are a container of one.
    Project,
    /// An encapsulated PostScript file, read for the preview it carries: an EPS holds a
    /// metafile or a TIFF of itself and is not interpreted here.
    Eps,
    /// A Windows metafile, replayed by the drawing layer rather than decoded.
    Metafile,
    /// A PDF, drawn a page at a time by the engine Windows has.
    Pdf,
    /// A comic: the first plate read out of the container it is published in.
    Comic,
    /// A listing of a zip, which is also what a comic's plate and a project's picture are read
    /// out of — the container, not the preview.
    ArchiveZip,
    /// A listing of a 7z.
    ArchiveSevenZ,
    /// A listing of a rar, read by the UnRAR sources the app links.
    ArchiveRar,
    /// A listing of a tar.
    ArchiveTar,
    /// A listing of a gzip stream of a tar, and of the `tgz` spelling of the same file.
    ArchiveTarGz,
    /// The text reader: the document decoded, a window of it styled on demand and painted,
    /// with the theme and the Markdown mode the configuration names.
    Text,
    /// An SVG document, which is measured here and drawn by the browser engine — both steps,
    /// and the second is what makes the first worth doing, so this is one job and not two.
    SvgDocument,
    /// A font specimen: the two tables a specimen is described by are read here, and the
    /// glyphs are the browser engine's — the same two steps as an SVG document's.
    FontSpecimen,
    /// A video played by the media engine Windows has, in this app's own window.
    VideoMediaFoundation,
}

/// What reads a file of `kind`, or nothing where this app has no reader for it.
///
/// The kind is handed in rather than worked out here: it is the answer to a different
/// question — whether a file is previewed at all, and as what — and it has already been given
/// by the time a route is resolved (see `routing::kind_of`). What this narrows is which of
/// this app's readers does the work, and the container formats are the reason it takes the
/// file rather than the name: a zip and a tarball are told apart by their own bytes, so a
/// renamed one is still listed by the reader that can open it.
///
/// The configuration is handed in as well, and nothing here takes that lock again: the hook
/// asks its questions with the configuration already held, and a question that went and read
/// it a second time would be a lock taken twice on one thread (see
/// `explorer_hook::is_media_file`).
///
/// A picture is the one job here the name does not settle, so it is the one that reads the
/// file: the front of it, through `head`, where the still `.gif` and the moving one are told
/// apart by their own frame blocks and a `.png` holding more than one frame is the animated
/// reader's. Nothing is decoded to answer it and nothing past the head is read.
///
/// There is no setting behind any of this and no gate: what a file's own bytes say is asked
/// always, which is also what the content tier does before this one (see `content_type::of`).
/// A file whose head could not be read — one that is not there — is answered by the still choice
/// below, the same way it was when the name was the whole of the answer.
pub fn job_for(path: &Path, kind: PreviewType, config: &AppConfig) -> Option<NativeJob> {
    match kind {
        PreviewType::Images => Some(picture_job(path)),
        PreviewType::Design => Some(design_job(path)),
        PreviewType::Vector => Some(drawing_job(path)),
        PreviewType::Ebook => Some(page_job(path, config)),
        PreviewType::Archives => archive_job(path),
        PreviewType::Text => Some(NativeJob::Text),
        PreviewType::Fonts => Some(NativeJob::FontSpecimen),
        PreviewType::Videos => Some(NativeJob::VideoMediaFoundation),

        // The kinds an engine draws, and the one kind that is a reader's own but has its job
        // asked where the engine is: a picture an image converter develops is a PNG of the
        // engine's, so there is nothing here for it to be.
        PreviewType::Peazip
        | PreviewType::Calibre
        | PreviewType::Document
        | PreviewType::Libre
        | PreviewType::Magick => None,
    }
}

/// A picture: the crate's decoder where it carries the format, the codec Windows has where it
/// does not, and the animated reader first of all where the file's own bytes say it moves.
///
/// This is the one job in this table that the name cannot settle. A `.gif`, a `.webp` and a
/// `.png` each cover a still file and a moving one, and what tells them apart is the file's own
/// structure — its frame blocks, its `ANIM` chunk, its `acTL` chunk — which `head` reads out of
/// the front of the file without decoding any of it (see `head::PictureNature`).
///
/// A file whose own bytes say *still* is the still job, which is what a `.gif` with one frame in
/// it is: the animated reader is not asked about it at all, and the frame it used to decode and
/// throw away before the still path drew the same picture is not decoded. A file whose head
/// could not be read — one that is not there, one whose content is still in the cloud — is the
/// still job too, which is the path that would have failed to read it in any case.
fn picture_job(path: &Path) -> NativeJob {
    use crate::formats::head::{PictureFamily, PictureForm};

    if let Some(nature) = crate::formats::head::picture_nature(path) {
        match nature.form {
            PictureForm::Plays(PictureFamily::Gif) => return NativeJob::AnimatedGif,
            PictureForm::Plays(PictureFamily::Webp) => return NativeJob::AnimatedWebp,
            PictureForm::Plays(PictureFamily::Apng) => return NativeJob::AnimatedApng,
            PictureForm::Still | PictureForm::Unplayable | PictureForm::Paged => {}
        }
    }

    if wic_image::is_codec_file(path) {
        return NativeJob::PictureCodec;
    }

    NativeJob::Picture
}

/// A design document: the merged picture a Photoshop document keeps, or the picture a project
/// container keeps of the whole document.
fn design_job(path: &Path) -> NativeJob {
    if psd_image::is_psd_file(path) {
        return NativeJob::Psd;
    }

    NativeJob::Project
}

/// A drawing: the browser engine's where the document is an SVG, the drawing layer's where it
/// is a metafile, and the encapsulated PostScript reader's for the rest.
fn drawing_job(path: &Path) -> NativeJob {
    if svg_preview::is_svg_file(path) {
        return NativeJob::SvgDocument;
    }

    if metafile_image::is_metafile_name(path) {
        return NativeJob::Metafile;
    }

    NativeJob::Eps
}

/// A page of the `Ebook` kind: a comic where the container is one, and a PDF otherwise.
///
/// The two are one kind because what they are shown as is one thing, and which of them a file
/// is, is asked exactly as the kind was: the same question the router asks to reach this kind
/// is the one that tells the two halves apart, so a PDF, a page spelling the list no longer
/// holds, and an Illustrator document carrying a PDF inside it are all the PDF reader's (see
/// `pdf_preview::is_pdf_file_in`), and everything else of the kind is a comic.
fn page_job(path: &Path, config: &AppConfig) -> NativeJob {
    if pdf_preview::is_pdf_file_in(path, &config.ebook_extensions) {
        return NativeJob::Pdf;
    }

    NativeJob::Comic
}

/// A container, in the reader the container itself names: its magic first and its spelling
/// after it, which is the answer the listing reader gives and the one this table borrows
/// rather than keeping a second copy of (see `archive_listing::kind_of`).
fn archive_job(path: &Path) -> Option<NativeJob> {
    let kind = archive_listing::kind_of(path)?;

    Some(match kind {
        archive_listing::ArchiveKind::Zip => NativeJob::ArchiveZip,
        archive_listing::ArchiveKind::SevenZ => NativeJob::ArchiveSevenZ,
        archive_listing::ArchiveKind::Rar => NativeJob::ArchiveRar,
        archive_listing::ArchiveKind::Tar => NativeJob::ArchiveTar,
        archive_listing::ArchiveKind::TarGz => NativeJob::ArchiveTarGz,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::config::AppConfig;
    use crate::formats::routing;
    use std::path::PathBuf;

    fn job_for_name(name: &str, kind: PreviewType) -> Option<NativeJob> {
        job_for(&PathBuf::from(name), kind, &AppConfig::default())
    }

    /// A picture is the crate where the crate carries the format and the codec where it does
    /// not. The three names that *can* move are not decided here: whether one of them moves is
    /// its own bytes' answer rather than its name's, and that is the test below.
    #[test]
    fn a_picture_is_the_decoder_that_carries_the_format() {
        for (name, job) in [
            ("photo.jpg", NativeJob::Picture),
            ("scan.tif", NativeJob::Picture),
            ("vector.bmp", NativeJob::Picture),
            ("texture.dds", NativeJob::PictureCodec),
            ("shot.heic", NativeJob::PictureCodec),
            ("picture.avif", NativeJob::PictureCodec),
            // A `.png` is a picture unless its own `acTL` chunk says otherwise, and a path
            // with no file behind it has no chunk to say it with.
            ("page.png", NativeJob::Picture),
            ("clip.gif", NativeJob::Picture),
        ] {
            assert_eq!(
                job_for_name(name, PreviewType::Images),
                Some(job),
                "`{name}` is read by {job:?}"
            );
        }
    }

    /// A container is opened by the reader its own bytes name, which is where a renamed one is
    /// still listed as what it is — and the dotted `tar.gz` is claimed by the end of the name
    /// rather than by the last extension, which is the one spelling no extension list holds.
    #[test]
    fn a_container_is_read_by_the_reader_its_own_bytes_name() {
        for (name, job) in [
            ("archive.zip", NativeJob::ArchiveZip),
            ("package.xpi", NativeJob::ArchiveZip),
            ("backup.7z", NativeJob::ArchiveSevenZ),
            ("bundle.rar", NativeJob::ArchiveRar),
            ("sources.tar", NativeJob::ArchiveTar),
            ("sources.tgz", NativeJob::ArchiveTarGz),
            ("sources.tar.gz", NativeJob::ArchiveTarGz),
        ] {
            assert_eq!(
                job_for_name(name, PreviewType::Archives),
                Some(job),
                "`{name}` is listed by {job:?}"
            );
        }

        assert_eq!(
            job_for_name("notes.txt", PreviewType::Archives),
            None,
            "and a name no container reader opens is not a container"
        );
    }

    /// A page is the PDF reader's or the comic reader's, and the name is what settles which:
    /// a comic is the half of the book list that is not a page.
    #[test]
    fn a_page_is_the_pdf_reader_or_the_comic_one() {
        for (name, job) in [
            ("book.pdf", NativeJob::Pdf),
            ("archived.pdfa", NativeJob::Pdf),
            ("chapter.cbz", NativeJob::Comic),
            ("chapter.cbr", NativeJob::Comic),
            ("chapter.cbc", NativeJob::Comic),
        ] {
            assert_eq!(
                job_for_name(name, PreviewType::Ebook),
                Some(job),
                "`{name}` is read by {job:?}"
            );
        }
    }

    /// A drawing is the browser engine's, the drawing layer's or the PostScript reader's, and
    /// a document is the merged picture of itself.
    #[test]
    fn a_drawing_and_a_document_reach_the_reader_that_replays_them() {
        for (name, job) in [
            ("art.svg", NativeJob::SvgDocument),
            ("clipart.emf", NativeJob::Metafile),
            ("sketch.wmf", NativeJob::Metafile),
            ("art.eps", NativeJob::Eps),
            ("artwork.ai", NativeJob::Eps),
        ] {
            assert_eq!(
                job_for_name(name, PreviewType::Vector),
                Some(job),
                "`{name}` is read by {job:?}"
            );
        }

        for (name, job) in [
            ("layered.psd", NativeJob::Psd),
            ("large.psb", NativeJob::Psd),
            ("project.kra", NativeJob::Project),
            ("project.ora", NativeJob::Project),
        ] {
            assert_eq!(
                job_for_name(name, PreviewType::Design),
                Some(job),
                "`{name}` is read by {job:?}"
            );
        }
    }

    /// The kinds an engine draws have no job here, which is what keeps an engine from being
    /// reached twice: the reader of one is named by the chain rather than by this table.
    #[test]
    fn the_engine_kinds_have_no_native_job() {
        for kind in [
            PreviewType::Peazip,
            PreviewType::Calibre,
            PreviewType::Document,
            PreviewType::Libre,
            PreviewType::Magick,
        ] {
            assert_eq!(
                job_for_name("anything.any", kind),
                None,
                "{kind:?} is drawn by an engine, which this table does not name"
            );
        }
    }

    /// The seam: what the router answers for a name is a kind, and every kind with a reader of
    /// this app's own has a job behind it — which is the whole route for a file, in one test.
    #[test]
    fn every_name_that_reaches_a_native_kind_reaches_a_job() {
        let config = AppConfig::default();

        for name in [
            "photo.jpg",
            "texture.dds",
            "clip.gif",
            "notes.txt",
            "sources.tar.gz",
            "chapter.cbz",
            "book.pdf",
            "art.svg",
            "layered.psd",
            "specimen.woff2",
            "film.mp4",
        ] {
            let path = PathBuf::from(name);
            let kind = routing::kind_of(&path, &config)
                .unwrap_or_else(|| panic!("`{name}` reaches a kind"));

            assert!(
                job_for(&path, kind, &config).is_some(),
                "`{name}` reaches {kind:?}, which has a reader of this app's own"
            );
        }
    }

    /// A picture's job is its own bytes' answer rather than its name's, which is the whole
    /// point of this table reading the head. A `.gif` with one frame in it is the still job, so
    /// the animated reader is never asked about one and the frame it used to decode on the way
    /// to finding that out is not decoded at all; a file whose name says *document* but whose
    /// bytes are an animation is the animated reader's, with no setting to switch it on; and a
    /// `.png` whose bytes are a container is still the picture job here, because what a
    /// container *is* is the kind's question and it is asked before this one.
    #[test]
    fn a_picture_job_is_what_the_bytes_say_rather_than_what_the_name_says() {
        use crate::formats::head::tests as sample;

        let config = AppConfig::default();
        let folder = std::env::temp_dir().join("rust-hover-preview-native-jobs");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let job_of = |name: &str, bytes: &[u8]| {
            let path = folder.join(name);
            std::fs::write(&path, bytes).expect("a test file");

            job_for(&path, PreviewType::Images, &config)
        };

        let still_gif = [sample::gif_open(), sample::gif_frame(), vec![0x3B]].concat();
        let moving_gif = [
            sample::gif_open(),
            sample::gif_frame(),
            sample::gif_frame(),
            vec![0x3B],
        ]
        .concat();
        let moving_png = [
            sample::png_open(),
            sample::png_chunk(b"IHDR", &[0u8; 13]),
            sample::png_chunk(b"acTL", &[0u8; 8]),
            sample::png_chunk(b"IDAT", &[0]),
        ]
        .concat();

        assert_eq!(
            job_of("clip.gif", &still_gif),
            Some(NativeJob::Picture),
            "a `.gif` with one frame in it is a still, and is decoded once rather than twice"
        );
        assert_eq!(
            job_of("named-document.docx", &moving_gif),
            Some(NativeJob::AnimatedGif),
            "the file's own bytes are the answer, whatever it is called"
        );
        assert_eq!(
            job_of("frames.png", &moving_png),
            Some(NativeJob::AnimatedApng)
        );
        assert_eq!(
            job_of("not-really.png", b"PK\x03\x04the rest of a container"),
            Some(NativeJob::Picture),
            "what this table answers for a picture kind, with the kind itself decided earlier"
        );
    }
}

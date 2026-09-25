//! Which kind claims a file, and who draws it: the one place the order is written down.
//!
//! A file's kind is asked in one order and one order only — a video, then a page, then the
//! boxes and the documents and the drawings, the text, the fonts and last the pictures — and
//! this module is that order. Every side that has to know what a file is asks it here: the
//! hook that raises a hover, the loader that fills it, the layout that places it, and the
//! content tier that has to turn a format's own names back into a kind (see
//! `content_type::kind_claiming`). Four chains used to answer this question separately, and
//! two of them had already drifted apart; the point of the table is that they cannot.
//!
//! Two questions are asked of it, and they differ in one place. [`kind_of`] is asked of a
//! file: the video list's two names that are also a text list's are settled by the file's own
//! content, which is a read. [`kind_of_name`] is asked of a name the content has already
//! answered with — the form `content_type` asks, where the file is a signature's own name and
//! has never been opened — and there is nothing left to read (see [`Asked`]).
//!
//! The lists themselves are not here: they are `config.ini`'s, and which names each one holds
//! is the user's to edit. What this module owns is the *order* they are asked in, and what
//! each kind is drawn by — see [`chain`]. A name that sits in two lists is settled by that
//! order rather than by the name, which is why the order has one author.

use crate::config::config::{AppConfig, PreviewType};
use crate::engines::{
    calibre_render, imagemagick_render, libreoffice_render, peazip_render, webview_preview,
};
use crate::formats::{
    archive_formats, calibre_formats, codecs, design_formats, ebook_formats, font_formats,
    image_formats, libre_formats, magick_formats, office_formats, peazip_formats, text_formats,
    vector_formats, video_formats,
};
use crate::readers::pdf_preview;
use crate::readers::svg_preview;
use std::path::Path;

/// What a claim is asked of: the file, or a name the content has already answered with.
///
/// The difference is worth a type because it is the difference between reading a file and not
/// reading it, and it exists for exactly one pair of names. A `.ts` is a transport stream in
/// one list and TypeScript in another, and which of the two it is, is the file's own content:
/// asked of a file, that content is read; asked of a name the content already spoke for, it is
/// not asked again (see `video_formats::claims_video_name`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Asked {
    /// The file itself, whose content has not been read.
    File,
    /// A name, where the caller already holds the content's own answer.
    Name,
}

/// One kind's claim on a file, and the kind it answers with.
///
/// The answer is the kind rather than a yes or a no because one list answers with a kind of
/// its own: a drawing is not a picture, and the image list is asked last of all — so a name
/// that list still holds and does not own is answered with the kind it belongs to. See
/// [`claim_images`].
type Claim = fn(&Path, &AppConfig, Asked) -> Option<PreviewType>;

/// The kinds, in the one order they are asked in.
///
/// The order is the one the hook has always used, which is the order the renderer's chain
/// already agreed with. Two of them are load-bearing rather than arbitrary:
///
/// * **A video is asked first**, because only its content settles the names it shares with the
///   text lists: a `.ts` carrying MPEG-TS packets is a video however the gates stand, and one
///   that does not is the TypeScript source the text lists claim (see [`claim_video`]).
/// * **The pictures are asked last**, because a name can reach them by being a picture's and
///   by being something else's — a `.dds` is a texture, an `svg` a hand-edited image list
///   still names is a drawing — and the kinds that can answer for one are asked before it.
const CLAIMS: &[Claim] = &[
    claim_video,
    claim_ebook,
    claim_archives,
    claim_peazip,
    claim_calibre,
    claim_document,
    claim_libre,
    claim_magick,
    claim_design,
    claim_vector,
    claim_text,
    claim_fonts,
    claim_images,
];

/// The kind that claims `path`, or nothing where no list does.
///
/// This is the question a hover asks: the file's own content is not in hand, so the two names
/// the video list shares with the text lists are settled by reading the file (see [`Asked`]).
/// The caller must not be holding the configuration itself — nothing here takes that lock
/// again, and every reader that would is handed the configuration instead.
pub fn kind_of(path: &Path, config: &AppConfig) -> Option<PreviewType> {
    claims(path, config, Asked::File)
}

/// The kind a name belongs to, where the file's content has already answered for it.
///
/// It is the form `content_type` asks in: what it holds is a list of names a signature
/// answered with — a Matroska file's `mkv` and `webm` — and those names have to be turned back
/// into a kind. A name asked this way is asked of the lists alone, because the read that would
/// settle it has already been made on the file it came from.
pub fn kind_of_name(path: &Path, config: &AppConfig) -> Option<PreviewType> {
    claims(path, config, Asked::Name)
}

fn claims(path: &Path, config: &AppConfig, asked: Asked) -> Option<PreviewType> {
    CLAIMS.iter().find_map(|claim| claim(path, config, asked))
}

/// A video: the configured list's names, and the two of them it shares with the text lists
/// settled by the file rather than by the name where the file is the thing being asked about.
fn claim_video(path: &Path, config: &AppConfig, asked: Asked) -> Option<PreviewType> {
    let claimed = match asked {
        Asked::File => video_formats::matches_video_list(path, &config.video_extensions),
        Asked::Name => video_formats::claims_video_name(path, &config.video_extensions),
    };

    claimed.then_some(PreviewType::Videos)
}

/// A page this app draws itself: a PDF, and a comic, which is the other half of the same kind.
///
/// The two are one kind because what they are shown as is one thing — a page — and what a user
/// turns off for either is books. Which of the two a file is, is asked by the reader rather
/// than here: a comic is a container of plates and a PDF is a page, and the loader knows the
/// difference (see `preview_window::load_media_of_kind`). The PDF's names are asked of the list
/// in hand rather than of the gate that reads it, since the caller holds the configuration
/// already (see `pdf_preview::is_pdf_file_in`).
fn claim_ebook(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    let page = pdf_preview::is_pdf_file_in(path, &config.ebook_extensions);
    let book = ebook_formats::matches_ebook_list(path, &config.ebook_extensions);

    (page || book).then_some(PreviewType::Ebook)
}

/// An archive this app reads itself, which is a listing rather than an engine's work.
fn claim_archives(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    archive_formats::matches_archive_list(path, &config.archive_extensions)
        .then_some(PreviewType::Archives)
}

/// An archive a listing engine reads, asked beside the archive list above it: a name in that
/// list is read by this app itself, and one in this list is read by an engine. A name in both
/// is the archive list's, since that list is asked first.
fn claim_peazip(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    peazip_formats::matches_peazip_list(path, &config.peazip_extensions)
        .then_some(PreviewType::Peazip)
}

/// A book a conversion engine reads: a name no reader here opens, and no list above claims.
fn claim_calibre(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    calibre_formats::matches_calibre_list(path, &config.calibre_extensions)
        .then_some(PreviewType::Calibre)
}

/// A document shown as a page: the names the application that owns the format draws, and the
/// names a render engine draws because no application of that family is installed. Which of
/// the two draws a file is the machine's answer rather than the list's (see [`chain`]).
fn claim_document(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    office_formats::matches_office_list(path, &config.office_extensions)
        .then_some(PreviewType::Document)
}

/// A document only a render engine draws, asked after the Office list because a name can sit
/// in both: what such a file keeps of itself is a thumbnail, and a thumbnail is not a preview.
fn claim_libre(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    libre_formats::matches_libre_list(path, &config.libre_extensions).then_some(PreviewType::Libre)
}

/// A picture an image converter develops — a camera raw above all — asked before the design
/// list, which it is a neighbour of rather than a member of.
fn claim_magick(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    magick_formats::matches_magick_list(path, &config.magick_extensions)
        .then_some(PreviewType::Magick)
}

/// A design document: a kind of its own, whose preview is made of the picture the file keeps
/// of itself rather than of a decoder its name names.
fn claim_design(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    design_formats::matches_design_list(path, &config.design_extensions)
        .then_some(PreviewType::Design)
}

/// A drawing: a metafile, an encapsulated PostScript file, or an SVG document.
fn claim_vector(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    vector_formats::matches_vector_list(path, &config.vector_extensions)
        .then_some(PreviewType::Vector)
}

/// Text: the extension lists and the names list, which is what makes a `makefile` a text file
/// and what the dot-file rule is for (see `text_formats::lookup_extension`).
fn claim_text(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    text_formats::matches_text_lists(path, &config.text_extensions, &config.text_names)
        .then_some(PreviewType::Text)
}

/// A font, asked ahead of the image list that would have turned a `.ttf` down: what draws one
/// is the browser engine rather than a decoder.
fn claim_fonts(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    font_formats::matches_font_list(path, &config.font_extensions).then_some(PreviewType::Fonts)
}

/// A picture — and the one entry that can answer with a kind that is not its own.
///
/// A drawing is not a picture, and the image list is where the drawings of that kind were until
/// they were given a kind of their own: an `svg` a hand-edited image list still names is
/// answered as the drawing it is rather than as the picture its name says (see
/// `svg_preview::is_svg_file`). It is asked inside this entry rather than as an entry of its
/// own so that the order the other kinds are asked in is not disturbed by it.
fn claim_images(path: &Path, config: &AppConfig, _asked: Asked) -> Option<PreviewType> {
    if !image_formats::matches_image_list(path, &config.image_extensions) {
        return None;
    }

    Some(if svg_preview::is_svg_file(path) {
        PreviewType::Vector
    } else {
        PreviewType::Images
    })
}

/// Who draws a file of one kind, in the order they are asked.
///
/// It is the readers a kind has rather than the reader it has, because a kind can be drawn by
/// more than one: a document is the application that owns its format where one is installed
/// and an installed render engine where none is, and which of the two is the machine's answer
/// rather than the user's (`office_formats::page_engine`). Most chains are one long today —
/// what writing them down buys is that the second reader a kind will grow is a line here
/// instead of an edit in six chains.
///
/// What resolves a chain is [`readers_for`] beside it: whether a reader is installed and
/// whether it has refused this file are questions about the machine and the run, and they are
/// asked where those answers are kept. This is the declaration; that is the resolution.
pub fn chain(kind: PreviewType) -> &'static [Reader] {
    match kind {
        // A video is played by FFmpeg where it is installed and by the engine Windows has
        // where it is not, which is one answer per file rather than per hover (see `codecs`).
        PreviewType::Videos => &[Reader::Ffmpeg, Reader::Native],

        PreviewType::Ebook => &[Reader::Native],
        PreviewType::Archives => &[Reader::Native],
        PreviewType::Peazip => &[Reader::PeaZip],
        PreviewType::Calibre => &[Reader::Calibre],

        // The application the format belongs to first, and the render engine where it is not
        // installed — the fallback the page tier has always had.
        PreviewType::Document => &[Reader::Office, Reader::LibreOffice],

        PreviewType::Libre => &[Reader::LibreOffice],
        PreviewType::Magick => &[Reader::ImageMagick],

        // A design document is read for the picture its own format keeps of the whole thing,
        // and a page an engine has already drawn of one is preferred to that thumbnail.
        PreviewType::Design => &[Reader::LibreOffice, Reader::Native],

        // An SVG document is drawn by the browser engine and a metafile or an encapsulated
        // PostScript file by the drawing layer; the name settles which of the two.
        PreviewType::Vector => &[Reader::WebView2, Reader::Native],

        PreviewType::Text => &[Reader::Native],
        PreviewType::Fonts => &[Reader::WebView2],
        PreviewType::Images => &[Reader::Native],
    }
}

/// The readers of a file of this kind that could answer for it, in the order they are asked.
///
/// It is [`chain`] narrowed to what is true of the machine and the file: a reader that is not
/// installed is not one to ask, one that was already asked about this file and turned it down
/// is not asked twice, and one that cannot be started here is not a wait to sit through. What
/// is left is asked in the chain's order, so the first reader that can answer is the one that
/// does — which is what the preview loop walks when it has a hover to fill and no page to fill
/// it with.
///
/// Whether a reader has *already* answered is not asked here: a page that has landed is in
/// hand, and what the loop asks of this is only which engine to ask next. That question is
/// asked where the answers are kept (see the render tiers in `preview_window`).
pub fn readers_for(kind: PreviewType, path: &Path) -> Vec<Reader> {
    chain(kind)
        .iter()
        .copied()
        .filter(|reader| can_answer(*reader, path))
        .collect()
}

/// Whether a reader is one to ask about `path` at all.
fn can_answer(reader: Reader, path: &Path) -> bool {
    match reader {
        // A reader of this app's own is always there, and what it can read is the file's own
        // question rather than the machine's (`native_formats`).
        Reader::Native => true,

        // The application that owns the format, where the machine has it: the answer is per
        // file rather than per engine, because which application a name belongs to is the
        // name's (see `office_formats::app_for`).
        Reader::Office => office_formats::app_installed(path),

        Reader::LibreOffice => {
            libreoffice_render::available() && !libreoffice_render::refused(path)
        }
        Reader::ImageMagick => {
            imagemagick_render::available() && !imagemagick_render::refused(path)
        }
        // Per name rather than per engine: the tools beside the console archiver read three of
        // the names in the list, and a machine with the archiver and without the tool is a
        // machine that cannot list those three (see `peazip_formats::Backend`).
        Reader::PeaZip => peazip_render::available_for(path) && !peazip_render::refused(path),
        Reader::Calibre => calibre_render::available() && !calibre_render::refused(path),

        // FFmpeg's player, where it is installed: the one reader preferred over the engine
        // Windows has, and the reason a video is a chain rather than a reader (see `codecs`).
        Reader::Ffmpeg => codecs::ffplay_available(),

        // The browser engine, which draws what neither of the two others can: a metafile is
        // not this reader's at all, but a document drawn by it is nothing if it is missing.
        Reader::WebView2 => webview_preview::is_available(),
    }
}

/// One reader of a file: a reader of this app's own, or an engine this app drives.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Reader {
    /// A reader of this app's own — the picture decoders, the text and archive readers, the
    /// PDF engine Windows has, the media engine that plays a video. Which one is a question
    /// of the file, and it is asked by `native_formats`.
    Native,
    /// The application that owns the format, driven through its own automation.
    Office,
    /// The installed render engine, which draws a page of what its import filters read.
    LibreOffice,
    /// The installed image converter, which develops a picture nothing else opens.
    ImageMagick,
    /// The installed archiver, which lists a container no reader here opens.
    PeaZip,
    /// The installed ebook converter, which writes a PDF of a book nothing here reads.
    Calibre,
    /// FFmpeg's player, where it is installed: the engine video previews are played by when
    /// it is, and the one engine preferred over the reader Windows has.
    Ffmpeg,
    /// The browser engine the machine already has, which draws SVG documents and font
    /// specimens — neither of them rasterized on this side.
    WebView2,
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    /// Every list this app ships, as the paths that reach it: an extension is reached by a
    /// name carrying it, and a text name is reached by the name itself, since that is the
    /// lookup the list is written for (see `text_formats::lookup_name`).
    fn shipped_lists(config: &AppConfig) -> Vec<(PreviewType, Vec<PathBuf>)> {
        let extensions = |names: &[String]| -> Vec<PathBuf> {
            names
                .iter()
                .map(|name| PathBuf::from(format!("preview.{name}")))
                .collect()
        };

        let mut lists = vec![
            (PreviewType::Videos, extensions(&config.video_extensions)),
            (PreviewType::Ebook, extensions(&config.ebook_extensions)),
            (
                PreviewType::Archives,
                extensions(&config.archive_extensions),
            ),
            (PreviewType::Peazip, extensions(&config.peazip_extensions)),
            (PreviewType::Calibre, extensions(&config.calibre_extensions)),
            (PreviewType::Document, extensions(&config.office_extensions)),
            (PreviewType::Libre, extensions(&config.libre_extensions)),
            (PreviewType::Magick, extensions(&config.magick_extensions)),
            (PreviewType::Design, extensions(&config.design_extensions)),
            (PreviewType::Vector, extensions(&config.vector_extensions)),
            (PreviewType::Text, extensions(&config.text_extensions)),
            (PreviewType::Fonts, extensions(&config.font_extensions)),
            (PreviewType::Images, extensions(&config.image_extensions)),
        ];

        let names = config
            .text_names
            .iter()
            .map(PathBuf::from)
            .collect::<Vec<PathBuf>>();
        lists.push((PreviewType::Text, names));

        lists
    }

    /// The names two lists hold, and the kind the order gives each of them.
    ///
    /// Each is a name that is deliberately in two lists, and the winner is the kind asked
    /// first. Nothing else may be in two lists: a name that reaches two kinds is a file whose
    /// preview depends on the order rather than on the name, which is the thing this table
    /// exists to keep written down and small.
    const SHARED: &[(&str, PreviewType)] = &[
        // A `.ts` and an `.mts` are a transport stream in the video list and TypeScript in the
        // text lists, and the content settles which: a file of either name that is not a
        // transport stream is the source the text lists claim.
        ("preview.ts", PreviewType::Text),
        ("preview.mts", PreviewType::Text),
        // A `.dif` is a DV stream in the video list and a Data Interchange Format spreadsheet
        // in the render engine's list. The video list is asked first, so the stream wins: a
        // spreadsheet of that name is previewed as the video it is not.
        ("preview.dif", PreviewType::Videos),
        // A `.cdr` is a drawing the render engine draws and a CorelDRAW container the design
        // list reads a thumbnail out of. The engine is asked first, because what a thumbnail
        // shows of a drawing is not a preview of it.
        ("preview.cdr", PreviewType::Libre),
        // A `.vhd` is a Virtual PC disk image the listing engine opens and VHDL source the
        // text lists read, and nothing in the app had ever written the pair down: the two
        // lists were asked in an order that settled it and no test said which. It is the
        // listing engine's, because that list is asked first — a name that is a language more
        // often than it is a disk image is the reading the order gets wrong, and moving it is
        // a line in this table rather than a reordering of the lists.
        ("preview.vhd", PreviewType::Peazip),
    ];

    /// Every name this app ships reaches the kind its list is written for, and a name two
    /// lists hold reaches the kind that table names.
    ///
    /// It is the test that would have caught the two chains that had drifted apart: the hook
    /// asked the image converter's list before the design list and the loader asked the design
    /// list first, so a name in both was gated as one kind and drawn as another. Nothing this
    /// app ships was in both — which is why it went unnoticed — and this is what says so.
    #[test]
    fn every_shipped_name_reaches_the_kind_its_list_is_written_for() {
        let config = AppConfig::default();
        let lists = shipped_lists(&config);
        let mut unexpected: Vec<String> = Vec::new();

        for (kind, paths) in &lists {
            for path in paths {
                let name = path
                    .file_name()
                    .and_then(|name| name.to_str())
                    .expect("a file name");

                let owners = lists.iter().filter(|(_, held)| held.contains(path)).count();

                let expected = match SHARED.iter().find(|(shared, _)| *shared == name) {
                    Some((_, winner)) => *winner,
                    None if owners > 1 => {
                        unexpected.push(format!(
                            "{name}: {kind:?} list and another list both claim it, and it is not \
                             written down in SHARED"
                        ));
                        continue;
                    }
                    None => *kind,
                };

                let reached = kind_of(path, &config);
                if reached != Some(expected) {
                    unexpected.push(format!(
                        "{name}: the {kind:?} list's, reached {reached:?}, expected {expected:?}"
                    ));
                }
            }
        }

        assert!(
            unexpected.is_empty(),
            "names reached a kind their list is not written for:\n  {}",
            unexpected.join("\n  ")
        );
    }

    /// The walk narrows a chain and never reorders it: what a file can be asked of is what its
    /// kind declares, less whatever this machine cannot answer with, in the order the kind
    /// names them — so the reader asked first is the first of the chain that can answer.
    #[test]
    fn the_readers_a_file_is_asked_of_are_its_chain_narrowed() {
        for (name, kind) in [
            ("letter.docx", PreviewType::Document),
            ("help.chm", PreviewType::Peazip),
            ("book.epub", PreviewType::Calibre),
            ("drawing.cdr", PreviewType::Libre),
            ("shot.nef", PreviewType::Magick),
            ("photo.jpg", PreviewType::Images),
            ("film.mp4", PreviewType::Videos),
            ("notes.txt", PreviewType::Text),
        ] {
            let path = Path::new(name);
            let asked = readers_for(kind, path);
            let declared = chain(kind);

            assert!(
                asked.len() <= declared.len(),
                "`{name}` cannot be asked of more readers than {kind:?} declares"
            );

            // A subsequence rather than a set: every reader asked is one the chain names, and
            // they come in the order it names them.
            let mut rest = declared.iter();
            for reader in &asked {
                assert!(
                    rest.any(|declared| declared == reader),
                    "`{name}` is asked of {reader:?} out of the order {kind:?} declares"
                );
            }
        }

        assert_eq!(
            readers_for(PreviewType::Document, Path::new("letter.docx")).first(),
            chain(PreviewType::Document).first(),
            "the first reader of a chain that can answer is the one asked"
        );
    }

    /// The two questions this module answers are the same question but in one place, and that
    /// place is the video list's two shared names: asked of a file that is not a transport
    /// stream, the content is read and the text lists win; asked of a name the content has
    /// already answered with, there is nothing to read and the video list's answer stands.
    #[test]
    fn a_name_the_content_answered_with_is_asked_of_the_lists_alone() {
        let config = AppConfig::default();
        let stream = Path::new("content.ts");

        assert_eq!(
            kind_of_name(stream, &config),
            Some(PreviewType::Videos),
            "a signature that named a transport stream has already settled what it is"
        );
        assert_eq!(
            kind_of(stream, &config),
            Some(PreviewType::Text),
            "and a file of that name which is not one is the source the text lists claim"
        );
    }

    /// The chain names the readers of each kind, and the two names whose preview was given up
    /// on purpose keep the single reader they have: a `.cdr` is a page an engine draws and
    /// nothing without one, and a `.chm` is a listing that is there immediately rather than a
    /// page that takes two seconds to draw.
    #[test]
    fn the_chain_names_the_readers_of_each_kind() {
        let config = AppConfig::default();

        assert_eq!(
            chain(PreviewType::Document),
            &[Reader::Office, Reader::LibreOffice],
            "a document is the application's to draw where it is installed, and the engine's where it is not"
        );
        assert_eq!(
            chain(PreviewType::Images),
            &[Reader::Native],
            "a picture has one reader, which is the file's own question"
        );

        for (name, kind) in [
            ("drawing.cdr", PreviewType::Libre),
            ("help.chm", PreviewType::Peazip),
        ] {
            let reached = kind_of(Path::new(name), &config);
            assert_eq!(reached, Some(kind), "`{name}` is the {kind:?} kind's");
            assert_eq!(
                chain(kind).len(),
                1,
                "`{name}` has the one reader it has always had: what it gives up is deliberate"
            );
        }
    }
}

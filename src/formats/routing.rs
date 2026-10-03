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
//! What is *not* here is a second answer to the same question. "Who draws this file?" was
//! asked twelve ways across five modules — the router's claim table, the reader chain, the
//! in-app job, six backdrop functions, and a predicate per kind in the window that composes the
//! preview — and a kind left out of one of them was a preview drawn at the wrong scale, which is
//! the failure this module's existence is for. [`resolve`] answers the whole of it in one place
//! as a [`Route`], and every caller that wanted more than the kind reads fields off that
//! instead of reaching for a predicate of its own. Which readers could answer is deliberately
//! still a separate question: it is about the machine and the run rather than about the file,
//! and it costs a registry lookup per engine.
//!
//! Which names each list holds is not written here either: it is `config.ini`'s, it is the user's
//! to edit, and it is a row of [`crate::formats::lists`] — one table every kind's list is a row
//! of, keyed by the section the file writes it under. Every claim below reads its list as that
//! row rather than as a field of the configuration, which is what makes the table the one owner of
//! a list rather than one of two places it is written down: a kind's names are the row's, and what
//! this module owns is the *order* the rows are asked in and what each kind is drawn by — see
//! [`chain`]. A name that sits in two lists is settled by that order rather than by the name,
//! which is why the order has one author.

use crate::config::config::{AppConfig, PreviewType, TransparentBackground};
use crate::formats::{
    archive_formats, audio_formats, lists, native_formats::NativeJob, video_formats,
};
use crate::readers::pdf_preview;
use crate::readers::svg_preview;
use std::path::Path;

mod readers;

pub use readers::{readers_for, Reader};

/// What a claim is asked of: the file, or a name the content has already answered with.
///
/// The difference is worth a type because it is the difference between reading a file and not
/// reading it, and it exists for exactly one pair of names. A `.ts` is a transport stream in
/// one list and TypeScript in another, and which of the two it is, is the file's own content:
/// asked of a file, that content is read; asked of a name the content already spoke for, it is
/// not asked again (see `video_formats::claims_any_video_name`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Asked {
    /// The file itself, whose content has not been read.
    File,
    /// A name, where the caller already holds the content's own answer.
    Name,
}

/// What a pin's own previous/next buttons count as the same kind of file as the one it is
/// showing.
///
/// A [`PreviewType`] is a claim, and claims are per reader: an SVG document is a drawing
/// while the image list still names it, a camera raw is a picture the converter develops, and
/// a `.docx` is a document whether Word or LibreOffice draws it. The buttons the pin carries
/// are not asking what would draw a file, though — they are asking whether a file in the
/// folder is worth stepping onto from the one on screen, and a folder of artwork should step
/// from a `.psd` to a `.svg` rather than past it. So the kinds are folded into the eight
/// things a user would call them (see [`nav_category`]), which is also what the
/// `Pin Mode → Nav File Types → Category` switch means.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum NavCategory {
    Images,
    Video,
    Audio,
    Documents,
    Archives,
    Text,
    Fonts,
    Design,
}

/// The category a kind belongs to, which is what one line of that fold is worth.
///
/// Exhaustive rather than answered with a default: a kind this app grows has to be told
/// where it goes, and the place that says so is the one that will not compile until it has
/// been said. Two groups are decisions rather than accidents. The pictures take the drawings
/// that are pictures in every sense but the list they are named in — a converter's camera
/// raw is a picture to a user, and this app already shares one switch with the pictures
/// (see `PreviewType::enabled_in`). The documents take the ebook, the Office and the render
/// engine's kinds together, and the archives the listing engine's with the ones read here:
/// both are what the file is, whichever program opened it.
pub fn nav_category(kind: PreviewType) -> NavCategory {
    match kind {
        PreviewType::Images | PreviewType::Magick => NavCategory::Images,
        PreviewType::Videos => NavCategory::Video,
        PreviewType::Audio => NavCategory::Audio,
        PreviewType::Ebook | PreviewType::Calibre | PreviewType::Document | PreviewType::Libre => {
            NavCategory::Documents
        }
        PreviewType::Archives | PreviewType::Peazip => NavCategory::Archives,
        PreviewType::Text => NavCategory::Text,
        PreviewType::Fonts => NavCategory::Fonts,
        PreviewType::Design | PreviewType::Vector => NavCategory::Design,
    }
}

/// One kind's claim on a file: a question about one file, with the version of that file already
/// read where reading it is what the question needs.
///
/// The answer is the kind rather than a yes or a no because one list answers with a kind of its
/// own: a drawing is not a picture, and the image list is asked last of all — so a name that list
/// still holds and does not own is answered with the kind it belongs to (see [`claim_images`]).
///
/// The entry is `None` for a question about a name alone, which never touches the disk: the
/// caller has already read the file, and the read that would settle the name is the one that was
/// made. It is a parameter rather than something each claim works out for itself so that one entry
/// read serves all fourteen (see [`claims`]), and it is the whole of what two of the fourteen need
/// it for: the probe's answer is held under the version of the file it was read at, and whether an
/// `.ai`'s content is on this machine is a question about its entry.
///
/// Every claim is a function rather than a row of data, and the reason is that ten of the fourteen
/// are a row's own answer and four are not: a video's two names the text lists share are the
/// file's bytes, a sound is a container a probe found a sound in, a book is a page the PDF reader
/// draws, and a picture is a drawing when the name is one. A table of `kind` plus `list` plus a
/// flag for each of those four would be fourteen arms of data and a dispatch over a five-armed
/// enum, which is more machinery than fourteen one-line functions whose prose says what each is
/// for. The one fact they all share — which row names which kind — is not in them at all: it is
/// [`named_as`], one exhaustive match, because that is the question eleven of the fourteen used to
/// be asked of a list of their own inside a module apiece.
type Claim =
    fn(&Path, &AppConfig, Asked, Option<&crate::formats::head::Facts>) -> Option<PreviewType>;

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
    claim_audio,
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

/// Everything one file's route is, answered once.
///
/// A file's route is one question asked twelve ways. Which kind claims it is the router's answer
/// ([`kind_of`]); what its own bytes say it is, which outranks that answer where they disagree,
/// is `content_type`'s; which reader draws it is [`chain`] narrowed to this machine; what of this
/// app's own reads it is `native_formats`'s job; and which of the three the browser draws is
/// [`DrawnBy`]'s. Each of those was reached from a different module, and a kind left out of one
/// of them was a preview drawn at the wrong scale — which is the defect this struct is for.
///
/// It is a record rather than a thirteenth function: every field is an answer something already
/// had, and what it buys is that a caller asks once and reads fields. Adding a kind is one arm of
/// the match that builds it, in the one file that owns the order, rather than an edit in six
/// predicates that each worked it out for themselves.
///
/// It is [`resolve`] that builds one, and it is deliberately the whole of what this module offers:
/// a caller that wants the kind alone is served by [`kind_of`], which is a lookup in the table
/// rather than the construction of everything above it, because most callers want one field and
/// the fields that cost something — an engine's availability, a reader's job — are not free to
/// ask.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Route {
    /// What the file's own bytes say it is, or nothing where they had no opinion.
    pub content: crate::formats::content_type::Content,
    /// The kind the file's name's lists claim, or nothing where no list does.
    pub named: Option<PreviewType>,
    /// Which half of the `Ebook` kind the file is, of the book list's names.
    pub page: Page,
    /// Which half of the `Vector` kind the file is, of its name.
    pub drawing: Drawing,
    /// Who draws it — the three the browser draws, a media engine, or the kind's own reader.
    pub drawn_by: DrawnBy,
}

/// Who draws a file, in the one exhaustive answer.
///
/// The question was eleven predicates, six of them about the backdrop a preview is composited
/// over and one of them a hit-test on a rectangle (see [`Backdrop`]). What is left is the one
/// thing they were all reaching for: which of this app's own hands the file to.
///
/// The engine kinds are not in it, and that is deliberate rather than a gap: a document, a book an
/// engine converts, an archive an engine lists and a picture an engine develops are drawn by
/// whatever the engine's window is, and which engine is a question about the machine and the
/// run rather than about the file (`readers_for`). What is here is what this file is *handed*,
/// which is what the preview window composes, lays out and shows a wait for.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DrawnBy {
    /// A reader of this app's own draws it into this app's window.
    Native(NativeJob),
    /// The media engine Windows has plays a video in a window of its own.
    MediaEngine,
    /// A card is painted rather than drawn: a sound's facts, a listing, a page of text.
    Card,
    /// The browser engine's window is the preview — an SVG document, a font specimen, a page
    /// of HTML (see [`WebPage`]).
    WebView(WebPage),
    /// Nothing draws it: no kind claims it, or the kind's gate is off.
    Nothing,
}

/// The three things the browser engine draws, which is all `DrawnBy::WebView` can be.
///
/// They are one field rather than three booleans because a file is at most one of them, and the
/// loader asks the question by matching on a kind rather than by consulting three predicates in
/// an order it had to remember.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WebPage {
    /// An SVG document, drawn by the engine at whatever box the layout came out with.
    Svg,
    /// A font specimen: the two tables are read by this app and the glyphs are the engine's.
    FontSpecimen,
    /// A page of HTML, which the engine lays out rather than this side painting.
    Html,
}

/// The backdrop a preview of a file is composited over, and the six tray settings that decide it.
///
/// The six `current_*_background` functions each took the configuration's lock to read one scalar
/// and were asked from two places, so the same picture's backdrop was read twice and a kind's
/// was asked of a predicate that had to remember which three names the engine draws. It is one
/// exhaustive match over the kind here instead, and the paint asks it of the answer it already
/// has.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Backdrop {
    /// A picture's, which is also the one every kind that is none of the others below keeps.
    Image,
    /// A specimen's, which is a page of its own.
    Font,
    /// A texture's: a `.dds`'s alpha is as often a mask or an unfilled channel as transparency.
    Dds,
    /// A design document's, which is the document's own rather than a photograph's.
    Design,
    /// A drawing's: what stands behind the marks is this app's, except where the engine brings
    /// its own page.
    Vector,
    /// A page of HTML's, which is a page the engine brings rather than a transparency.
    Html,
}

/// The backdrop a preview of `kind` is composited over, out of the six the tray keeps apart.
///
/// It is exhaustive rather than answered with a default, and the reason it is worth a type is
/// that it was six functions before it: a kind added to this app had to be added to six places
/// to be drawn over the right thing, and nothing said so. The kinds that share a backdrop do so
/// because a user turns off one thing — a picture and a picture a converter developed, a drawing
/// and a drawing the engine drew (see [`nav_category`] for the fold the pin's own buttons walk by).
pub fn backdrop_of(kind: PreviewType, web: Option<WebPage>) -> Backdrop {
    // The browser's three first, because they are the three that are asked about by name rather
    // than by kind: a font file and a page of HTML are both text-or-drawing by kind, and what
    // stands behind either is the engine's page rather than this app's transparency.
    match web {
        Some(WebPage::FontSpecimen) => return Backdrop::Font,
        Some(WebPage::Html) => return Backdrop::Html,
        Some(WebPage::Svg) => return Backdrop::Vector,
        None => {}
    }

    match kind {
        PreviewType::Images | PreviewType::Magick => Backdrop::Image,
        PreviewType::Fonts => Backdrop::Font,
        PreviewType::Design => Backdrop::Design,
        PreviewType::Vector => Backdrop::Vector,
        // The rest are pages and cards rather than bitmaps: what is on screen for a document is
        // the engine's own window over its own backdrop, and for a text page or a listing or a
        // sound's card it is painted with a background this side never composites through.
        PreviewType::Videos
        | PreviewType::Audio
        | PreviewType::Ebook
        | PreviewType::Archives
        | PreviewType::Peazip
        | PreviewType::Calibre
        | PreviewType::Document
        | PreviewType::Libre
        | PreviewType::Text => Backdrop::Image,
    }
}

/// The backdrop of one of the kinds above, as the configuration has it.
pub fn backdrop_value(backdrop: Backdrop, config: &AppConfig) -> TransparentBackground {
    match backdrop {
        Backdrop::Image => config.image_background,
        Backdrop::Font => config.font_background,
        Backdrop::Dds => config.dds_background,
        Backdrop::Design => config.design_background,
        Backdrop::Vector => config.vector_background,
        Backdrop::Html => config.html_background,
    }
}

/// Everything one file's route is, asked once and answered once.
///
/// It is [`kind_of`] and the four questions beside it, read together so that a caller who wants
/// the whole of a route pays for the parts it has not been asked for once rather than each of
/// them finding the file for itself — and so that the twelve answers cannot drift apart, because
/// there is one place they are written down.
///
/// The entry is a parameter rather than something read here, for the same reason `content_type`
/// takes one: a hover reads a file's directory entry once for a dozen questions and hands it
/// down, and every one of those questions used to read it for itself.
pub fn resolve(
    path: &Path,
    config: &AppConfig,
    probe: &crate::formats::content_type::Probe,
) -> Route {
    let content = crate::formats::content_type::answer(probe, config);
    let named = kind_of_with_facts(path, config, probe.facts());

    let page = named
        .filter(|kind| *kind == PreviewType::Ebook)
        .map(|_| page_of(path, config, probe.facts()))
        .unwrap_or(Page::Comic);

    let drawing = drawing_of(path);
    let drawn_by = drawn_by_of(content, named, page, drawing, path);

    Route {
        content,
        named,
        page,
        drawing,
        drawn_by,
    }
}

/// Who draws a file, from what it is and what it is called.
///
/// It is one match because that is the shape of the question: a file is drawn one way or another,
/// and a kind this function does not mention has no answer. The five terms are the five ways a
/// preview reaches a screen — a reader of this app's own, the media engine, a card this side
/// paints, the browser's own window, or nothing at all — and each is asked of the two answers
/// above it rather than worked out from the file a second time.
///
/// The name is asked for three of them and the bytes for two, which is the same order the loader
/// asks in: what the file's own bytes say outranks what its name says (a `.docx` whose bytes are
/// an MP4 is a video), and a file whose bytes said nothing is its name's.
fn drawn_by_of(
    content: crate::formats::content_type::Content,
    named: Option<PreviewType>,
    page: Page,
    drawing: Drawing,
    path: &Path,
) -> DrawnBy {
    use crate::formats::content_type::Content;

    // The browser's three are asked of both answers, because a drawing is a document by name and
    // a specimen is a specimen whatever a file's bytes are: the browser draws an SVG document,
    // a font file and a page of HTML, and nothing else this app previews is handed to it.
    let web = |named: Option<PreviewType>, content: Content| match named {
        Some(PreviewType::Vector) if drawing == Drawing::Svg => Some(WebPage::Svg),
        Some(PreviewType::Fonts) => Some(WebPage::FontSpecimen),
        Some(PreviewType::Text) if crate::formats::text_formats::is_html_extension(path) => {
            Some(WebPage::Html)
        }
        // A file the bytes named as one of the three is one of them whatever it is called, and a
        // file the bytes named as something else is not one of them however it is named.
        _ => match content {
            Content::Kind(PreviewType::Vector) if drawing == Drawing::Svg => Some(WebPage::Svg),
            Content::Kind(PreviewType::Fonts) => Some(WebPage::FontSpecimen),
            Content::Kind(PreviewType::Text)
                if crate::formats::text_formats::is_html_extension(path) =>
            {
                Some(WebPage::Html)
            }
            _ => None,
        },
    };

    let web = web(named, content);
    if let Some(web) = web {
        return DrawnBy::WebView(web);
    }

    // And then the kind, where the answer is what this app's own reader is. A video is the one
    // that is not: it is played by the media engine in a window of its own, which is why a video
    // preview is a wait rather than a frame (see `video_player`).
    let kind = match content {
        Content::Kind(kind) => Some(kind),
        Content::Foreign => None,
        Content::Unknown => named,
    };

    match kind {
        Some(PreviewType::Videos) => DrawnBy::MediaEngine,
        // The four that are painted into the box the layout planned rather than scaled within
        // it: a listing, a page of text, and a sound's card. A font is not one of them — a
        // specimen is a page of its own size, laid out like a page.
        Some(PreviewType::Archives)
        | Some(PreviewType::Peazip)
        | Some(PreviewType::Text)
        | Some(PreviewType::Audio) => DrawnBy::Card,
        Some(PreviewType::Images) => DrawnBy::Native(NativeJob::Picture),
        Some(PreviewType::Design) => DrawnBy::Native(NativeJob::Project),
        Some(PreviewType::Vector) => match drawing {
            Drawing::Svg => DrawnBy::WebView(WebPage::Svg),
            Drawing::Replayed => DrawnBy::Native(NativeJob::Metafile),
        },
        Some(PreviewType::Ebook) => DrawnBy::Native(match page {
            Page::Pdf => NativeJob::Pdf,
            Page::Comic => NativeJob::Comic,
        }),
        // A specimen is reached here only as itself — a file whose bytes are a font's and whose
        // name is not, which the browser still draws.
        Some(PreviewType::Fonts) => DrawnBy::WebView(WebPage::FontSpecimen),
        // The four kinds an engine draws, and nothing at all: their answer is the engine's own
        // window, which is not a kind of reader of this app's own, and a file no list claims is
        // the picture the loader's own chain ends at.
        Some(PreviewType::Calibre)
        | Some(PreviewType::Document)
        | Some(PreviewType::Libre)
        | Some(PreviewType::Magick)
        | None => DrawnBy::Nothing,
    }
}

/// The kind that claims `path`, or nothing where no list does.
///
/// This is the question a hover asks: the file's own content is not in hand, so the two names
/// the video list shares with the text lists are settled by reading the file (see [`Asked`]).
/// The caller must not be holding the configuration itself — nothing here takes that lock
/// again, and every reader that would is handed the configuration instead.
///
/// A caller that has read the file's own entry already — a hover, which reads it for half a
/// dozen questions of its own — asks [`kind_of_with_facts`] instead, and pays no second
/// `fs::metadata` for the same file. A caller that wants the whole of a file's route asks
/// [`resolve`], which is this question and the four beside it read together.
pub fn kind_of(path: &Path, config: &AppConfig) -> Option<PreviewType> {
    claims(path, config, Asked::File, None)
}

/// The same question, of a file whose directory entry the caller has already read.
///
/// It is [`kind_of`] with the one thing it would have read for itself handed in, and it exists
/// because that thing is an `fs::metadata` on the thread that pumps this window's messages: a
/// hover reads the entry once for the layout, the loader and the eleven predicates that all ask
/// what a file is, and each of them was reading it again for itself (see
/// [`crate::formats::content_type::Probe`]).
pub fn kind_of_with_facts(
    path: &Path,
    config: &AppConfig,
    facts: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    claims(path, config, Asked::File, facts)
}

/// The kind a name belongs to, where the file's content has already answered for it.
///
/// It is the form `content_type` asks in: what it holds is a list of names a signature
/// answered with — a Matroska file's `mkv` and `webm` — and those names have to be turned back
/// into a kind. A name asked this way is asked of the lists alone, because the read that would
/// settle it has already been made on the file it came from.
pub fn kind_of_name(path: &Path, config: &AppConfig) -> Option<PreviewType> {
    claims(path, config, Asked::Name, None)
}

/// Whether `kind`'s own list claims `path` — the question asked of one kind rather than of the
/// order, which is what the eleven `is_<kind>_file` functions each answered for themselves.
///
/// The difference from [`kind_of`] is the whole of it: this asks one row, so a name two rows hold
/// is both kinds' here and only the earlier one's there. A layout that measures each kind in turn
/// needs that — a `.cdr` is a drawing to this app's own reader and a page to the render engine,
/// and both measurements are wanted — while the loader, which wants one answer, asks
/// [`kind_of`].
///
/// It is exhaustive over the kinds rather than answered with a default, for the reason
/// [`nav_category`] is: a kind this app grows has to be told which rows are its own, and the one
/// place that says so is the one that will not compile until it has.
///
/// **The video kind is the one that reads the file.** Its two rows share `ts` and `mts` with the
/// text lists, so a name either of them carries is settled by whether the file holds MPEG-TS
/// packets — a `File::open` — and a caller that is holding the configuration's lock must not ask
/// this of a video. Every other kind's row is a comparison in memory.
pub fn named_as(path: &Path, config: &AppConfig, kind: PreviewType) -> bool {
    match kind {
        // The two video lists together, because which of them carries a name is which engine
        // plays it and not what the file is (see `video_formats::matches_any_video_list`).
        PreviewType::Videos => video_formats::matches_any_video_list(path, config),
        // The text kind's two rows: an extension, or a whole name for a repository file that has
        // none — which is what `matches_text_lists` used to be, before the two rows became the
        // only place either list is read from.
        PreviewType::Text => lists::TEXT.claims(path, config) || lists::NAMES.claims(path, config),
        PreviewType::Audio => lists::AUDIO.claims(path, config),
        PreviewType::Ebook => lists::EBOOK.claims(path, config),
        PreviewType::Archives => archive_formats::claims_in(path, lists::ARCHIVE.entries(config)),
        PreviewType::Peazip => lists::PEAZIP.claims(path, config),
        PreviewType::Calibre => lists::CALIBRE.claims(path, config),
        PreviewType::Document => lists::OFFICE.claims(path, config),
        PreviewType::Libre => lists::LIBRE.claims(path, config),
        PreviewType::Magick => lists::MAGICK.claims(path, config),
        PreviewType::Design => lists::DESIGN.claims(path, config),
        PreviewType::Vector => lists::VECTOR.claims(path, config),
        PreviewType::Fonts => lists::FONT.claims(path, config),
        PreviewType::Images => lists::IMAGE.claims(path, config),
    }
}

/// Whether a preview of `path` may be shown as `kind`: its own list claims it, and the tray has
/// that kind switched on.
///
/// It is the question the fourteen `is_<kind>_preview` predicates each asked of a list of their
/// own and a switch of their own, and the two halves of it are one question because a kind turned
/// off leaves its list exactly as it is and turning it back on restores it — which is what a
/// user's own edit to a list and the tray's own switch are: two switches over one kind.
///
/// The configuration is a parameter rather than read here, for the reason it is one everywhere
/// else in this layer: a caller that has it in hand pays no lock, and one that has not must not
/// hold the one it takes across the video kind's read (see [`named_as`]).
pub fn previewed_as(path: &Path, config: &AppConfig, kind: PreviewType) -> bool {
    named_as(path, config, kind) && kind.enabled_in(config)
}

fn claims(
    path: &Path,
    config: &AppConfig,
    asked: Asked,
    entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    // One reading of the directory entry for the whole walk, because the video claim below
    // needs the version of the file to look the probe's answer up. Every claim that reads
    // anything else reads it from `config`, and the one claim that opens the file is settled
    // by `asked` rather than here.
    //
    // `Asked::Name` is a question about a name and never touches the disk at all: the caller
    // has already read the file, and the read that would settle the name is the one that was
    // made. Asking for the entry anyway cost a `fs::metadata` on a synthetic path, which is a
    // metadata call for a file that is not there.
    let read;
    let entry = match asked {
        Asked::File => match entry {
            Some(entry) => Some(entry),
            None => {
                read = crate::formats::head::Facts::read(path);
                read.as_ref()
            }
        },
        // A file that is not there has no version to key the probe's answer by, and the two
        // claims that ask it fall back to asking of the name — which is the same miss a file
        // that is not there gets either way.
        Asked::Name => None,
    };

    CLAIMS
        .iter()
        .find_map(|claim| claim(path, config, asked, entry))
}

/// A video: the configured list's names, and the two of them it shares with the text lists
/// settled by the file rather than by the name where the file is the thing being asked about.
///
/// The one thing that takes a video away from this claim is the probe: a container of a name
/// this list carries — an `.mp4`, an `.mka`, an `.ogg` — whose own streams turned out to hold
/// a sound and no picture is not a video at all, and the sound claim below is where it is
/// answered instead (see `audio_formats::probed_audio_only`).
fn claim_video(
    path: &Path,
    config: &AppConfig,
    asked: Asked,
    entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    // Asked of the key the walk already has, and of the path only where there is none. The
    // probe's answer is a set in memory, so the entry that keys it is the only thing the
    // question needs from the disk, and the walk has read it once for all fourteen claims.
    let probed = match entry.map(crate::formats::head::Facts::key) {
        Some(key) => audio_formats::probed_audio_only_in(key),
        None => audio_formats::probed_audio_only(path),
    };
    if probed {
        return None;
    }

    let claimed = match asked {
        Asked::File => video_formats::matches_any_video_list(path, config),
        Asked::Name => video_formats::claims_any_video_name(path, config),
    };

    claimed.then_some(PreviewType::Videos)
}

/// A sound: a name the `[audio]` list carries, or a container of a name another list carries
/// that a probe found a sound in and no picture — see `audio_formats` for the list and for the
/// verdict, and `audio_track` for the probe that reaches it.
///
/// It is asked beside the video claim rather than after the kinds below it, because the two
/// are the pair that share containers: what a `.mka` is, is whichever of the two its streams
/// say it is, and nothing else in the table has an opinion about one.
fn claim_audio(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    let named = lists::AUDIO.claims(path, config);
    // Of the key the walk already has where there is one: this is the second of the two claims
    // that asks the probe's answer, and the entry it was keyed by was read once for both.
    let probed = match entry.map(crate::formats::head::Facts::key) {
        Some(key) => audio_formats::probed_audio_only_in(key),
        None => audio_formats::probed_audio_only(path),
    };

    (named || probed).then_some(PreviewType::Audio)
}

/// Which half of the `Ebook` kind a file is: a page the PDF engine reads, or a comic, which
/// is the first plate out of the container it is published in.
///
/// The two are one kind because what they are shown as is one thing — a page — and what a user
/// turns off for either is books, which is what makes them one switch and one kind. The reader
/// that draws one is not the same for the two, so the half has to be settled by somebody: this
/// is that somebody, because the loader used to ask the same question a second time for itself
/// and the answer cost a read of the file (see `native_formats::page_job`).
///
/// The names are asked of the row in hand rather than of the gate that reads it, since the caller
/// holds the configuration already. Three names make a page outright and an Illustrator document
/// carrying a PDF inside it makes one of its own bytes, which is the only half here that is not
/// the name's to answer (see `pdf_preview::is_pdf_file_in_of`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Page {
    /// A page: a PDF, or a document of another name carrying one, which the PDF engine reads.
    Pdf,
    /// A comic, which is the container read for its first plate rather than for the book.
    Comic,
}

/// The half of the `Ebook` kind `path` is, of the list in hand.
///
/// A file whose own name the book list does not carry is not a book of either half, so this is
/// asked of the page's own question and answered as the comic for everything else rather than
/// as nothing: which of the two a file is cannot be asked without knowing it is one of them.
///
/// The entry is `None` where the caller has none, and that is the only question here that
/// touches the disk at all: an `.ai` is a page by what it keeps at an offset, and whether its
/// content is on this machine is a question about the directory entry rather than about the
/// name — which is `pdf_preview`'s own fix generalised to the walk that asks it (see
/// [`pdf_preview::is_pdf_file_in_of`]).
pub fn page_of(
    path: &Path,
    config: &AppConfig,
    entry: Option<&crate::formats::head::Facts>,
) -> Page {
    let list = lists::EBOOK.entries(config);
    let is_page = match entry {
        Some(facts) => pdf_preview::is_pdf_file_in_of(path, list, facts.needs_download()),
        None => pdf_preview::is_pdf_file_in(path, list),
    };

    if is_page {
        Page::Pdf
    } else {
        Page::Comic
    }
}

/// The same question of a file with no entry in hand, which is the form a caller with no hover
/// facts of its own asks.
/// A page this app draws itself: a PDF, and a comic, which is the other half of the same kind.
///
/// Which of the two a file is is [`page_of`]'s answer rather than this entry's, because the
/// reader that draws the one is not the reader that draws the other and the loader asks for a
/// job rather than a kind.
fn claim_ebook(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    let book = lists::EBOOK.claims(path, config);

    (page_of(path, config, entry) == Page::Pdf || book).then_some(PreviewType::Ebook)
}

/// An archive this app reads itself, which is a listing rather than an engine's work.
fn claim_archives(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::ARCHIVE
        .claims(path, config)
        .then_some(PreviewType::Archives)
}

/// An archive a listing engine reads, asked beside the archive list above it: a name in that
/// list is read by this app itself, and one in this list is read by an engine. A name in both
/// is the archive list's, since that list is asked first.
fn claim_peazip(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::PEAZIP
        .claims(path, config)
        .then_some(PreviewType::Peazip)
}

/// A book a conversion engine reads: a name no reader here opens, and no list above claims.
fn claim_calibre(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::CALIBRE
        .claims(path, config)
        .then_some(PreviewType::Calibre)
}

/// A document shown as a page: the names the application that owns the format draws, and the
/// names a render engine draws because no application of that family is installed. Which of
/// the two draws a file is the machine's answer rather than the list's (see [`chain`]).
fn claim_document(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::OFFICE
        .claims(path, config)
        .then_some(PreviewType::Document)
}

/// A document only a render engine draws, asked after the Office list because a name can sit
/// in both: what such a file keeps of itself is a thumbnail, and a thumbnail is not a preview.
fn claim_libre(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::LIBRE
        .claims(path, config)
        .then_some(PreviewType::Libre)
}

/// A picture an image converter develops — a camera raw above all — asked before the design
/// list, which it is a neighbour of rather than a member of.
fn claim_magick(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::MAGICK
        .claims(path, config)
        .then_some(PreviewType::Magick)
}

/// A design document: a kind of its own, whose preview is made of the picture the file keeps
/// of itself rather than of a decoder its name names.
fn claim_design(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::DESIGN
        .claims(path, config)
        .then_some(PreviewType::Design)
}

/// A drawing: a metafile, an encapsulated PostScript file, or an SVG document.
fn claim_vector(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::VECTOR
        .claims(path, config)
        .then_some(PreviewType::Vector)
}

/// Text: the extension lists and the names list, which is what makes a `makefile` a text file
/// and what the dot-file rule is for (see `text_formats::lookup_extension`).
fn claim_text(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    (lists::TEXT.claims(path, config) || lists::NAMES.claims(path, config))
        .then_some(PreviewType::Text)
}

/// A font, asked ahead of the image list that would have turned a `.ttf` down: what draws one
/// is the browser engine rather than a decoder.
fn claim_fonts(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    lists::FONT
        .claims(path, config)
        .then_some(PreviewType::Fonts)
}

/// Which half of the `Vector` kind a file is: an SVG document the browser engine draws, or a
/// Windows metafile or an encapsulated PostScript file the drawing layer replays.
///
/// The two are one kind because a drawing is a drawing whichever way it is drawn, and which of
/// the two is settled by the name in both directions: an `svg` a hand-edited image list still
/// names is a drawing rather than the picture its name says (see [`claim_images`]), and a
/// metafile is the drawing layer's rather than the browser engine's whatever engine is
/// installed (see `native_formats::drawing_job`). Both of those used to ask the name for
/// themselves, which is one question with two owners and a drift nobody would see — a
/// metafile answered as an SVG document is a window the browser never fills.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Drawing {
    /// An SVG document, including the gzipped spelling of the same document, which is drawn by
    /// the browser engine in a window of its own.
    Svg,
    /// A metafile or an encapsulated PostScript file, replayed by the drawing layer into a
    /// frame this app composites itself.
    Replayed,
}

/// The half of the `Vector` kind `path` is, which is its name and nothing else.
///
/// The name is the whole of it, and deliberately so: an `svg` that is not a document is a
/// drawing the browser refuses rather than a metafile, and the renderer has the last word on
/// whether a document is a document at all (see `svg_preview::is_svg_file`). Reading the front
/// of the file here would settle the wrong half of the wrong question.
pub fn drawing_of(path: &Path) -> Drawing {
    if svg_preview::is_svg_file(path) {
        Drawing::Svg
    } else {
        Drawing::Replayed
    }
}

/// A picture — and the one entry that can answer with a kind that is not its own.
///
/// A drawing is not a picture, and the image list is where the drawings of that kind were until
/// they were given a kind of their own: an `svg` a hand-edited image list still names is
/// answered as the drawing it is rather than as the picture its name says. It is asked inside
/// this entry rather than as an entry of its own so that the order the other kinds are asked in
/// is not disturbed by it — which is the one thing that decides which of the two halves of a
/// drawing a name in this list is, so it is [`drawing_of`]'s answer and not a second reading.
fn claim_images(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    if !lists::IMAGE.claims(path, config) {
        return None;
    }

    Some(if drawing_of(path) == Drawing::Svg {
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

        // A sound is the other way round, and deliberately: the engine Windows has plays it
        // inside this app's own process — nothing to start, nothing to supervise, and the
        // position it reports is the position the card draws — while FFmpeg's player is the
        // answer for the formats that engine has no decoder for. Which of the two a file is
        // is the probe's own answer rather than this chain's (see `audio_track`).
        PreviewType::Audio => &[Reader::Native, Reader::Ffmpeg],

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

#[cfg(test)]
mod tests;

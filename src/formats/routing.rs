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
//! The lists themselves are not here: they are `config.ini`'s, and which names each one holds
//! is the user's to edit. What this module owns is the *order* they are asked in, and what
//! each kind is drawn by — see [`chain`]. A name that sits in two lists is settled by that
//! order rather than by the name, which is why the order has one author.

use crate::config::config::{AppConfig, PreviewType, TransparentBackground};
use crate::engines::{
    calibre_render, imagemagick_render, libreoffice_render, peazip_render, webview_preview,
};
use crate::formats::{
    archive_formats, audio_formats, calibre_formats, codecs, design_formats, ebook_formats,
    font_formats, image_formats, libre_formats, magick_formats, native_formats::NativeJob,
    office_formats, peazip_formats, text_formats, vector_formats, video_formats,
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

/// One kind's claim on a file, and the kind it answers with.
///
/// The answer is the kind rather than a yes or a no because one list answers with a kind of
/// its own: a drawing is not a picture, and the image list is asked last of all — so a name
/// that list still holds and does not own is answered with the kind it belongs to. See
/// [`claim_images`].
/// A question about one file, with the version of that file already read where reading it is
/// what the question needs.
///
/// The entry is `None` for a question about a name alone, which never touches the disk: the
/// caller has already read the file, and the read that would settle the name is the one that
/// was made. It is a parameter rather than something each claim works out for itself so that
/// one entry read serves all fourteen (see `claims`), and it is the whole of what two of the
/// fourteen need it for: the probe's answer is held under the version of the file it was read
/// at, and whether an `.ai`'s content is on this machine is a question about its entry.
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
    let named = audio_formats::matches_audio_list(path, &config.audio_extensions);
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
/// The names are asked of the list in hand rather than of the gate that reads it, since the
/// caller holds the configuration already. Three names make a page outright and an Illustrator
/// document carrying a PDF inside it makes one of its own bytes, which is the only half here
/// that is not the name's to answer (see `pdf_preview::is_pdf_file_in`).
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
    let is_page = match entry {
        Some(facts) => {
            pdf_preview::is_pdf_file_in_of(path, &config.ebook_extensions, facts.needs_download())
        }
        None => pdf_preview::is_pdf_file_in(path, &config.ebook_extensions),
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
    let book = ebook_formats::matches_ebook_list(path, &config.ebook_extensions);

    (page_of(path, config, entry) == Page::Pdf || book).then_some(PreviewType::Ebook)
}

/// An archive this app reads itself, which is a listing rather than an engine's work.
fn claim_archives(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    archive_formats::matches_archive_list(path, &config.archive_extensions)
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
    peazip_formats::matches_peazip_list(path, &config.peazip_extensions)
        .then_some(PreviewType::Peazip)
}

/// A book a conversion engine reads: a name no reader here opens, and no list above claims.
fn claim_calibre(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    calibre_formats::matches_calibre_list(path, &config.calibre_extensions)
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
    office_formats::matches_office_list(path, &config.office_extensions)
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
    libre_formats::matches_libre_list(path, &config.libre_extensions).then_some(PreviewType::Libre)
}

/// A picture an image converter develops — a camera raw above all — asked before the design
/// list, which it is a neighbour of rather than a member of.
fn claim_magick(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    magick_formats::matches_magick_list(path, &config.magick_extensions)
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
    design_formats::matches_design_list(path, &config.design_extensions)
        .then_some(PreviewType::Design)
}

/// A drawing: a metafile, an encapsulated PostScript file, or an SVG document.
fn claim_vector(
    path: &Path,
    config: &AppConfig,
    _asked: Asked,
    _entry: Option<&crate::formats::head::Facts>,
) -> Option<PreviewType> {
    vector_formats::matches_vector_list(path, &config.vector_extensions)
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
    text_formats::matches_text_lists(path, &config.text_extensions, &config.text_names)
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
    font_formats::matches_font_list(path, &config.font_extensions).then_some(PreviewType::Fonts)
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
    if !image_formats::matches_image_list(path, &config.image_extensions) {
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

        // A video is the two lists together: the kind is what a file is, and which of the two
        // lists its name is in is only which engine plays it.
        let mut video_names = extensions(&config.video_extensions);
        video_names.extend(extensions(&config.ffmpeg_extensions));

        let mut lists = vec![
            (PreviewType::Videos, video_names),
            (PreviewType::Audio, extensions(&config.audio_extensions)),
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
        // A `.mpc` is a Musepack sound and the persistent cache ImageMagick keeps for itself,
        // and the two are told apart from the front of the file: an image cache opens with the
        // `id=ImageMagick` the converter's own reader answers for (see the signature that names
        // `miff`), and a Musepack file with its own `MPCK` or `MP+`. A name-level answer is
        // what this table settles, and it is the sound's, because the sound list is asked
        // first — a file of neither shape is the one the order gets wrong, and that is a file
        // neither reader could draw.
        ("preview.mpc", PreviewType::Audio),
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

    /// The half of the `Ebook` kind is the one answer two tables used to work out for
    /// themselves, and the drift it can suffer is a file whose kind is the book's and whose
    /// reader is the comic's — which nothing above either would notice, because the kind was
    /// right and only the page underneath it was wrong.
    ///
    /// It is asserted over every name the book list ships rather than over a sample, and both
    /// ways round: each name reaches the half its own spelling settles, and the two spellings
    /// that are a page and a comic of the same kind are told apart from each other. The first
    /// alone would pass a table that answered every name with a comic.
    #[test]
    fn every_shipped_book_name_reaches_the_half_its_spelling_settles() {
        let config = AppConfig::default();

        // The names the shipped list holds that are a page and a comic, read off the list
        // rather than written here, so a name added to either list is covered by this without
        // being added to this.
        let expected: Vec<(String, Page)> = config
            .ebook_extensions
            .iter()
            .map(|name| {
                let page = matches!(name.trim_start_matches('.'), "pdf" | "pdfa" | "epdf");
                (name.clone(), if page { Page::Pdf } else { Page::Comic })
            })
            .collect();

        assert!(
            expected.len() >= 2,
            "the book list has to carry a page and a comic for this to be worth saying"
        );

        let mut wrong: Vec<String> = Vec::new();
        let mut saw_page = false;
        let mut saw_comic = false;

        for (name, expected) in &expected {
            let path = PathBuf::from(format!("book.{name}"));
            let reached = page_of(&path, &config, None);

            match reached {
                Page::Pdf => saw_page = true,
                Page::Comic => saw_comic = true,
            }

            if reached != *expected {
                wrong.push(format!(
                    "`{name}` is a page's spelling, read as {reached:?}"
                ));
            }

            // And the kind the router reaches is the book's either way: the half is what
            // picks the reader, not whether the file is a book at all.
            assert_eq!(
                kind_of(&path, &config),
                Some(PreviewType::Ebook),
                "`{name}` is the book kind's whichever half it is"
            );
        }

        assert!(
            saw_page && saw_comic,
            "the list has to hold both halves or this proves nothing about telling them apart"
        );
        assert!(
            wrong.is_empty(),
            "a book's spelling read as the wrong half of the kind:\n  {}",
            wrong.join("\n  ")
        );
    }

    /// The half of a drawing is the second answer two tables used to work out for themselves,
    /// and it is asked in both directions: the image list asks it to reach the drawing's kind
    /// rather than the picture's, and the loader asks it to pick the reader. A name the two
    /// answered differently about is a metafile handed to the browser engine, so the assertion
    /// is that both directions agree for every shipped name — and that a name reaching either
    /// kind from the *other* list is still the drawing it is.
    #[test]
    fn every_shipped_drawing_name_reaches_one_half_and_the_list_agrees() {
        let config = AppConfig::default();

        let mut wrong: Vec<String> = Vec::new();
        let mut saw_svg = false;
        let mut saw_replayed = false;

        // The drawings this app ships, read off the two lists that hold them rather than
        // written here, so a name added to either is covered by this without being added here.
        let shipped: Vec<(String, PreviewType)> = config
            .vector_extensions
            .iter()
            .map(|name| (name.clone(), PreviewType::Vector))
            .chain(
                config
                    .image_extensions
                    .iter()
                    .map(|name| (name.clone(), PreviewType::Images)),
            )
            .collect();

        for (name, listed_as) in shipped {
            let path = PathBuf::from(format!("drawing.{name}"));
            let half = drawing_of(&path);
            let reached = kind_of(&path, &config);

            let (expected_half, expected_kind) = match half {
                Drawing::Svg => {
                    saw_svg = true;
                    (Drawing::Svg, PreviewType::Vector)
                }
                Drawing::Replayed => {
                    saw_replayed = true;
                    (Drawing::Replayed, listed_as)
                }
            };

            assert_eq!(
                half, expected_half,
                "`{name}`'s own half, which is what both tables read"
            );

            if let Some(reached) = reached {
                if reached != expected_kind {
                    wrong.push(format!(
                        "`{name}` is the {expected_kind:?} kind's half, reached {reached:?}"
                    ));
                }
            }

            // A drawing is reached from the vector list, and from the image list only where a
            // hand-edited entry left it there — never the other way round, which is the whole
            // of what the image entry is for.
            if reached == Some(PreviewType::Vector) && listed_as == PreviewType::Images {
                wrong.push(format!(
                    "`{name}` was reached as a drawing through the image list, which is the \
                     one direction that order allows"
                ));
            }
        }

        assert!(
            saw_svg && saw_replayed,
            "the shipped names have to carry both halves or this proves nothing about telling \
             them apart"
        );
        assert!(
            wrong.is_empty(),
            "a drawing's name reached a kind its half is not:\n  {}",
            wrong.join("\n  ")
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

    /// The fold the pin's own buttons walk by is the one a user would draw: a picture a
    /// converter develops is a picture, a book a converter converted and a page an engine
    /// drew are both a document, and a drawing is a drawing whether it was reached through
    /// the image list or the vector one.
    #[test]
    fn a_kind_folds_into_the_thing_a_user_would_call_it() {
        for (kind, expected) in [
            (PreviewType::Images, NavCategory::Images),
            (PreviewType::Magick, NavCategory::Images),
            (PreviewType::Videos, NavCategory::Video),
            (PreviewType::Audio, NavCategory::Audio),
            (PreviewType::Ebook, NavCategory::Documents),
            (PreviewType::Calibre, NavCategory::Documents),
            (PreviewType::Document, NavCategory::Documents),
            (PreviewType::Libre, NavCategory::Documents),
            (PreviewType::Archives, NavCategory::Archives),
            (PreviewType::Peazip, NavCategory::Archives),
            (PreviewType::Text, NavCategory::Text),
            (PreviewType::Fonts, NavCategory::Fonts),
            (PreviewType::Design, NavCategory::Design),
            (PreviewType::Vector, NavCategory::Design),
        ] {
            assert_eq!(nav_category(kind), expected, "{kind:?} is a {expected:?}");
        }
    }

    /// A kind and the kind it shares a switch with are one category, always: the two answers
    /// `All` and `Category` give a folder are supposed to differ, and a pair of kinds the tray
    /// cannot switch apart is a difference they cannot have.
    #[test]
    fn the_kinds_that_share_a_switch_share_a_category() {
        let config = AppConfig::default();
        let kinds = [
            PreviewType::Images,
            PreviewType::Magick,
            PreviewType::Videos,
            PreviewType::Audio,
            PreviewType::Ebook,
            PreviewType::Calibre,
            PreviewType::Document,
            PreviewType::Libre,
            PreviewType::Archives,
            PreviewType::Peazip,
            PreviewType::Text,
            PreviewType::Fonts,
            PreviewType::Design,
            PreviewType::Vector,
        ];

        // The whole set is switched off one kind at a time, and every kind whose own gate
        // went down with it is the kind that shares its switch — which is the fold's own
        // table, read back off the gates rather than off the arms that wrote it.
        for kind in kinds {
            let mut probe = config.clone();
            probe.image_preview_enabled = false;
            probe.video_preview_enabled = false;
            probe.audio_preview_enabled = false;
            probe.text_preview_enabled = false;
            probe.ebook_preview_enabled = false;
            probe.archive_preview_enabled = false;
            probe.document_preview_enabled = false;
            probe.font_preview_enabled = false;
            probe.design_preview_enabled = false;
            probe.vector_preview_enabled = false;
            kind.set_enabled_in(&mut probe, true);

            let switched = kind.enabled_in(&probe);
            assert!(
                switched,
                "{kind:?} is switched back on by the switch it is written for"
            );

            // What one switch turns on is one category: the tray can only narrow a walk by
            // something it can also switch, so a gate that brought up a kind of another
            // category would put a file in a walk the user had turned off.
            for other in kinds {
                if other.enabled_in(&probe) {
                    assert_eq!(
                        nav_category(other),
                        nav_category(kind),
                        "{other:?} is switched by {kind:?}'s gate, and so has to be walked with it"
                    );
                }
            }
        }
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

    /// A sound is reached by its name, and a container whose streams a probe found a sound in
    /// and no picture is reached by that verdict — the one thing a list cannot say, and the
    /// one thing that takes a file away from the video claim above it.
    #[test]
    fn a_sound_is_the_name_it_carries_or_the_streams_a_probe_found() {
        let config = AppConfig::default();

        for name in [
            "track.flac",
            "song.mp3",
            "book.m4b",
            "radio.mka",
            "album.opus",
        ] {
            assert_eq!(
                kind_of(Path::new(name), &config),
                Some(PreviewType::Audio),
                "`{name}` is a sound"
            );
        }

        assert_eq!(
            chain(PreviewType::Audio),
            &[Reader::Native, Reader::Ffmpeg],
            "a sound is played inside this app's own process before FFmpeg's player is asked"
        );

        // A container the video list claims, whose own streams the probe found no picture in:
        // the verdict is what answers it, and it is remembered under the file it was read from
        // rather than under its name — so the path here is one no other test writes.
        let probed = Path::new("routing-test-audio-only-container.mkv");
        assert_eq!(
            kind_of(probed, &config),
            Some(PreviewType::Videos),
            "a container of a video's name is a video until a probe says otherwise"
        );

        crate::formats::audio_formats::remember_audio_only(probed);
        assert_eq!(
            kind_of(probed, &config),
            Some(PreviewType::Audio),
            "and a sound once its own streams have answered for it"
        );
        assert_eq!(
            kind_of_name(probed, &config),
            Some(PreviewType::Audio),
            "the same answer through the name half of the question, which is the form the \
             content tier asks in"
        );
    }

    /// One route and the four questions it is made of are the same four answers, for every
    /// name this app ships.
    ///
    /// `resolve` exists so that a caller asks once and reads fields, and what that is worth is
    /// only true if the record cannot say something the functions it replaced would have
    /// disagreed with — so this asks both, over the shipped lists rather than over a sample, and
    /// compares. The kind and the half of each kind are the two that were reached from different
    /// modules, which is how a book came to be drawn by the reader the comic's or a drawing by
    /// the browser engine's: the kind was right and the reader was another kind's (see
    /// [`every_shipped_drawing_name_reaches_one_half_and_the_list_agrees`]).
    ///
    /// Who draws it is asserted too, since that is the answer none of the four has on its own:
    /// a kind this app ships is drawn by something for every name, and the four engine kinds
    /// are the ones whose answer is `Nothing` — the engine's own window is not a reader of this
    /// app's, which is what `chain` is for (see [`DrawnBy`]).
    #[test]
    fn a_route_says_what_the_four_questions_it_is_made_of_say() {
        let config = AppConfig::default();

        for (kind, paths) in shipped_lists(&config) {
            for path in paths {
                let probe = crate::formats::content_type::Probe::read(&path);
                let route = resolve(&path, &config, &probe);

                assert_eq!(
                    route.named,
                    kind_of(&path, &config),
                    "`{}` is reached as one kind here and another by the router",
                    path.display()
                );
                assert_eq!(
                    route.drawing,
                    drawing_of(&path),
                    "`{}` is one half of the drawing kind here and another by the name",
                    path.display()
                );

                if route.named == Some(PreviewType::Ebook) {
                    assert_eq!(
                        route.page,
                        page_of(&path, &config, probe.facts()),
                        "`{}` is one half of the book kind here and another by the list",
                        path.display()
                    );
                } else {
                    assert_eq!(
                        route.page,
                        Page::Comic,
                        "`{}` is not a book, so which half of one it is has no answer to be \
                         wrong about",
                        path.display()
                    );
                }

                // The four engine kinds and nothing else are drawn by an engine's own window
                // rather than by a reader of this app's, so they are the four arms that answer
                // `Nothing` — the only way a claimed kind has no hand to be drawn by.
                let engine_kind = matches!(
                    route.named,
                    Some(PreviewType::Document)
                        | Some(PreviewType::Libre)
                        | Some(PreviewType::Calibre)
                        | Some(PreviewType::Magick)
                );

                assert_eq!(
                    route.drawn_by == DrawnBy::Nothing,
                    engine_kind,
                    "`{}` is the {kind:?} list's, and its reader is {}",
                    path.display(),
                    if engine_kind {
                        "an engine's own window rather than a reader of this app's"
                    } else {
                        "a reader of this app's own"
                    }
                );
            }
        }
    }
}

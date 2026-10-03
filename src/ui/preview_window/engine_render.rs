//! The renderers this app asks for a picture of a file it cannot draw itself, and the route a
//! file takes to one of them: the kind it is read as, whether it is painted already, and which
//! engine answers for it.

use super::*;

/// The backdrop a preview of one of the six is drawn over, out of the one configuration.
///
/// It was six functions, each of which took the lock for one scalar, and each of which was asked
/// from one place that had to remember which of the six a file fell in. `routing::Backdrop` is
/// the one exhaustive answer and this is the one read that turns it into a colour — see
/// `routing::backdrop_of` for the judgement behind each of the six, and for why a texture and a
/// design document keep a backdrop a picture does not.
pub(super) fn current_background(
    backdrop: crate::formats::routing::Backdrop,
) -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| crate::formats::routing::backdrop_value(backdrop, &cfg))
        .unwrap_or(match backdrop {
            crate::formats::routing::Backdrop::Image => DEFAULT_IMAGE_BACKGROUND,
            crate::formats::routing::Backdrop::Font => DEFAULT_FONT_BACKGROUND,
            crate::formats::routing::Backdrop::Dds => DEFAULT_DDS_BACKGROUND,
            crate::formats::routing::Backdrop::Design => DEFAULT_DESIGN_BACKGROUND,
            crate::formats::routing::Backdrop::Vector => DEFAULT_VECTOR_BACKGROUND,
            crate::formats::routing::Backdrop::Html => DEFAULT_HTML_BACKGROUND,
        })
}

/// Which of the three the browser draws, for a file whose own bytes have not already said.
///
/// It is asked of the name because that is what these three are told apart by, and it is one
/// function rather than three predicates because the loader asks the same question when it hands
/// the hover over, the layout asks it when it decides the size, and the paint asks it when it
/// picks the backdrop — three places that had each been picking out of the same three names.
pub(super) fn web_page_of(path: &Path) -> Option<crate::formats::routing::WebPage> {
    use crate::formats::routing::WebPage;

    if svg_preview::is_svg_file(path) {
        Some(WebPage::Svg)
    } else if named_as(path, PreviewType::Fonts) {
        Some(WebPage::FontSpecimen)
    } else if crate::formats::text_formats::is_html_extension(path) {
        Some(WebPage::Html)
    } else {
        None
    }
}

/// The backdrop an engine-drawn preview of `path` is drawn over: the kind decides it, the same
/// way it decides everything else about a document. The one engine draws all the kinds this app
/// hands it, and each of the three answers for itself — a font file's specimen is a page of its
/// own, an SVG document is a vector drawing, and a page of HTML is a page, so the backdrop the
/// tray keeps for the kind is the one it is given.
///
/// It asks the router's one answer for which of the three a file is rather than repeating the
/// three-name test, so the backdrop, the layout's size and the loader's hand-over cannot be
/// three readings of the same question (see `routing::WebPage`).
pub(super) fn engine_background(path: &Path) -> TransparentBackground {
    current_background(crate::formats::routing::backdrop_of(
        PreviewType::Vector,
        web_page_of(path),
    ))
}

/// The kind of engine-drawn preview `path` would get, when it is one of the three the browser
/// draws: a document, a font file's specimen, or a page of HTML.
///
/// It stands in for the file's name where a hover is replayed or taken down: what is on
/// screen for any of them is the engine's window rather than anything this app composed, so
/// what the loop asks about one it asks about the other — the same way the loader asks the
/// name gates in one order.
pub(super) fn engine_kind_of(path: &Path) -> Option<PreviewType> {
    // What the file's own bytes say comes first, as it does for the loader that draws it: a
    // picture left under a font's name is drawn here as the picture it is rather than by the
    // engine, and a font left under a document's name is the engine's. A drawing is the one
    // kind whose two halves have to be told apart by the name as well — the browser draws a
    // document, and the drawing layer replays a metafile or a PostScript program, which is
    // no engine window at all — so the name answers which half of that kind a file is.
    //
    // All of that is now one read of the router's answer rather than four questions asked of
    // the file in this order, which is what it used to be: the content, then the router, then
    // three names. A kind this app grows is an arm of `routing::drawn_by_of` rather than a fourth
    // question here (see `HoverFacts`).
    let hover = HoverFacts::read(path);

    match hover.route.content {
        // A file the bytes named as another kind is that kind, whatever the name says, and a
        // kind that is not one of the browser's three is none of this app's business: a picture
        // under a font's name is the picture it is, and nothing is handed to the engine for it.
        crate::formats::content_type::Content::Kind(_) => match hover.routed_kind() {
            Some(PreviewType::Vector) => Some(PreviewType::Vector),
            Some(PreviewType::Fonts) => Some(PreviewType::Fonts),
            // A page of HTML is a text file, so the kind it answers with is the text kind's: the
            // gate over it is the one a text preview is switched by, and the loader reaches the
            // engine through that same arm (see `load_media_of_kind`).
            Some(PreviewType::Text) if hover.html_drawn_by_the_engine() => Some(PreviewType::Text),
            _ => None,
        },

        // Where the bytes named nothing, the name's answer is what the browser draws, and it is
        // the router's — so a file whose name is a specimen's and whose bytes are nothing in
        // particular is a specimen, and a file no list claims is nothing at all.
        crate::formats::content_type::Content::Unknown => match hover.route.drawn_by {
            crate::formats::routing::DrawnBy::WebView(web) => match web {
                crate::formats::routing::WebPage::Svg => Some(PreviewType::Vector),
                crate::formats::routing::WebPage::FontSpecimen => Some(PreviewType::Fonts),
                crate::formats::routing::WebPage::Html => Some(PreviewType::Text),
            },
            _ => None,
        },

        // A format no kind of this app previews: there is no kind to draw it as, so there is
        // no engine to hand it to either.
        crate::formats::content_type::Content::Foreign => None,
    }
}

/// Whether `path` is a page of HTML the browser engine draws: the name is one of the two a
/// page goes by, and the engine is the thing that draws it (see `webview_preview::draws`,
/// which answers for a machine with no runtime by not drawing at all).
///
/// It is the one of the four questions `engine_kind_of` asks that needs nothing but the name
/// and the machine, so it stays a function of the path: the engine's availability is asked of
/// the run rather than of the file, and a caller that has a hover's answer in hand asks that
/// one instead (`HoverFacts::html_drawn_by_the_engine`).
pub(super) fn html_is_engine_drawn(path: &Path) -> bool {
    crate::formats::text_formats::is_html_extension(path) && webview_preview::draws(path)
}

/// Whether a preview of `path` may be shown as `kind`: that kind's own list claims it, and the
/// tray has that kind switched on.
///
/// It is one function for the eleven `is_<kind>_preview` predicates it replaces — one per module,
/// each of which reached for the configuration's lock for itself, which is the split-lock form the
/// deleted grep test could not see because the lock and the read were never in one place to be
/// matched. The lock is taken here, once, around two comparisons in memory: the row that names the
/// kind, and the switch that hides it. Nothing on this side of it opens a file, which is the rule
/// the whole of C4 is about (see `HoverFacts`).
///
/// The video kind is deliberately not asked of this. Its two lists share `ts` and `mts` with the
/// text lists, so answering it reads the file — and a guard held across a read is the defect this
/// form exists to be free of. A caller that wants to know about a video asks the hover's own
/// answer, which has already read what that question needs (`HoverFacts::is_video`).
pub(super) fn previewed_as(path: &Path, kind: PreviewType) -> bool {
    debug_assert_ne!(
        kind,
        PreviewType::Videos,
        "asked under a lock, and this kind reads the file"
    );

    CONFIG
        .lock()
        .map(|config| crate::formats::routing::previewed_as(path, &config, kind))
        .unwrap_or(false)
}

/// Whether `kind`'s own list claims `path`, without asking whether that kind is switched on.
///
/// It is the same question as [`previewed_as`] with the switch left off, for the two callers that
/// ask what a file *is* rather than whether a preview of it may be shown: the backdrop's own
/// question about a specimen, and the manual probes' printed table of what each list says.
pub(super) fn named_as(path: &Path, kind: PreviewType) -> bool {
    CONFIG
        .lock()
        .map(|config| crate::formats::routing::named_as(path, &config, kind))
        .unwrap_or(false)
}

pub(super) fn current_webp_playback_fps() -> u32 {
    CONFIG
        .lock()
        .map(|cfg| sanitize_webp_playback_fps(cfg.webp_playback_fps))
        .unwrap_or(DEFAULT_WEBP_PLAYBACK_FPS)
}

/// Whether this hover is owed a render: an Office document with no page in the
/// cache yet — or one whose page was exported narrower than a render asked for this
/// hover's room would be, which is a deck that was first previewed on a smaller
/// display — with the render tier switched on.
pub(super) fn office_render_is_due(path: &Path, width: u32) -> bool {
    if !previewed_as(path, PreviewType::Document) || !office_render::enabled() {
        return false;
    }

    // A file whose bytes are another kind is not a document to render, whatever it is
    // called: a picture left under a `.docx` name is drawn as the picture it is, and asking
    // Office for a page would start an engine for a file that is not its own — which is the
    // one thing the question of content exists to prevent (see `content_type`).
    if content_names_another_kind(path, PreviewType::Document) {
        return false;
    }

    // Where the render engine is the one that draws this document's page — because the tray
    // has asked it for every Office document, or because the family's application is not
    // installed — there is no engine of this tier's to ask, and the request goes there
    // instead (see `libre_render_is_due`, and `office_formats::page_engine` for the question
    // both sides ask).
    if office_formats::page_engine(path) != Some(OfficeEngine::MicrosoftOffice) {
        return false;
    }

    match office_render::held_page(path) {
        Some(page) => office_render::page_is_narrower_than(path, &page, width),
        None => true,
    }
}

/// Ask the render tier for the page a hover needs, at the moment that hover is
/// installed, and answer what is now being waited on.
///
/// Every other format has something to draw the moment a hover is up, because it
/// is read or decoded on this side. A document's page is the one thing that does
/// not exist until Office has drawn it, and none of that work can begin before
/// it is asked for — so asking late is waiting twice, once for the timer and once
/// for the render. A hover is only ever raised for the file the pointer is on, so
/// there is nothing to wait for: the page is asked for as soon as there is a
/// hover to ask for it, and the engine that request starts is kept warm for the
/// documents asked for after it.
pub(super) fn request_office_render(
    path: &Path,
    generation: u64,
    width: u32,
    height: u32,
) -> Option<(PathBuf, u64)> {
    if !office_render_is_due(path, width) {
        return None;
    }

    office_render::request(path, width, height, generation);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a page by the render engine: a document the engine draws — one
/// of its own lists, or an Office document whose own application is not installed — with an
/// engine installed to draw it and no page drawn for this version of it yet.
///
/// It is the question `office_render_is_due` asks of a document whose application is there,
/// asked of the documents whose engine is a whole application rather than an automation
/// server. Three things ask it: the layout, which measures a document like this as the wait
/// for a page; the loader, which answers with it that a hover is still waiting rather than
/// failed; and the loop, which asks the engine for the page only where there is one to ask
/// for.
///
/// Which kind the page is shown under is part of the answer rather than a second question:
/// a document of the engine's own is shown under `Libre`, and an Office document the engine
/// draws where Office cannot is shown under `Office`, at that kind's scale and over that
/// kind's backdrop — the file is what it is whichever engine drew it (see
/// `libre_formats::engine_page_kind`).
pub(super) fn libre_render_is_due(path: &Path) -> bool {
    libre_formats::engine_page_kind(path).is_some_and(PreviewType::enabled)
        && libreoffice_render::available()
        && libreoffice_render::rendered_page(path).is_none()
        && !libreoffice_render::refused(path)
}

/// Ask the engine for the page this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_office_render`: a page does not
/// exist until an engine has drawn it, and asking late is waiting twice. Nothing is waited
/// on here either — the conversion runs on the engine's own thread — so what comes back is
/// the wait, and the loop watches the folder the page lands in for it.
pub(super) fn request_libre_render(path: &Path, generation: u64) -> Option<(PathBuf, u64)> {
    if !libre_render_is_due(path) {
        return None;
    }

    libreoffice_render::request(path);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a picture by the ImageMagick engine: a file the engine develops,
/// with an engine installed to develop it and nothing developed for this version of it yet —
/// in hand for the hover that asked, or written down for the hovers after it.
///
/// It is the same question `libre_render_is_due` is, asked of an engine that is a converter
/// rather than an application — one that reads a file, writes one and exits, which is why
/// there is a process to wait for and a page rather than an instance to keep. Three things ask
/// it: the layout, which measures a file like this as the wait for a picture; the loader, which
/// answers with it that a hover is still waiting rather than failed; and the loop, which asks
/// the engine for the picture only where there is one to ask for.
pub(super) fn magick_render_is_due(path: &Path) -> bool {
    // The file's own bytes first, the name after them, exactly as the render engine's own
    // question asks it: a picture renamed to a name no list holds is still the engine's to
    // develop, and one whose bytes are another kind is not a file to start it for (see
    // `magick_formats::is_engine_picture`).
    magick_formats::is_engine_picture(path)
        && PreviewType::Magick.enabled()
        && imagemagick_render::available()
        && !imagemagick_render::refused(path)
        && !imagemagick_render::developed(path)
        && !imagemagick_render::has_page(path)
        // A raw sample dump whose own length does not settle a shape is not a file to ask about:
        // the engine would answer that it must be told a size, which is a launch spent on nothing
        // (see `raw_geometry`).
        && (!imagemagick_render::is_raw_sample(path)
            || imagemagick_render::raw_geometry(path).is_some())
}

/// Ask the engine for the picture this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_libre_render`: there is no picture
/// until the engine has developed one, and asking late is waiting twice. Nothing is waited on
/// here either — the conversion runs on the engine's own thread — so what comes back is the
/// wait, and the hover is replayed when the engine answers. The room is part of the request
/// rather than of the wait, and it is the room the display has rather than the one the wait
/// was laid out at: what the engine is told is how large a picture it may write, and a
/// picture written into too small a box is one no later layout can draw any larger (see
/// `PendingLoad::room`).
pub(super) fn request_magick_render(
    path: &Path,
    generation: u64,
    room: (u32, u32),
) -> Option<(PathBuf, u64)> {
    if !magick_render_is_due(path) {
        return None;
    }

    imagemagick_render::request(path, room, generation);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a listing by the PeaZip engine: an archive the engine lists, with an
/// engine installed to list it and nothing listed for this version of it yet.
///
/// It is the same question `magick_render_is_due` is, asked of an engine that reports rather than
/// draws — one that reads an archive, prints what is inside it and exits, which is why there is a
/// process to wait for and nothing to keep. Four things ask it: the layout, which measures a file
/// like this as the wait for a listing; the loader, which answers with it that a hover is still
/// waiting rather than failed; the loop, which asks the engine for the listing only where there is
/// one to ask for; and the layout's own placement question, which decides whether a hover is a
/// wait for something rather than a preview of it (see `page_is_on_the_way`).
pub(super) fn peazip_render_is_due(path: &Path) -> bool {
    // The file's own bytes first, the name after them, exactly as the image converter's own
    // question asks it: an archive renamed to a name no list holds is still the engine's to list,
    // and one whose bytes are another kind is not a file to start it for (see
    // `peazip_formats::is_engine_archive`).
    peazip_formats::is_engine_archive(path)
        && PreviewType::Peazip.enabled()
        && peazip_render::available_for(path)
        && !peazip_render::refused(path)
        && !peazip_render::listed(path)
}

/// Ask the engine for the listing this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_libre_render` and
/// `request_magick_render`: a listing does not exist until the engine has produced it, and asking
/// late is waiting twice. Nothing is waited on here either — the run happens on the engine's own
/// thread — so what comes back is the wait, and the hover is replayed when the engine answers.
pub(super) fn request_peazip_render(path: &Path, generation: u64) -> Option<(PathBuf, u64)> {
    if !peazip_render_is_due(path) {
        return None;
    }

    peazip_render::request(path, generation);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a page by the ebook engine: a book the engine reads, with an engine
/// installed to convert it and nothing converted for this version of it yet.
///
/// It is the same question `libre_render_is_due` is, asked of an engine that converts a whole book
/// rather than drawing a page of one — one that reads a file, writes a PDF and exits, which is why
/// there is a process to wait for and nothing to keep. Four things ask it: the layout, which
/// measures a book like this as the wait for a page; the loader, which answers with it that a hover
/// is still waiting rather than failed; the loop, which asks the engine for the page only where
/// there is one to ask for; and the layout's own placement question, which decides whether a hover
/// is a wait for something rather than a preview of it (see `page_is_on_the_way`).
pub(super) fn calibre_render_is_due(path: &Path) -> bool {
    // The file's own bytes first, the name after them, exactly as the render engine's own question
    // asks it: a book renamed to a name no list holds is still the engine's to convert, and one
    // whose bytes are another kind is not a file to start it for (see
    // `calibre_formats::is_engine_ebook`).
    calibre_formats::is_engine_ebook(path)
        && PreviewType::Calibre.enabled()
        && calibre_render::available()
        && !calibre_render::refused(path)
        && calibre_render::rendered_page(path).is_none()
}

/// Ask the engine for the page this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_libre_render`: a page does not exist until
/// the engine has converted the book, and asking late is waiting twice. Nothing is waited on here
/// either — the conversion runs on the engine's own thread — so what comes back is the wait, and
/// the loop watches the folder the page lands in for it.
pub(super) fn request_calibre_render(path: &Path, generation: u64) -> Option<(PathBuf, u64)> {
    if !calibre_render_is_due(path) {
        return None;
    }

    calibre_render::request(path);
    Some((path.to_path_buf(), generation))
}

/// Ask the first reader of this file's kind that can answer for it, and answer what is now being
/// waited on.
///
/// It is the chain a hover's page is owed by, walked once rather than spelled out: which readers
/// a kind has is `routing::chain`, which of them can answer for this file is
/// `routing::readers_for`, and this asks them in that order. Nothing is waited on in this thread
/// — every request starts work on an engine's own — so what comes back is the wait the loop
/// watches for, and the first reader that asked for one is the reader being waited on.
pub(super) fn request_engine_render(
    path: &Path,
    generation: u64,
    room: (u32, u32),
) -> Option<(PathBuf, u64)> {
    let kind = CONFIG
        .lock()
        .ok()
        .and_then(|config| crate::formats::routing::kind_of(path, &config))?;

    for reader in crate::formats::routing::readers_for(kind, path) {
        let requested = match reader {
            crate::formats::routing::Reader::Office => {
                request_office_render(path, generation, room.0, room.1)
            }
            crate::formats::routing::Reader::LibreOffice => request_libre_render(path, generation),
            crate::formats::routing::Reader::ImageMagick => {
                request_magick_render(path, generation, room)
            }
            crate::formats::routing::Reader::PeaZip => request_peazip_render(path, generation),
            crate::formats::routing::Reader::Calibre => request_calibre_render(path, generation),

            // A reader of this app's own owes the loop no wait, and neither do the two the loop
            // does not ask: a video's player is started where the video is shown, and a document
            // the browser draws is a hover handed over rather than a page waited on.
            crate::formats::routing::Reader::Native
            | crate::formats::routing::Reader::Ffmpeg
            | crate::formats::routing::Reader::WebView2 => None,
        };

        if requested.is_some() {
            return requested;
        }
    }

    None
}

/// Ask the engine that draws this file's page to be up, where the pointer has settled on it.
///
/// It is the question `request_libre_render` asks a moment later, asked a moment early: a page
/// this file is owed, by an engine that is installed, whose kind is switched on and has not
/// turned the file down. Nothing else is warmed — an engine started for a preview the user has
/// switched off is a process on the machine for nothing, and one started for a file that has its
/// page already is a launch nobody was waiting for. The ask itself is the tier's, and it is a
/// no-op for an engine that is up (see `libreoffice_render::warm`).
///
/// Office is not warmed, and that is not an omission: its worker is asked for a *document* to
/// render — the slot it takes is a `RenderRequest`, and the application is created inside the
/// render it makes — so "start the application and hold it with nothing open" is not a request
/// that tier can be handed without taking it apart, which is not what this is for.
///
/// The ebook engine is not warmed either, and there is nothing of it to warm: every book is a
/// conversion of its own and no instance is kept between them (see `calibre_render`).
pub(super) fn warm_engines_for(path: &Path) {
    if libre_render_is_due(path) {
        libreoffice_render::warm();
    }
}

/// What an engine that answers by writing a page into the app's own folder has said about this
/// file: `Some(true)` where the page has landed, `Some(false)` where the engine has answered that
/// it will not draw the file at all, and `None` where neither engine is the one being waited on or
/// where one of them is and nothing has come back yet.
///
/// Two engines answer this way — the render engine and the ebook engine — and neither sends a
/// message when it is done: what says a page is there is the page, read out of the cache it was
/// kept in (see `libre_render_is_due` and `calibre_render_is_due`). One question for both, so that
/// the wait and the replay that takes the answer up are one code path whichever engine produced it.
///
/// The question is floored rather than asked every tick. What it costs is not a comparison but a
/// content probe, a folder index and two cache keys — two of which are `fs::metadata` of the
/// document — and it was being paid sixty times a second for as long as a render took, which is a
/// couple of seconds for a LibreOffice document. There is nothing to gain by asking sooner: a
/// page that has not been written yet is not written yet however often it is looked for, and the
/// only cost of waiting longer is that the spinner turns for another tenth of a second.
pub(super) fn engine_page_answer(path: &Path) -> Option<bool> {
    // Read through one clock for both callers, so a hover waiting and a pin waiting for the same
    // document in the same tick are one read of the disk rather than two.
    static LAST_ASKED: Lazy<Mutex<Instant>> =
        Lazy::new(|| Mutex::new(Instant::now() - ENGINE_PAGE_POLL));

    let mut last = match LAST_ASKED.lock() {
        Ok(last) => last,
        Err(poisoned) => poisoned.into_inner(),
    };
    if last.elapsed() < ENGINE_PAGE_POLL {
        return None;
    }
    // The guard is given up before the question is asked, for the same reason every other
    // question on this thread gives it up: what follows reads the disk, and a lock held across
    // a read is a lock every other thread of the app waits on for as long as the disk takes.
    *last = Instant::now();
    drop(last);

    // Which engine owes the page is asked of the file, and it is the same question the request side
    // asked before there was anything to ask for: an engine that was never asked has no page for the
    // file, and waiting on it would be waiting for nothing.
    let (drawn, refused) = if calibre_formats::is_engine_ebook(path) {
        (
            calibre_render::rendered_page(path).is_some(),
            calibre_render::refused(path),
        )
    } else if libre_formats::engine_page_kind(path).is_some() {
        (
            libreoffice_render::rendered_page(path).is_some(),
            libreoffice_render::refused(path),
        )
    } else {
        return None;
    };

    match (drawn, refused) {
        (true, _) => Some(true),
        (false, true) => Some(false),
        (false, false) => None,
    }
}

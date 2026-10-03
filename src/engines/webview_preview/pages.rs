use std::path::{Path, PathBuf};
use std::sync::atomic::Ordering;

use super::api::{
    draws, drop_placement, owed, showing_path, wake_engine_thread, wanted, Area, Placement, Wanted,
    ENGINE, PLACED, PLACE_ASKED, WANTED, WANTED_GENERATION,
};
use super::engine::{trace, Command, Engine};
use super::environment::{file_url, user_data_folder};

use crate::app::engine_processes;
use crate::config::config::TransparentBackground;
use crate::formats::text_formats;
use crate::readers::font_preview;

/// The page the engine is pointed at, written beside this run's state.
///
/// The document is not the page. A browser draws a standalone SVG at the size it asks
/// for — it does not stretch one to the window it was given — so a document given to the
/// engine as the page is drawn small in a large window, and the window this app lays out
/// is the share of the display `vector_scale` asked for: a document asking for 120 pixels is
/// a 120-pixel picture in a window half a display wide, or in a display-sized one at
/// fit-to-screen. What is given to the engine instead is this page: an image of the
/// document, in a box that is the whole page. An image *is* scaled to the box it is
/// given, whatever its own size is, which is the one thing that makes the window and the
/// document the same size at every scale.
///
/// Nothing is given up for that. An SVG drawn as an image is animated and not scripted,
/// and here nothing runs either — an image is a picture a browser shows, and a picture
/// cannot run code, take the pointer, or reach anything outside itself — not a file beside
/// it, not a URL — so a document that links to the world is drawn without it. A page of
/// HTML is the one document this engine runs, and it is not this page: that one is a frame
/// around the file (see `html_page`), with its own rule, which is the same reach and one
/// more thing given. Chromium parses a document as XML, in the mode built for animated
/// images, so the document is the document rather than a copy of it in a page of our own.
///
/// The version is the document's own modification time, which is what keeps an edited
/// file from being answered out of the browser's image cache: the URL changes when the
/// file does, and the same file at the same version is drawn again from memory.
///
/// The backdrop is the page's business for one of the four kinds: a checkerboard is
/// drawn by whatever composites the frame, and this engine composites its own — it can be
/// given a colour and nothing else — so the page paints the same squares this app's own
/// compositing draws, and the controller's colour stands behind them for the moment
/// before the page is up.
pub(super) fn frame_page(
    path: &Path,
    version: u64,
    background: TransparentBackground,
) -> Option<(PathBuf, String)> {
    let document = file_url(path)?;
    // The backdrop is in the page's name as well as in its content, of the four kinds the
    // one that is a page's own — a checkerboard — is painted here rather than by the
    // controller: a change of backdrop is then a page the browser has not seen, rather than
    // the same URL answered out of its cache with the squares of the last one.
    let page = user_data_folder().join(format!("frame-{}.html", background.as_str()));

    let html = frame_html(&document, version, background);

    write_page(&page, &html, version)
}

/// What every page of this app's opens with: the doctype, the encoding, and a page that fills
/// the window it is given rather than scrolling inside it — the arrangement `html_page` exists
/// for, and the one a document drawn as an image needs as well.
///
/// The style is left open: each of the three kinds of page carries its own rules after it and
/// closes the same element.
const PAGE_HEAD: &str = "<!doctype html><meta charset=\"utf-8\"><title>preview</title>\
     <style>html,body{margin:0;padding:0;height:100%;overflow:hidden}";

/// The markup a document is drawn in, as a function of its URL rather than of its path, so
/// that what the refusal on it is can be asked of the markup without a page being written for
/// the question — one page per backdrop is one file, so two hovers of the same kind answer
/// from the same name by design, and two tests asking the same question of it at once would
/// be reading each other's.
///
/// The version is in the image's URL and in the page's, so neither is answered out of the
/// browser's cache with a document that has been written since it was read.
pub(super) fn frame_html(
    document: &str,
    version: u64,
    background: TransparentBackground,
) -> String {
    format!(
        "{PAGE_HEAD}\
         img{{display:block;width:100%;height:100%;object-fit:contain}}</style>\
         <style>{no_interaction}</style>\
         {checkerboard}\
         <img src=\"{}?v={version}\" draggable=\"false\" alt=\"\">",
        escape_attribute(document),
        checkerboard = checkerboard_style(background),
        no_interaction = NO_INTERACTION_STYLE
    )
}

/// The page a page of HTML is drawn in, and the one it runs in: the file itself, whole, in a
/// frame that fills the window it is given — the arrangement `frame_page` reaches the same
/// end by with an image, for a document that has a size of its own.
///
/// The frame is given two allowances and nothing else. It is given its own origin, without
/// which the page itself would not arrive: a page's stylesheets and its pictures are relative
/// to it, so a frame loaded as a document of an opaque origin of its own comes up unstyled.
/// And it is given script, which is the one thing a page of HTML is handed the browser's own
/// engine for: a page that draws itself with WebGL, or lays itself out from a script, is a
/// page nothing but a run can show, and withheld it comes up a blank rectangle.
///
/// What is still withheld is everything that is a way *out* of the frame — popups, forms, and
/// any navigation of the top frame — which is the same reach a document drawn as an image has,
/// since a picture cannot pop up, submit, or navigate either, and which `BROWSER_ARGUMENTS`
/// reaches from the outside for the links a page keeps. Nor is the frame given the three
/// things a run would otherwise bring with it: a browser that plays sound without a gesture, a
/// page that reads the files beside it, and a page that takes the whole screen — no autoplay
/// policy is passed, no `--allow-file-access-from-files`, and no `allowfullscreen` for a page
/// asking to be shown full screen to be refused.
///
/// The two allowances together are not the loosening they would be for same-origin content,
/// where an allowance to keep one's own origin alongside one to run would let a framed document
/// reach out of itself and take the frame with it. They are not same-origin here: the wrapper
/// is a file this app wrote into this run's own profile folder and the document is another
/// file altogether, so the two are separate origins whatever the sandbox is told, and a
/// document that runs reaches its own file and no further out of it.
pub(super) fn html_page(
    path: &Path,
    version: u64,
    background: TransparentBackground,
) -> Option<(PathBuf, String)> {
    let page_url = file_url(path)?;
    let page = user_data_folder().join(format!(
        "html-{}-{}.html",
        background.as_str(),
        path_identity(path)
    ));

    let html = format!(
        "{PAGE_HEAD}\
         iframe{{display:block;width:100%;height:100%;border:0}}</style>\
         {checkerboard}\
         <iframe src=\"{}?v={version}\" sandbox=\"allow-same-origin allow-scripts\" title=\"\"></iframe>",
        escape_attribute(&page_url),
        checkerboard = checkerboard_style(background)
    );

    write_page(&page, &html, version)
}

/// What a page of HTML is called in the page that draws it, from its own path.
///
/// A wrapper is one page however many targets it has been written for, so the target is
/// hashed into the name the browser caches by: a hover on one page and then on another is two
/// pages rather than one page rewritten under the browser that has already seen the first.
/// The version is what distinguishes one file from itself, not two files from each other.
fn path_identity(path: &Path) -> u64 {
    use std::collections::hash_map::DefaultHasher;
    use std::hash::{Hash, Hasher};

    let mut hasher = DefaultHasher::new();
    path.hash(&mut hasher);
    hasher.finish()
}

/// The page a font specimen is drawn in: the font itself, in the page through
/// `@font-face`, with the lines the file's own character map covers under a heading the
/// font's `name` table supplies.
///
/// One page per font, pointed at the file the engine can actually read — the font itself, or
/// the face `font_preview::browser_source` wrote out for a collection, which is the face the
/// specimen was read at — and every size in it is a viewport unit, so the window the layout
/// planned is the size the type is drawn at: the share of the display `font_scale` names is a
/// share of the specimen's size, the same relationship `object-fit: contain` gives a document.
pub(super) fn font_page(
    path: &Path,
    version: u64,
    background: TransparentBackground,
    face: usize,
) -> Option<(PathBuf, String)> {
    let specimen = font_preview::probe_face(path, face)?;
    let source = font_preview::browser_source(path, &specimen, &user_data_folder())?;
    let font = file_url(&source)?;
    // The face is in the page's name as well as in its content, for the reason the backdrop
    // is: a specimen read at another face is a page the browser has not seen, rather than the
    // same URL answered out of its cache with the face before it.
    let page = user_data_folder().join(format!(
        "font-{}-{}.html",
        background.as_str(),
        specimen.face
    ));

    // The first line is the specimen's headline — the pangram wherever the font has Latin —
    // and the lines under it are the scripts the font also holds. Each line carries the
    // direction its own script is written in, which is what a browser settles it from anyway:
    // an Arabic or Hebrew line is then laid out from the right, its full stop ending it where
    // the script ends it rather than where a left-to-right page would, and a line of any
    // other script is drawn as it was.
    let html = specimen_html(
        &font,
        version,
        background,
        &specimen.title,
        &specimen.samples,
    );

    write_page(&page, &html, version)
}

/// The markup a specimen is drawn in, apart from what it is read from: the font's own URL,
/// the name its `name` table gave and the lines its character map covers. Split out for the
/// reason `frame_html` is — one page per backdrop and face is one file, so a question asked of
/// the markup should not be asked of a file two hovers are also writing.
pub(super) fn specimen_html(
    font: &str,
    version: u64,
    background: TransparentBackground,
    title: &str,
    samples: &[String],
) -> String {
    let lines: String = samples
        .iter()
        .enumerate()
        .map(|(index, line)| {
            let class = if index == 0 { "pangram" } else { "script" };
            format!(
                "<p class=\"line {class}\" dir=\"auto\">{}</p>",
                escape_text(line)
            )
        })
        .collect();

    // A font that covers more of the sample lines than the box was shaped for — a pan-script
    // one, which is rare — is drawn smaller rather than past the bottom of it: the sizes above
    // are what a pangram and the handful of script lines a font usually holds want.
    let (pangram_size, script_size, script_margin) = if samples.len() > SPECIMEN_FULL_LINES {
        ("6vh", "3.6vh", "1vh")
    } else {
        ("8vh", "5vh", "1.6vh")
    };

    let (ink, shadow) = specimen_ink(background);
    format!(
        "{PAGE_HEAD}\
         body{{display:flex;flex-direction:column;justify-content:center;\
         padding:6vh 6vw;box-sizing:border-box;color:{ink};{shadow}}}\
         .title{{font-family:\"Segoe UI\",system-ui,sans-serif;font-size:2.4vh;\
         font-weight:600;opacity:.6;margin:0 0 2.4vh;white-space:nowrap;\
         overflow:hidden;text-overflow:ellipsis}}\
         .line{{font-family:\"RHPPreviewFont\",\"Segoe UI\",sans-serif;margin:0;\
         line-height:1.15}}\
         .pangram{{font-size:{pangram_size}}}\
         .script{{font-size:{script_size};margin-top:{script_margin}}}\
         </style>\
         <style>{no_interaction}</style>\
         <style>@font-face{{font-family:\"RHPPreviewFont\";\
         src:url(\"{font}?v={version}\")}}</style>\
         {checkerboard}\
         <div class=\"title\">{title}</div>{lines}",
        font = escape_attribute(font),
        title = escape_text(title),
        checkerboard = checkerboard_style(background),
        no_interaction = NO_INTERACTION_STYLE
    )
}

/// How many lines a specimen is drawn at the sizes above: the pangram and six lines under it
/// are what the specimen's box holds, and a font that covers more of the lines this app knows
/// than that — a pan-script font, which covers most of them — is drawn smaller instead.
const SPECIMEN_FULL_LINES: usize = 7;

/// Put a page where the engine will find it, and answer with the URL it is navigated to.
///
/// The version — the file's own modification time — goes into that URL, which is what keeps
/// an edited file from being answered out of the browser's cache: the URL changes when the
/// file does, and the same file at the same version is drawn again from memory.
fn write_page(page: &Path, html: &str, version: u64) -> Option<(PathBuf, String)> {
    std::fs::create_dir_all(user_data_folder()).ok()?;
    std::fs::write(page, html).ok()?;

    let url = format!("{}?v={version}", file_url(page)?);

    Some((page.to_path_buf(), url))
}

/// The squares a checkerboard backdrop is drawn as, for the one backdrop the engine cannot be
/// given: it takes a colour and nothing else, so the squares are painted by the page — over
/// the mid grey the controller is given, which is what the two square colours average to (see
/// `background_color`).
fn checkerboard_style(background: TransparentBackground) -> &'static str {
    const SQUARES: &str = "<style>html{background:#e0e0e0;background-image:\
         conic-gradient(#909090 25%,transparent 0 50%,#909090 0 75%,transparent 0);\
         background-size:32px 32px}</style>";

    match background {
        TransparentBackground::Checkerboard => SQUARES,
        _ => "",
    }
}

/// The refusal a looked-at document is drawn under: the page takes no pointer at all.
///
/// This is the second of the two refusals a picture in a browser needs, and the first is
/// `draggable="false"` on the drawing itself (`frame_page`). Neither is a setting of the
/// browser's — WebView2 has none for it — and both are the page's own, because the page is
/// the only part of this a hand can be kept away from.
///
/// What is refused is a *drag*, and the drag is what hangs a pinned preview. A browser drag
/// of an image is not a message this app is given: Chromium begins a drag session of its own
/// on the engine's thread, with the drawing's own silhouette following the pointer, and that
/// session is a modal loop holding the pointer. The pin takes the same pointer for its own
/// drag of the window at almost the same moment, and one pointer with two owners is a press
/// whose release reaches neither — the pin left holding a capture nothing will release, and
/// a browser's drag that never ends. So the page is drawn so that there is no press to
/// begin one with.
///
/// A specimen is refused the same way for the same reason, and is the one place the selection
/// matters: text a pointer can sweep across is text a pointer can begin a drag out of. It is
/// put here rather than written into either page because the two of them are the two kinds of
/// document that are *looked at*, and a page of HTML — the third kind, and the only one this
/// engine runs — is exactly the one that keeps its pointer (`html_page`).
const NO_INTERACTION_STYLE: &str = "*{pointer-events:none;user-select:none;\
     -webkit-user-select:none;-webkit-user-drag:none}";

/// The colour a specimen's text is drawn in over each backdrop, and the shadow that goes with
/// it.
///
/// Three of the four are a colour: light text on black, dark on white and on the checkerboard's
/// light squares. The one that is not is transparency, which has no colour to be read
/// against — so the glyphs are drawn light with a soft dark shadow behind them, and what a
/// specimen looks like over whatever the desktop happens to be is still readable.
fn specimen_ink(background: TransparentBackground) -> (&'static str, &'static str) {
    match background {
        TransparentBackground::Black => ("#f2f2f2", ""),
        TransparentBackground::White | TransparentBackground::Checkerboard => ("#1a1a1a", ""),
        TransparentBackground::Transparent => ("#f2f2f2", "text-shadow:0 0 .45vh rgba(0,0,0,.6);"),
    }
}

/// A URL as an attribute value: the two characters that would end it early.
pub(super) fn escape_attribute(url: &str) -> String {
    url.replace('&', "&amp;").replace('"', "&quot;")
}

/// A string as the page's text: the characters that would end it, or open a tag in it. A
/// specimen's title is the font's own business rather than this app's — it is a name out of a
/// file — so it is escaped rather than trusted.
fn escape_text(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Ask the engine to draw `path` in a window at `area`.
///
/// The answer is immediate and says nothing about whether the document arrived: the
/// engine works on its own thread, and what it does with this is navigates, waits for
/// the document, and puts its window up — `showing_path` is what says which document got
/// there. What a caller puts on screen while it waits is the waiting spinner and nothing
/// else: this app draws no document, so there is no still frame to hold the place.
///
/// A file that is not a document at all is not the engine's to draw, and everything else
/// is the want in `WANTED`: the newest one wins, and what was asked for before it is not
/// drawn at all. The *same* document asked for a second time — which is what a wait that
/// follows the pointer does, sixty times a second — is the same want with the box it has
/// moved to, and it keeps the generation it was made under, so a pointer that keeps moving
/// while the document is on its way does not call off the navigation it is waiting for.
pub(crate) fn show(path: &Path, area: Area, background: TransparentBackground) {
    if !draws(path) {
        trace(&format!(
            "show({}): not a document the engine draws",
            path.display()
        ));
        return;
    }

    let generation = {
        let Ok(mut wanted) = WANTED.lock() else {
            trace("show: the wanted cell is poisoned");
            return;
        };

        let same_document = wanted
            .as_ref()
            .is_some_and(|wanted| wanted.path == path && wanted.background == background);

        let generation = if same_document {
            wanted.as_ref().map(|wanted| wanted.generation).unwrap_or(0)
        } else {
            WANTED_GENERATION.fetch_add(1, Ordering::AcqRel) + 1
        };

        *wanted = Some(Wanted {
            generation,
            path: path.to_path_buf(),
            background,
            area,
        });

        generation
    };

    let Ok(mut engine) = ENGINE.lock() else {
        trace("show: the engine's lock is poisoned");
        return;
    };

    let sender = engine.get_or_insert_with(Engine::start).sender.clone();

    if let Err(error) = sender.send(Command::Show { generation }) {
        trace(&format!(
            "show({}): the engine's thread is gone: {error}",
            path.display()
        ));
    }

    // A navigation the engine is in the middle of is waiting for a page that is no longer
    // wanted: the thread is woken rather than left to finish it (see `wake_engine_thread`).
    wake_engine_thread();
}

/// Take the engine's window down. The engine itself is kept warm: what it costs to
/// begin is a browser start, and what it costs to point at another document is a few
/// milliseconds, so a hover that follows another one pays almost nothing.
///
/// Kept warm is not the same as left working, and this is the other half of it: the
/// browser is told to stop while its window is off screen, so a document that runs costs
/// nothing until the next one is asked for (see `Host::hide`, `suspend`).
///
/// Nothing is wanted once this returns, and the generation goes with it: a navigation the
/// engine is in the middle of is one whose file the pointer has left, so it is dropped
/// rather than put up, and the window comes down without waiting for it.
pub(crate) fn hide() {
    if let Ok(mut wanted) = WANTED.lock() {
        *wanted = None;
    }
    WANTED_GENERATION.fetch_add(1, Ordering::AcqRel);

    let Ok(engine) = ENGINE.lock() else {
        return;
    };

    if let Some(engine) = engine.as_ref() {
        let _ = engine.sender.send(Command::Hide);
    }

    // Nothing is wanted any more, so a navigation in the middle of arriving is one nobody
    // is waiting for: the thread is woken rather than left to finish it.
    wake_engine_thread();
}

/// Take the engine's window down when the file it is drawing is a page of HTML: the switch
/// that asks for the page has just been turned off, and a page left standing would be a
/// window nothing takes down again until the pointer leaves (see `renders_html`).
///
/// The want is read as well as what is up: a switch turned while the engine is still
/// navigating to a page has no window yet, and the page that lands afterwards would stay.
pub(crate) fn hide_html_preview() {
    let watching_html = wanted().is_some_and(|want| text_formats::is_html_extension(&want.path));
    let showing_html = showing_path().is_some_and(|path| text_formats::is_html_extension(&path));

    if watching_html || showing_html {
        hide();
    }
}

/// Move a want that is already in hand to the box the wait has ended up in: what a pointer
/// that kept moving while the document was on its way asks for.
///
/// Nothing is sent to the engine — a document that is being navigated to is already drawn
/// in the box the want carries when it lands — and nothing is asked where the file is not
/// the one that is wanted: a box belongs to the hover that is waiting on it, and a hover
/// for another file is a `show`, which is the ask that takes the place of this want
/// altogether.
pub(crate) fn wanted_here(path: &Path, area: Area) {
    if let Ok(mut wanted) = WANTED.lock() {
        if let Some(wanted) = wanted.as_mut().filter(|wanted| wanted.path == path) {
            wanted.area = area;
        }
    }
}

/// Move the engine's window to another box, for a preview whose *own* window has changed
/// rather than for a pointer that moved: a pinned document is dragged, resized, maximized,
/// restored or carried to another display, and what stands in its media band has to travel
/// with the box.
///
/// It is both of the asks above in one, and either half applies (`box_change`). A document
/// still on its way has only a want to move — moving one sends no command, exactly as
/// `wanted_here` does — while one the engine is already holding has the window put in the new
/// box by a placement rather than by a `show`, which would navigate again for a document the
/// engine already has. The two are told apart by the window rather than by the want: a
/// document that has landed keeps its want, so a box that changed under one would be read as
/// a box still on its way and moved nowhere (see `PLACED`, `Host::place`).
///
/// Nothing is asked for another file: a box belongs to the preview it was measured for, and a
/// preview for another file is a `show`, which is the ask that takes this want's place
/// altogether. A box that is already the one asked for is nothing to do at all — a drag is
/// many of these.
pub(crate) fn place(path: &Path, area: Area, background: TransparentBackground) {
    // What is done with the box is read from the two cells the engine keeps and nothing else:
    // whether the document is the one *wanted*, whether it is the one *held*, and whether the
    // box is already the one on record for it. A drag is many of these, so the last of the three
    // is what keeps a box that has not moved from asking anything at all.
    let owed = owed(path);
    let holds = showing_path().is_some_and(|shown| shown == path);
    let same_box = WANTED
        .lock()
        .ok()
        .and_then(|wanted| {
            wanted
                .as_ref()
                .map(|wanted| wanted.path == path && wanted.area == area)
        })
        .unwrap_or(false);

    match box_change(owed, holds, same_box) {
        // A document still on its way has only a want to move, and moving one sends nothing: what
        // is drawn is drawn in the box the newest want asks for when it lands, so a box that
        // changed while a browser was coming up is a document that arrives in the right place and
        // an engine that is not asked again (see `wanted_here`).
        BoxChange::Want => wanted_here(path, area),
        // A document the engine is holding has a window that has to move with the box, and that
        // is a *move* and not a document being put on screen: the want is moved too, so a
        // navigation that comes later lands in the box the window is in, and the window itself is
        // asked for by a placement rather than by a `show` (see `ask_place`, `Host::place`).
        BoxChange::Window => {
            wanted_here(path, area);
            ask_place(path, area, background);
        }
        // A box that did not move, or a file the engine neither holds nor is owed: nothing to
        // move, and nothing asked (see `box_change`).
        BoxChange::Nothing => {}
    }
}

/// Publish a box for the engine's thread to move a held document's window into, and ask for it
/// once however many boxes have been published.
///
/// The cell is what makes a drag of any length cost one move: each of these replaces what was
/// there, so the placement the engine eventually takes up is the box the hand ended in. The
/// flag is what makes it cost that one move rather than one per pointer move — a drag can
/// publish boxes far faster than a single-threaded engine takes commands, and a document that
/// animates is a browser already busy enough, so an ask that finds one already outstanding
/// leaves the newer box in the cell and is answered by the ask already in the channel.
///
/// A send that finds no engine, or an engine whose thread has gone, is a placement nothing will
/// ever take up, so the flag is given back rather than left owed: a flag left owed is a window
/// that stops following the hand for the rest of the run (see `PLACE_ASKED`).
pub(super) fn ask_place(path: &Path, area: Area, background: TransparentBackground) {
    if let Ok(mut placed) = PLACED.lock() {
        *placed = Some(Placement {
            path: path.to_path_buf(),
            background,
            area,
        });
    } else {
        return;
    }

    // One outstanding placement at a time, and the newest box travels in the cell above rather
    // than in the command, so this is the same one command however long the drag is.
    if PLACE_ASKED.swap(true, Ordering::AcqRel) {
        return;
    }

    let Ok(engine) = ENGINE.lock() else {
        PLACE_ASKED.store(false, Ordering::Release);
        return;
    };

    match engine.as_ref() {
        Some(engine) if engine.sender.send(Command::Place).is_ok() => {}
        _ => PLACE_ASKED.store(false, Ordering::Release),
    }
}

/// Which half of a `place` a box that changed under a document belongs to.
///
/// "Owed" cannot answer it on its own: a want is what the engine is owed, a document that
/// lands keeps it until the next file takes its place, and both a document still on its way and
/// the one being drawn are owed — so a box that changed under the second was answered as if it
/// were the first, and a pinned window left its document behind the moment it was dragged. What
/// separates them is the window, which is what `holds` is.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum BoxChange {
    /// Nothing to do: the box is the one on record, or the file is not the engine's.
    Nothing,
    /// The want is moved, and the engine is not asked: the document has not landed.
    Want,
    /// The engine's window is put in the new box, and the document is not navigated again.
    Window,
}

pub(super) fn box_change(owed: bool, holds: bool, same_box: bool) -> BoxChange {
    if same_box {
        return BoxChange::Nothing;
    }

    if owed {
        // Both halves of what a window is: the document is on screen, and it is the one this
        // preview is for — a held document nobody wants is a want that has already moved on.
        return if holds {
            BoxChange::Window
        } else {
            BoxChange::Want
        };
    }

    BoxChange::Nothing
}

/// Publish a want for a document and take it back again, without asking for a browser.
///
/// What `owed` and `is_behind` answer is a question about the *want*, and a want is otherwise
/// only ever made by asking the engine for one — which in a test means a browser, and there is no
/// browser in a test. This is the want those questions are about and nothing else: what is owed,
/// and no engine behind it yet.
#[cfg(test)]
pub(crate) fn publish_want_for_test(path: &Path) {
    let generation = WANTED_GENERATION.fetch_add(1, Ordering::AcqRel) + 1;

    if let Ok(mut wanted) = WANTED.lock() {
        *wanted = Some(Wanted {
            generation,
            path: path.to_path_buf(),
            background: TransparentBackground::Black,
            area: Area {
                x: 0,
                y: 0,
                width: 8,
                height: 8,
            },
        });
    }
}

/// And the same want taken back, for the same reason (see `publish_want_for_test`).
#[cfg(test)]
pub(crate) fn clear_want_for_test() {
    if let Ok(mut wanted) = WANTED.lock() {
        *wanted = None;
    }

    WANTED_GENERATION.fetch_add(1, Ordering::AcqRel);
}

/// Let the engine go, window, browser process and thread together. Called when the app
/// ends.
pub(crate) fn shutdown() {
    let engine = ENGINE.lock().ok().and_then(|mut engine| engine.take());

    // A placement published for a window this call is about to destroy belongs to nothing, and
    // the flag has to go back with it — the thread that would have taken it up is joined below
    // and the next engine begins with no placement owed (see `drop_placement`).
    drop_placement();

    if let Some(engine) = engine {
        let _ = engine.sender.send(Command::Shutdown);
        let _ = engine.thread.join();
    }

    // The engine thread ends its browser as it goes. What this is for is the browser
    // that is still there anyway — the runtime's process is not one this app gets to
    // assume about — and the one this run started that could not be told apart from
    // an earlier engine's, and so was never recorded. A browser started by anything
    // else is not a child of this process, which is what keeps this from reaching
    // past this app's own.
    engine_processes::end_our_browsers();
}

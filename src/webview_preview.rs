//! Drawing a document — or a font specimen — in the browser engine that is already on the
//! machine.
//!
//! Every document this app previews is drawn here — still, gzipped or animated: this app
//! rasterizes none of them, so what a document needs is the WebView2 runtime Windows 11
//! ships with, which is Chromium, and which is the only complete implementation of SVG
//! that is on the machine without installing anything. A machine without the runtime has
//! no SVG preview at all, which is the answer a file that will not decode gets.
//!
//! A font is drawn by the same engine for the same reason, and it is one of the two kinds
//! this app draws rather than rasterizes: the five formats a font goes by are one container
//! or another around the same outlines, and the browser reads all of them. What the page
//! carries for one is a specimen — the font's own name and the lines its character map
//! covers, worked out on this side by `font_preview` — and a collection is answered with the
//! face the setting names written out as a font of its own, because no page can name a face
//! inside one.
//!
//! What lives here is the engine and the window it draws in, not the preview loop: the
//! loop measures the document for the layout, hands it over with the box the layout came
//! out with, and this answers with a window of its own. That window is its own because
//! the preview window is a layered one, and a layered window has no window tree to put a
//! child in — the same shape as the video path, where the player's own window is the
//! preview.
//!
//! One engine is kept warm between documents and let go after `webview_idle`, ten
//! minutes by default: beginning one costs a browser start, and pointing a warm one at
//! another file costs a few milliseconds, so what a hover pays for a second document is
//! nothing worth measuring — and what a hover pays for the *first* one is a browser
//! start, which is what the waiting spinner is shown for. What is let go of is the
//! engine and not the thread that holds it, which stays parked on its channel for the
//! run and begins a new engine for the next document: a thread that ended with its
//! browser would leave every document after the first idle timeout with nobody to draw
//! it, which is no preview for the rest of the run. An app left alone has no browser
//! process and one thread asleep — and the settings it is given are the app's own rules
//! rather than a browser's: a document is drawn and not run, and nothing about it is a
//! way out of the preview.

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{self, Receiver, RecvTimeoutError, Sender};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use webview2_com::Microsoft::Web::WebView2::Win32::{
    CreateCoreWebView2EnvironmentWithOptions, GetAvailableCoreWebView2BrowserVersionString,
    ICoreWebView2, ICoreWebView2Controller, ICoreWebView2Environment,
    ICoreWebView2EnvironmentOptions, COREWEBVIEW2_COLOR,
};
use webview2_com::{
    CoreWebView2EnvironmentOptions, CreateCoreWebView2ControllerCompletedHandler,
    CreateCoreWebView2EnvironmentCompletedHandler, NavigationCompletedEventHandler,
};
use windows::core::{w, Interface, PCWSTR, PWSTR};
use windows::Win32::Foundation::{E_POINTER, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::System::Com::{CoInitializeEx, CoTaskMemFree, COINIT_APARTMENTTHREADED};
use windows::Win32::System::WinRT::EventRegistrationToken;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, PeekMessageW, RegisterClassW,
    SetWindowPos, ShowWindow, TranslateMessage, HWND_TOPMOST, MSG, PM_REMOVE, SWP_NOACTIVATE,
    SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE, WM_MOUSEACTIVATE, WNDCLASSW, WS_EX_NOACTIVATE,
    WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

use crate::config::{
    EngineIdle, PreviewType, TransparentBackground, DEFAULT_WEBVIEW_IDLE_SECS,
};
use crate::engine_processes;
use crate::font_formats;
use crate::font_preview;
use crate::{svg_preview, CONFIG};

/// The window class the engine's window is made from. It exists to refuse activation: a
/// preview never takes the keyboard away from what the pointer is over, and a browser
/// hosted in one is no different.
const WEBVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWebView");

/// The browser arguments this app passes. The first answers for no host at all, so
/// nothing a document links to is fetched from anywhere (see `create_environment`); the
/// second keeps a scrollbar out of a preview if a document is ever a pixel larger than
/// the window it was given. Both are quoted because the browser's command line is a
/// command line: unquoted, an argument with a space in it arrives as several.
const BROWSER_ARGUMENTS: &str = "--host-resolver-rules=\"MAP * ~NOTFOUND\" --hide-scrollbars";

/// What one engine cost to begin and to point at a document, in milliseconds. It is
/// kept for the same reason the other probes exist: "the preview is slow" is answered
/// by a number, and this is the module the number belongs to.
#[derive(Clone, Copy, Debug, Default)]
pub struct Timings {
    pub environment_ms: u64,
    pub controller_ms: u64,
    pub navigate_ms: u64,
}

static LAST_TIMINGS: Lazy<Mutex<Timings>> = Lazy::new(|| Mutex::new(Timings::default()));

/// Whether the runtime is on this machine, answered once: the check reads the version
/// of the installed runtime, and that does not change while the app runs.
static RUNTIME: Lazy<Option<String>> = Lazy::new(runtime_version);

/// Whether the engine's window is on screen. The preview loop reads this to know when
/// to take its own window down, so it is an atomic rather than a message.
static SHOWING: AtomicBool = AtomicBool::new(false);

/// Where the engine's window is while it is on screen, in screen coordinates: the box
/// the document was last handed over in. Kept as a rectangle rather than as the window
/// handle, because a rectangle is what the preview loop asks about a preview.
static SHOWING_RECT: Lazy<Mutex<Option<ScreenRect>>> = Lazy::new(|| Mutex::new(None));

/// A rectangle in screen coordinates, as the preview loop asks about one.
type ScreenRect = (i32, i32, i32, i32);

/// The engine's thread, once one has been started.
static ENGINE: Lazy<Mutex<Option<Engine>>> = Lazy::new(|| Mutex::new(None));

/// Where the engine is asked to put its window, in screen coordinates.
#[derive(Clone, Copy)]
pub struct Area {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

/// The version of the WebView2 runtime on this machine, or nothing when it is not
/// installed — which is what decides whether an SVG document can be previewed at all.
pub fn runtime_version() -> Option<String> {
    unsafe {
        let mut version = PWSTR::null();
        GetAvailableCoreWebView2BrowserVersionString(PCWSTR::null(), &mut version).ok()?;

        let text = pwstr_to_string(version);
        CoTaskMemFree(Some(version.0 as *const _));

        text
    }
}

/// Whether a document can be handed to the engine at all.
pub fn is_available() -> bool {
    RUNTIME.is_some()
}

/// Whether the engine has a window on screen.
pub fn is_showing() -> bool {
    SHOWING.load(Ordering::Acquire)
}

/// Where the engine's window is, when it has one on screen: the box the document was
/// last handed over in, in screen coordinates.
///
/// The preview loop reads it for the reason it reads its own window's rectangle: a
/// preview that is under the parked pointer is one a keyboard hover placed there, and a
/// document the engine draws is a preview of this app's even though the window is not.
pub fn screen_rect() -> Option<ScreenRect> {
    SHOWING_RECT.lock().ok().and_then(|rect| *rect)
}

/// Say where the engine's window is — or that it is nowhere — for `screen_rect` to
/// answer with.
fn publish_rect(area: Option<Area>) {
    if let Ok(mut rect) = SHOWING_RECT.lock() {
        *rect = area.map(|area| {
            (
                area.x,
                area.y,
                area.x + area.width,
                area.y + area.height,
            )
        });
    }
}

/// What the engine cost last time it was asked for a document. Read by the probe: it
/// is the number that says whether the engine is worth keeping warm at all.
#[cfg(test)]
pub fn last_timings() -> Timings {
    LAST_TIMINGS
        .lock()
        .map(|timings| *timings)
        .unwrap_or_default()
}

/// Whether a file is one the engine draws: the runtime is on the machine, the file is a
/// document or a font, and the engine is not in one of its own bad spells.
pub fn draws(path: &Path) -> bool {
    can_draw() && (svg_preview::is_svg_file(path) || font_formats::is_font_file(path))
}

/// Whether the engine can draw anything at all.
///
/// The runtime is what draws a document or a specimen, so a machine without it — or one the
/// engine has stood down on — has neither: this is the question the layout asks before it
/// measures one, and answering no is what keeps a hover from opening a box that nothing
/// would be drawn into.
pub fn can_draw() -> bool {
    is_available() && !is_failing()
}

/// How long the engine is left alone after it has failed to come up.
///
/// The reason it can fail is a folder, not the document: one user data folder is one
/// browser at a time, and a browser left behind by an earlier run — one whose app was
/// ended before it could take its browser with it — holds it until it goes. What that
/// must not cost is more than it has to: an engine that cannot be had draws no document
/// at all, so for this long afterwards an SVG hover opens nothing, and the engine is
/// asked again once the window has passed.
const ENGINE_RETRY_AFTER: Duration = Duration::from_secs(60);

/// When the engine last failed to come up.
static ENGINE_FAILED_AT: Lazy<Mutex<Option<Instant>>> = Lazy::new(|| Mutex::new(None));

/// Whether the preview loop has not yet been told that the engine has something to
/// answer for.
static FAILURE_NOTICE: AtomicBool = AtomicBool::new(false);

/// Whether the engine has failed since this was last asked.
///
/// One signal for the two ways it can: an engine that could not be started at all, which
/// stands it down for a while, and one that could not put a document up, which costs
/// that hover and nothing else. What the preview loop does with either is the same,
/// because this app draws no document itself: the wait a hover was given goes, rather
/// than staying a spinner over nothing.
pub fn take_failure_notice() -> bool {
    FAILURE_NOTICE.swap(false, Ordering::AcqRel)
}

fn is_failing() -> bool {
    ENGINE_FAILED_AT
        .lock()
        .ok()
        .and_then(|failed| *failed)
        .is_some_and(|failed| failed.elapsed() < ENGINE_RETRY_AFTER)
}

/// The engine could not be started at all: the runtime is on the machine, but no browser
/// could be had for it, which is the profile folder being held by one that is not this
/// app's.
///
/// It stands the engine down for `ENGINE_RETRY_AFTER` — no document is handed over in
/// that window, because every one of them would fail the same way — and tells the
/// preview loop, whose hover has nothing left to wait for.
fn note_failure() {
    if let Ok(mut failed) = ENGINE_FAILED_AT.lock() {
        *failed = Some(Instant::now());
    }

    FAILURE_NOTICE.store(true, Ordering::Release);
}

/// One document could not be put up: its page could not be written, or the engine did
/// not arrive at it.
///
/// The engine itself is fine, so this is not a failure to stand down for — the next
/// document is drawn as this one was meant to be. What it costs is the hover that was
/// waiting on it, and the notice is what takes that wait down.
fn note_document_failed() {
    FAILURE_NOTICE.store(true, Ordering::Release);
}

fn note_engine_up() {
    if let Ok(mut failed) = ENGINE_FAILED_AT.lock() {
        *failed = None;
    }
}

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
/// which is the rule this engine runs under anyway: in an image a document cannot run
/// code, cannot take the pointer, and cannot reach anything outside itself — not a file
/// beside it, not a URL — so a document that links to the world is drawn without it.
/// Chromium parses it as XML, in the mode built for animated images, so the document is
/// the document rather than a copy of it in a page of our own.
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
fn frame_page(
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

    // The version is in the image's URL and in the page's, so neither is answered out
    // of the browser's cache with a document that has been written since it was read.
    let html = format!(
        "<!doctype html><meta charset=\"utf-8\"><title>preview</title>\
         <style>html,body{{margin:0;padding:0;height:100%;overflow:hidden}}\
         img{{display:block;width:100%;height:100%;object-fit:contain}}</style>\
         {checkerboard}\
         <img src=\"{}?v={version}\" alt=\"\">",
        escape_attribute(&document),
        checkerboard = checkerboard_style(background)
    );

    write_page(&page, &html, version)
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
fn font_page(
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
    let lines: String = specimen
        .samples
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
    let (pangram_size, script_size, script_margin) = if specimen.samples.len() > SPECIMEN_FULL_LINES
    {
        ("6vh", "3.6vh", "1vh")
    } else {
        ("8vh", "5vh", "1.6vh")
    };

    let (ink, shadow) = specimen_ink(background);
    let html = format!(
        "<!doctype html><meta charset=\"utf-8\"><title>preview</title>\
         <style>\
         html,body{{margin:0;padding:0;height:100%;overflow:hidden}}\
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
         <style>@font-face{{font-family:\"RHPPreviewFont\";\
         src:url(\"{font}?v={version}\")}}</style>\
         {checkerboard}\
         <div class=\"title\">{title}</div>{lines}",
        font = escape_attribute(&font),
        title = escape_text(&specimen.title),
        checkerboard = checkerboard_style(background)
    );

    write_page(&page, &html, version)
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
fn escape_attribute(url: &str) -> String {
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
/// the document, and puts its window up — `is_showing` is what says it got there. What a
/// caller puts on screen while it waits is the waiting spinner and nothing else: this app
/// draws no document, so there is no still frame to hold the place.
///
/// The same document asked for a second time — which is what a wait that follows the
/// pointer does — is that window moved rather than the document navigated to again, and
/// a file that is not a document at all is not the engine's to draw.
pub fn show(path: &Path, area: Area, background: TransparentBackground) {
    if !draws(path) {
        trace(&format!(
            "show({}): not a document the engine draws",
            path.display()
        ));
        return;
    }

    let Ok(mut engine) = ENGINE.lock() else {
        trace("show: the engine's lock is poisoned");
        return;
    };

    let sender = engine.get_or_insert_with(Engine::start).sender.clone();

    if let Err(error) = sender.send(Command::Show {
        path: path.to_path_buf(),
        area,
        background,
    }) {
        trace(&format!(
            "show({}): the engine's thread is gone: {error}",
            path.display()
        ));
    }
}

/// Take the engine's window down. The engine itself is kept warm: what it costs to
/// begin is a browser start, and what it costs to point at another document is a few
/// milliseconds, so a hover that follows another one pays almost nothing.
pub fn hide() {
    let Ok(engine) = ENGINE.lock() else {
        return;
    };

    if let Some(engine) = engine.as_ref() {
        let _ = engine.sender.send(Command::Hide);
    }
}

/// Let the engine go, window, browser process and thread together. Called when the app
/// ends.
pub fn shutdown() {
    let engine = ENGINE.lock().ok().and_then(|mut engine| engine.take());

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

/// What the preview thread asks the engine's thread to do.
enum Command {
    Show {
        path: PathBuf,
        area: Area,
        background: TransparentBackground,
    },
    Hide,
    Shutdown,
}

struct Engine {
    sender: Sender<Command>,
    thread: std::thread::JoinHandle<()>,
}

impl Engine {
    /// Start the engine's thread. Nothing is created until the first document is asked
    /// for: a machine that never hovers one never starts a browser.
    fn start() -> Self {
        let (sender, receiver) = mpsc::channel();
        let thread = std::thread::spawn(move || engine_thread(receiver));

        Self { sender, thread }
    }
}

/// How long the engine is kept after its last document. It is read from the
/// configuration each time rather than captured, so an edit applies to the engine that
/// is already warm.
fn idle_timeout() -> Option<Duration> {
    CONFIG
        .lock()
        .map(|config| config.webview_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_WEBVIEW_IDLE_SECS))
        .as_duration()
}

/// Write one line to the trace file when `RHP_WEBVIEW_TRACE` is set, for the same
/// reason the probes exist: an engine that does not come up says nothing on its own,
/// and this is what it says. Nothing is written when the variable is not set, so an
/// ordinary run leaves no file anywhere.
fn trace(message: &str) {
    if std::env::var_os("RHP_WEBVIEW_TRACE").is_none() {
        return;
    }

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(std::env::temp_dir().join("rhp-webview-trace.log"))
    {
        use std::io::Write;
        let _ = writeln!(file, "{message}");
    }
}

fn engine_thread(commands: Receiver<Command>) {
    // The engine's own apartment, and its own thread: WebView2 must be created on a
    // thread that is pumping messages, and this is that thread.
    let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

    let mut host: Option<Host> = None;
    let mut idle_since = Instant::now();

    loop {
        // While there is nothing on screen the wait is a poll of the idle clock; while
        // a document is up it is a poll of the message queue, because a browser needs
        // the thread that made it to keep retrieving messages. With no engine at all
        // there is nothing to watch for, so the wait is long enough that only a document
        // wakes the thread, and an app left alone is one thread asleep on a channel.
        let wait = if SHOWING.load(Ordering::Acquire) {
            Duration::from_millis(5)
        } else if host.is_some() {
            Duration::from_millis(250)
        } else {
            Duration::from_secs(60 * 60)
        };

        match commands.recv_timeout(wait) {
            Ok(Command::Shutdown) => break,
            Ok(Command::Hide) => {
                if let Some(host) = host.as_mut() {
                    host.hide();
                }
                idle_since = Instant::now();
            }
            Ok(Command::Show {
                path,
                area,
                background,
            }) => {
                trace(&format!("engine: show {}", path.display()));

                if host.is_none() {
                    host = Host::create();

                    // An engine that could not be had is noted, so that hovers stop
                    // opening a box nothing will be drawn into until the folder it
                    // could not have is free again.
                    if host.is_some() {
                        note_engine_up();
                    } else {
                        note_failure();
                    }
                    trace(&format!("engine: host created: {}", host.is_some()));
                }

                if let Some(host) = host.as_mut() {
                    host.show(&path, area, background);
                    trace(&format!("engine: shown: {}", is_showing()));
                }
                idle_since = Instant::now();
            }
            Err(RecvTimeoutError::Timeout) => {}
            Err(RecvTimeoutError::Disconnected) => break,
        }

        pump_messages();

        // The engine draws two kinds, and both gates switched off is a browser held for
        // nothing: while neither can be shown no hover can be answered with a document or a
        // specimen at all, so there is nothing warm to keep. It is read here rather than being
        // told because a gate can be closed either way — in the tray, or in `config.ini` for
        // the watcher to reload — and because the thread that would be told is this one,
        // parked on its channel. The browser's own children are its business: ending it ends
        // them.
        if host.is_some() && !PreviewType::Vector.enabled() && !PreviewType::Fonts.enabled() {
            trace("engine: let go, the kinds are switched off");

            if let Some(mut host) = host.take() {
                host.close();
            }
            idle_since = Instant::now();
        }

        // A document that has been off screen for longer than the setting asks for is
        // one the engine is let go of, browser process and all.
        //
        // The thread is not let go of with it, and that is the whole of it: the channel
        // it holds is the one a hover sends into, and a thread that ended here would
        // leave every document after the first idle timeout with a message nobody reads.
        // What the app would show is no document at all — the engine's window is the
        // whole of an SVG preview — for the rest of the run.
        let expired = match (host.as_ref(), idle_timeout()) {
            (Some(_), Some(limit)) => {
                !SHOWING.load(Ordering::Acquire) && idle_since.elapsed() >= limit
            }
            _ => false,
        };

        if expired {
            trace("engine: let go after idle");

            if let Some(mut host) = host.take() {
                host.close();
            }
        }
    }

    if let Some(mut host) = host {
        host.close();
    }

    unsafe {
        windows::Win32::System::Com::CoUninitialize();
    }
}

/// Everything one engine is: the window it draws in, the environment, and the
/// controller over it.
struct Host {
    hwnd: HWND,
    environment: ICoreWebView2Environment,
    controller: ICoreWebView2Controller,
    webview: ICoreWebView2,
    /// The file the engine is holding, the backdrop its page was written for, and — for a font —
    /// the face of a collection it was read at, so a second hover on the same file is a window
    /// that is put back up rather than a navigation. A change of backdrop or of face is a page to
    /// write again, since three of the four backdrops are partly the page's own (the
    /// checkerboard's squares, a specimen's ink) and the page is what the browser caches, and a
    /// specimen of another face is another page and another font in it. The face is the index the
    /// page was written for, and `0` for a file that is not a collection.
    current: Option<(PathBuf, TransparentBackground, usize)>,
    /// The browser process this engine started, when it could be told which one it
    /// was. The runtime owns the browser, but the process is this app's own child —
    /// started by the loader in this process — and it is what is ended if it is
    /// somehow still there when the engine goes.
    browser_pid: u32,
}

impl Host {
    fn create() -> Option<Self> {
        register_class();

        // What the browser is running as before this engine is begun, so the one it
        // starts can be told from one an earlier engine of this same run left
        // closing: both are children of this process, and only the one that was not
        // there a moment ago is this engine's.
        let before = engine_processes::processes_named_by_parent(
            engine_processes::BROWSER_IMAGE,
            std::process::id(),
        );

        let folder = user_data_folder();
        std::fs::create_dir_all(&folder).ok()?;

        let window = create_host_window();
        trace(&format!(
            "host: window {:?} (last error {})",
            window.is_some(),
            windows::core::Error::from_win32().code().0
        ));
        let hwnd = window?;

        let started = Instant::now();
        let environment = create_environment(&folder);
        trace(&format!(
            "host: environment {:?} in {} ms",
            environment.is_some(),
            started.elapsed().as_millis()
        ));
        let environment = environment?;
        let environment_ms = started.elapsed().as_millis() as u64;

        let started = Instant::now();
        // A controller is refused while the folder this engine keeps its state in is
        // held by another browser — which is what a browser left behind by an earlier
        // run looks like, and it is usually gone within a moment. Asking again a few
        // times is worth more than the wait it costs the engine's own thread, and the
        // caller falls back to this app's reader if even that comes to nothing.
        let mut controller = None;
        for attempt in 0..4 {
            controller = create_controller(environment.clone(), hwnd);

            if controller.is_some() {
                break;
            }

            if attempt < 3 {
                std::thread::sleep(Duration::from_millis(250 * (attempt + 1)));
            }
        }

        trace(&format!(
            "host: controller {:?} in {} ms",
            controller.is_some(),
            started.elapsed().as_millis()
        ));
        let controller = controller?;
        let controller_ms = started.elapsed().as_millis() as u64;

        let webview = unsafe { controller.CoreWebView2().ok()? };
        configure(&webview);

        // The browser is the runtime's, but the process is this app's: the loader
        // started it here, which is what makes it findable by its parent, and what it
        // is put in the job and written down for is the same thing the Office engines
        // are — a process that must not outlive the app that started it, however the
        // app ends.
        let browser_pid = engine_processes::processes_named_by_parent(
            engine_processes::BROWSER_IMAGE,
            std::process::id(),
        )
        .into_iter()
        .find(|pid| !before.contains(pid))
        .unwrap_or(0);

        if browser_pid != 0 {
            engine_processes::record(engine_processes::BROWSER_IMAGE, browser_pid);
        }
        trace(&format!("host: browser {browser_pid}"));

        if let Ok(mut timings) = LAST_TIMINGS.lock() {
            timings.environment_ms = environment_ms;
            timings.controller_ms = controller_ms;
        }

        let host = Self {
            hwnd,
            environment,
            controller,
            webview,
            current: None,
            browser_pid,
        };

        Some(host)
    }

    fn show(&mut self, path: &Path, area: Area, background: TransparentBackground) {
        // Which face of a collection a specimen of this file is drawn from, read here rather
        // than inside the page: it is part of what the engine is holding, so a setting changed
        // between two hovers of one file is a page to write and navigate to again. A file that
        // is not a font is drawn from no face at all.
        let face = if font_formats::is_font_file(path) {
            font_preview::configured_face()
        } else {
            0
        };

        unsafe {
            // The background is a setting of the controller rather than of the page,
            // and it belongs to the interface that added it.
            if let Ok(controller) = self
                .controller
                .cast::<webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Controller2>()
            {
                let _ = controller.SetDefaultBackgroundColor(background_color(background));
            }

            let _ = self.controller.SetBounds(RECT {
                left: 0,
                top: 0,
                right: area.width,
                bottom: area.height,
            });
        }

        // A file the engine is not already holding, or one whose backdrop or face has changed,
        // is navigated to *before* the window is put up: a window shown first would be the file
        // before it, and what is on screen a moment ago belongs to another hover.
        if self.current.as_ref() != Some(&(path.to_path_buf(), background, face)) {
            let started = Instant::now();
            let arrived = self.navigate(path, background, face);
            trace(&format!(
                "engine: navigate {} arrived={arrived} in {} ms",
                path.display(),
                started.elapsed().as_millis()
            ));

            if !arrived {
                // The page was not put up. The engine is fine — the next file is drawn as
                // this one was meant to be — but the hover waiting on this one has nothing
                // left to wait for.
                self.hide();
                note_document_failed();
                return;
            }

            if let Ok(mut timings) = LAST_TIMINGS.lock() {
                timings.navigate_ms = started.elapsed().as_millis() as u64;
            }

            self.current = Some((path.to_path_buf(), background, face));
        }

        unsafe {
            let _ = self.controller.SetIsVisible(true);
            let _ = SetWindowPos(
                self.hwnd,
                HWND_TOPMOST,
                area.x,
                area.y,
                area.width,
                area.height,
                SWP_NOACTIVATE | SWP_SHOWWINDOW,
            );
            let _ = ShowWindow(self.hwnd, SW_SHOWNOACTIVATE);
        }

        SHOWING.store(true, Ordering::Release);
        publish_rect(Some(area));
    }

    fn hide(&mut self) {
        unsafe {
            let _ = ShowWindow(self.hwnd, SW_HIDE);
        }
        SHOWING.store(false, Ordering::Release);
        publish_rect(None);
    }

    fn close(&mut self) {
        self.hide();
        self.current = None;

        unsafe {
            let _ = self.controller.Close();
            let _ = DestroyWindow(self.hwnd);
        }

        // The environment goes with the host when it is dropped, and the browser
        // process it owns goes with the last controller over it.
        let _ = &self.environment;
    }

    /// Point the engine at `path` and wait for the page to arrive, pumping the thread's
    /// messages while it does.
    ///
    /// What it is pointed at is a page of this app's rather than the file itself, and which
    /// page is what the file is: a document's is the document as an image, which is what makes
    /// it the size of the window, and a font's is the font in the page with its own lines under
    /// it — read at `face`, which is which of a collection's faces is written out; see
    /// `frame_page` and `font_page`.
    fn navigate(&self, path: &Path, background: TransparentBackground, face: usize) -> bool {
        let version = file_version(path);
        let page = if font_formats::is_font_file(path) {
            font_page(path, version, background, face)
        } else {
            frame_page(path, version, background)
        };

        let Some((page, url)) = page else {
            return false;
        };

        trace(&format!(
            "engine: page {} for {}",
            page.display(),
            path.display()
        ));
        let url = wide(&url);
        let (sender, receiver) = mpsc::channel();

        unsafe {
            let handler =
                NavigationCompletedEventHandler::create(Box::new(move |_sender, _args| {
                    let _ = sender.send(());
                    Ok(())
                }));

            let mut token = EventRegistrationToken::default();
            if self
                .webview
                .add_NavigationCompleted(&handler, &mut token)
                .is_err()
            {
                return false;
            }

            let started = self.webview.Navigate(PCWSTR(url.as_ptr()));
            let arrived = started.is_ok() && webview2_com::wait_with_pump(receiver).is_ok();

            let _ = self.webview.remove_NavigationCompleted(token);

            arrived
        }
    }
}

impl Drop for Host {
    fn drop(&mut self) {
        SHOWING.store(false, Ordering::Release);
        publish_rect(None);

        // Closing the engine drops the environment, and the browser goes with the
        // last controller over it — usually. A browser that does not is the leftover
        // the profile folders are named for, still holding the folder of a run whose
        // app is gone, so the process is asked about here and ended if it is still
        // there: it is one this app started, and nothing of it should outlive the
        // app. Its own children are its business — ending it ends them.
        if self.browser_pid != 0 {
            if engine_processes::is_running(self.browser_pid) {
                engine_processes::terminate_owned(self.browser_pid);
            } else {
                engine_processes::forget(self.browser_pid);
            }
        }
    }
}

/// The window the engine draws into: a popup of its own, topmost, tool-windowed and
/// never activated.
fn create_host_window() -> Option<HWND> {
    unsafe {
        CreateWindowExW(
            WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
            WEBVIEW_CLASS,
            w!("Rust Hover Preview"),
            WS_POPUP,
            0,
            0,
            1,
            1,
            None,
            None,
            None,
            None,
        )
        .ok()
    }
}

fn register_class() {
    static REGISTERED: AtomicBool = AtomicBool::new(false);

    if REGISTERED.swap(true, Ordering::AcqRel) {
        return;
    }

    let class = WNDCLASSW {
        lpfnWndProc: Some(window_proc),
        lpszClassName: WEBVIEW_CLASS,
        ..Default::default()
    };

    unsafe {
        RegisterClassW(&class);
    }
}

/// A preview never takes the keyboard: the pointer may be over it, but what is being
/// worked in is Explorer.
extern "system" fn window_proc(
    hwnd: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    const MA_NOACTIVATE: LRESULT = LRESULT(3);

    if message == WM_MOUSEACTIVATE {
        return MA_NOACTIVATE;
    }

    unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
}

/// Retrieve and dispatch whatever is waiting, without blocking.
fn pump_messages() {
    let mut message = MSG::default();

    unsafe {
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

/// The folder the engine keeps its own state in — under the local profile, because a
/// browser profile is not something to synchronize between machines.
///
/// One folder is one browser at a time, and that is the whole reason this is a folder
/// *per run* rather than one folder for the app. An environment pointed at a folder
/// another browser is already holding is answered with `ERROR_INVALID_STATE` when it
/// asks for its controller, and the browser that holds it may be one left behind by a
/// run that ended badly — a zombie whose app is gone and which holds the folder for as
/// long as it lives, which can be indefinitely. A run that keeps its state in a folder
/// of its own cannot be blocked by any of that: the worst a leftover can do is take up
/// space, and the next start clears it away.
///
/// `RHP_WEBVIEW_PROFILE` names a folder for one run of the probes, so that a probe can
/// run beside a running app without either of them noticing the other.
fn user_data_folder() -> PathBuf {
    if let Some(folder) = std::env::var_os("RHP_WEBVIEW_PROFILE") {
        return PathBuf::from(folder);
    }

    profile_root().join(std::process::id().to_string())
}

/// The runs that left a profile folder behind.
///
/// Every folder under the browser's profile root is named for the run that made it,
/// so a folder that is not this run's names a run whose browser may still be holding
/// it — which is what reaches a browser left by a version of this app that wrote no
/// record of it. Read at startup, before this run has a folder of its own.
pub fn stale_profile_pids() -> Vec<u32> {
    engine_processes::stale_run_pids(&profile_root())
}

/// The folder the per-run folders live in.
fn profile_root() -> PathBuf {
    directories::BaseDirs::new()
        .map(|dirs| {
            dirs.data_local_dir()
                .join("rust-hover-preview")
                .join("webview")
        })
        .unwrap_or_else(std::env::temp_dir)
}

/// Clear away the folders earlier runs left, which nothing is using any more.
///
/// A folder a browser is still holding is left where it is: it cannot be removed while
/// it is open, and it is not this run's to insist on. What is left is taken by the next
/// start, once whatever held it is gone. Called once, before anything can create a
/// folder of this run's own.
pub fn clear_stale_profiles() {
    let root = profile_root();
    let ours = std::process::id().to_string();

    let Ok(entries) = std::fs::read_dir(&root) else {
        return;
    };

    for entry in entries.flatten() {
        if entry.file_name() == std::ffi::OsStr::new(&ours) {
            continue;
        }

        if entry.path().is_dir() {
            let _ = std::fs::remove_dir_all(entry.path());
        }
    }
}

fn create_environment(user_data_folder: &Path) -> Option<ICoreWebView2Environment> {
    let folder = wide(&user_data_folder.to_string_lossy());
    let (sender, receiver) = mpsc::channel();

    // What a document names is never fetched: the engine is given a resolver rule that
    // answers for no host at all, so an `<image href="http://…">` is a picture that
    // does not arrive rather than a request this app made. That is the promise the rest
    // of it keeps — a hover reads the file under the pointer and nothing else — and a
    // browser would otherwise break it without a line of anything being written.
    let options = CoreWebView2EnvironmentOptions::default();
    unsafe {
        options.set_additional_browser_arguments(BROWSER_ARGUMENTS.to_string());
    }
    let options: ICoreWebView2EnvironmentOptions = options.into();

    CreateCoreWebView2EnvironmentCompletedHandler::wait_for_async_operation(
        Box::new(move |handler| unsafe {
            CreateCoreWebView2EnvironmentWithOptions(
                PCWSTR::null(),
                PCWSTR(folder.as_ptr()),
                Some(&options),
                &handler,
            )
            .map_err(webview2_com::Error::WindowsError)
        }),
        Box::new(move |error_code, environment| {
            error_code?;
            sender
                .send(environment.ok_or_else(|| windows::core::Error::from(E_POINTER)))
                .expect("the waiting thread is gone");
            Ok(())
        }),
    )
    .ok()?;

    receiver.recv().ok()?.ok()
}

fn create_controller(
    environment: ICoreWebView2Environment,
    hwnd: HWND,
) -> Option<ICoreWebView2Controller> {
    let (sender, receiver) = mpsc::channel();

    CreateCoreWebView2ControllerCompletedHandler::wait_for_async_operation(
        Box::new(move |handler| unsafe {
            environment
                .CreateCoreWebView2Controller(hwnd, &handler)
                .map_err(webview2_com::Error::WindowsError)
        }),
        Box::new(move |error_code, controller| {
            trace(&format!(
                "controller callback: error={error_code:?} controller={}",
                controller.is_some()
            ));
            error_code?;
            sender
                .send(controller.ok_or_else(|| windows::core::Error::from(E_POINTER)))
                .expect("the waiting thread is gone");
            Ok(())
        }),
    )
    .ok()?;

    receiver.recv().ok()?.ok()
}

/// What the engine is and is not allowed to do, all of it the app's own rules rather
/// than a browser's: a document is drawn and not run, and nothing about it is a way out
/// of the preview.
fn configure(webview: &ICoreWebView2) {
    unsafe {
        if let Ok(settings) = webview.Settings() {
            let _ = settings.SetIsScriptEnabled(false);
            let _ = settings.SetAreDefaultContextMenusEnabled(false);
            let _ = settings.SetAreDevToolsEnabled(false);
            let _ = settings.SetIsZoomControlEnabled(false);
            let _ = settings.SetIsStatusBarEnabled(false);
            let _ = settings.SetIsWebMessageEnabled(false);
        }

        // The keys a browser answers for itself — a find bar, a print dialog — belong
        // to no preview.
        if let Ok(settings3) = webview.Settings().and_then(|settings| {
            settings.cast::<webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Settings3>()
        }) {
            let _ = settings3.SetAreBrowserAcceleratorKeysEnabled(false);
        }
    }
}

/// A file's version: when it was last written, in milliseconds since the epoch, and
/// nothing at all when that cannot be read.
///
/// It goes into the URL the document is drawn by, so an edited file is a different URL
/// and the browser's cache cannot answer a hover with the document as it was. A version
/// that cannot be read is a URL a hover shares with the previous one, which is no worse
/// than not asking: what comes back is the cache's copy of a document that has not been
/// changed as far as this app can tell.
fn file_version(path: &Path) -> u64 {
    std::fs::metadata(path)
        .and_then(|meta| meta.modified())
        .ok()
        .and_then(|modified| modified.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|since| since.as_millis() as u64)
        .unwrap_or(0)
}

/// The engine's background, which is what the preview's own backdrop setting means to a
/// window that composites for itself.
fn background_color(background: TransparentBackground) -> COREWEBVIEW2_COLOR {
    match background {
        TransparentBackground::Transparent => COREWEBVIEW2_COLOR {
            A: 0,
            R: 0,
            G: 0,
            B: 0,
        },
        TransparentBackground::Black => COREWEBVIEW2_COLOR {
            A: 255,
            R: 0,
            G: 0,
            B: 0,
        },
        TransparentBackground::White => COREWEBVIEW2_COLOR {
            A: 255,
            R: 255,
            G: 255,
            B: 255,
        },
        // A checkerboard is the one backdrop the engine cannot be given: it is drawn by
        // whatever composites the frame, and this window composites its own. The page
        // paints it instead — see `frame_page` — so what this colour is is what stands
        // behind the page until it has been drawn: mid grey, which is what the squares
        // average to.
        TransparentBackground::Checkerboard => COREWEBVIEW2_COLOR {
            A: 255,
            R: 184,
            G: 184,
            B: 184,
        },
    }
}

/// A local file as a URL, which is what the engine is pointed at.
///
/// The path a hover carries is the Shell's, and the Shell canonicalizes paths to the
/// verbatim form — `\\?\C:\…`, with `\\?\UNC\` in front of a share — which is not
/// something a URL may contain: a browser pointed at it fails at once, silently, and
/// what a hover shows is nothing. So the prefix comes off first, and a share keeps its
/// server: `\\?\UNC\server\share` becomes `file://server/share`, a drive becomes
/// `file:///C:/…`. The characters that would end the path early — a space, a hash, a
/// question mark, a percent — are escaped; anything else is left as it is written.
fn file_url(path: &Path) -> Option<String> {
    if !path.is_absolute() {
        return None;
    }

    let text = path.to_string_lossy();
    let local = match text.strip_prefix(r"\\?\UNC\") {
        Some(share) => format!(r"\\{share}"),
        None => text
            .strip_prefix(r"\\?\")
            .map(str::to_string)
            .unwrap_or_else(|| text.to_string()),
    };

    let (mut url, rest) = match local.strip_prefix(r"\\") {
        Some(share) => (String::from("file://"), share.to_string()),
        None => (String::from("file:///"), local),
    };

    for character in rest.chars() {
        match character {
            '\\' => url.push('/'),
            ' ' => url.push_str("%20"),
            '#' => url.push_str("%23"),
            '?' => url.push_str("%3F"),
            '%' => url.push_str("%25"),
            other => url.push(other),
        }
    }

    Some(url)
}

fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

fn pwstr_to_string(value: PWSTR) -> Option<String> {
    if value.is_null() {
        return None;
    }

    let mut length = 0usize;
    unsafe {
        while *value.0.add(length) != 0 {
            length += 1;
        }

        Some(String::from_utf16_lossy(std::slice::from_raw_parts(
            value.0, length,
        )))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The page the engine is pointed at is what makes a document the size of the window
    /// it was given: the document goes in an image, and an image is scaled to its box
    /// whatever size it asks for. What a document may not do — run code, leave the page —
    /// is what an image may not do either, which is why it is drawn this way rather than
    /// being put in the page as markup.
    ///
    /// A document whose name needs escaping is the case the URL builder and the attribute
    /// have to agree on: an `&` written as itself ends the attribute early and the image
    /// is a document the browser never found.
    #[test]
    fn the_page_draws_the_document_as_an_image_that_fills_it() {
        let folder = std::env::temp_dir().join("rust-hover-preview-frame-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let document = folder.join("a document & one.svg");
        std::fs::write(
            &document,
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"></svg>"#,
        )
        .expect("a written file");

        let (page, url) = frame_page(&document, 42, TransparentBackground::Black)
            .expect("a page for a document");
        let html = std::fs::read_to_string(&page).expect("a written page");

        assert!(html.contains("width:100%;height:100%;object-fit:contain"));
        assert!(html.contains("file:///"));
        assert!(html.contains("a%20document%20&amp;%20one.svg"));
        assert!(html.contains("?v=42"));
        assert!(
            !html.contains("conic-gradient"),
            "a backdrop the controller can be given is the controller's"
        );
        assert!(url.ends_with("?v=42"));
        assert!(url.starts_with("file:///"));

        // A checkerboard is the one backdrop it cannot be given, because it is drawn by
        // whatever composites the frame and the page is what composites this one.
        let (page, _) = frame_page(&document, 42, TransparentBackground::Checkerboard)
            .expect("a page for a document");
        let html = std::fs::read_to_string(&page).expect("a written page");

        assert!(html.contains("conic-gradient"));
        assert!(html.contains("#e0e0e0"));
        assert!(html.contains("background-size:32px 32px"));
    }

    /// A specimen's page is the font itself in the page, the lines the file's own character
    /// map covers, and a heading the font's `name` table supplies — the three things the page
    /// carries rather than asks the engine for. Two of the four backdrops are the page's too:
    /// the checkerboard's squares, and the ink a specimen is drawn in, which is dark on the
    /// light backdrops, light on black, and light with a shadow where there is no backdrop at
    /// all to be read against.
    #[test]
    fn the_page_draws_the_font_with_its_own_lines_and_a_name() {
        let folder = std::env::temp_dir().join("rust-hover-preview-font-page-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let font = folder.join("specimen.ttf");
        std::fs::write(&font, specimen_font()).expect("a written font");

        let (page, url) =
            font_page(&font, 7, TransparentBackground::Black, 0).expect("a page for a font");
        let html = std::fs::read_to_string(&page).expect("a written page");

        assert!(html.contains("@font-face"));
        assert!(html.contains("font-family:\"RHPPreviewFont\""));
        assert!(
            html.contains("specimen.ttf?v=7"),
            "the font is in the page as the file it is, at the version that was read"
        );
        assert!(html.contains("Test Family Regular"));
        assert!(html.contains("The quick brown fox jumps over the lazy dog."));
        assert!(
            html.contains("<p class=\"line pangram\" dir=\"auto\">"),
            "a line is drawn in its own script's direction, Arabic and Hebrew being written right to left"
        );
        assert!(html.contains("#f2f2f2"), "light ink on a black backdrop");
        assert!(
            !html.contains("conic-gradient"),
            "a backdrop the controller can be given is the controller's"
        );
        assert!(
            !html.contains("text-shadow"),
            "and so is the colour the text is drawn in"
        );
        assert!(url.ends_with("?v=7"));
        assert!(
            url.contains("font-black-0.html"),
            "the page is named for the backdrop and the face it was written for"
        );

        // The two backdrops the page owns: the squares, and the ink that reads on them.
        let (page, _) =
            font_page(&font, 7, TransparentBackground::Checkerboard, 0).expect("a page");
        let html = std::fs::read_to_string(&page).expect("a written page");

        assert!(html.contains("conic-gradient"));
        assert!(html.contains("#1a1a1a"), "dark ink on the light squares");

        // And the one with no backdrop to be read against: the ink carries its own shadow.
        let (page, _) = font_page(&font, 7, TransparentBackground::Transparent, 0).expect("a page");
        let html = std::fs::read_to_string(&page).expect("a written page");

        assert!(html.contains("text-shadow"));
        assert!(!html.contains("conic-gradient"));

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// A font of this test's own making: an sfnt with a `cmap` covering the pangram and a
    /// `name` table naming it, which is everything a specimen's page is built from.
    fn specimen_font() -> Vec<u8> {
        let mut codes: Vec<u32> = "The quick brown fox jumps over the lazy dog."
            .chars()
            .map(u32::from)
            .collect();
        codes.sort_unstable();
        codes.dedup();

        let mut cmap = Vec::new();
        cmap.extend_from_slice(&12u16.to_be_bytes()); // format
        cmap.extend_from_slice(&0u16.to_be_bytes()); // reserved
        cmap.extend_from_slice(&(16u32 + codes.len() as u32 * 12).to_be_bytes()); // length
        cmap.extend_from_slice(&0u32.to_be_bytes()); // language
        cmap.extend_from_slice(&(codes.len() as u32).to_be_bytes()); // numGroups
        for (index, code) in codes.iter().enumerate() {
            cmap.extend_from_slice(&code.to_be_bytes());
            cmap.extend_from_slice(&code.to_be_bytes());
            cmap.extend_from_slice(&(index as u32 + 1).to_be_bytes());
        }

        let mut cmap_table = Vec::new();
        cmap_table.extend_from_slice(&0u16.to_be_bytes()); // version
        cmap_table.extend_from_slice(&1u16.to_be_bytes()); // numTables
        cmap_table.extend_from_slice(&3u16.to_be_bytes()); // Windows
        cmap_table.extend_from_slice(&10u16.to_be_bytes()); // UCS-4
        cmap_table.extend_from_slice(&12u32.to_be_bytes()); // offset
        cmap_table.extend_from_slice(&cmap);

        let mut strings = Vec::new();
        let mut records = Vec::new();
        for (name_id, text) in [(1u16, "Test Family"), (2u16, "Regular")] {
            let offset = strings.len();
            for unit in text.encode_utf16() {
                strings.extend_from_slice(&unit.to_be_bytes());
            }

            records.extend_from_slice(&3u16.to_be_bytes()); // Windows
            records.extend_from_slice(&1u16.to_be_bytes()); // Unicode BMP
            records.extend_from_slice(&0x0409u16.to_be_bytes()); // English (United States)
            records.extend_from_slice(&name_id.to_be_bytes());
            records.extend_from_slice(&((strings.len() - offset) as u16).to_be_bytes());
            records.extend_from_slice(&(offset as u16).to_be_bytes());
        }

        let mut name_table = Vec::new();
        name_table.extend_from_slice(&0u16.to_be_bytes()); // format
        name_table.extend_from_slice(&2u16.to_be_bytes()); // count
        name_table.extend_from_slice(&((6 + 2 * 12) as u16).to_be_bytes()); // stringOffset
        name_table.extend_from_slice(&records);
        name_table.extend_from_slice(&strings);

        let tables: [(&[u8; 4], &[u8]); 2] = [
            (b"cmap", cmap_table.as_slice()),
            (b"name", name_table.as_slice()),
        ];
        let count = tables.len();
        let mut directory = Vec::new();
        let mut data = Vec::new();

        for (tag, table) in tables {
            directory.extend_from_slice(tag);
            directory.extend_from_slice(&0u32.to_be_bytes()); // checksum, which nothing reads
            directory.extend_from_slice(&((12 + count * 16 + data.len()) as u32).to_be_bytes());
            directory.extend_from_slice(&(table.len() as u32).to_be_bytes());
            data.extend_from_slice(table);
            while !data.len().is_multiple_of(4) {
                data.push(0);
            }
        }

        let mut font = Vec::new();
        font.extend_from_slice(b"\x00\x01\x00\x00");
        font.extend_from_slice(&(count as u16).to_be_bytes());
        font.extend_from_slice(&[0u8; 6]); // the binary-search fields
        font.extend_from_slice(&directory);
        font.extend_from_slice(&data);
        font
    }

    /// What a hover's path becomes when it is handed to the engine: the Shell's
    /// verbatim form is not a URL, and a document reached through it was a document the
    /// engine never opened.
    #[test]
    fn turns_a_verbatim_path_into_a_url_a_browser_opens() {
        for (path, expected) in [
            (
                r"C:\art\clock.svg",
                "file:///C:/art/clock.svg",
            ),
            (
                r"\\?\C:\art\clock.svg",
                "file:///C:/art/clock.svg",
            ),
            (r"\\?\C:\a b\c#d.svg", "file:///C:/a%20b/c%23d.svg"),
            (r"\\?\UNC\server\share\a.svg", "file://server/share/a.svg"),
            (r"\\server\share\a.svg", "file://server/share/a.svg"),
        ] {
            assert_eq!(
                file_url(Path::new(path)).as_deref(),
                Some(expected),
                "{path}"
            );
        }

        assert_eq!(file_url(Path::new(r"relative\a.svg")), None);
    }

    /// What the engine costs on this machine: beginning it, pointing it at a document,
    /// and pointing it at another one once it is warm. Ignored, and driven by
    /// `RHP_WEBVIEW_PROBE` — `$env:RHP_WEBVIEW_PROBE = "C:\art\one.svg; C:\art\two.svg";
    /// cargo test --release -- --ignored --nocapture webview_probe` — because it puts a
    /// window on the screen and starts a browser.
    #[test]
    #[ignore = "starts the WebView2 runtime and shows a window"]
    fn webview_probe() {
        let paths: Vec<PathBuf> = std::env::var("RHP_WEBVIEW_PROBE")
            .unwrap_or_default()
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
            .map(PathBuf::from)
            .collect();

        if paths.is_empty() {
            println!("set RHP_WEBVIEW_PROBE to one or more paths, separated by ';'");
            return;
        }

        println!("runtime: {:?}", runtime_version());
        println!("available: {}", is_available());

        let area = std::env::var("RHP_WEBVIEW_PROBE_AREA")
            .ok()
            .and_then(|size| {
                let (width, height) = size.split_once('x')?;
                Some(Area {
                    x: 60,
                    y: 60,
                    width: width.trim().parse().ok()?,
                    height: height.trim().parse().ok()?,
                })
            })
            .unwrap_or(Area {
                x: 60,
                y: 60,
                width: 800,
                height: 800,
            });

        for path in &paths {
            // Each document is measured from nothing on screen, so what the wait below
            // measures is this document rather than the window the last one left up.
            hide();
            let mut cleared = Duration::ZERO;
            while is_showing() && cleared < Duration::from_secs(2) {
                std::thread::sleep(Duration::from_millis(10));
                cleared += Duration::from_millis(10);
            }

            let drawn = draws(path);
            println!("{}: drawn={drawn}", path.display());

            let started = Instant::now();
            // Black by default, because a probe that measures what a document is drawn
            // at reads the screen and a transparent window shows the desktop through it,
            // which reads as a document that fills its window whatever it actually
            // draws. White is for a document drawn in dark strokes.
            let background = match std::env::var("RHP_WEBVIEW_PROBE_BACKGROUND").as_deref() {
                Ok("white") => TransparentBackground::White,
                _ => TransparentBackground::Black,
            };
            show(path, area, background);

            // The engine answers on its own thread; this is the wait for it to have
            // arrived rather than a measurement of the navigation itself.
            let mut waited = Duration::ZERO;
            while drawn && !is_showing() && waited < Duration::from_secs(5) {
                std::thread::sleep(Duration::from_millis(20));
                waited = started.elapsed();
            }

            let timings = last_timings();
            println!(
                "  showing={} after {} ms (environment {} ms, controller {} ms, navigation {} ms)",
                is_showing(),
                started.elapsed().as_millis(),
                timings.environment_ms,
                timings.controller_ms,
                timings.navigate_ms
            );

            // Kept up long enough for the screen to be looked at, which is what a probe
            // measuring what a document is drawn at needs and what a probe measuring a
            // navigation does not.
            let hold = std::env::var("RHP_WEBVIEW_PROBE_HOLD_MS")
                .ok()
                .and_then(|ms| ms.trim().parse().ok())
                .unwrap_or(1500);
            std::thread::sleep(Duration::from_millis(hold));
        }

        hide();
        std::thread::sleep(Duration::from_millis(200));
        println!("after hide: showing={}", is_showing());

        shutdown();
        println!("after shutdown: showing={}", is_showing());
    }
}

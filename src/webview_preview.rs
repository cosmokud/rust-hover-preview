//! Playing a document in the browser engine that is already on the machine.
//!
//! usvg drops animation — its own documentation says "no events and no animations" —
//! so `svg_animation` plays what it can read, and this plays the rest: the WebView2
//! runtime Windows 11 ships with, which is Chromium, and which is the only complete
//! implementation of SMIL and CSS animation that is on the machine without installing
//! anything. It is asked for a document that *moves* only. A still document is drawn by
//! `svg_preview`, costs no browser at all, and stays sharp at any size.
//!
//! What lives here is the engine and the window it draws in, not the preview loop: the
//! loop asks whether a document is one for the engine, hands it over with the box the
//! layout came out with, and this answers with a window of its own. That window is its
//! own because the preview window is a layered one, and a layered window has no window
//! tree to put a child in — the same shape as the video path, where the player's own
//! window is the preview.
//!
//! One engine is kept warm between documents and let go after `webview_idle`, ten
//! minutes by default: beginning one costs a browser start, and pointing a warm one at
//! another file costs a few milliseconds, so what a hover pays for a second animated
//! document is nothing worth measuring. What is let go of is the engine and not the
//! thread that holds it, which stays parked on its channel for the run and begins a new
//! engine for the next document: a thread that ended with its browser would leave every
//! document after the first idle timeout with nobody to play it, which is a still frame
//! for the rest of the run. An app left alone has no browser process and one thread
//! asleep — and the settings it is given are the app's own rules rather than a
//! browser's: a document is drawn and not run, and nothing about it is a way out of the
//! preview.

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
/// installed — which is what decides whether a document that moves is played by the
/// engine or by this app's own reader.
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

/// What the engine cost last time it was asked for a document. Read by the probe: it
/// is the number that says whether the engine is worth keeping warm at all.
#[cfg(test)]
pub fn last_timings() -> Timings {
    LAST_TIMINGS
        .lock()
        .map(|timings| *timings)
        .unwrap_or_default()
}

/// Whether a document is one the engine should play: the runtime is on the machine, the
/// document says it moves, and the engine is not in one of its own bad spells.
///
/// The declaration is what decides it rather than this app's own reader: the engine
/// plays the whole of SMIL and CSS, so a document that moves in a way `svg_animation`
/// cannot follow is still one to hand over. The answer is held with the parsed
/// document, so asking it again costs nothing.
pub fn moves(path: &Path) -> bool {
    is_available() && !is_failing() && svg_preview::moves(path)
}

/// How long a document that moves is played by this app rather than by the engine after
/// the engine has failed to come up.
///
/// The reason it can fail is a folder, not the document: one user data folder is one
/// browser at a time, and a browser left behind by an earlier run — one whose app was
/// ended before it could take its browser with it — holds it until it goes. What that
/// must not cost is the preview: for this long afterwards the reader plays what it can,
/// which is a picture that always plays, and the engine is asked again once the window
/// has passed.
const ENGINE_RETRY_AFTER: Duration = Duration::from_secs(300);

/// When the engine last failed to come up.
static ENGINE_FAILED_AT: Lazy<Mutex<Option<Instant>>> = Lazy::new(|| Mutex::new(None));

/// Whether the preview loop has not yet been told about the failure above.
static FAILURE_NOTICE: AtomicBool = AtomicBool::new(false);

/// Whether the engine has failed since this was last asked, which is what the preview
/// loop reads to lay the document out again for this app's own reader.
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

fn note_failure() {
    if let Ok(mut failed) = ENGINE_FAILED_AT.lock() {
        *failed = Some(Instant::now());
    }

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
/// is the share of the display `svg_scale` asked for: a document asking for 120 pixels is
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
fn frame_page(path: &Path, version: u64) -> Option<(PathBuf, String)> {
    let document = file_url(path)?;
    let page = user_data_folder().join("frame.html");

    // The version is in the image's URL and in the page's, so neither is answered out
    // of the browser's cache with a document that has been written since it was read.
    let html = format!(
        "<!doctype html><meta charset=\"utf-8\"><title>preview</title>\
         <style>html,body{{margin:0;padding:0;height:100%;overflow:hidden}}\
         img{{display:block;width:100%;height:100%;object-fit:contain}}</style>\
         <img src=\"{}?v={version}\" alt=\"\">",
        escape_attribute(&document)
    );

    std::fs::create_dir_all(user_data_folder()).ok()?;
    std::fs::write(&page, html).ok()?;

    let url = format!("{}?v={version}", file_url(&page)?);

    Some((page, url))
}

/// A URL as an attribute value: the two characters that would end it early.
fn escape_attribute(url: &str) -> String {
    url.replace('&', "&amp;").replace('"', "&quot;")
}

/// Ask the engine to play `path` in a window at `area`.
///
/// The answer is immediate and says nothing about whether the document arrived: the
/// engine works on its own thread, and what it does with this is navigates, waits for
/// the document, and puts its window up — `is_showing` is what says it got there. A
/// caller that wants something on screen in the meantime has one: it is the still frame
/// `svg_preview` drew, which this lands on top of.
///
/// A document that does not move is not the engine's: a browser is not started for a
/// picture, and a hover onto one is drawn by this app as it always was.
pub fn show(path: &Path, area: Area, background: TransparentBackground) {
    if !moves(path) {
        trace(&format!("show({}): not the engine's", path.display()));
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
    /// for: a machine that never hovers an animated document never starts a browser.
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

                    // An engine that could not be had is noted, so that a document that
                    // moves is played by this app's own reader rather than left as a
                    // still frame until the folder it could not have is free again.
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

        // A document's kind switched off is a browser held for nothing: while that gate
        // is off no hover can be answered with an animation at all, so there is nothing
        // warm to keep. It is read here rather than being told because the gate can be
        // closed either way — in the tray, or in `config.ini` for the watcher to reload
        // — and because the thread that would be told is this one, parked on its
        // channel. The browser's own children are its business: ending it ends them.
        if host.is_some() && !PreviewType::Svg.enabled() {
            trace("engine: let go, the kind is switched off");

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
        // What the app would show is the still frame — the engine's window is what the
        // animation was — for the rest of the run.
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
    /// The document the engine is holding, so a second hover on the same file is a
    /// window that is put back up rather than a navigation.
    current: Option<PathBuf>,
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

        Some(Self {
            hwnd,
            environment,
            controller,
            webview,
            current: None,
            browser_pid,
        })
    }

    fn show(&mut self, path: &Path, area: Area, background: TransparentBackground) {
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

        // A document the engine is not already holding is navigated to *before* the
        // window is put up: a window shown first would be the document before it, and
        // what is on screen a moment ago is the still frame this lands on top of.
        if self.current.as_deref() != Some(path) {
            let started = Instant::now();
            let arrived = self.navigate(path);
            trace(&format!(
                "engine: navigate {} arrived={arrived} in {} ms",
                path.display(),
                started.elapsed().as_millis()
            ));

            if !arrived {
                self.hide();
                return;
            }

            if let Ok(mut timings) = LAST_TIMINGS.lock() {
                timings.navigate_ms = started.elapsed().as_millis() as u64;
            }

            self.current = Some(path.to_path_buf());
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
    }

    fn hide(&mut self) {
        unsafe {
            let _ = ShowWindow(self.hwnd, SW_HIDE);
        }
        SHOWING.store(false, Ordering::Release);
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

    /// Point the engine at the document and wait for the page to arrive, pumping the
    /// thread's messages while it does.
    ///
    /// What it is pointed at is the page `frame_page` writes rather than the document
    /// itself, which is what makes the document the size of the window: see there.
    fn navigate(&self, path: &Path) -> bool {
        let version = file_version(path);
        let Some((page, url)) = frame_page(path, version) else {
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
        // the window that composites the frame, and this window composites its own.
        // Mid grey is what its squares average to.
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
/// verbatim form — `\\?\G:\…`, with `\\?\UNC\` in front of a share — which is not
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

        let (page, url) = frame_page(&document, 42).expect("a page for a document");
        let html = std::fs::read_to_string(&page).expect("a written page");

        assert!(html.contains("width:100%;height:100%;object-fit:contain"));
        assert!(html.contains("file:///"));
        assert!(html.contains("a%20document%20&amp;%20one.svg"));
        assert!(html.contains("?v=42"));
        assert!(url.ends_with("?v=42"));
        assert!(url.starts_with("file:///"));
    }

    /// What a hover's path becomes when it is handed to the engine: the Shell's
    /// verbatim form is not a URL, and a document reached through it was a document the
    /// engine never opened.
    #[test]
    fn turns_a_verbatim_path_into_a_url_a_browser_opens() {
        for (path, expected) in [
            (
                r"G:\Downloads\Stash\Animated_clock.svg",
                "file:///G:/Downloads/Stash/Animated_clock.svg",
            ),
            (
                r"\\?\G:\Downloads\Stash\Animated_clock.svg",
                "file:///G:/Downloads/Stash/Animated_clock.svg",
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

            let moving = moves(path);
            println!("{}: moves={moving}", path.display());

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
            while moving && !is_showing() && waited < Duration::from_secs(5) {
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

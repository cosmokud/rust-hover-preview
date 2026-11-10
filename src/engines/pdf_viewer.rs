//! A pinned PDF shown by the viewer the WebView2 runtime already carries.
//!
//! The hover's PDF road reads a raster of page one (`readers::pdf_preview`), and it stays that
//! road: a hover is a card the pointer is on, and one page of it is the whole of what a card can
//! say. A *pin* is the other thing — a document somebody works in — so a pinned PDF is handed to
//! Chromium instead: the runtime's own PDF viewer, pointed at the file itself, which is what makes
//! pages, continuous scroll, zoom, thumbnails, search and text selection Chromium's rather than
//! this app's (see ARCHITECTURE.md, "The two PDF roads").
//!
//! **Shape: one thread, one environment, one controller per viewer.** WebView2 wants the thread that
//! made a controller to keep retrieving messages for it, so every ask about a controller is made
//! here and nowhere else, and the rest of the app talks to this module rather than to a controller.
//! A viewer's window is a plain popup of this app's — no layered surface and no alpha, which is the
//! one visual difference from a picture pin and is deliberate: Chromium's pixels are the window's
//! own, and the Shell's DWM look comes for free. It is *owned* by the pin's window, so it stands
//! above it, is hidden and minimized with it, and needs no placement road of its own beyond the box
//! the pin's content band has.
//!
//! **Nothing here is the hover engine's.** `SHOWING`, `ENGINE_HWND`, `SHOWING_RECT` and
//! `webview_idle` describe the one hover page and are left alone: a PDF pin has its own refcount and
//! no TTL at all. The last pin closed closes the controller, closes the host window and ends the
//! thread; the next pin begins a fresh browser.
//!
//! **A viewer that never arrives shows nothing.** The window is created hidden and revealed only
//! when `NavigationCompleted` says the document loaded. What is on screen until then — and what
//! stays on screen for good where the engine or the navigation fails — is the static raster the pin
//! was already showing, so a pin is never a white window.

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use std::sync::mpsc::{self, Receiver, RecvTimeoutError, Sender};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use webview2_com::Microsoft::Web::WebView2::Win32::{
    ICoreWebView2, ICoreWebView2Controller, ICoreWebView2Environment, ICoreWebView2_3,
};
use webview2_com::NavigationCompletedEventHandler;
use windows::core::{w, Interface, PCWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, POINT, RECT, WPARAM};
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::System::WinRT::EventRegistrationToken;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, GetAncestor, GetClassNameW,
    GetClientRect, GetForegroundWindow, IsWindowVisible, PeekMessageW, PostMessageW,
    RegisterClassW, SetWindowPos, ShowWindow, TranslateMessage, WindowFromPoint, GA_ROOT, MSG,
    PM_NOREMOVE, PM_REMOVE, SWP_NOACTIVATE, SWP_NOZORDER, SW_SHOWNOACTIVATE, WHEEL_DELTA,
    WINDOW_EX_STYLE, WNDCLASSW, WS_POPUP,
};

use crate::app::engine_processes;
use crate::engines::webview_preview::{
    create_controller, create_environment, file_url, trace, wide, Area,
};

/// The window class every viewer window is made of. It is asked of the desktop by the wheel seam
/// (see `the_window_under`), which may take no lock, so the answer is a class name rather than a
/// table: one `GetClassNameW` on the root above the point.
const VIEWER_CLASS: PCWSTR = w!("RhpPdfViewer");

/// What the browser is given, which is nothing: the default runtime, in its own PDF viewer. The
/// hover engine's arguments are deliberately not reused — `--hide-scrollbars` would take the
/// scrollbar out of a document the user is meant to scroll.
const BROWSER_ARGUMENTS: &str = "";

/// How long a viewer may take to arrive before the pin is read as one that will not: a local PDF
/// through the runtime's own reader is a browser start and a file read, so what is past this is a
/// navigation that is not coming (see `Host::settle`).
const NAVIGATION_TIMEOUT: Duration = Duration::from_secs(15);

/// Whether a viewer is on screen — revealed, so the pin's band is Chromium's rather than the static
/// raster's. Read by the paint and zoom roads, which skip the band while it is set.
static VIEWER_OPEN: AtomicBool = AtomicBool::new(false);

/// How many viewers the app has asked for and not yet let go of. The thread lives while it is not
/// zero, and that is the whole lifetime rule: refcount, no TTL.
static OPEN: AtomicUsize = AtomicUsize::new(0);

/// The engine: the thread and the channel into it, begun by the first `open` and ended by the last
/// `close`.
static ENGINE: Lazy<Mutex<Option<Engine>>> = Lazy::new(|| Mutex::new(None));

/// The channel into the viewer thread. The handle to the thread is deliberately not kept: a thread
/// whose channel is dropped ends on its own — it closes what it still holds and returns — so a pin
/// closing never waits on a browser on the thread that draws the previews.
struct Engine {
    sender: Sender<Command>,
}

/// What the app asks of a viewer.
enum Command {
    Open {
        path: PathBuf,
        /// The pin's window as a number: the viewer's owner, and the window whose content band the
        /// viewer fills. A handle rather than the type because this crosses a thread.
        owner: isize,
        area: Area,
    },
    /// The newest box the pin's band has, published rather than queued per pointer move: a drag
    /// asks on every move and only the last one is a box to be in (the same shape as the hover
    /// engine's `PLACED`, for the same reason).
    Place(Area),
    Close(PathBuf),
    Shutdown,
}

/// Whether a viewer stands in a pin's band, which is what stops the pin painting the raster it still
/// holds underneath it.
pub(crate) fn is_up() -> bool {
    VIEWER_OPEN.load(Ordering::Acquire)
}

/// Ask for a viewer on `path`, in the pin `owner`, at `area`, counting one more viewer owed.
///
/// The ask is taken as soon as there is a thread to take it: what it becomes is a browser start and
/// a navigation, both of which are the viewer thread's and neither of which the caller waits for.
pub(crate) fn open(path: &Path, owner: isize, area: Area) -> bool {
    if path.as_os_str().is_empty() || owner == 0 {
        return false;
    }

    let mut engine = match ENGINE.lock() {
        Ok(engine) => engine,
        Err(_) => return false,
    };

    if engine.is_none() {
        let (sender, receiver) = mpsc::channel();
        std::thread::spawn(move || viewer_thread(receiver));

        *engine = Some(Engine { sender });
    }

    let taken = engine.as_ref().is_some_and(|engine| {
        engine
            .sender
            .send(Command::Open {
                path: path.to_path_buf(),
                owner,
                area,
            })
            .is_ok()
    });

    if taken {
        OPEN.fetch_add(1, Ordering::AcqRel);
    }

    taken
}

/// Move every viewer to `area` — the pin's content band, which is the whole of what a viewer is
/// placed by.
pub(crate) fn place(area: Area) {
    if !is_up() {
        return;
    }

    if let Ok(engine) = ENGINE.lock() {
        if let Some(engine) = engine.as_ref() {
            let _ = engine.sender.send(Command::Place(area));
        }
    }
}

/// Let go of `path`'s viewer, and — where it was the last one owed — of the browser and the thread.
pub(crate) fn close(path: &Path) {
    let mut engine = match ENGINE.lock() {
        Ok(engine) => engine,
        Err(_) => return,
    };

    let Some(held) = engine.as_mut() else {
        return;
    };

    let _ = held.sender.send(Command::Close(path.to_path_buf()));

    let left = OPEN.load(Ordering::Acquire).saturating_sub(1);
    OPEN.store(left, Ordering::Release);

    if left == 0 {
        // Dropping the channel is what ends the thread: it closes what it still holds — controller,
        // window, environment — on its own way out, and nothing here waits for it (see `Engine`).
        if let Some(engine) = engine.take() {
            let _ = engine.sender.send(Command::Shutdown);
        }

        VIEWER_OPEN.store(false, Ordering::Release);
    }
}

/// Let go of every viewer at once: the run is ending, or a pin went away without naming a file.
pub(crate) fn close_all() {
    OPEN.store(0, Ordering::Release);

    let mut engine = match ENGINE.lock() {
        Ok(engine) => engine,
        Err(_) => return,
    };

    if let Some(engine) = engine.take() {
        let _ = engine.sender.send(Command::Shutdown);
    }

    VIEWER_OPEN.store(false, Ordering::Release);
}

/// The viewer's own lifecycle: the four states it can be in, and the call each transition owes.
///
/// It is a state machine rather than a handful of flags because its three questions are asked from
/// different roads — a navigation that arrives, a pin that loses the foreground, a pin that goes
/// away — and every one of them is a *transition* rather than a setting: a viewer is revealed once,
/// suspended while it is not the window the user is in, resumed rather than begun again, and closed
/// exactly once however many of those arrive behind it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Life {
    /// Asked for, created hidden, navigated: nothing of it is on screen.
    Arriving,
    /// The document arrived and the window is up, which is the only state a notch or a click is
    /// answered in.
    Shown,
    /// Up but put away (`SetIsVisible(false)` then `TrySuspend`): occluded, unfocused or minimized.
    Suspended,
    /// Closed, and nothing more is owed to it.
    Closed,
}

/// What a transition of `Life` owes, which is the whole of what the viewer thread does about it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Step {
    /// Put the window up: the document arrived.
    Reveal,
    /// Stop drawing: the pin is not the window the user is in.
    Suspend,
    /// Draw again.
    Resume,
    /// Let go of it, which is also what a document that never arrived is owed: the pin's static
    /// raster is what the user is left with (see the module docs).
    Close,
    /// Nothing at all.
    Nothing,
}

impl Life {
    /// A navigation that answered. Only a document that arrived is revealed.
    fn navigated(&mut self, arrived: bool) -> Step {
        if *self != Life::Arriving {
            return Step::Nothing;
        }

        if arrived {
            *self = Life::Shown;
            Step::Reveal
        } else {
            *self = Life::Closed;
            Step::Close
        }
    }

    /// Whether the pin is the window the user is in, asked once a pass while a document is up.
    fn settle(&mut self, the_pin_is_the_window_the_user_is_in: bool) -> Step {
        match (*self, the_pin_is_the_window_the_user_is_in) {
            (Life::Shown, false) => {
                *self = Life::Suspended;
                Step::Suspend
            }
            (Life::Suspended, true) => {
                *self = Life::Shown;
                Step::Resume
            }
            _ => Step::Nothing,
        }
    }

    /// Let go of: one close, and every state after it is the same state.
    fn closed(&mut self) -> Step {
        if *self == Life::Closed {
            return Step::Nothing;
        }

        *self = Life::Closed;
        Step::Close
    }

    fn is_arriving(self) -> bool {
        self == Life::Arriving
    }

    fn is_closed(self) -> bool {
        self == Life::Closed
    }

    fn is_up(self) -> bool {
        matches!(self, Life::Shown | Life::Suspended)
    }
}

/// One viewer: its window, the controller over it, and where it has got to.
struct Host {
    path: PathBuf,
    hwnd: HWND,
    owner: HWND,
    controller: ICoreWebView2Controller,
    webview: ICoreWebView2,
    life: Life,
    area: Area,
    deadline: Instant,
    /// What `NavigationCompleted` said, sent by the handler on this same thread and taken by the
    /// thread's own pass (see `settle`).
    answer: Receiver<bool>,
    /// Whether the browser has been asked to suspend, so the ask is made once (see `suspend`).
    suspended: bool,
}

impl Host {
    fn create(
        environment: &ICoreWebView2Environment,
        path: &Path,
        owner: isize,
        area: Area,
    ) -> Option<Self> {
        register_class();

        let owner = HWND(owner as *mut _);
        let deadline = Instant::now() + NAVIGATION_TIMEOUT;

        // The window is created hidden and revealed only by a document that arrived: a window put
        // up first would be the pin's band gone white for as long as a browser takes to load a
        // file, which is the one thing a preview may never do.
        let hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                VIEWER_CLASS,
                w!("Rust Hover Preview"),
                WS_POPUP,
                area.x,
                area.y,
                area.width,
                area.height,
                owner,
                None,
                None,
                None,
            )
        }
        .ok()?;

        // A viewer is closed on every road that gives up on it, and the window it was made with is
        // the one it closes — so a failure below has one teardown rather than four.
        let Some(controller) = create_controller(environment.clone(), hwnd, deadline) else {
            trace("pdf: no controller");
            unsafe {
                let _ = DestroyWindow(hwnd);
            }
            return None;
        };

        let webview = unsafe { controller.CoreWebView2().ok() };

        let Some(webview) = webview else {
            unsafe {
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
            }
            return None;
        };

        configure(&webview);

        let Some(url) = file_url(path) else {
            unsafe {
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
            }
            return None;
        };

        let (sender, answer) = mpsc::channel::<bool>();

        unsafe {
            let handler =
                NavigationCompletedEventHandler::create(Box::new(move |_sender, args| {
                    let mut arrived = windows::Win32::Foundation::BOOL(0);
                    let _ = args.is_some_and(|args| args.IsSuccess(&mut arrived).is_ok());
                    let arrived = arrived.as_bool();
                    let _ = sender.send(arrived);
                    Ok(())
                }));

            let mut token = EventRegistrationToken::default();

            if webview
                .add_NavigationCompleted(&handler, &mut token)
                .is_err()
            {
                trace("pdf: no navigation handler");
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
                return None;
            }

            let url = wide(&url);
            let started = webview.Navigate(PCWSTR(url.as_ptr()));

            if started.is_err() {
                trace("pdf: would not navigate");
                let _ = controller.Close();
                let _ = DestroyWindow(hwnd);
                return None;
            }
        }

        trace(&format!("pdf: navigating to {}", path.display()));

        Some(Self {
            path: path.to_path_buf(),
            hwnd,
            owner,
            controller,
            webview,
            life: Life::Arriving,
            area,
            deadline,
            answer,
            suspended: false,
        })
    }

    /// What this viewer is owed on this pass: a document that arrived, one that did not answer in
    /// time, and the suspension the point of the user's attention decides.
    fn settle(&mut self) {
        if self.life.is_arriving() {
            match self.answer.try_recv() {
                Ok(arrived) => match self.life.navigated(arrived) {
                    Step::Reveal => self.reveal(),
                    Step::Close => {
                        trace(&format!("pdf: nothing arrived for {}", self.path.display()));
                        self.close();
                        return;
                    }
                    _ => {}
                },
                Err(mpsc::TryRecvError::Empty) => {
                    if Instant::now() >= self.deadline {
                        trace(&format!(
                            "pdf: gave up on {} after {} s",
                            self.path.display(),
                            NAVIGATION_TIMEOUT.as_secs()
                        ));
                        self.close();
                        return;
                    }
                }
                Err(mpsc::TryRecvError::Disconnected) => {
                    self.close();
                    return;
                }
            }
        }

        if !self.life.is_up() {
            return;
        }

        let running = self.the_pin_is_the_window_the_user_is_in();

        match self.life.settle(running) {
            Step::Suspend => self.suspend(),
            Step::Resume => self.wake(),
            _ => {}
        }
    }

    /// Whether this viewer's window — or the pin that owns it — is the window the user is in.
    ///
    /// It is asked of the foreground rather than of `WM_ACTIVATE` because a pin is put away by the
    /// Shell as well as by its own buttons (a minimize, a taskbar visit, another program's click),
    /// and a window that is not visible is one no compositor is drawing: both are one question the
    /// pass can ask, and it is the answer `SetIsVisible`/`TrySuspend` want.
    fn the_pin_is_the_window_the_user_is_in(&self) -> bool {
        unsafe {
            if !IsWindowVisible(self.hwnd).as_bool() {
                return false;
            }

            let foreground = GetForegroundWindow();

            if foreground == self.hwnd || foreground == self.owner {
                return true;
            }

            // The caret may be inside Chromium's own child of the viewer, which is still this
            // viewer's window: the root above it is the one the question is about.
            !foreground.is_invalid() && GetAncestor(foreground, GA_ROOT) == self.hwnd
        }
    }

    fn reveal(&mut self) {
        self.set_bounds();

        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                HWND::default(),
                self.area.x,
                self.area.y,
                self.area.width,
                self.area.height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
            let _ = self.controller.SetIsVisible(true);
            let _ = ShowWindow(self.hwnd, SW_SHOWNOACTIVATE);
        }

        VIEWER_OPEN.store(true, Ordering::Release);
        trace(&format!(
            "pdf: shown {} at {}x{}+{}+{}",
            self.path.display(),
            self.area.width,
            self.area.height,
            self.area.x,
            self.area.y
        ));
    }

    /// The controller's own box, which is the window's client: a page is laid out at the size it is
    /// drawn in, so the browser is told the same box the window has.
    fn set_bounds(&mut self) {
        let mut client = RECT::default();

        if unsafe { GetClientRect(self.hwnd, &mut client) }.is_err() {
            return;
        }

        unsafe {
            let _ = self.controller.SetBounds(RECT {
                left: 0,
                top: 0,
                right: client.right.max(1),
                bottom: client.bottom.max(1),
            });
        }
    }

    fn place(&mut self, area: Area) {
        self.area = area;

        if !self.life.is_up() {
            // A box asked for before the document arrived is the one it arrives in (see `reveal`).
            return;
        }

        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                HWND::default(),
                area.x,
                area.y,
                area.width,
                area.height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }

        self.set_bounds();
    }

    /// Stop drawing, and keep the document: a suspended browser is one still there to be resumed,
    /// which is what makes a pin the user works away from cheap rather than begun again.
    fn suspend(&mut self) {
        if self.suspended {
            return;
        }

        unsafe {
            let _ = self.controller.SetIsVisible(false);
        }

        // `TrySuspend` states a precondition rather than a preference: the controller must already
        // be hidden (see above), and a runtime older than suspension is answered by the hidden
        // controller alone.
        if let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() {
            unsafe {
                let _ = webview.TrySuspend(&webview2_com::TrySuspendCompletedHandler::create(
                    Box::new(|_code, _successful| Ok(())),
                ));
            }
        }

        self.suspended = true;
        trace(&format!("pdf: suspended {}", self.path.display()));
    }

    fn wake(&mut self) {
        if !self.suspended {
            return;
        }

        self.suspended = false;

        if let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() {
            unsafe {
                let _ = webview.Resume();
            }
        }

        unsafe {
            let _ = self.controller.SetIsVisible(true);
        }

        trace(&format!("pdf: resumed {}", self.path.display()));
    }

    fn close(&mut self) {
        if self.life.closed() == Step::Nothing {
            return;
        }

        unsafe {
            let _ = self.controller.Close();
            let _ = DestroyWindow(self.hwnd);
        }

        VIEWER_OPEN.store(false, Ordering::Release);
        trace(&format!("pdf: closed {}", self.path.display()));
    }
}

/// The browser's own furniture for a document to be worked in: the developer tools, the status bar
/// and the messages a page could reach this app's threads through are all taken away, and nothing
/// else is. The context menu, the accelerators (`Ctrl+F`, `Ctrl+wheel`, `Ctrl+C`) and the zoom stay
/// Chromium's, because the whole point of this road is that the document is used as a document (see
/// the module docs).
fn configure(webview: &ICoreWebView2) {
    unsafe {
        if let Ok(settings) = webview.Settings() {
            let _ = settings.SetAreDevToolsEnabled(false);
            let _ = settings.SetIsStatusBarEnabled(false);
            let _ = settings.SetIsWebMessageEnabled(false);
        }
    }
}

fn register_class() {
    static REGISTERED: AtomicBool = AtomicBool::new(false);

    if REGISTERED.swap(true, Ordering::AcqRel) {
        return;
    }

    let class = WNDCLASSW {
        lpfnWndProc: Some(window_proc),
        lpszClassName: VIEWER_CLASS,
        ..Default::default()
    };

    unsafe {
        let _ = RegisterClassW(&class);
    }
}

/// A viewer window answers nothing of its own: the pin that owns it is the window with the buttons,
/// and every click, key and drag inside the document is Chromium's (see the module docs).
extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> windows::Win32::Foundation::LRESULT {
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

fn viewer_thread(commands: Receiver<Command>) {
    // The browser's own apartment, and its own thread: WebView2 must be created on a thread that is
    // pumping messages, and this is that thread for every viewer this app has open.
    let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

    // The thread's message queue, made before anything can be posted to it.
    {
        let mut message = MSG::default();
        unsafe {
            let _ = PeekMessageW(&mut message, None, 0, 0, PM_NOREMOVE);
        }
    }

    let folder = crate::engines::webview_preview::a_pins_profile_folder();
    let _ = std::fs::create_dir_all(&folder);

    let mut environment: Option<ICoreWebView2Environment> = None;
    let mut hosts: Vec<Host> = Vec::new();

    loop {
        // While a document is up the thread polls, because a browser needs the thread that made it
        // to keep retrieving messages; with nothing open there is nothing to watch for and the wait
        // is the channel's own.
        let wait = if hosts.is_empty() {
            Duration::from_secs(60 * 60)
        } else {
            Duration::from_millis(5)
        };

        match commands.recv_timeout(wait) {
            Ok(Command::Shutdown) => break,
            Ok(Command::Open { path, owner, area }) => {
                if environment.is_none() {
                    // The browser the runtime starts is this app's child whatever engine asked for
                    // it: one begun here is written down like the hover engine's, so a run that
                    // ends badly leaves no browser of its own behind (see `app::engine_processes`).
                    let before = engine_processes::processes_named_by_parent(
                        engine_processes::BROWSER_IMAGE,
                        std::process::id(),
                    );

                    let deadline = Instant::now() + NAVIGATION_TIMEOUT;
                    environment = create_environment(&folder, deadline, BROWSER_ARGUMENTS);

                    if environment.is_none() {
                        trace("pdf: no environment");
                    }

                    for pid in engine_processes::processes_named_by_parent(
                        engine_processes::BROWSER_IMAGE,
                        std::process::id(),
                    ) {
                        if !before.contains(&pid) {
                            trace(&format!("pdf: browser {pid}"));
                            engine_processes::record(engine_processes::BROWSER_IMAGE, pid);
                        }
                    }
                }

                let host = environment
                    .as_ref()
                    .and_then(|environment| Host::create(environment, &path, owner, area));

                if host.is_none() {
                    trace(&format!("pdf: no viewer for {}", path.display()));
                }

                if let Some(host) = host {
                    hosts.push(host);
                }
            }
            Ok(Command::Place(area)) => {
                for host in &mut hosts {
                    host.place(area);
                }
            }
            Ok(Command::Close(path)) => {
                for host in hosts.iter_mut().filter(|host| host.path == path) {
                    host.close();
                }
            }
            Err(RecvTimeoutError::Timeout) => {}
            Err(RecvTimeoutError::Disconnected) => break,
        }

        for host in &mut hosts {
            host.settle();
        }

        hosts.retain(|host| !host.life.is_closed());

        pump_messages();
    }

    for host in &mut hosts {
        host.close();
    }

    drop(environment);

    VIEWER_OPEN.store(false, Ordering::Release);

    unsafe {
        CoUninitialize();
    }
}

/// Retrieve and dispatch whatever is waiting, without blocking: the thread a controller was made on
/// is the thread its callbacks arrive on.
fn pump_messages() {
    let mut message = MSG::default();

    unsafe {
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

/// The window a notch over a viewer is posted to: Chromium's own child of it, or the viewer window
/// itself where the point is on its furniture.
///
/// Asked by the wheel hook, which may take no lock and runs for every wheel message on the desktop:
/// the answer is one `WindowFromPoint`, one `GetAncestor` and one class name, and nothing is kept in
/// a table this side would have to be told about. A window of another class is three calls and a
/// comparison, which is what a notch over anything else costs.
pub(crate) fn the_window_under(point: POINT) -> Option<HWND> {
    unsafe {
        let window = WindowFromPoint(point);

        if window.is_invalid() {
            return None;
        }

        let root = GetAncestor(window, GA_ROOT);
        let root = if root.is_invalid() { window } else { root };

        // Chromium's child is the window that scrolls: a notch posted to the container would be a
        // message nobody in the browser is listening for.
        class_is_the_viewers(root).then_some(window)
    }
}

unsafe fn class_is_the_viewers(window: HWND) -> bool {
    let mut name = [0u16; 32];
    let len = unsafe { GetClassNameW(window, &mut name) };

    if len <= 0 {
        return false;
    }

    let name = &name[..(len as usize).min(name.len())];

    name == VIEWER_CLASS.as_wide()
}

/// Post one wheel notch to `window`, which is Chromium's, at `point` in screen coordinates.
///
/// The message is the one Windows would have handed the viewer had the keyboard been in it: a notch
/// is `WHEEL_DELTA` a line, and the point is in the message because that is what the document
/// scrolls under (`WM_MOUSEWHEEL`'s `lParam` is screen, not client). `message` is either wheel: a
/// document may be scrolled either way, and both are named here rather than by the caller.
pub(crate) fn post_a_notch(window: HWND, message: u32, notches: i32, point: POINT) -> bool {
    let wparam = WPARAM(a_notch_wparam(notches));
    let lparam = LPARAM(a_notch_lparam(point));

    unsafe { PostMessageW(window, message, wparam, lparam).is_ok() }
}

/// The delta a notch message carries, in the high word: what Windows itself puts there.
fn a_notch_wparam(notches: i32) -> usize {
    ((((notches * WHEEL_DELTA as i32) as i16 as u16) as u32) << 16) as usize
}

/// The screen point a notch message carries, in the packed form Windows uses.
fn a_notch_lparam(point: POINT) -> isize {
    (point.y as u32 as isize) << 16 | (point.x as u16 as isize)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The road a pin's PDF is on may not become a white window: a viewer that never arrived is
    /// closed, and the pin's static raster is what stays on screen.
    #[test]
    fn a_viewer_that_never_arrives_is_closed_and_never_revealed() {
        let mut life = Life::Arriving;

        assert_eq!(life.navigated(false), Step::Close);
        assert!(life.is_closed());
        assert!(!life.is_up());

        let mut life = Life::Arriving;
        assert_eq!(life.navigated(false), Step::Close);
        assert!(life.is_closed());
    }

    /// It is revealed exactly once, and only by a document that arrived.
    #[test]
    fn a_viewer_is_revealed_once_by_a_document_that_arrived() {
        let mut life = Life::Arriving;

        assert_eq!(life.navigated(true), Step::Reveal);
        assert_eq!(life.navigated(true), Step::Nothing);
        assert!(life.is_up());
    }

    /// A pin the user leaves is suspended rather than closed, and a pin they come back to is
    /// resumed: the document is kept, which is the whole of what a pin is for.
    #[test]
    fn a_viewer_the_user_leaves_is_suspended_and_the_one_they_return_to_is_resumed() {
        let mut life = Life::Arriving;
        let _ = life.navigated(true);

        assert_eq!(life.settle(false), Step::Suspend);
        assert_eq!(life.settle(false), Step::Nothing);
        assert_eq!(life.settle(true), Step::Resume);
        assert_eq!(life.settle(true), Step::Nothing);
        assert!(life.is_up());
    }

    /// Nothing is suspended before it is up, and a viewer is closed exactly once however many roads
    /// reach the close.
    #[test]
    fn nothing_is_suspended_before_it_is_up_and_a_close_is_owed_once() {
        let mut life = Life::Arriving;
        assert_eq!(life.settle(false), Step::Nothing);

        let mut life = Life::Suspended;
        assert_eq!(life.closed(), Step::Close);
        assert_eq!(life.closed(), Step::Nothing);
        assert!(life.is_closed());
    }

    /// The notch a viewer is handed is the message Windows would have made: the delta in the high
    /// word of `wParam`, and the point in screen coordinates in `lParam`.
    #[test]
    fn a_notch_carries_its_delta_and_the_point_it_was_turned_at() {
        let point = POINT { x: 300, y: 200 };

        assert_eq!(
            (a_notch_wparam(3) as u32 >> 16) as i16,
            3 * WHEEL_DELTA as i16
        );
        assert_eq!(
            (a_notch_wparam(-1) as u32 >> 16) as i16,
            -(WHEEL_DELTA as i16)
        );

        let packed = a_notch_lparam(point) as u32;
        assert_eq!((packed & 0xffff) as i16, 300);
        assert_eq!((packed >> 16) as i16, 200);
    }
}

use std::path::{Path, PathBuf};
use std::sync::atomic::Ordering;
use std::sync::mpsc::{self, Receiver};
use std::time::{Duration, Instant};

use webview2_com::Microsoft::Web::WebView2::Win32::{
    ICoreWebView2, ICoreWebView2Controller, ICoreWebView2Environment, ICoreWebView2_3,
};
use webview2_com::{NavigationCompletedEventHandler, TrySuspendCompletedHandler};
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::{HWND, RECT};
use windows::Win32::System::WinRT::EventRegistrationToken;
use windows::Win32::UI::WindowsAndMessaging::{
    DestroyWindow, DispatchMessageW, GetMessageW, GetWindowLongPtrW, KillTimer, SetTimer,
    SetWindowLongPtrW, SetWindowPos, ShowWindow, TranslateMessage, GWL_EXSTYLE, HWND_TOPMOST, MSG,
    SWP_NOACTIVATE, SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE,
};

use super::api::{
    drop_placement, is_wanted, note_document_failed, page_runs, publish, publish_rect, wanted_area,
    Area, Placement, Wanted, HOST_HWND, LAST_TIMINGS, RUNNING_DOCUMENT,
};
use super::engine::trace;
use super::environment::{
    background_color, configure, create_controller, create_environment, create_host_window,
    ex_style_for, file_version, register_class, remaining, set_scripts, user_data_folder, wide,
    HOST_CREATION_TIMEOUT,
};
use super::pages::{font_page, frame_page, html_page};

use crate::app::engine_processes;
use crate::config::config::TransparentBackground;
use crate::formats::font_formats;
use crate::formats::text_formats;
use crate::readers::font_preview;

/// Everything one engine is: the window it draws in, the environment, and the
/// controller over it.
pub(super) struct Host {
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
    /// The width and height the controller was last told to lay the page out at, and the
    /// window's own last box, so that a box which merely moved is not also re-laid-out.
    ///
    /// This is what makes a carried window cost one `SetWindowPos` rather than a bounds change
    /// as well: a page is laid out at the size of the window it is drawn in, so `SetBounds` is
    /// owed only by a box that changed *size*, and a drag that changes only where the window
    /// is has no reason to ask the browser to lay the document out again (see `place`).
    last_area: Option<(i32, i32)>,
    /// The browser process this engine started, when it could be told which one it
    /// was. The runtime owns the browser, but the process is this app's own child —
    /// started by the loader in this process — and it is what is ended if it is
    /// somehow still there when the engine goes.
    browser_pid: u32,
    /// Whether the browser has stopped answering a navigation for longer than any document
    /// takes. Such an engine is let go of by the thread that holds it rather than handed
    /// the next document: what it would answer with is the file before this one, and what
    /// its thread would do is wait on a page that is never coming (see `NAVIGATION_TIMEOUT`).
    hung: bool,
    /// Whether the browser has been asked to suspend and is to be resumed before the next
    /// document is drawn. It is kept here rather than read back from the browser because
    /// both questions are about this host's own last word: a runtime older than
    /// suspension is a host that never suspends, and a resume of a WebView that is not
    /// suspended is harmless (see `suspend`, `wake`).
    ///
    /// It is also the record of the invariant `hide` keeps: after a hide returns, the
    /// browser has been asked to stop, whatever was on screen and whatever was. Nothing
    /// is read to decide whether the ask is owed, so a hide that has already been had
    /// does not ask again, and a browser that is stopped is a browser `wake` knows to put
    /// back to work before it navigates.
    suspended: bool,
}

impl Host {
    pub(super) fn create() -> Option<Self> {
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

        // One deadline for the whole attempt at having an engine: what is waited on is the
        // runtime answering at all, and one that has not answered by then is not going to
        // (see `HOST_CREATION_TIMEOUT`).
        let deadline = Instant::now() + HOST_CREATION_TIMEOUT;

        let started = Instant::now();
        let environment = create_environment(&folder, deadline);
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
        // caller falls back to this app's reader if even that comes to nothing. A
        // browser that is *silent* rather than refusing spends the deadline of the
        // attempt instead of being asked again on a fresh clock, which is why the
        // deadline is one of its own and not a per-call timeout.
        //
        // The first ask is made whatever the deadline says, since `create_controller` is
        // what spends it; the ones after it are put off a quarter of a second further
        // each time, and never past the end of the attempt, up to four asks.
        let mut controller = create_controller(environment.clone(), hwnd, deadline);
        let mut asked = 1;

        while controller.is_none() && asked < 4 && Instant::now() < deadline {
            std::thread::sleep(Duration::from_millis(250 * asked).min(remaining(deadline)));

            controller = create_controller(environment.clone(), hwnd, deadline);
            asked += 1;
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
            last_area: None,
            browser_pid,
            hung: false,
            suspended: false,
        };

        // The window a hit test compares against, for the preview loop and the Explorer
        // hook: it is the window a document is drawn in, so it is the one the pointer
        // touches when it is taken onto a document's preview.
        HOST_HWND.store(hwnd.0 as isize, Ordering::Release);

        Some(host)
    }

    /// Draw the document this want names in the box it asks for, or move the window to it
    /// where the engine already holds that document.
    ///
    /// A file the engine is not already holding, or one whose backdrop or face has changed,
    /// is navigated to *before* the window is put up: a window shown first would be the file
    /// before it, and what is on screen a moment ago belongs to another hover. What the
    /// navigation ends with is the whole of what happens next — the four ways it can end are
    /// [`Arrival`], and three of them put nothing up at all:
    ///
    /// - the page arrived, and the window is put up in the box the *newest* want asks for,
    ///   which is where a wait that followed the pointer ended up rather than where it
    ///   started;
    /// - a newer want took this one's place while the page was on its way: a file the
    ///   pointer has left is not shown, and the window comes down, because what the window
    ///   holds is the file before it and the loop is waiting for another one;
    /// - the page could not be written or the engine would not navigate: the engine is fine
    ///   and the hover waiting on this one has nothing left to wait for;
    /// - the browser never answered at all, which is the one that leaves the engine unfit to
    ///   be handed the next document (see `NAVIGATION_TIMEOUT`).
    pub(super) fn show(&mut self, wanted: &Wanted) {
        let Wanted {
            generation,
            path,
            background,
            area,
        } = wanted;
        let (background, area) = (*background, *area);

        // Which face of a collection a specimen of this file is drawn from, read here rather
        // than inside the page: it is part of what the engine is holding, so a setting changed
        // between two hovers of one file is a page to write and navigate to again. A file that
        // is not a font is drawn from no face at all.
        let face = if font_formats::is_font_file(path) {
            font_preview::configured_face()
        } else {
            0
        };

        // A browser that was told to stop when the last document was taken down is put back
        // to work here: before anything is navigated to, and before the controller is told
        // it is on screen, which is the order the runtime documents for a resume. Either of
        // those two would wake it anyway — a `Navigate` resumes a suspended WebView, and so
        // does making it visible — so nothing about the document turns on the order. The
        // resume is made explicitly rather than left to either of them so that the state this
        // app keeps is the state the browser is in, and the flag on `Host` cannot come to
        // disagree with it (see `suspend`, `wake`).
        self.wake();

        unsafe {
            // The background is a setting of the controller rather than of the page,
            // and it belongs to the interface that added it.
            if let Ok(controller) = self
                .controller
                .cast::<webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Controller2>()
            {
                let _ = controller.SetDefaultBackgroundColor(background_color(background));
            }

            // The box the page is rendered in, set before it is navigated to: a document
            // is drawn at this size, so what arrives is what the layout asked for rather
            // than something drawn small and resized after the fact. It is set again
            // below from the box the *newest* want asks for, which is where a wait that
            // followed a moving hand ended up.
            self.set_bounds(area);
        }

        if self.current.as_ref() != Some(&(path.clone(), background, face)) {
            // What the window is holding is *kept* while the next document is on its way:
            // a window taken down first would be a preview that vanishes and comes back,
            // which is worse than the file before it standing there for the few
            // milliseconds a navigation takes — and the file before it is one the pointer
            // has left only if the hook has said so, which is a hide of its own that
            // arrives here and takes the window down without waiting for this navigation.
            let started = Instant::now();
            let arrival = self.navigate(path, background, face, *generation);
            trace(&format!(
                "engine: navigate {} {arrival:?} in {} ms",
                path.display(),
                started.elapsed().as_millis()
            ));

            match arrival {
                Arrival::Arrived => {
                    if let Ok(mut timings) = LAST_TIMINGS.lock() {
                        timings.navigate_ms = started.elapsed().as_millis() as u64;
                    }

                    self.current = Some((path.clone(), background, face));
                }
                Arrival::Superseded => {
                    // Another want has taken this one's place, so what was navigated to is
                    // a file the pointer has left: nothing of it goes on screen. What the
                    // engine holds is nothing either — the page this navigation was made
                    // against has been written over with another document's — so the next
                    // want is navigated to rather than moved to.
                    self.current = None;
                    return;
                }
                Arrival::Failed => {
                    self.current = None;
                    note_document_failed();
                    return;
                }
                Arrival::TimedOut => {
                    // A browser that never answered the navigation is not a browser to
                    // hand the next document to: the hover waiting on this one is answered
                    // with nothing, and the engine is let go of and begun again (see
                    // `NAVIGATION_TIMEOUT` and `is_hung`).
                    self.hung = true;
                    self.current = None;
                    note_document_failed();
                    return;
                }
            }
        }

        // The box the *newest* want asks for, which is the one the wait has ended up in:
        // a wait that followed a moving pointer keeps following it, and a document that
        // takes a moment to be drawn lands where the hand is rather than where it was.
        let area = wanted_area(*generation).unwrap_or(area);

        // What the window is being asked to show decides what it may be activated for: a
        // page of HTML that runs is a program the user clicks into, and every other
        // document is a picture the caret is not to be taken out of for. The style is set
        // here rather than at the window's making, because the window is made once and
        // drawn in for every document this engine is given — and it is set before the
        // window is put up, so a click that lands on the first frame already has the
        // answer it is going to get.
        let runs = page_runs(path);
        RUNNING_DOCUMENT.store(runs, Ordering::Release);

        // The bounds this window was given are recorded, so that a box which only moves from
        // here is not also re-laid-out (see `place`).
        self.last_area = Some((area.width, area.height));

        unsafe {
            self.set_bounds(area);

            let current = GetWindowLongPtrW(self.hwnd, GWL_EXSTYLE);
            let style = ex_style_for(current, runs);
            if style != current {
                SetWindowLongPtrW(self.hwnd, GWL_EXSTYLE, style);
            }

            let _ = self.controller.SetIsVisible(true);

            self.put_window(area);
            let _ = ShowWindow(self.hwnd, SW_SHOWNOACTIVATE);
        }

        publish(Some((path.clone(), area)));
    }

    /// Put the window in `area`, topmost, and without ever activating it.
    ///
    /// The placement never activates, whatever the document is: a hover is not a click, and a
    /// preview that appeared over a window being named would put the caret somewhere else
    /// than where it was.
    pub(super) fn put_window(&self, area: Area) {
        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                HWND_TOPMOST,
                area.x,
                area.y,
                area.width,
                area.height,
                SWP_NOACTIVATE | SWP_SHOWWINDOW,
            );
        }
    }

    /// Tell the controller the box the page is rendered in, which is what the browser lays the
    /// page out at: a page is drawn at the size of the window it is drawn in, so this is owed
    /// by a box that changed *size* and by nothing else (see `place`).
    pub(super) fn set_bounds(&self, area: Area) {
        unsafe {
            let _ = self.controller.SetBounds(RECT {
                left: 0,
                top: 0,
                right: area.width,
                bottom: area.height,
            });
        }
    }

    /// Whether the document this host is holding is `path`, which is what a placement is
    /// answered for: a box published for a document the engine has since swapped away from
    /// belongs to no window it has, and moving whatever window it does have would put a
    /// document the pointer has left where the hand let go.
    pub(super) fn holds(&self, path: &Path) -> bool {
        self.current
            .as_ref()
            .is_some_and(|(held, _, _)| held == path)
    }

    /// Whether a placement under this backdrop is a page to write again rather than a window
    /// to move.
    ///
    /// Three of the four backdrops are partly the page's own — the checkerboard's squares, a
    /// specimen's ink — and the page is what the browser caches, so a backdrop the page was
    /// not written for is a different page and a different URL rather than a different colour
    /// on the window. A placement is published with whatever backdrop is on record at the time
    /// it is asked for, so a tray switch behind a pin arrives here as one placement that is a
    /// page to write, and as a move for every pointer move after it.
    pub(super) fn needs_page(&self, placement: &Placement) -> bool {
        self.current
            .as_ref()
            .is_some_and(|(_, held, _)| *held != placement.background)
    }

    /// Move the window of the document already on screen into a box that changed under it.
    ///
    /// This is deliberately the *whole* of what a moved box costs, and the difference from
    /// `show` is the point of it. A `show` is a document being put up: it wakes a suspended
    /// browser, sets the controller's colour, sets its bounds, tells it the window is on
    /// screen and shows the window. A drag of a pinned window asked for that on every pointer
    /// move, and the engine's thread takes one command per pass — so the queue of full shows
    /// grew faster than it could be drained and never was. A document that animates is what
    /// made that fatal rather than merely slow: its compositor is already working, so the
    /// thread fell so far behind that the drawing stopped following the window, and then
    /// stopped answering at all.
    ///
    /// So a box that merely moved costs one `SetWindowPos` and nothing else, which is all a
    /// window being carried has ever cost on this side (see `apply_pin_drag` in
    /// `preview_window`). `SetBounds` is owed only by a box that changed *size*, because a
    /// page is laid out at the size of the window it is drawn in; and nothing here wakes,
    /// re-shows or re-asserts visibility, all of which describe a document arriving rather
    /// than a window being carried. The window is kept above the pin's own, which is what a
    /// drag needs — the pin is raised on every move and would otherwise cover the drawing it
    /// is carrying.
    pub(super) fn place(&mut self, placement: &Placement) {
        let size = (placement.area.width, placement.area.height);
        let resized = self.last_area != Some(size);
        self.last_area = Some(size);

        if resized {
            self.set_bounds(placement.area);
        }

        self.put_window(placement.area);
        publish_rect(Some(placement.area));
    }

    /// Whether the browser has stopped answering, which is a host to be let go of rather
    /// than one to hand another document to.
    pub(super) fn is_hung(&self) -> bool {
        self.hung
    }

    /// Take the window down, and tell the browser that nothing is looking at it.
    ///
    /// `ShowWindow(SW_HIDE)` is only half of what a window leaving the screen means to a
    /// browser: a controller still told that it is on screen keeps the compositor making
    /// frames for it, and a document that runs goes on running at it for the rest of the
    /// session, so the browser is asked to suspend as well (see `suspend`). That is asked
    /// of every kind of document and not only of a page of HTML, because the frames are
    /// the browser's rather than the page's: an animated SVG is composited just as hard.
    ///
    /// There is deliberately no check of whether the window was ever put up, because a
    /// browser can be awake with its window hidden: a `show` that puts the browser back to
    /// work and then ends as a navigation that was superseded or failed returns before the
    /// controller is told it is on screen, so what is left is a browser drawing a window
    /// nobody is looking at (see `show`). A hide that took the absence of a visible window
    /// as an answer to the question of what the browser is doing would skip the one ask
    /// that stops it, and the engine is kept warm between documents, so nothing else would
    /// ever come along and do it. What not guarding costs is two Win32 calls and the one
    /// `Interface::cast` in `suspend`.
    pub(super) fn hide(&mut self) {
        unsafe {
            let _ = ShowWindow(self.hwnd, SW_HIDE);
            // `TrySuspend` states a precondition rather than a preference: the controller's
            // `IsVisible` must already be false when it is called, and otherwise the call
            // fails outright with `HRESULT_FROM_WIN32(ERROR_INVALID_STATE)` rather than
            // answering no. So the controller is hidden first, and this half of the hide is
            // never made conditional on a runtime that can be asked to suspend — it is the
            // half that stops the frames whether or not the ask is ever answered, including
            // on a runtime too old to be asked at all (see `suspend`).
            let _ = self.controller.SetIsVisible(false);
        }

        // These are what the preview loop reads as the preview still being there — through
        // `is_showing`, `showing_path` and `screen_rect` — so they are published before the
        // ask rather than after it: `suspend` waits, and can wait for up to
        // `SUSPEND_TIMEOUT`, and publishing first is what keeps a window in which the
        // preview has been taken down and the app is still reporting one under the pointer.
        publish(None);

        // The window is off screen, so the box a drag last asked for is a box nothing is in:
        // taking it back is what stops a placement published before this hide being carried out
        // against a window that is no longer showing, and it is the same reason the publishes
        // above happen first.
        drop_placement();

        self.suspend();
    }

    /// Ask the browser to stop working while nothing is looking at its window.
    ///
    /// What this is for is the engine being kept warm between documents (see the module
    /// docs): the browser is there to be pointed at the next one, and a browser left
    /// rendering a window that is off screen spends CPU and GPU on a picture no one is
    /// seeing. `TrySuspend` is the runtime's own answer to that — it halts rendering and
    /// throttles the page's script timers, which is what a page that runs needs, and it
    /// does it for an animated document as much as for a page.
    ///
    /// The ask is waited for, and waited for within `SUSPEND_TIMEOUT`, rather than left to
    /// complete whenever it completes (see `wait_for_suspend`). A runtime older than
    /// suspension is answered by leaving the controller hidden and doing nothing else —
    /// `SetIsVisible(false)` has already stopped the frames, and that half is the one that
    /// is not optional.
    ///
    /// It is called on every hide rather than only on a hide of something on screen, so the
    /// guard at the top of it is what keeps that cheap: a browser that has already been asked
    /// to stop is asked nothing further, however many hides arrive behind it and whether or
    /// not a window was ever put up in between.
    pub(super) fn suspend(&mut self) {
        if self.suspended {
            return;
        }

        let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() else {
            trace("host: this runtime does not suspend; the controller is hidden and no more");
            return;
        };

        trace("host: asked the browser to suspend");
        let (sender, receiver) = mpsc::channel::<bool>();

        unsafe {
            // The handler is let go of when this returns, and that is safe: the browser
            // holds a reference of its own to it, so an answer that arrives after the wait
            // has given up goes into a channel nobody is reading, and a send into a
            // receiver that has been dropped is the error this ignores.
            let handler =
                TrySuspendCompletedHandler::create(Box::new(move |_code, is_successful| {
                    let _ = sender.send(is_successful);
                    Ok(())
                }));

            if webview.TrySuspend(&handler).is_err() {
                trace("host: the browser would not be asked to suspend");
                return;
            }
        }

        // Whether the browser said yes, said no, or said nothing at all, the WebView is
        // recorded as suspended: what the flag is for is not asking again while the browser
        // stays where this left it, however many hides arrive behind it, and `wake` resumes
        // a browser that may not really have suspended — harmlessly, its result being
        // ignored. A refusal is therefore retried on the next hide, which is deliberate: a
        // browser that refused because of something transient — a script dialog left open,
        // say — is quite likely to say yes the next time it is asked (see `wake`).
        match wait_for_suspend(&receiver) {
            Some(true) => trace("host: the browser is suspended"),
            Some(false) => trace("host: the browser refused to suspend"),
            None => trace("host: the browser did not answer the suspend in time"),
        }

        self.suspended = true;
    }

    /// Put the browser back to work, which is what a page that runs needs before the next
    /// document is drawn: a resume rather than a browser begun again, and a page that picks
    /// up where it left off rather than an engine thrown away and started (see `suspend`).
    ///
    /// The order it is called in — resume first, and the controller told it is on screen
    /// after it — is the runtime's documented one, and nothing turns on it here: `Navigate`
    /// resumes a suspended WebView of its own accord, and so does making it visible. The
    /// resume is made explicitly rather than left to either of those so that the state this
    /// app keeps is the state the browser is in, and the flag on `Host` cannot come to
    /// disagree with it (see `show`).
    ///
    /// The flag is cleared before the call rather than after it, so a runtime with nothing to
    /// resume cannot leave the host believing it has a suspended WebView to wake before every
    /// document — and a resume of a WebView that was never suspended is harmless, its result
    /// being ignored either way.
    ///
    /// Nothing asserts that order: `suspend`, this and `wait_for_suspend` all need a live
    /// COM object to be asked against, so the invariant is on whoever changes this next rather
    /// than on the suite.
    pub(super) fn wake(&mut self) {
        if !self.suspended {
            return;
        }

        self.suspended = false;

        if let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() {
            unsafe {
                let _ = webview.Resume();
            }
        }
    }

    pub(super) fn close(&mut self) {
        self.hide();
        self.current = None;

        // A browser that stopped answering is one whose close may never return: the call is
        // a message to that browser, and a browser that has stopped taking messages is what
        // a hung host is. So the process is ended first — by the id it was recorded under,
        // which is the same verified end every other engine of this app gets — and the close
        // that follows is the close of something that is already gone rather than a wait on
        // it. A browser `hide` suspended is closed the same way as one that was not, and
        // nothing here waits for it to be woken first: a suspend is a browser that has
        // stopped drawing rather than one that has stopped answering, and either way the
        // close is this engine's last word — what did not take the browser down is ended
        // by `Drop`, which is the path a suspended browser is answered on as well.
        if self.hung && self.browser_pid != 0 {
            if engine_processes::is_running(self.browser_pid) {
                engine_processes::terminate_owned(self.browser_pid);
            }
            engine_processes::forget(self.browser_pid);
        }

        unsafe {
            let _ = self.controller.Close();
            let _ = DestroyWindow(self.hwnd);
        }

        // The window is gone, so nothing under the pointer is this engine's any more: a
        // hit test that still held the handle would be reading a window that was destroyed.
        HOST_HWND.store(0, Ordering::Release);

        // The environment goes with the host when it is dropped, and the browser
        // process it owns goes with the last controller over it.
        let _ = &self.environment;
    }

    /// Point the engine at `path` and wait for the page to arrive, pumping the thread's
    /// messages while it does.
    ///
    /// What it is pointed at is a page of this app's rather than the file itself, and which
    /// page is what the file is: a document's is the document as an image, which is what makes
    /// it the size of the window, a font's is the font in the page with its own lines under
    /// it — read at `face`, which is which of a collection's faces is written out; see
    /// `frame_page` and `font_page`. A page of HTML is the third of those kinds of document:
    /// its page is a frame of its own around the file, so the page is laid out by the browser
    /// rather than drawn by the app, and what it asks for is whatever box it is given (see
    /// `html_page`).
    ///
    /// The wait is asked under the generation this navigation was made for, and it ends on one
    /// of the four [`Arrival`]s: the page arriving, a newer want taking this one's place, the
    /// browser not answering within `NAVIGATION_TIMEOUT`, or the navigation not being made at
    /// all. What is *not* done is waiting without a bound — see `wait_for_navigation`.
    pub(super) fn navigate(
        &self,
        path: &Path,
        background: TransparentBackground,
        face: usize,
        generation: u64,
    ) -> Arrival {
        let version = file_version(path);
        let page = if font_formats::is_font_file(path) {
            font_page(path, version, background, face)
        } else if text_formats::is_html_extension(path) {
            html_page(path, version, background)
        } else {
            frame_page(path, version, background)
        };

        let Some((page, url)) = page else {
            return Arrival::Failed;
        };

        trace(&format!(
            "engine: page {} for {}",
            page.display(),
            path.display()
        ));
        let url = wide(&url);
        let (sender, receiver) = mpsc::channel();

        // Whether the browser runs what it is about to be given is settled here, before the
        // navigation rather than after it, because the document is read as it loads: a page
        // that draws itself would be a blank canvas by the time a setting asked for after
        // the fact could reach it (see `set_scripts`).
        set_scripts(&self.webview, page_runs(path));

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
                return Arrival::Failed;
            }

            let started = self.webview.Navigate(PCWSTR(url.as_ptr()));
            let arrival = if started.is_err() {
                Arrival::Failed
            } else {
                wait_for_navigation(&receiver, generation)
            };

            let _ = self.webview.remove_NavigationCompleted(token);

            arrival
        }
    }
}

/// How a navigation ended, as the four things `Host::show` does something about.
#[derive(Clone, Copy, Debug)]
pub(super) enum Arrival {
    /// The page arrived: the document is the one the engine is holding.
    Arrived,
    /// The page could not be written, or the engine would not navigate at all.
    Failed,
    /// A newer want took this one's place while the navigation ran: what is being navigated
    /// to is a file the pointer has left, so nothing of it goes up.
    Superseded,
    /// The browser stopped answering: the navigation was given longer than any document
    /// takes and nothing came back (see `NAVIGATION_TIMEOUT`).
    TimedOut,
}

/// How long a navigation may run before the browser is read as having stopped answering.
///
/// What the wait is for is `NavigationCompleted`, and that event is the one thing in this
/// module with no bound of its own: a navigation it never fires for leaves the engine's
/// thread waiting on it for the rest of the run, with every document after it queued behind
/// a wait that never ends, the idle timeout unable to fire because the thread that holds it
/// is the thread that is waiting, and the window standing there with the file before this
/// one on it. A page of this app's own is a file on the machine and a document a browser has
/// to lay out — milliseconds warm, a few hundred with the browser being begun — so what is
/// past this is not a document still being drawn: it is a browser that has stopped
/// answering, and it is let go of and begun again by the next document.
const NAVIGATION_TIMEOUT: Duration = Duration::from_secs(15);

/// How long the browser is given to answer a suspend before the engine reads it as having
/// stopped rather than as having answered.
///
/// What the wait is for is `TrySuspendCompleted`, which the runtime calls as soon as it has
/// stopped, so a quarter of a second is longer than a suspend takes on any machine and short
/// enough that a pointer moving on over a document never notices it. A browser that does not
/// answer is recorded as suspended anyway — the answer is what says it has stopped, not the
/// asking — and a resume of a WebView that is not suspended is harmless (see `pump_until`).
const SUSPEND_TIMEOUT: Duration = Duration::from_millis(250);

/// Why a wait ended without an answer: the bound ran out, or the message queue is gone.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Unanswered {
    /// The wait ran longer than the answer takes.
    TimedOut,
    /// There was no queue left to wait on.
    Gone,
}

/// Pump this thread's messages until `probe` answers, `timeout` runs out, or the queue is
/// gone — whichever comes first.
///
/// The messages have to be pumped rather than waited on: a controller is created on this
/// thread and stops drawing when the thread stops retrieving messages, and both of the events
/// this waits for arrive through that same queue — so the wait *is* a `GetMessage`, and an
/// answer is noticed the moment it lands, with no interval between and nothing polled.
///
/// The timer is armed on the thread rather than on a window, and under an id of its own so
/// that two waits cannot answer one another's timer: its whole job is to bring `GetMessage`
/// back so that the elapsed check is read again. The bound is therefore read on every pass,
/// which is what stops a browser that never answers from parking the engine's thread for the
/// rest of the run.
fn pump_until<T>(
    timeout: Duration,
    timer: usize,
    mut probe: impl FnMut() -> Option<T>,
) -> Result<T, Unanswered> {
    let started = Instant::now();
    let mut message = MSG::default();

    unsafe {
        let _ = SetTimer(HWND::default(), timer, timeout.as_millis() as u32, None);
    }

    let answer = loop {
        if let Some(answer) = probe() {
            break Ok(answer);
        }

        if started.elapsed() >= timeout {
            break Err(Unanswered::TimedOut);
        }

        let mut answered = false;
        unsafe {
            let retrieved = GetMessageW(&mut message, HWND::default(), 0, 0);
            if retrieved.0 > 0 {
                let _ = TranslateMessage(&message);
                DispatchMessageW(&message);
                answered = true;
            }
        }

        if !answered {
            // `GetMessage` answered -1, an error, or 0, a quit: there is no queue left to
            // wait on, and an answer that has not arrived by then is not coming.
            break Err(Unanswered::Gone);
        }
    };

    unsafe {
        let _ = KillTimer(HWND::default(), timer);
    }

    answer
}

/// Wait for the page to arrive, pumping the thread's messages while it does, and ending on
/// one of the three things that can end the wait.
///
/// The three are the page arriving, a want that is no longer this one — which
/// `wake_engine_thread` brings this thread back to look at — and a browser that has not
/// answered within `NAVIGATION_TIMEOUT` (see `pump_until`).
fn wait_for_navigation(receiver: &Receiver<()>, generation: u64) -> Arrival {
    const NAVIGATION_TIMER: usize = 1;

    match pump_until(NAVIGATION_TIMEOUT, NAVIGATION_TIMER, || {
        if receiver.try_recv().is_ok() {
            Some(Arrival::Arrived)
        } else if !is_wanted(generation) {
            Some(Arrival::Superseded)
        } else {
            None
        }
    }) {
        Ok(arrival) => arrival,
        Err(Unanswered::TimedOut) => Arrival::TimedOut,
        Err(Unanswered::Gone) => Arrival::Failed,
    }
}

/// Wait for the browser to answer a suspend, pumping the thread's messages while it does.
///
/// The wait is a `GetMessage` for the reason the one in `wait_for_navigation` is (see
/// `pump_until`). The answer is `Some` when the browser gave one, and `None` when it gave
/// nothing before `SUSPEND_TIMEOUT` or when the queue is gone, both of which the caller reads
/// as the same thing (see `suspend`).
fn wait_for_suspend(receiver: &Receiver<bool>) -> Option<bool> {
    const SUSPEND_TIMER: usize = 2;

    pump_until(SUSPEND_TIMEOUT, SUSPEND_TIMER, || receiver.try_recv().ok()).ok()
}

impl Drop for Host {
    fn drop(&mut self) {
        publish(None);
        HOST_HWND.store(0, Ordering::Release);

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

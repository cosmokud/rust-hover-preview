use crate::app::engine_processes;
use crate::config::config::{EngineIdle, DEFAULT_OFFICE_ENGINE_IDLE_SECS};
use crate::formats::office_formats::OfficeApp;
use crate::CONFIG;
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::time::{Duration, Instant};

use windows::core::VARIANT;
use windows::Win32::Foundation::{BOOL, HWND, LPARAM};
use windows::Win32::UI::Accessibility::{SetWinEventHook, UnhookWinEvent, HWINEVENTHOOK};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetClassNameW, GetWindowThreadProcessId, IsWindowVisible, ShowWindow,
    EVENT_OBJECT_SHOW, SW_HIDE, WINEVENT_OUTOFCONTEXT,
};

use super::com::Object;
use super::renderers::{alerts_off, render_excel, render_powerpoint, render_word, RenderTarget};
use super::worker::{pump_messages, wait_for_message};

/// How long the engines are kept after each family's last page, which is the
/// tray's `Engine → Microsoft Office TTL` setting.
///
/// Read rather than captured, so a change applies to engines that are already
/// warm: the worker wakes twice a second and asks this of every engine it holds,
/// so a setting lowered from an hour to nothing drops them within a tick of it
/// being made.
fn engine_idle() -> EngineIdle {
    CONFIG
        .lock()
        .map(|config| config.office_engine_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_OFFICE_ENGINE_IDLE_SECS))
}

/// Whether the engines are kept whatever the user is doing, which is the `Persistent`
/// toggle at the top of the same submenu.
///
/// Read rather than captured, and for the reason the idle time is: it is what decides
/// whether an engine is let go at all, so it is asked at the moment that is decided rather
/// than held from whenever the engine was started.
fn engine_persistent() -> bool {
    CONFIG
        .lock()
        .map(|config| config.office_engine_persistent)
        .unwrap_or(false)
}

/// The value that switches macro execution off entirely.
const MSO_AUTOMATION_SECURITY_FORCE_DISABLE: i32 = 3;

/// The window class Office shows its progress bars in — the wide, short
/// "Publishing…" window an export puts on the desktop. See
/// [`ProgressWindowHider`].
///
/// It is the class that is matched and not the title: a title is the machine's
/// language, and this is not.
const MSO_PROGRESS_WINDOW_CLASS: &str = "CMsoProgressBarWindow";
/// How long the progress hider waits on its thread's queue between looks at the
/// render it is watching. It is woken by the events it is hooked to, so this is
/// only how promptly it notices that the render is over.
const PROGRESS_WAIT_MS: u32 = 5;
/// How often the progress hider sweeps the process's windows as well as watching
/// for them being shown. The event is what makes the bar never part of a frame;
/// the sweep is what does not depend on an event arriving.
const PROGRESS_SWEEP_INTERVAL: Duration = Duration::from_millis(100);
/// The longest class name `GetClassNameW` is given room for.
const MAX_CLASS_NAME: usize = 256;

// ------------------------------------------------------------------ engines

/// An automation instance, and what it takes to leave it as it was found.
pub(super) struct Engine {
    app_kind: OfficeApp,
    app: Object,
    /// Whether the instance was already running for the user when it was reached.
    /// Such an instance is never hidden, never quit, and never ended.
    pub(super) attached: bool,
    /// The process this app started, when it started one: what may be ended if the
    /// engine stops answering, and what must never be ended otherwise.
    pub(super) owned_pid: u32,
    /// When this engine last drew a page. Each family's engine is on its own
    /// clock, so the family asked for most recently is the one that outlives the
    /// others rather than one idle time standing for all of them.
    pub(super) idle_since: Instant,
    /// Whether the automation settings a render needs are on the instance right
    /// now. They are taken for a render and put back when it is over, so this is
    /// `false` whenever the engine is sitting idle — which is what an engine kept
    /// for an hour, or for good, spends almost all of its life doing.
    settings_taken: bool,
    /// What the instance said about those settings before the render took them,
    /// read fresh for each render rather than once at creation.
    previous_alerts: Option<VARIANT>,
    previous_security: Option<VARIANT>,
    /// Whether this engine is being let go because it refused a page rather than
    /// because it has been idle. A refusal is what an instance that has lost its
    /// automation state answers with, and it is ended rather than quit — see
    /// [`Self::abandon`] and the drop below.
    abandoned: bool,
}

impl Engine {
    pub(super) fn create(app_kind: OfficeApp) -> Option<Self> {
        // What is running before the instance is created, so the process this app
        // starts can be told from one the user already had open.
        let before = engine_processes::processes_named(app_kind.image_name());

        let app = Object::create(app_kind.prog_id())?;

        // An instance this app created is hidden already; the user's is visible,
        // and hiding it would take their windows away.
        let attached = app
            .value("Visible")
            .and_then(|value| bool::try_from(&value).ok())
            .unwrap_or(false);

        let owned_pid = if attached {
            0
        } else {
            engine_processes::processes_named(app_kind.image_name())
                .into_iter()
                .find(|pid| !before.contains(pid))
                .unwrap_or(0)
        };
        if owned_pid != 0 {
            // From here on the process is this app's: it is put in the job, so that
            // whatever ends this app ends it, and written down, so that a run which
            // never gets to end it can be answered for by the next one.
            engine_processes::record(app_kind.image_name(), owned_pid);
        }

        if !attached {
            let _ = app.set("Visible", VARIANT::from(false));
        }

        Some(Self {
            app_kind,
            app,
            attached,
            owned_pid,
            idle_since: Instant::now(),
            settings_taken: false,
            previous_alerts: None,
            previous_security: None,
            abandoned: false,
        })
    }

    pub(super) fn render(
        &mut self,
        source: &Path,
        target: &RenderTarget,
        width: u32,
        height: u32,
        retrying: bool,
    ) -> bool {
        // What a render needs of the instance it runs in is taken for the render
        // and put back as soon as it is over. Nothing of the user's is therefore
        // held reconfigured while the engine sits warm between documents, which is
        // what an engine kept for an hour — or for good — would otherwise be doing
        // for almost the whole of its life.
        self.take_settings();

        // The one part of a render that is on screen is Office's own progress
        // window, and only for a process this app started — see
        // `ProgressWindowHider`.
        let _hider = (self.owned_pid != 0).then(|| ProgressWindowHider::new(self.owned_pid));

        let rendered = match self.app_kind {
            OfficeApp::Word => render_word(&self.app, source, target),
            OfficeApp::Excel => render_excel(&self.app, source, target, retrying),
            OfficeApp::PowerPoint => render_powerpoint(&self.app, source, target, width, height),
        };

        self.restore_settings();
        rendered
    }

    /// Suppress dialogs and switch macros off for the render that is about to run,
    /// keeping what the instance said so it can be put back.
    ///
    /// These are read here rather than once at creation because they are not the
    /// engine's to hold. An instance may be the user's own Word or Excel, which
    /// they may have changed themselves since the engine was created, and what is
    /// put back afterwards should be what the instance says now rather than what it
    /// said an hour ago.
    fn take_settings(&mut self) {
        if self.settings_taken {
            return;
        }

        self.previous_alerts = self.app.value("DisplayAlerts");
        self.previous_security = self.app.value("AutomationSecurity");

        let _ = self.app.set(
            "AutomationSecurity",
            VARIANT::from(MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
        );
        let _ = self.app.set("DisplayAlerts", alerts_off(self.app_kind));

        self.settings_taken = true;
    }

    /// Put back what [`Self::take_settings`] kept, leaving the process warm.
    ///
    /// This is the state an engine kept for a long time rests in: the settings
    /// belong to a render rather than to a process, and an instance that is doing
    /// nothing this app's business has no business holding another application's
    /// dialogs off or its macros switched off.
    fn restore_settings(&mut self) {
        if !self.settings_taken {
            return;
        }

        if let Some(alerts) = self.previous_alerts.take() {
            let _ = self.app.set("DisplayAlerts", alerts);
        }
        if let Some(security) = self.previous_security.take() {
            let _ = self.app.set("AutomationSecurity", security);
        }
        self.settings_taken = false;
    }

    /// Whether this engine has been kept as long as the settings say it may be.
    ///
    /// Two rules, one per `Persistent` setting at the top of the family's TTL submenu. An
    /// engine that is marked persistent is kept for its idle time whatever the user is
    /// doing, and one kept indefinitely never has; an engine that is not is let go once no
    /// Explorer window has been reachable for the AFK timer, with its idle time not
    /// consulted at all in that mode — what an idle time is for is the memory an engine
    /// holds while the user is elsewhere, and that is the question the AFK timer asks.
    ///
    /// What is let go of is the engine rather than what it drew: a rendered page is held in
    /// `RENDERS` and not in the automation object, so a preview on screen stays on screen
    /// while the engine that produced it goes.
    fn is_idle(&self) -> bool {
        if engine_persistent() {
            engine_idle().has_expired(self.idle_since.elapsed())
        } else {
            crate::app::afk::expired()
        }
    }

    /// Let go of an engine that has just refused a page.
    ///
    /// The process this app started is ended rather than asked to quit: a quit is
    /// another call into the instance that refused — and an unanswered one holds
    /// the worker for as long as it is waited on — while the instance is one
    /// nothing is owed to, since it holds no document of the user's. An instance
    /// the user is working in is never one this app created, and is not ended by
    /// this: `attached` is what stops it, the same as everywhere else.
    pub(super) fn abandon(mut self) {
        self.abandoned = true;
    }
}

impl Drop for Engine {
    fn drop(&mut self) {
        // An engine that refused a page and is this app's own process is ended
        // where it stands: asking it to quit costs seconds when it does not
        // answer — the quit is pumped for a while before the process is ended
        // anyway — and there is nothing to put back in a process that is about to
        // go. Everything else, the user's own Office included, is left as it was
        // found below.
        if self.abandoned && !self.attached && self.owned_pid != 0 {
            engine_processes::terminate_owned(self.owned_pid);
            return;
        }

        // A render puts these back itself, so this is for one that did not: a
        // panic inside it, or a render the abandonment of this worker left applied.
        self.restore_settings();

        if !self.attached {
            let _ = self.app.call("Quit", &[]);
        }

        if self.owned_pid != 0 {
            // An Office application finishes quitting while its automation client
            // is pumping: one that is told to quit and then abandoned can sit there
            // for good, which is what an Excel instance left behind every time
            // looks like. So the quit is given a moment of pumping, and a process
            // that is still there after it — one this app started, and verified as
            // still being that application — is ended rather than left running.
            if wait_for_exit(self.owned_pid) {
                engine_processes::forget(self.owned_pid);
            } else {
                // Still there after being asked to quit, so it is ended rather than
                // left running. What it costs if it outlasts that too is a record
                // kept until something sees it gone, which is what stops the next
                // engine of this family from being started beside it.
                engine_processes::terminate_owned(self.owned_pid);
            }
        }
    }
}

/// The engines the tier is holding, one per family.
///
/// A family's engine is only ever asked for its own documents — Word cannot open
/// a workbook and Excel cannot open a deck — so one slot for the whole tier meant
/// a folder holding one of each cost an Office start every time the pointer
/// crossed between them, with the family it had just left quit on the way out.
/// Each family keeps its own engine instead: three at the very most, since three
/// is how many families a preview can ask about, and only for the ones actually
/// being asked for pages.
///
/// Each is let go on its own clock rather than the tier having one idle time
/// between them, so the family last asked for is the one that outlives the rest.
pub(super) struct Engines {
    word: Option<Engine>,
    excel: Option<Engine>,
    powerpoint: Option<Engine>,
}

impl Engines {
    pub(super) fn new() -> Self {
        Self {
            word: None,
            excel: None,
            powerpoint: None,
        }
    }

    pub(super) fn slot(&mut self, app_kind: OfficeApp) -> &mut Option<Engine> {
        match app_kind {
            OfficeApp::Word => &mut self.word,
            OfficeApp::Excel => &mut self.excel,
            OfficeApp::PowerPoint => &mut self.powerpoint,
        }
    }

    pub(super) fn get(&self, app_kind: OfficeApp) -> Option<&Engine> {
        match app_kind {
            OfficeApp::Word => self.word.as_ref(),
            OfficeApp::Excel => self.excel.as_ref(),
            OfficeApp::PowerPoint => self.powerpoint.as_ref(),
        }
    }

    /// The three slots, for the sweeps that act on every engine rather than on
    /// one family's.
    fn all_mut(&mut self) -> [&mut Option<Engine>; 3] {
        [&mut self.word, &mut self.excel, &mut self.powerpoint]
    }

    pub(super) fn iter(&self) -> impl Iterator<Item = &Engine> {
        [&self.word, &self.excel, &self.powerpoint]
            .into_iter()
            .flatten()
    }

    /// Whether any engine has gone long enough without a page to be let go.
    pub(super) fn has_idle(&self) -> bool {
        self.iter().any(Engine::is_idle)
    }

    /// Let go of every engine that has been idle long enough. Each drop is a COM
    /// call of its own, which is why the caller does this inside the work marker.
    pub(super) fn drop_idle(&mut self) {
        for engine in self.all_mut() {
            if engine.as_ref().is_some_and(Engine::is_idle) {
                *engine = None;
            }
        }
    }

    pub(super) fn is_empty(&self) -> bool {
        self.iter().next().is_none()
    }

    /// Let go of every engine that is left, which is what the worker does as it
    /// ends: the thread owns them, so they go with it.
    pub(super) fn drop_all(&mut self) {
        for engine in self.all_mut() {
            *engine = None;
        }
    }
}

/// Give a process a moment to exit, pumping messages while it does — which is what
/// an Office application that has been asked to quit may be waiting for. `true`
/// when it is gone.
pub(super) fn wait_for_exit(pid: u32) -> bool {
    let deadline = Instant::now() + Duration::from_secs(2);

    while Instant::now() < deadline {
        if !engine_processes::is_running(pid) {
            return true;
        }
        wait_for_message(50);
    }

    !engine_processes::is_running(pid)
}

/// Take Office's own progress window off the screen for as long as a render lasts.
///
/// Exporting a page is work Office shows a progress bar for — the wide, short
/// "Publishing…" window in the middle of the desktop — and it shows it whether or
/// not the application it belongs to is visible, and through every setting this
/// engine holds: alerts, screen updating and the automation security mode were all
/// measured against it, and none of them keeps it off the screen. What the window
/// is, though, is a progress bar — nothing about the export depends on it being
/// seen — so it is hidden while the render runs. It is hidden only in a process
/// this app started: an instance the user is working in keeps every window it
/// owns, the same as it keeps its dialogs and its macros.
struct ProgressWindowHider {
    stop: Arc<AtomicBool>,
    thread: Option<std::thread::JoinHandle<()>>,
}

impl ProgressWindowHider {
    fn new(pid: u32) -> Self {
        let stop = Arc::new(AtomicBool::new(false));
        let watching = Arc::clone(&stop);
        let thread = std::thread::spawn(move || {
            unsafe {
                // The bar is taken down the moment it is put up, so it is never
                // part of a frame a display is showing: a sweep of the process's
                // windows cannot be that prompt — the bar is up for a moment
                // between the two, and that moment is what a hover used to show —
                // and the sweep below is only what does not depend on an event
                // arriving.
                let hook = SetWinEventHook(
                    EVENT_OBJECT_SHOW,
                    EVENT_OBJECT_SHOW,
                    None,
                    Some(progress_window_shown),
                    pid,
                    0,
                    WINEVENT_OUTOFCONTEXT,
                );

                let mut swept = Instant::now();
                while !watching.load(Ordering::Acquire) {
                    // The thread has to pump messages for the hook to be called at
                    // all, and what it pumps is nothing else: no window belongs to
                    // this thread.
                    pump_messages();
                    if swept.elapsed() >= PROGRESS_SWEEP_INTERVAL {
                        swept = Instant::now();
                        let _ = EnumWindows(Some(hide_progress_window), LPARAM(pid as isize));
                    }
                    wait_for_message(PROGRESS_WAIT_MS);
                }

                if !hook.0.is_null() {
                    let _ = UnhookWinEvent(hook);
                }
            }
        });

        Self {
            stop,
            thread: Some(thread),
        }
    }
}

impl Drop for ProgressWindowHider {
    fn drop(&mut self) {
        self.stop.store(true, Ordering::Release);
        if let Some(thread) = self.thread.take() {
            let _ = thread.join();
        }
    }
}

/// Hide the window an event is about, when it is Office's progress bar.
///
/// The hook this is called by was made for one process's windows, so nothing else
/// can reach it, and only a window being *shown* is reported.
unsafe extern "system" fn progress_window_shown(
    _hook: HWINEVENTHOOK,
    _event: u32,
    window: HWND,
    object: i32,
    child: i32,
    _thread: u32,
    _time: u32,
) {
    // `OBJID_WINDOW` and `CHILDID_SELF`: the bar itself, rather than one of the
    // accessibility objects inside it.
    if object != 0 || child != 0 {
        return;
    }

    hide_when_progress_bar(window);
}

/// Hide one window, when it is the progress bar of the process being watched.
/// Every other window is left where it is and the walk goes on: a window that is
/// not this one says nothing about the windows after it.
unsafe extern "system" fn hide_progress_window(window: HWND, pid: LPARAM) -> BOOL {
    let mut owner = 0u32;
    GetWindowThreadProcessId(window, Some(&mut owner));
    if owner == pid.0 as u32 {
        hide_when_progress_bar(window);
    }

    BOOL(1)
}

/// Take a window off the screen when it is Office's progress bar and it is on it.
unsafe fn hide_when_progress_bar(window: HWND) {
    if !IsWindowVisible(window).as_bool() {
        return;
    }

    let mut class = [0u16; MAX_CLASS_NAME];
    let length = GetClassNameW(window, &mut class);
    if length > 0
        && String::from_utf16_lossy(&class[..length as usize]) == MSO_PROGRESS_WINDOW_CLASS
    {
        let _ = ShowWindow(window, SW_HIDE);
    }
}

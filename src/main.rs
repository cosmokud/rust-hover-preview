#![windows_subsystem = "windows"]

mod app;
mod config;
mod engines;
mod formats;
mod readers;
mod shell;
mod text;
mod ui;

use once_cell::sync::Lazy;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::UI::HiDpi::{
    SetProcessDpiAwareness, SetProcessDpiAwarenessContext,
    DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, PROCESS_PER_MONITOR_DPI_AWARE,
};

// Global state
pub static RUNNING: AtomicBool = AtomicBool::new(true);
pub static CONFIG: Lazy<Mutex<config::config::AppConfig>> =
    Lazy::new(|| Mutex::new(config::config::AppConfig::load()));

/// Whether `RHP_STARTUP_TRACE` asks for what the start cost to be written out, and where.
///
/// A run without the variable set writes nothing anywhere, the way `RHP_HOOK_TRACE` writes no
/// probe counts: what the steps below are for is a start that is being measured, and where a
/// launch is spent is otherwise only reasoned about. The file is a handful of tiny appends,
/// written once per step and only by the run that asked for them.
static STARTUP_TRACE: Lazy<Option<PathBuf>> = Lazy::new(|| {
    std::env::var_os("RHP_STARTUP_TRACE")
        .map(|_| std::env::temp_dir().join("rhp-startup-trace.log"))
});

/// The start, as the steps it is made of.
///
/// Each step is noted with what it took and what the process has taken by then, and the two
/// are read together: a step that is slow by itself and a step that is only slow because
/// everything in front of it was are told apart by the pair, and the total is the number the
/// user feels — the moment the tray icon is there is a step of its own (see
/// `shell::tray::run_tray`).
///
/// It is carried from `main` rather than read from a clock of its own, so that what is
/// measured is the process's own start rather than the moment the first step happened to be
/// noticed. The housekeeping thread carries a copy, which is what gives that thread's steps
/// the same anchor as the ones on this path.
#[derive(Clone)]
pub(crate) struct StartupTrace {
    started: Instant,
    last: Instant,
}

impl StartupTrace {
    /// The trace to note a start against, taken as the first thing `main` does.
    pub(crate) fn new() -> Self {
        // The file holds this run's steps rather than every run's: what is read afterwards is
        // one start, and a file that was never cleared would make the second reading of one a
        // walk through the first.
        if let Some(path) = &*STARTUP_TRACE {
            let _ = fs::write(path, b"");
        }

        let now = Instant::now();
        Self {
            started: now,
            last: now,
        }
    }

    /// Note that the start has reached `step`.
    pub(crate) fn step(&mut self, step: &str) {
        let Some(path) = &*STARTUP_TRACE else {
            return;
        };

        let now = Instant::now();
        let line = format!(
            "{step} +{}ms (total {}ms)\n",
            now.duration_since(self.last).as_millis(),
            now.duration_since(self.started).as_millis(),
        );
        self.last = now;
        write_startup_step(path, &line);
    }
}

/// Append one line to the trace file, or nothing at all when it cannot be opened — a trace is
/// something a run is asked for, and a run whose trace cannot be written is a run that has
/// nothing to say about the machine's drive rather than a start that must fail.
fn write_startup_step(path: &Path, line: &str) {
    use std::io::Write;

    if let Ok(mut file) = fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = file.write_all(line.as_bytes());
    }
}

fn main() {
    // Taken first, so that what the steps below are measured from is the process's own start
    // rather than the first moment anything looked at the clock (see `StartupTrace`).
    let mut trace = StartupTrace::new();

    // Bail out before any hook, window, or thread is created when the app is
    // already running, so only the first instance stays in the tray.
    let _instance_guard = match app::single_instance::acquire() {
        Some(guard) => guard,
        None => return,
    };
    trace.step("instance guard");

    configure_dpi_awareness();
    trace.step("dpi awareness");

    sync_startup_setting();
    trace.step("startup setting");

    // The Office engines and browsers earlier runs started are ended here, before
    // this run can start one of its own: a run that was killed, or crashed, never
    // ends what it started, and a leftover engine is both a process nobody is using
    // and the thing a new engine would be a duplicate of. Only what a run that is
    // gone was holding is ended — a run that is still alive is another session, and
    // its engines are its own.
    app::engine_processes::reap_leftovers();
    trace.step("engine leftovers");

    // And the browsers of the runs that left no record: every profile folder under
    // the engine's own folder is named for the run that made it, so a folder whose
    // browser is still holding it names a browser to end. What the record above
    // catches for the runs that wrote one, this catches for the versions of this app
    // that did not.
    for pid in engines::webview_preview::stale_profile_pids() {
        app::engine_processes::end_browsers_started_by(pid);
    }
    trace.step("browser sweep");

    // Everything that is left is files rather than processes, and none of it is anything the
    // app has to be able to preview with: the folders earlier versions kept their pages in,
    // the log one of them wrote for every video hover, the installer the update check fetched
    // before it learned to fetch one only for a click, and the browser profiles of the runs
    // that are gone. What that costs grows with what those versions left in the temp folder —
    // thousands of files is a few thousand deletions — and what the user is waiting for is a
    // tray icon, which is a window and an icon and nothing else, so it is done on a thread of
    // its own rather than between the launch and the icon.
    //
    // The two sweeps above are deliberately not on that thread: they end processes, and what
    // they are for is not leaving a leftover engine beside a new one, which is a question
    // this run's first hover asks. What moves here can race the first hover's writes instead
    // — the page cache's folder and this app's temp folder are the tree being walked — and
    // what that can cost is a page rendered again: the cache's index stats a page before it
    // hands one back and drops an entry whose file is gone, so a swept page is a page that
    // is drawn a second time and never a preview answered from nothing.
    let mut housekeeping = trace.clone();
    std::thread::spawn(move || {
        // What an earlier version cached on disk goes, and whatever a render that was ended
        // mid-flight left in the temp folder goes with it — the pages this version keeps are not
        // that, and are left where they are so that a document drawn before this run is a read
        // rather than another render.
        engines::document_cache::discard_leftovers();

        // The log an earlier version appended a line to for every video hover is no
        // longer written; the file it left behind goes the same way, as its own user,
        // before a video preview could start adding to it again.
        let _ = fs::remove_file(std::env::temp_dir().join("rust-hover-preview-video.log"));

        // The update check keeps nothing on disk any more — the hour is counted in
        // memory, and the installer is fetched for the click that asks for it — so
        // what earlier versions left for it goes with the other files those versions
        // left behind.
        app::updates::discard_old_files();

        // The browser that draws a document keeps its state in a folder of its
        // own, one per run; what earlier runs left behind is cleared away here, before this
        // run has a folder for something to hold.
        engines::webview_preview::clear_stale_profiles();

        housekeeping.step("housekeeping");
    });
    trace.step("housekeeping started");

    // Initialize COM
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
    trace.step("com");

    // Start the preview window in a separate thread
    let preview_handle = std::thread::spawn(|| {
        ui::preview_window::run_preview_window();
    });

    // Watch config.ini changes off the hover hot path.
    let config_watch_handle = std::thread::spawn(|| {
        let config_path = config::config::AppConfig::config_path();
        let mut last_modified = config_path
            .as_ref()
            .and_then(|path| fs::metadata(path).ok())
            .and_then(|meta| meta.modified().ok());

        while RUNNING.load(Ordering::Acquire) {
            std::thread::sleep(Duration::from_millis(1000));

            let modified = config_path
                .as_ref()
                .and_then(|path| fs::metadata(path).ok())
                .and_then(|meta| meta.modified().ok());

            if modified != last_modified {
                last_modified = modified;
                if let Ok(mut config) = CONFIG.lock() {
                    config.reload_from_disk();
                }
            }
        }
    });

    // Start the explorer hook in a separate thread
    let hook_handle = std::thread::spawn(|| {
        shell::explorer_hook::run_explorer_hook();
    });

    // Watch system-wide wheel input so scrolling Explorer refreshes the preview
    // of the item that lands under the parked cursor.
    let wheel_handle = shell::wheel_input::spawn_wheel_watcher();
    trace.step("threads");

    // The check for a newer release is asked for here, as the run starts: the row
    // it may put in the menu should be there the first time the menu is opened
    // rather than only after an opening of its own asked for a check. It runs on a
    // thread of its own, so nothing here waits on GitHub.
    app::updates::request_check();
    trace.step("update check");

    // Run the system tray (this blocks until exit), noting the two steps between the launch
    // and the icon the user is waiting for: the window, and the icon that goes in it.
    shell::tray::run_tray(&mut trace);

    // Signal other threads to stop
    RUNNING.store(false, Ordering::SeqCst);
    shell::wheel_input::request_stop();

    // The engine thread is joined only when it is idle: a COM call into Office
    // cannot be cancelled, and the app's exit must not wait on one.
    engines::office_render::shutdown();

    // The browser engine is this app's own process tree rather than an application a
    // user may also be working in, so it is ended here: nothing of it should outlive
    // the app.
    engines::webview_preview::shutdown();

    // Wait for the threads that end by themselves. The hook is not one of them: every probe
    // it makes crosses into Explorer, and a shell that has stopped answering holds one of
    // those crossings for as long as it likes — which is the wait the Office engine's own
    // thread is not given, for the same reason. It goes with the process.
    let _ = preview_handle.join();
    drop(hook_handle);
    let _ = wheel_handle.join();
    let _ = config_watch_handle.join();

    // Cleanup COM
    unsafe {
        CoUninitialize();
    }
}

fn configure_dpi_awareness() {
    unsafe {
        // Prefer per-monitor v2 to avoid DPI scaling artifacts on layered windows.
        if SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2).is_err() {
            let _ = SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
        }
    }
}

fn sync_startup_setting() {
    let should_enable_startup = CONFIG
        .lock()
        .map(|config| config.is_first_run && config.run_at_startup)
        .unwrap_or(false);

    if should_enable_startup {
        app::startup::enable_startup();
    }

    // And the configuration is brought into step with what the registry holds, which is what
    // actually starts this app: a `run_at_startup` that says yes while the entry is gone is a
    // value that lies about the machine, and the tray's own toggle is where the choice is
    // made after the first run anyway. Nothing else reads this value but the tray.
    let registered = app::startup::is_startup_enabled();

    if let Ok(mut config) = CONFIG.lock() {
        if config.run_at_startup != registered {
            config.run_at_startup = registered;
            config.save();
        }
    }
}

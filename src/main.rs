#![windows_subsystem = "windows"]

mod app;
mod config;
mod engines;
mod formats;
mod paths;
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
    app::local_env::load();

    // Taken first, so that what the steps below are measured from is the process's own start
    // rather than the first moment anything looked at the clock (see `StartupTrace`).
    let mut trace = StartupTrace::new();

    // **Which of the two runs this is, before anything is claimed or swept.** A detached window is
    // a process of this same executable begun by the tray app's own hand (see `app::detach`), and
    // the three things a second process must not do are all below: claim the single-instance mutex,
    // reap what earlier runs left behind, and put an icon in the tray.
    match app::detach::startup_mode() {
        app::detach::Startup::Tray => {}
        // A hand-off this app cannot answer starts nothing at all rather than a second tray icon.
        app::detach::Startup::Refused => return,
        app::detach::Startup::Detached(request) => {
            app::detach::begin(request);
            run_a_detached_window(&mut trace);
            return;
        }
        app::detach::Startup::Viewer(tray) => {
            run_a_viewer(&mut trace, tray);
            return;
        }
    }

    // Bail out before any hook, window, or thread is created when the app is
    // already running, so only the first instance stays in the tray.
    let _instance_guard = match app::single_instance::acquire() {
        Some(guard) => guard,
        None => return,
    };
    trace.step("instance guard");

    // The facts this run publishes are a section of its own from here on (see `app::seam::the_facts`),
    // opened before the threads that publish are begun so that no thread starts on this process's own
    // copy and finds the facts moved under it. A run that could not open one has nothing a second
    // process could read, and draws its own previews below.
    let the_seam_is_open = app::seam::open_the_seam_section();
    trace.step("seam section");

    configure_dpi_awareness();
    trace.step("dpi awareness");

    sync_startup_setting(&mut trace);
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
    // app has to be able to preview with: what a render that was ended mid-flight left in the
    // temp folder, and the browser profiles of the runs that are gone. What that costs grows
    // with what those runs left behind — thousands of files is a few thousand deletions — and
    // what the user is waiting for is a tray icon, which is a window and an icon and nothing
    // else, so it is done on a thread of its own rather than between the launch and the icon.
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
        // Whatever a render that was ended mid-flight left in the temp folder goes — the pages
        // this version keeps are not that, and are left where they are so that a document drawn
        // before this run is a read rather than another render.
        engines::document_cache::discard_leftovers();

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

    // **Where a preview is drawn, which is the one thing this run decides about itself.** A viewer of its
    // own where one could be begun — the frame of every preview from here on, in a process that opens the
    // file on its own thread and owns the media it opened — and this process where it could not, which is
    // exactly the run this app was before the split: one process, its own channel, its own preview thread.
    // The road is taken once and never again: there is no switching back, and no second attempt at a viewer
    // a run has given up on (see `app::viewer`).
    let two_processes = the_seam_is_open && app::viewer::begin_the_viewer();
    app::viewer::note_the_road(two_processes);
    let mut preview_handle = None;
    if two_processes {
        app::seam::carry_messages_in_the_rings();
    } else {
        // The fallback road: this process renders, which is exactly the run this app was before the
        // split (see `app::viewer`).
        preview_handle = Some(std::thread::spawn(|| {
            ui::preview_window::run_preview_window();
        }));
    }

    // Watch config.ini changes off the hover hot path.
    let config_watch_handle = std::thread::spawn(watch_the_configuration);

    // Start the explorer hook in a separate thread
    let hook_handle = std::thread::spawn(|| {
        shell::explorer_hook::run_explorer_hook();
    });

    // Watch system-wide wheel input so scrolling Explorer refreshes the preview
    // of the item that lands under the parked cursor. **The viewer's hook where there is a viewer**:
    // what a notch does — a pin's zoom, a pin's volume, a model's band, a text preview's scroll — is
    // all of it drawn there, and a hook is a process's own (see `shell::wheel_input`).
    let mut wheel_handle = None;
    if !two_processes {
        wheel_handle = Some(shell::wheel_input::spawn_wheel_watcher());
    }

    // And watch the pin key, which is the one key this app acts on itself: a preview
    // on screen when it is pressed becomes a window of its own. The watch is bound
    // before the thread is started, so the first press is already the configured one.
    shell::key_input::refresh();
    let key_handle = shell::key_input::spawn_key_watcher();
    trace.step("threads");

    // The check for a newer release is asked for here, as the run starts: the row
    // it may put in the menu should be there the first time the menu is opened
    // rather than only after an opening of its own asked for a check. It runs on a
    // thread of its own, so nothing here waits on GitHub — and nothing is asked
    // at all where the checks are switched off (see `check_for_updates`). What the
    // tray loop's own timer and every menu opening ask for are the rest of the run's
    // checks, one period apart; what a check does with its answer is no longer a
    // start-only thing either — a silent install happens wherever the app is idle, and
    // the one that found an update over a pin waits for it (see `the_app_is_idle`).
    app::updates::request_check();
    trace.step("update check");

    // Run the system tray (this blocks until exit), noting the two steps between the launch
    // and the icon the user is waiting for: the window, and the icon that goes in it.
    shell::tray::run_tray(&mut trace);

    // Signal other threads to stop
    RUNNING.store(false, Ordering::SeqCst);
    shell::wheel_input::request_stop();
    shell::key_input::request_stop();

    // The engine thread is joined only when it is idle: a COM call into Office
    // cannot be cancelled, and the app's exit must not wait on one.
    engines::office_render::shutdown();

    // The browser engine is this app's own process tree rather than an application a
    // user may also be working in, so it is ended here: nothing of it should outlive
    // the app.
    engines::webview_preview::shutdown();

    // The mpv a hover's film was loaded into is kept between hovers and is this app's own
    // process: a run that ends leaves nothing of it behind (see `mpv_resident`).
    ui::preview_window::shut_down_the_resident();

    // Wait for the threads that end by themselves. The hook is not one of them: every probe
    // it makes crosses into Explorer, and a shell that has stopped answering holds one of
    // those crossings for as long as it likes — which is the wait the Office engine's own
    // thread is not given, for the same reason. It goes with the process.
    if let Some(preview_handle) = preview_handle {
        let _ = preview_handle.join();
    }
    drop(hook_handle);
    if let Some(wheel_handle) = wheel_handle {
        let _ = wheel_handle.join();
    }
    let _ = key_handle.join();
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

fn sync_startup_setting(trace: &mut StartupTrace) {
    let should_enable_startup = CONFIG
        .lock()
        .map(|config| config.is_first_run && config.run_at_startup)
        .unwrap_or(false);

    if should_enable_startup {
        app::startup::enable_startup();
    }

    // And an entry that is there but names another copy of this app is put right, before
    // anything reads it as settled: a portable copy run once from a folder of its own, an
    // older version still installed where it put itself, an app moved since — each of them
    // writes its own path into the one entry there is, and that path is what the next logon
    // starts. The entry decides only *whether* this app starts, which is left to the user;
    // what it points at is this app's own business, and this is the copy running now.
    if app::startup::repair_startup_entry() {
        trace.step("startup entry repaired");
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

/// The run a detached window has: one preview thread, one configuration watch, and nothing that
/// watches a listing.
///
/// It is deliberately not the tray app's start with pieces skipped at their own doors: the pieces
/// that are missing here are the *whole* of what makes a hover possible — the Explorer hook, the
/// two key watches, the sweep of what earlier runs left behind, the update check and the tray — and
/// a detached window is a window no listing can reach (see `app::detach`).
fn run_a_detached_window(trace: &mut StartupTrace) {
    configure_dpi_awareness();
    trace.step("dpi awareness");

    // The apartment a page's own measurements are taken in before any file is read: a page is laid
    // out by this thread, and a PDF page is rendered through Windows.Data.Pdf and a picture of a
    // format this app has no decoder for is decoded by the codec Windows has (see
    // `run_preview_window`).
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
    trace.step("com");

    let preview_handle = std::thread::spawn(|| {
        ui::preview_window::run_preview_window();
    });

    // What the tray would have said out loud in its own process is said here off the same file (see
    // `watch_the_configuration`).
    let config_watch_handle = std::thread::spawn(watch_the_configuration);

    // **The one watch a detached window keeps.** The wheel over a pinned picture, a pinned page and
    // a sound's knob is the low-level hook's rather than the window's — a pin is nobody's keyboard
    // until it is pressed, so Windows hands its notches to Explorer — and that hook is a process's
    // own (see `shell::wheel_input`). The Explorer's hook and the two key watches are exactly what
    // this window does not have.
    // The pin key's chords, bound in this process as well: only the mouse half of the key is counted
    // here, and a chord table nothing bound counts no press at all (see `shell::key_input::refresh`).
    shell::key_input::refresh();
    let wheel_handle = shell::wheel_input::spawn_wheel_watcher();
    trace.step("threads");

    // And what ends it: the tray app going — which the job every hand-off joins answers on its own —
    // and the process that began it, waited on a second time for the machine where a job could not be
    // joined (see `app::detach::watch_the_parent`).
    app::detach::watch_the_parent();

    // The window draws itself on its own thread, and it ends itself the moment its pin does (see
    // `ui::preview_window::detach_open`): what is left here is only the waiting.
    while RUNNING.load(Ordering::Acquire) {
        std::thread::sleep(Duration::from_millis(50));
    }

    shell::wheel_input::request_stop();

    // The engines are this process's own and are ended with it the way the tray app's are ended
    // with it: nothing of this app's should outlive the window it was begun for.
    engines::office_render::shutdown();
    engines::webview_preview::shutdown();
    ui::preview_window::shut_down_the_resident();

    let _ = preview_handle.join();
    let _ = wheel_handle.join();
    let _ = config_watch_handle.join();

    unsafe {
        CoUninitialize();
    }
}

/// The run a viewer has: the preview loop, the one low-level hook a preview keeps, and nothing that
/// watches a listing.
///
/// **This is the process every preview of a run is drawn in.** It renders: the layered window, the pins,
/// the players, the engines, the pages. It does not sense: the Explorer hook, the two key watches, the
/// tray icon and the sweeps of what earlier runs left behind are the tray app's, and none of them is begun
/// here. It was begun by that tray app and is told its pid, which is what every name the seam is known by
/// is made of — the section it maps rather than makes, and the two events its messages are announced on
/// (see `app::seam` and `app::viewer`).
fn run_a_viewer(trace: &mut StartupTrace, tray: u32) {
    // The facts are the tray's from here on. A viewer that cannot find the section has nothing to be told
    // and nothing to tell, and stops rather than standing as a process that draws nothing: the tray's next
    // hover begins another one (see `app::viewer`).
    if !app::seam::join_the_seam_section(tray) {
        return;
    }
    // And the rings are the road from here on: every hover, every pin and every notice crosses there
    // rather than through a channel of this process's own.
    app::seam::carry_messages_in_the_rings();
    trace.step("seam section");

    configure_dpi_awareness();
    trace.step("dpi awareness");

    // The apartment a page's own measurements are taken in before any file is read: a page is laid out by
    // this thread, and a PDF page is rendered through Windows.Data.Pdf and a picture of a format this app
    // has no decoder for is decoded by the codec Windows has (see `run_preview_window`).
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
    trace.step("com");

    let preview_handle = std::thread::spawn(|| {
        ui::preview_window::run_preview_window();
    });

    // The configuration is watched here as well, and the two rows that are *spoken* to a window already
    // up are said by the tray's own process rather than by this one — so which of them this watcher says
    // to itself is asked of the detach flag on every reload, because this window can become one of its
    // own while it stands (see `watch_the_configuration` and `pin_detach::detach_the_pinned_window`).
    let config_watch_handle = std::thread::spawn(watch_the_configuration);

    // **The one watch a viewer keeps.** The wheel over a pinned picture, a pinned page, a model's band and
    // a page of text is the low-level hook's rather than the window's — a pin is nobody's keyboard until it
    // is pressed, so Windows hands its notches to Explorer — and that hook is a process's own (see
    // `shell::wheel_input`). The Explorer's hook and the two key watches are exactly what this process does
    // not have.
    // **The pin key is a number the hook reads, and this process's hook reads its own copy.** The mouse
    // half of the pin key (`mouse3` by default) never reaches a keyboard hook, so it is counted here,
    // where the mouse watcher lives — and a viewer that never bound the chords counts nothing at all:
    // the middle press is then a press no window of this app hears, while the `Space` half goes on
    // working because the tray's own keyboard hook was refreshed at its start (see
    // `shell::key_input::refresh`).
    shell::key_input::refresh();
    if std::env::var_os("RHP_SEAM_TRACE").is_some() {
        eprintln!(
            "seam: the viewer's pin key is bound {}",
            shell::key_input::pin_key_is_bound() as u8
        );
    }
    let wheel_handle = shell::wheel_input::spawn_wheel_watcher();
    trace.step("threads");

    // And what ends it: the tray app going, which the job it began this process in answers on its own, and
    // the process that began it, waited on for the machine where a job could not be joined (see
    // `app::detach::watch_the_process`).
    app::detach::watch_the_process(tray);

    while RUNNING.load(Ordering::Acquire) {
        std::thread::sleep(Duration::from_millis(50));
    }

    shell::wheel_input::request_stop();

    // The engines are this process's own and are ended with it the way the tray app's are ended with it:
    // nothing of this app's should outlive the window it was begun for.
    engines::office_render::shutdown();
    engines::webview_preview::shutdown();
    ui::preview_window::shut_down_the_resident();

    let _ = preview_handle.join();
    let _ = wheel_handle.join();
    let _ = config_watch_handle.join();

    unsafe {
        CoUninitialize();
    }
}

/// Watch `config.ini` for changes, off both runs' hot paths.
///
/// **Whether this process is a window of its own is asked of the flag on every reload rather than taken
/// once from the launch**, because a pin can now become one inside a viewer that was begun as the run's
/// own preview thread: that window keeps this watcher for the life of the process, and an answer copied
/// from the launch is the wrong one from the moment it flips (see `app::detach::is_detached` and
/// `pin_detach::detach_the_pinned_window`).
///
/// A detached window is a process the tray's own rows are not spoken to. What the tray writes down reaches
/// it through this watch the way it reaches a pin; the two rows that are *said* to a window already up —
/// the text size and the band — are said to the tray's own preview thread and to no other, so this window
/// says them to itself here, off the same reload. A row that is a window's own question at its take-up —
/// the engine a film is played by, the scale its box was measured at — is not one of them and reaches a
/// detached window at its next start, which is what a pin does.
fn watch_the_configuration() {
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

        if modified == last_modified {
            continue;
        }
        last_modified = modified;

        // The two values a detached window has to be *told* about rather than only read, taken
        // before the reload: what changed is the difference between the two readings.
        let before = CONFIG
            .lock()
            .map(|config| (config.text_font_scale_percent, config.pin_always_on_top))
            .unwrap_or_default();

        // A level edited on disk reaches a live pin from here, or opening the menu is the first
        // place its change is noticed (see `show_context_menu`). Unchanged halves — and a level this
        // process wrote itself — publish nothing.
        let changed = CONFIG
            .lock()
            .map(|mut config| config.reload_from_disk_keeping_our_own_levels())
            .unwrap_or((false, false));

        // Lock released first: the message a publish sends locks `CONFIG` again.
        if changed.0 {
            ui::preview_window::publish_tray_volume(false);
        }
        if changed.1 {
            ui::preview_window::publish_tray_volume(true);
        }

        if crate::app::detach::is_detached() {
            let after = CONFIG
                .lock()
                .map(|config| (config.text_font_scale_percent, config.pin_always_on_top))
                .unwrap_or_default();
            if after.0 != before.0 {
                ui::preview_window::publish_text_font_size(after.0);
            }
            // The band, and only while pinning is switched on: the message is the same one the tray's
            // own row sends, and its answer for a window with **Pin Mode → Enable** off is to end it —
            // which is right for a pin the tray owns and wrong for a window of its own, because that
            // switch is about taking pins from the listing and this window was taken long ago (see
            // `PreviewMessage::PinChanged`).
            if after.1 != before.1 && crate::ui::preview_window::pin_mode_is_enabled() {
                ui::preview_window::refresh_pin();
            }
        }

        // The pin key is a number the keyboard hook reads rather than the configuration, so a
        // spelling that changed on disk is handed to it here — off the hook's own path, which may
        // not take a lock.
        shell::key_input::refresh();
    }
}

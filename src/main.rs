#![windows_subsystem = "windows"]

mod archive_formats;
mod archive_listing;
mod archive_preview;
mod cloud_files;
mod config;
mod explorer_hook;
mod image_formats;
mod office_formats;
mod office_preview;
mod office_render;
mod pdf_preview;
mod preview_window;
mod single_instance;
mod startup;
mod text_formats;
mod text_paint;
mod text_preview;
mod text_theme;
mod theme_files;
mod tray;
mod video_formats;
mod wheel_input;

use once_cell::sync::Lazy;
use std::fs;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::Duration;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::UI::HiDpi::{
    SetProcessDpiAwareness, SetProcessDpiAwarenessContext,
    DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, PROCESS_PER_MONITOR_DPI_AWARE,
};

// Global state
pub static RUNNING: AtomicBool = AtomicBool::new(true);
pub static CONFIG: Lazy<Mutex<config::AppConfig>> =
    Lazy::new(|| Mutex::new(config::AppConfig::load()));

fn main() {
    // Bail out before any hook, window, or thread is created when the app is
    // already running, so only the first instance stays in the tray.
    let _instance_guard = match single_instance::acquire() {
        Some(guard) => guard,
        None => return,
    };

    configure_dpi_awareness();
    sync_startup_setting();

    // A page is held in memory and nowhere else, so whatever an earlier version left
    // in the cache folder — and whatever a render that was ended mid-flight left in
    // the temp folder — is dropped before anything starts writing there again.
    office_render::discard_old_disk_cache();

    // Initialize COM
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    // Start the preview window in a separate thread
    let preview_handle = std::thread::spawn(|| {
        preview_window::run_preview_window();
    });

    // Watch config.ini changes off the hover hot path.
    let config_watch_handle = std::thread::spawn(|| {
        let config_path = config::AppConfig::config_path();
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
        explorer_hook::run_explorer_hook();
    });

    // Watch system-wide wheel input so scrolling Explorer refreshes the preview
    // of the item that lands under the parked cursor.
    let wheel_handle = wheel_input::spawn_wheel_watcher();

    // Run the system tray (this blocks until exit)
    tray::run_tray();

    // Signal other threads to stop
    RUNNING.store(false, Ordering::SeqCst);
    wheel_input::request_stop();

    // The engine thread is joined only when it is idle: a COM call into Office
    // cannot be cancelled, and the app's exit must not wait on one.
    office_render::shutdown();

    // Wait for threads to finish (with timeout)
    let _ = preview_handle.join();
    let _ = hook_handle.join();
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
        startup::enable_startup();
    }
}

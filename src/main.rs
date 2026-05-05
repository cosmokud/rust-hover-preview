#![windows_subsystem = "windows"]

mod config;
mod explorer_hook;
mod preview_window;
mod startup;
mod tray;

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
    configure_dpi_awareness();
    sync_startup_setting();

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

    // Run the system tray (this blocks until exit)
    tray::run_tray();

    // Signal other threads to stop
    RUNNING.store(false, Ordering::SeqCst);

    // Wait for threads to finish (with timeout)
    let _ = preview_handle.join();
    let _ = hook_handle.join();
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

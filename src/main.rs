#![windows_subsystem = "windows"]

mod config;
mod explorer_hook;
mod preview_window;
mod startup;
mod tray;

use once_cell::sync::Lazy;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

// Global state
pub static RUNNING: AtomicBool = AtomicBool::new(true);
pub static CONFIG: Lazy<Mutex<config::AppConfig>> =
    Lazy::new(|| Mutex::new(config::AppConfig::load()));

fn main() {
    // Initialize COM
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    // Start the preview window in a separate thread
    let preview_handle = std::thread::spawn(|| {
        preview_window::run_preview_window();
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

    // Cleanup COM
    unsafe {
        CoUninitialize();
    }
}

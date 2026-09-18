//! Playing a document in the browser engine that is already on the machine.
//!
//! usvg drops animation — its own documentation says "no events and no animations" —
//! so `svg_animation` plays what it can read, and this is what plays the rest: the
//! WebView2 runtime that Windows 11 ships with, which is Chromium, and which is the
//! only complete implementation of SMIL and CSS animation that exists on this machine
//! without installing anything. It is used for documents that *move* only: a still
//! document is drawn by `svg_preview`, which costs no browser at all.
//!
//! What lives here is the engine, not the plumbing: creating the runtime, holding the
//! controller it draws through, pointing it at a file and taking it down again. The
//! window it draws into is its own — a layered window cannot host a child, since what
//! `UpdateLayeredWindow` is given is a surface and not a window tree — so this is the
//! same shape as the video path, where the player's own window is the preview.
//!
//! A document's script is not run: the setting this hands the engine is the same
//! promise the rest of the app makes, that a document is drawn and not executed.

// Nothing calls the engine yet: the dispatcher that sends a document that moves to it,
// the idle timer that lets it go, and the settings that turn it on are the next piece,
// and every part of this is exercised by the probe in its tests until then.
#![allow(dead_code)]

use std::path::{Path, PathBuf};
use std::sync::mpsc;
use std::time::{Duration, Instant};

use webview2_com::Microsoft::Web::WebView2::Win32::{
    CreateCoreWebView2EnvironmentWithOptions, GetAvailableCoreWebView2BrowserVersionString,
    ICoreWebView2Controller,
};
use webview2_com::{
    CreateCoreWebView2ControllerCompletedHandler, CreateCoreWebView2EnvironmentCompletedHandler,
    NavigationCompletedEventHandler,
};
use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{E_POINTER, HWND, RECT};
use windows::Win32::System::Com::{CoInitializeEx, CoTaskMemFree, COINIT_APARTMENTTHREADED};
use windows::Win32::System::WinRT::EventRegistrationToken;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, SetWindowPos, ShowWindow, HWND_TOPMOST, SWP_NOACTIVATE,
    SWP_NOSIZE, SWP_SHOWWINDOW, SW_SHOWNOACTIVATE, WS_POPUP,
};

/// What one run of the engine cost, which is the number the design turns on.
pub struct Report {
    /// The version of the runtime that answered, or nothing when none did.
    pub runtime_version: Option<String>,
    /// Where the engine keeps its own state, which is a folder this app owns.
    pub user_data_folder: PathBuf,
    /// Beginning the engine: the environment it needs and the controller it draws
    /// through, in milliseconds.
    pub environment_ms: u128,
    pub controller_ms: u128,
    /// Each document that was pointed at, and how long it took to arrive.
    pub navigations: Vec<(PathBuf, u128)>,
}

/// The folder the engine keeps its own state in. It is under the local profile rather
/// than the roaming one because a browser profile is not something to synchronize
/// between machines.
fn user_data_folder() -> PathBuf {
    directories::BaseDirs::new()
        .map(|dirs| {
            dirs.data_local_dir()
                .join("rust-hover-preview")
                .join("webview")
        })
        .unwrap_or_else(std::env::temp_dir)
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

/// Point the engine at each path in turn and report what each step cost.
///
/// The engine runs on a thread of its own, with its own window and a message pump,
/// because that is what WebView2 requires of the thread that creates it: everything it
/// does is delivered by posted message, and a thread that is not retrieving them never
/// hears back. The calls below are therefore made there rather than here, and the
/// answer is handed back through a channel.
pub fn measure(paths: Vec<PathBuf>, width: i32, height: i32) -> Option<Report> {
    let (sender, receiver) = mpsc::channel();

    std::thread::spawn(move || {
        let _ = sender.send(measure_on_its_own_thread(paths, width, height));
    });

    receiver.recv().ok().flatten()
}

fn measure_on_its_own_thread(paths: Vec<PathBuf>, width: i32, height: i32) -> Option<Report> {
    // An apartment of its own: WebView2 must be created on a thread that is pumping
    // messages, and the thread this runs on is that thread.
    let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

    let report = run(&paths, width, height);

    unsafe {
        windows::Win32::System::Com::CoUninitialize();
    }

    report
}

fn run(paths: &[PathBuf], width: i32, height: i32) -> Option<Report> {
    let folder = user_data_folder();
    std::fs::create_dir_all(&folder).ok()?;

    let hwnd = create_host_window(width, height)?;

    let started = Instant::now();
    let environment = create_environment(&folder)?;
    let environment_ms = started.elapsed().as_millis();

    let started = Instant::now();
    let controller = create_controller(environment, hwnd)?;
    let controller_ms = started.elapsed().as_millis();

    unsafe {
        let _ = controller.SetBounds(RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: height,
        });
        let _ = controller.SetIsVisible(true);
    }

    let webview = unsafe { controller.CoreWebView2().ok()? };
    configure(&webview);

    unsafe {
        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            60,
            60,
            0,
            0,
            SWP_NOACTIVATE | SWP_NOSIZE | SWP_SHOWWINDOW,
        );
    }

    let mut navigations = Vec::new();
    for path in paths {
        let started = Instant::now();
        let arrived = navigate(&webview, path);
        navigations.push((path.clone(), started.elapsed().as_millis()));

        if !arrived {
            break;
        }

        // Long enough for a frame to have been painted and for the engine to settle,
        // so the next navigation is measured against a warm engine rather than against
        // one still starting.
        std::thread::sleep(Duration::from_millis(1500));
    }

    let report = Report {
        runtime_version: runtime_version(),
        user_data_folder: folder,
        environment_ms,
        controller_ms,
        navigations,
    };

    unsafe {
        let _ = controller.Close();
    }
    unsafe {
        let _ = DestroyWindow(hwnd);
    }

    Some(report)
}

/// The window the engine draws into. It is a popup of its own rather than a child of
/// the preview window, because the preview window is a layered one and a layered
/// window has no window tree to put a child in.
fn create_host_window(width: i32, height: i32) -> Option<HWND> {
    unsafe {
        CreateWindowExW(
            Default::default(),
            w!("STATIC"),
            w!("Rust Hover Preview"),
            WS_POPUP,
            60,
            60,
            width,
            height,
            None,
            None,
            None,
            None,
        )
        .ok()
    }
}

fn create_environment(
    user_data_folder: &Path,
) -> Option<webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Environment> {
    let folder = wide(&user_data_folder.to_string_lossy());
    let (sender, receiver) = mpsc::channel();

    CreateCoreWebView2EnvironmentCompletedHandler::wait_for_async_operation(
        Box::new(move |handler| unsafe {
            CreateCoreWebView2EnvironmentWithOptions(
                PCWSTR::null(),
                PCWSTR(folder.as_ptr()),
                None,
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
    environment: webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Environment,
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
/// than a browser's: a document is drawn and not run, and nothing about it is a way
/// out of the preview.
fn configure(webview: &webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2) {
    unsafe {
        if let Ok(settings) = webview.Settings() {
            let _ = settings.SetIsScriptEnabled(false);
            let _ = settings.SetAreDefaultContextMenusEnabled(false);
            let _ = settings.SetAreDevToolsEnabled(false);
            let _ = settings.SetIsZoomControlEnabled(false);
            let _ = settings.SetIsStatusBarEnabled(false);
            let _ = settings.SetIsWebMessageEnabled(false);
        }
    }
}

/// Point the engine at a file and wait for it to arrive.
fn navigate(
    webview: &webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2,
    path: &Path,
) -> bool {
    let Some(url) = file_url(path) else {
        return false;
    };

    let url = wide(&url);
    let (sender, receiver) = mpsc::channel();

    unsafe {
        let handler = NavigationCompletedEventHandler::create(Box::new(move |_sender, _args| {
            let _ = sender.send(());
            Ok(())
        }));

        let mut token = EventRegistrationToken::default();
        if webview
            .add_NavigationCompleted(&handler, &mut token)
            .is_err()
        {
            return false;
        }

        let started = webview.Navigate(PCWSTR(url.as_ptr()));
        let arrived = started.is_ok() && webview2_com::wait_with_pump(receiver).is_ok();

        let _ = webview.remove_NavigationCompleted(token);

        arrived
    }
}

/// A local file as a URL, which is what the engine is pointed at. The characters that
/// would end the path early — a space, a hash, a question mark, a percent — are the
/// ones escaped; anything else is left as it is written.
fn file_url(path: &Path) -> Option<String> {
    if !path.is_absolute() {
        return None;
    }

    let mut url = String::from("file:///");

    for character in path.to_string_lossy().chars() {
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

        match measure(paths, 800, 800) {
            Some(report) => {
                println!("user data: {}", report.user_data_folder.display());
                println!("environment: {} ms", report.environment_ms);
                println!("controller: {} ms", report.controller_ms);

                for (path, elapsed) in &report.navigations {
                    println!("navigated {} in {} ms", path.display(), elapsed);
                }
            }
            None => println!("no engine answered"),
        }
    }
}

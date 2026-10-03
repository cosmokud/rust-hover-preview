use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::time::{Duration, Instant};

use webview2_com::Microsoft::Web::WebView2::Win32::{
    CreateCoreWebView2EnvironmentWithOptions, ICoreWebView2, ICoreWebView2Controller,
    ICoreWebView2Environment, ICoreWebView2EnvironmentOptions, COREWEBVIEW2_COLOR,
};
use webview2_com::{
    CoreWebView2EnvironmentOptions, CreateCoreWebView2ControllerCompletedHandler,
    CreateCoreWebView2EnvironmentCompletedHandler,
};
use windows::core::{w, Interface, PCWSTR};
use windows::Win32::Foundation::{E_POINTER, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, PeekMessageW, RegisterClassW,
    TranslateMessage, MSG, PM_REMOVE, WM_MOUSEACTIVATE, WNDCLASSW, WS_EX_NOACTIVATE,
    WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

use super::api::{BROWSER_ARGUMENTS, RUNNING_DOCUMENT, WEBVIEW_CLASS};
use super::engine::trace;

use crate::app::engine_processes;
use crate::config::config::TransparentBackground;
use crate::paths::plain_path;

/// The window the engine draws into: a popup of its own, topmost, tool-windowed and
/// refusing activation to begin with. The refusal is put on here rather than answered here
/// alone because it is the style that makes a click refuse as well, and it comes off for the
/// one document that runs (see `ex_style_for`).
pub(super) fn create_host_window() -> Option<HWND> {
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

pub(super) fn register_class() {
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

/// A preview never takes the keyboard by itself: the pointer may be over it, but what is
/// being worked in is Explorer — unless the thing on screen is a page that runs, which is
/// the one document whose whole reading is a script answering keys, and a page that is
/// never activated is a page nothing can be typed into (see `mouse_activate_answers`).
extern "system" fn window_proc(
    hwnd: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if message == WM_MOUSEACTIVATE {
        return mouse_activate_answers(RUNNING_DOCUMENT.load(Ordering::Acquire));
    }

    unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
}

/// What a click into the engine's window is answered with, given what the window is showing.
///
/// Two answers, because a preview is two things. A document and a specimen are looked at, and
/// the pointer being over one of them says nothing about the user having left the window they
/// are working in, so a click is refused activation and the caret stays where it was. A page
/// of HTML that runs is interacted with, and a click that refused to activate it would leave
/// a page on screen that no key could reach, which is a page the user is looking at a picture
/// of. So the one document that is a program rather than a picture takes the keyboard, and it
/// takes it only on a click — showing it is still `SW_SHOWNOACTIVATE`, so a hover never
/// steals the caret from what the hand is on.
pub(super) fn mouse_activate_answers(runs: bool) -> LRESULT {
    const MA_ACTIVATE: LRESULT = LRESULT(1);
    const MA_NOACTIVATE: LRESULT = LRESULT(3);

    if runs {
        MA_ACTIVATE
    } else {
        MA_NOACTIVATE
    }
}

/// The window's extended style as it is for a document that runs, or for one that is only
/// looked at: `WS_EX_NOACTIVATE` off for the first and on for the second.
///
/// The style is the standing part of the refusal — a window that carries it cannot be
/// activated by anything, the click included, which is why it has to come off before
/// `WM_MOUSEACTIVATE` can answer `MA_ACTIVATE` for a page that runs. Nothing else in the
/// style is touched: the window stays a tool window so it is nothing in the taskbar or on
/// the alt-tab, and stays topmost, since it is a preview the pointer is on top of. A style
/// that has come off for a page is put back for the next document, which is not one — so the
/// refusal is the default and the allowance is the exception made and taken away again.
pub(super) fn ex_style_for(style: isize, runs: bool) -> isize {
    let noactivate = WS_EX_NOACTIVATE.0 as isize;

    if runs {
        style & !noactivate
    } else {
        style | noactivate
    }
}

/// Retrieve and dispatch whatever is waiting, without blocking.
pub(super) fn pump_messages() {
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
pub(super) fn user_data_folder() -> PathBuf {
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
pub(crate) fn stale_profile_pids() -> Vec<u32> {
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
pub(crate) fn clear_stale_profiles() {
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

pub(super) fn create_environment(
    user_data_folder: &Path,
    deadline: Instant,
) -> Option<ICoreWebView2Environment> {
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

    receiver.recv_timeout(remaining(deadline)).ok()?.ok()
}

/// How long the engine may take to be had at all — the environment and the controller
/// together — before the attempt is read as one that will not answer.
///
/// Both calls are asynchronous and both of them were waited on without a bound: the
/// completion handlers this thread parks on are the runtime's to fire, and one that never
/// fires leaves this thread waiting for the rest of the run — with every document after it
/// queued behind a wait nobody ends, and no failure notice to take a hover's spinner down
/// with, because the thread that would post one is the thread that is waiting. Twenty
/// seconds is far past what having an engine costs (a quarter of a second measured warm,
/// and the retries a held profile folder asks for are inside it), so what is past it is
/// silence rather than work.
pub(super) const HOST_CREATION_TIMEOUT: Duration = Duration::from_secs(20);

/// What is left of a deadline, as a wait: `recv_timeout` given nothing waits nothing and
/// answers `Timeout` at once, which is the answer a wait past its deadline wants.
pub(super) fn remaining(deadline: Instant) -> Duration {
    deadline.saturating_duration_since(Instant::now())
}

pub(super) fn create_controller(
    environment: ICoreWebView2Environment,
    hwnd: HWND,
    deadline: Instant,
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

    receiver.recv_timeout(remaining(deadline)).ok()?.ok()
}

/// The browser's own furniture, all of it taken away: the context menu, the tools, the zoom,
/// the status bar, the web messages a page could reach this app's threads through, and the
/// accelerator keys a browser answers for itself — a find bar, a print dialog — which belong
/// to no preview.
///
/// None of these says anything about the document. Whether it runs is not decided here, and
/// that is the one setting of the browser's that is not the same for every document this
/// engine is given (see `set_scripts`).
pub(super) fn configure(webview: &ICoreWebView2) {
    unsafe {
        if let Ok(settings) = webview.Settings() {
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

/// Whether the engine is to run the code in the document it is pointed at: the one setting of
/// the browser's that is a property of the document in it rather than of the engine, and so is
/// made as the document is pointed at (see `Host::navigate`).
pub(super) fn set_scripts(webview: &ICoreWebView2, on: bool) {
    if let Ok(settings) = unsafe { webview.Settings() } {
        let _ = unsafe { settings.SetIsScriptEnabled(on) };
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
pub(super) fn file_version(path: &Path) -> u64 {
    std::fs::metadata(path)
        .and_then(|meta| meta.modified())
        .ok()
        .and_then(|modified| modified.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|since| since.as_millis() as u64)
        .unwrap_or(0)
}

/// The engine's background, which is what the preview's own backdrop setting means to a
/// window that composites for itself.
pub(super) fn background_color(background: TransparentBackground) -> COREWEBVIEW2_COLOR {
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
        // whatever composites the frame, and this window composites its own. The page
        // paints it instead — see `frame_page` — so what this colour is is what stands
        // behind the page until it has been drawn: mid grey, which is what the squares
        // average to.
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
/// A verbatim path is not something a URL may contain: a browser pointed at one fails at
/// once, silently, and what a hover shows is nothing. So the Shell's prefix comes off
/// first (`crate::paths::plain_path`), and a share keeps its server: `\\?\UNC\server\share`
/// becomes `file://server/share`, a drive becomes `file:///C:/…`. The characters that
/// would end the path early — a space, a hash, a question mark, a percent — are escaped;
/// anything else is left as it is written.
pub(super) fn file_url(path: &Path) -> Option<String> {
    if !path.is_absolute() {
        return None;
    }

    let local = plain_path(path);

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

pub(super) fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

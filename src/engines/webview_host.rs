//! Shared WebView2 STA host infrastructure for both `webview_preview` and `pdf_viewer`.
//!
//! Owns the COM STA apartment guard, Win32 message-pump helpers, host HWND class registration and
//! creation, WebView2 environment/controller creation, controller bounds sync, and the
//! `NavigationCompleted` reveal state machine.

use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::time::{Duration, Instant};

use webview2_com::Microsoft::Web::WebView2::Win32::{
    CreateCoreWebView2EnvironmentWithOptions, ICoreWebView2, ICoreWebView2Controller,
    ICoreWebView2Environment, ICoreWebView2EnvironmentOptions,
    ICoreWebView2NavigationCompletedEventHandler, ICoreWebView2_3,
};
use webview2_com::{
    CoreWebView2EnvironmentOptions, CreateCoreWebView2ControllerCompletedHandler,
    CreateCoreWebView2EnvironmentCompletedHandler, NavigationCompletedEventHandler,
};
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::{E_POINTER, HWND, RECT};
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::System::WinRT::EventRegistrationToken;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, DispatchMessageW, GetClientRect, GetMessageW, KillTimer,
    PeekMessageW, RegisterClassW, SetTimer, SetWindowPos, ShowWindow, TranslateMessage, MSG,
    PM_NOREMOVE, PM_REMOVE, SET_WINDOW_POS_FLAGS, SW_HIDE, SW_SHOWNOACTIVATE, WINDOW_EX_STYLE,
    WNDCLASSW, WNDPROC, WS_POPUP,
};

use crate::app::engine_processes;

/// Screen-space rectangle a WebView2 host window is placed in.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) struct Area {
    pub(crate) x: i32,
    pub(crate) y: i32,
    pub(crate) width: i32,
    pub(crate) height: i32,
}

/// RAII COM STA apartment + primed Win32 message queue for a WebView2 host thread.
pub(crate) struct StaApartment;

impl StaApartment {
    /// Initialize COM as STA and ensure the thread message queue exists before anything posts to it.
    pub(crate) fn enter() -> Self {
        let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };
        let mut message = MSG::default();
        unsafe {
            let _ = PeekMessageW(&mut message, None, 0, 0, PM_NOREMOVE);
        }
        Self
    }
}

impl Drop for StaApartment {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

/// Non-blocking Win32 message pump for the thread that owns WebView2 controllers.
pub(crate) fn pump_messages() {
    let mut message = MSG::default();
    unsafe {
        while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

/// Why a bounded message-pump wait ended without an answer.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Unanswered {
    TimedOut,
    Gone,
}

/// Pump this thread's messages until `probe` returns `Some`, `timeout` expires, or the queue ends.
pub(crate) fn pump_until<T>(
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
            break Err(Unanswered::Gone);
        }
    };

    unsafe {
        let _ = KillTimer(HWND::default(), timer);
    }

    answer
}

/// Register a WebView2 host window class once.
pub(crate) fn register_host_class(registered: &AtomicBool, class_name: PCWSTR, wnd_proc: WNDPROC) {
    if registered.swap(true, Ordering::AcqRel) {
        return;
    }

    let class = WNDCLASSW {
        lpfnWndProc: wnd_proc,
        lpszClassName: class_name,
        ..Default::default()
    };

    unsafe {
        let _ = RegisterClassW(&class);
    }
}

/// Create a hidden `WS_POPUP` host window for a WebView2 controller.
pub(crate) fn create_host_window(
    ex_style: WINDOW_EX_STYLE,
    class_name: PCWSTR,
    title: PCWSTR,
    area: Area,
    owner: Option<HWND>,
) -> Option<HWND> {
    unsafe {
        CreateWindowExW(
            ex_style,
            class_name,
            title,
            WS_POPUP,
            area.x,
            area.y,
            area.width,
            area.height,
            owner.unwrap_or_default(),
            None,
            None,
            None,
        )
        .ok()
    }
}

/// Remaining duration until `deadline` (`Duration::ZERO` once expired).
pub(crate) fn remaining(deadline: Instant) -> Duration {
    deadline.saturating_duration_since(Instant::now())
}

/// Snapshot current browser child PIDs of this process before starting a new WebView2 instance.
pub(crate) fn snapshot_browser_pids() -> Vec<u32> {
    engine_processes::processes_named_by_parent(engine_processes::BROWSER_IMAGE, std::process::id())
}

/// Browser child PIDs that appeared after `before` was snapshotted.
pub(crate) fn new_browser_pids(before: &[u32]) -> Vec<u32> {
    snapshot_browser_pids()
        .into_iter()
        .filter(|pid| !before.contains(pid))
        .collect()
}

/// Create a WebView2 environment with `arguments` under `user_data_folder`, bounded by `deadline`.
///
/// **An elevated run is told to turn the browser's sandbox off, or it has no browser at all.**
/// WebView2 refuses to start its own process from a token at a higher integrity level than the one
/// the browser was meant to run at — the check is Chromium's, and it fails whichever account ran
/// the UAC prompt — so a document, an ebook, a page or a font would be drawn by nothing on the one
/// machine where this app was started as an administrator. `--no-sandbox` is what turns that check
/// off, and there is nothing left for it to protect here: the browser draws a file the user already
/// opened, in a window of this app's that is never navigated anywhere else, with the resolver rule
/// above making sure nothing it links to is fetched. It is added here rather than at each caller so
/// that every engine of this app's — the hover's, the PDF pin's — is answered the same way.
pub(crate) fn create_environment(
    user_data_folder: &Path,
    deadline: Instant,
    arguments: &str,
) -> Option<ICoreWebView2Environment> {
    let folder = wide(&user_data_folder.to_string_lossy());
    let (sender, receiver) = mpsc::channel();

    let arguments = if crate::app::elevation::is_elevated() {
        format!("{arguments} --no-sandbox")
    } else {
        arguments.to_string()
    };

    let options = CoreWebView2EnvironmentOptions::default();
    unsafe {
        options.set_additional_browser_arguments(arguments);
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
            let _ = sender.send(environment.ok_or_else(|| windows::core::Error::from(E_POINTER)));
            Ok(())
        }),
    )
    .ok()?;

    receiver.recv_timeout(remaining(deadline)).ok()?.ok()
}

/// Create a WebView2 controller bound to `hwnd`, bounded by `deadline`.
pub(crate) fn create_controller(
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
            error_code?;
            let _ = sender.send(controller.ok_or_else(|| windows::core::Error::from(E_POINTER)));
            Ok(())
        }),
    )
    .ok()?;

    receiver.recv_timeout(remaining(deadline)).ok()?.ok()
}

/// Register a `NavigationCompleted` handler delivering `args.IsSuccess()` to `callback`.
pub(crate) fn register_navigation_completed(
    webview: &ICoreWebView2,
    mut callback: impl FnMut(bool) + 'static,
) -> Option<(
    ICoreWebView2NavigationCompletedEventHandler,
    EventRegistrationToken,
)> {
    let handler = NavigationCompletedEventHandler::create(Box::new(move |_sender, args| {
        let mut arrived = windows::Win32::Foundation::BOOL(0);
        unsafe {
            let _ = args.is_some_and(|args| args.IsSuccess(&mut arrived).is_ok());
        }
        callback(arrived.as_bool());
        Ok(())
    }));

    let mut token = EventRegistrationToken::default();
    unsafe {
        webview.add_NavigationCompleted(&handler, &mut token).ok()?;
    }

    Some((handler, token))
}

/// Local absolute path encoded as a `file://` URL (stripping `\\?\` verbatim prefixes).
///
/// The one spelling a browser and a custom preview's program both take, so it is asked of
/// [`crate::paths::file_url`] here rather than written twice.
pub(crate) fn file_url(path: &Path) -> Option<String> {
    crate::paths::file_url(path)
}

/// NUL-terminated UTF-16 buffer for Win32 / WebView2 APIs.
pub(crate) fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Shared host surface owning a Win32 host HWND paired with a WebView2 controller and WebView.
pub(crate) struct WebView2Host {
    hwnd: HWND,
    controller: ICoreWebView2Controller,
    webview: ICoreWebView2,
    closed: bool,
}

impl WebView2Host {
    pub(crate) fn new(
        hwnd: HWND,
        controller: ICoreWebView2Controller,
        webview: ICoreWebView2,
    ) -> Self {
        Self {
            hwnd,
            controller,
            webview,
            closed: false,
        }
    }

    pub(crate) fn hwnd(&self) -> HWND {
        self.hwnd
    }

    pub(crate) fn controller(&self) -> &ICoreWebView2Controller {
        &self.controller
    }

    pub(crate) fn webview(&self) -> &ICoreWebView2 {
        &self.webview
    }

    /// Synchronize controller bounds to an explicit `Area` size.
    pub(crate) fn set_bounds(&self, area: Area) {
        unsafe {
            let _ = self.controller.SetBounds(RECT {
                left: 0,
                top: 0,
                right: area.width,
                bottom: area.height,
            });
        }
    }

    /// Synchronize controller bounds to the host window's current client rect (min 1x1).
    pub(crate) fn set_bounds_to_client(&self) {
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

    pub(crate) fn set_visible(&self, visible: bool) {
        unsafe {
            let _ = self.controller.SetIsVisible(visible);
        }
    }

    pub(crate) fn show_window(&self, visible: bool) {
        unsafe {
            let _ = ShowWindow(self.hwnd, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE });
        }
    }

    pub(crate) fn move_window(&self, area: Area, insert_after: HWND, flags: SET_WINDOW_POS_FLAGS) {
        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                insert_after,
                area.x,
                area.y,
                area.width,
                area.height,
                flags,
            );
        }
    }

    pub(crate) fn resume(&self) {
        if let Ok(webview) = self.webview.cast::<ICoreWebView2_3>() {
            unsafe {
                let _ = webview.Resume();
            }
        }
    }

    /// Idempotently close the controller and destroy the host HWND.
    pub(crate) fn close(&mut self) {
        if self.closed {
            return;
        }
        self.closed = true;

        unsafe {
            let _ = self.controller.Close();
            let _ = DestroyWindow(self.hwnd);
        }
    }
}

/// Tracks the latest placement box and whether size changed or a box arrived since navigation.
#[derive(Clone, Copy, Debug)]
pub(crate) struct PlacementSync {
    area: Area,
    last_size: Option<(i32, i32)>,
    placed: bool,
}

impl PlacementSync {
    pub(crate) fn new(area: Area) -> Self {
        Self {
            area,
            last_size: None,
            placed: false,
        }
    }

    pub(crate) fn area(&self) -> Area {
        self.area
    }

    pub(crate) fn placed(&self) -> bool {
        self.placed
    }

    pub(crate) fn reset_placed(&mut self) {
        self.placed = false;
    }

    /// Record a new placement box; returns `true` if `(width, height)` changed since last recorded size.
    pub(crate) fn record(&mut self, area: Area) -> bool {
        let size = (area.width, area.height);
        let resized = self.last_size != Some(size);
        self.area = area;
        self.last_size = Some(size);
        self.placed = true;
        resized
    }
}

/// `NavigationCompleted` reveal state machine shared by hosted WebView2 viewers.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum RevealState {
    /// Created hidden and navigating; not yet revealed.
    Arriving,
    /// Document arrived and window is visible.
    Shown,
    /// Window hidden/minimized and controller suspended.
    Suspended,
    /// Closed; terminal state.
    Closed,
}

/// Action owed by a [`RevealState`] transition.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum RevealStep {
    Reveal,
    Suspend,
    Resume,
    Close,
    Nothing,
}

impl RevealState {
    /// Handle `NavigationCompleted` outcome; reveals only once both `arrived` and `placed` hold.
    pub(crate) fn navigated(&mut self, arrived: bool, placed: bool) -> RevealStep {
        if *self != Self::Arriving {
            return RevealStep::Nothing;
        }

        if !arrived {
            *self = Self::Closed;
            return RevealStep::Close;
        }

        if !placed {
            return RevealStep::Nothing;
        }

        *self = Self::Shown;
        RevealStep::Reveal
    }

    /// Transition between `Shown` and `Suspended` based on whether the host is on screen.
    pub(crate) fn settle(&mut self, on_screen: bool) -> RevealStep {
        match (*self, on_screen) {
            (Self::Shown, false) => {
                *self = Self::Suspended;
                RevealStep::Suspend
            }
            (Self::Suspended, true) => {
                *self = Self::Shown;
                RevealStep::Resume
            }
            _ => RevealStep::Nothing,
        }
    }

    /// Transition to `Closed` once; subsequent calls return `RevealStep::Nothing`.
    pub(crate) fn closed(&mut self) -> RevealStep {
        if *self == Self::Closed {
            return RevealStep::Nothing;
        }

        *self = Self::Closed;
        RevealStep::Close
    }

    pub(crate) fn is_arriving(self) -> bool {
        self == Self::Arriving
    }

    pub(crate) fn is_closed(self) -> bool {
        self == Self::Closed
    }

    pub(crate) fn is_up(self) -> bool {
        matches!(self, Self::Shown | Self::Suspended)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reveal_state_waits_for_both_navigation_and_placement() {
        let mut state = RevealState::Arriving;
        assert_eq!(state.navigated(true, false), RevealStep::Nothing);
        assert!(state.is_arriving());

        assert_eq!(state.navigated(true, true), RevealStep::Reveal);
        assert_eq!(state, RevealState::Shown);
        assert!(state.is_up());

        assert_eq!(state.settle(false), RevealStep::Suspend);
        assert_eq!(state, RevealState::Suspended);
        assert!(state.is_up());

        assert_eq!(state.settle(true), RevealStep::Resume);
        assert_eq!(state, RevealState::Shown);

        assert_eq!(state.closed(), RevealStep::Close);
        assert!(state.is_closed());
        assert_eq!(state.closed(), RevealStep::Nothing);
    }

    #[test]
    fn reveal_state_closes_immediately_on_failed_navigation() {
        let mut state = RevealState::Arriving;
        assert_eq!(state.navigated(false, false), RevealStep::Close);
        assert!(state.is_closed());
    }

    #[test]
    fn placement_sync_detects_resizes_and_tracks_placed_flag() {
        let initial = Area {
            x: 10,
            y: 20,
            width: 400,
            height: 300,
        };
        let mut sync = PlacementSync::new(initial);
        assert!(!sync.placed());

        assert!(sync.record(initial));
        assert!(sync.placed());

        let moved = Area {
            x: 50,
            y: 80,
            ..initial
        };
        assert!(!sync.record(moved));
        assert_eq!(sync.area(), moved);

        let resized = Area {
            width: 600,
            ..moved
        };
        assert!(sync.record(resized));
        assert_eq!(sync.area(), resized);

        sync.reset_placed();
        assert!(!sync.placed());
    }
}

//! A preview drawn by a program of the user's own: the custom preview's hover.
//!
//! **What a custom preview is.** A preset in `config.ini` names a program and the extensions it
//! claims (`config::config::custom_presets`), and a file it claims is not decoded, read or drawn
//! by this app at all: the program is started with the file and its window is what the hover
//! shows. The preset outranks every list this app ships, so a name it claims never reaches a
//! reader or an engine.
//!
//! **Why the picture is a thumbnail and not a capture.** The window belongs to another process,
//! and a window of another process cannot be composited into a layered window of this app's —
//! so the picture is the Shell's: the program's window is stripped of its frame, parked past the
//! right edge of every display at the box the preview is drawn at, and registered into a host
//! window of this app's with `DwmRegisterThumbnail`. The compositor draws it, live, at whatever
//! rate the program presents — which is the whole of why this is not a `PrintWindow` capture: a
//! captured frame is one frame, it costs a copy per tick, and a program that draws with the GPU
//! answers a capture with black. It is the road a film of this app's own already takes (see
//! `pin_film_host` and `pin_film_thumbnail`), used here for a window nothing of this app's owns.
//!
//! **Which window is the program's.** Not simply the launched process's: a packaged program —
//! Windows 11's Paint among them — is begun by an activator that hands the file to a process of
//! its own, so the window that is the picture belongs to a process this app never started, and a
//! search by the launched id alone finds the activator's stub instead: an empty host beside the
//! program's real window, standing on the desktop where nobody put it. So the window is the
//! **new** one — one that was not on the desktop when the program was begun — whose process is
//! either in the launched tree or running the same image as the preset's program (see `census`),
//! and of those the one the program actually *shows*: the window the user would have seen, which
//! is the only window a program can be assumed to draw its document in. A window that is never
//! shown (the launch asked for that) is taken after a moment on the same terms.
//!
//! **And it is a document window rather than every window a program owns.** A window with no
//! title, or one the frame makes a few pixels square, is a handle a program's thread is woken on
//! rather than a place a document is drawn: winit — the window library PhotoCraft is written on —
//! makes a `Winit Thread Event Target` of 16 pixels at 0,0 and **that is the only one of its
//! windows that is visible at birth**, the real `Window Class` window being born hidden because
//! this app's launch asked it to be. A picture taken by visibility alone was that event target:
//! the hover drew a blank 16-pixel window while the photograph's own window was never adopted at
//! all. So the picture is a titled, unowned, non-tool window of a document's size (see
//! `THE_SMALLEST_PICTURE`), and where the program has several, the largest is the one it draws
//! its file in.
//!
//! **Why the program outlives the hover.** A program makes its own window and shows it at its own
//! size, and it does so *after* this app has started it — so for the first moment of its life a
//! window of somebody else's stands on the desktop, and often under the hand that asked for the
//! preview. That is a window the pointer has left the file for: the hook takes the hover down, and
//! a hover that killed its program on the way down would start it again on the way back up, and
//! again, and again — a preview that blinks instead of one that is read, and a program killed
//! before it ever opens the file it was given. So the two lifetimes are not the same one: the
//! *host* goes with the hover at once, and the *program* is kept for [`LINGER`] after it, so that
//! a hover the pointer gives back is the same program shown again rather than a second launch. It
//! is ended when the moment passes, when another file is hovered, or when the pin opens it.
//!
//! **And why it is hidden before it is ever seen, and parked before it is ever drawn.** The launch
//! carries `SW_HIDE` in its `STARTUPINFO` (the road a model's viewer already takes, see
//! `never_shown`), and a program that shows its window anyway — one that calls `ShowWindow` itself
//! with a state of its own — is hidden again on every poll of the wait, every 10 ms, from the
//! moment its first window exists. A window that opens maximized is *restored* rather than
//! resized: a `SetWindowPos` on a maximized window changes the size it would have been restored to
//! and leaves it covering the display it is maximized on. And it is shown only where it
//! **verifiably stands off the desktop at the parking corner** (see `park`): a window this app could
//! not move is one it must not put in front of the user, because a frame of somebody's program over
//! the listing is worse than no picture at all. The window may stand there at a size of its own —
//! the corner is what says it is off the desktop, and its size is what `the_picture_in` composes
//! into the box.
//!
//! **Where the box is, and how big.** The placement is the one every other preview takes: the
//! pointer's own (`compute_mouse_layout`) or the focused item's (`compute_keyboard_layout`),
//! which is what reads `Tray → Placement` — the position mode, the pointer standoff and the
//! `Avoid` region — and the room of the display the hover is on. The **shape** of the box is the
//! file's own: a picture says how wide and how tall it is in its own header, so a picture that is
//! taller than it is wide is a tall box (see `base_size`). Its **size** is the preset's `scaling`,
//! a share of the room the placement leaves — `100%` is the whole of that room cut to the file's
//! shape, `75%` three quarters of it, `Fit to Screen` the same as `100%` — which is the rule the
//! preset's own window states in its `Screen 100%`/`Screen 75%` rows (see `layout_for`). A
//! 1024x1024 picture on a 1440-tall display is a 1440x1440 box, whatever program draws it, which
//! is what makes a hover a preview of the file rather than a picture at its own pixel size.
//! **And the box a picture's shape cuts is one a window may refuse.** A window is not obliged to
//! draw itself in the box a preview asks for: Windows 11's Paint answers `WM_GETMINMAXINFO` with a
//! floor of 704x500 and takes anything smaller as 704x500 — measured with its frame stripped, as it
//! is here, and by a `SetWindowPlacement` as well as by a `SetWindowPos` — so a tall picture's box,
//! narrow by its shape, is refused, and a window that stands at its floor inside that box has a
//! shape of its own: composed into the box as it stands it left the picture in a band of background
//! above and below it. So the window is asked again for the box's own shape at the smallest size it
//! will take (see `the_ask_for`), and composed into the box at exactly that shape (see
//! `the_picture_in`): the hover is the size the preset's `scaling` says it is, and the program's
//! picture fills it. It is one rule for every program rather than a list of programs: what the window
//! will take is read off the window.
//! **A program's window is not that size**: it opens at whatever size it was last left at — one
//! left maximized is the whole display — so a box read off it is the size of the last file
//! somebody opened with the program, and a tall picture is shown in a wide window that fills the
//! screen. A file that declares no size of its own — a document, an archive, a sound — is the one
//! case where the program's window is the only shape there is, and that shape is read **once**,
//! before the window is parked, because parking writes the parked box into the window's own
//! placement and a second reading would be the shape of the last preview rather than of the
//! program. The room its share is of is the same room every preset's is (see `layout_for`): such a
//! file is cut to the window's shape where a picture is cut to its own.
//! `background` is what the host paints where the program's window is smaller than the box.
//! `volume` is kept in the preset and written back; a program of the user's own is not told what
//! level to play at — there is no portable way to set another process's audio — so it is a setting
//! a future engine-backed preset can use and nothing on this road reads it.
//!
//! **Why the host is a window of the hover's own.** The preview window paints a card and must
//! never take the focus; the host is the same shape as a hovered film's (`hover_film_host`):
//! a borderless, non-activating, topmost window that is shown only once the Shell has the picture
//! composed in it, and taken down with the hover. It is disabled, so a press on it is the hover's
//! press and lands on the listing behind it. It is made once and kept hidden between hovers:
//! making it again for every hover is a second window where one will do, and one more thing that
//! can flicker.
//!
//! **And the pin is not a pin.** Pressing the pin key on one of these does not freeze a picture
//! of the program: it starts the program properly, with its own frame, at its own place on the
//! desktop, and this app's preview goes away — the user asked for the file, not for a picture of
//! it (see [`open_fully`]).

use std::cell::RefCell;
use std::collections::HashMap;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::os::windows::io::{AsRawHandle, FromRawHandle, OwnedHandle};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::atomic::{AtomicBool, AtomicI32, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex, Once};
use std::time::{Duration, Instant};

use windows::core::{w, PCSTR, PCWSTR, PWSTR};
use windows::Win32::Foundation::{
    CloseHandle, FreeLibrary, BOOL, COLORREF, FILETIME, HANDLE, HINSTANCE, HMODULE, HWND,
    INVALID_HANDLE_VALUE, LPARAM, LRESULT, RECT, WAIT_OBJECT_0, WAIT_TIMEOUT, WPARAM,
};
use windows::Win32::Graphics::Dwm::{
    DwmFlush, DwmGetWindowAttribute, DwmQueryThumbnailSourceSize, DwmRegisterThumbnail,
    DwmSetWindowAttribute, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWMWA_EXTENDED_FRAME_BOUNDS, DWMWA_TRANSITIONS_FORCEDISABLED, DWM_THUMBNAIL_PROPERTIES,
    DWM_TNP_OPACITY, DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{
    CreateSolidBrush, DeleteObject, FillRect, GetMonitorInfoW, GetStockObject, InvalidateRect,
    MonitorFromRect, UpdateWindow, HBRUSH, HGDIOBJ, MONITORINFO, MONITOR_DEFAULTTONEAREST,
    NULL_BRUSH,
};
use windows::Win32::System::JobObjects::{
    AssignProcessToJobObject, CreateJobObjectW, JobObjectExtendedLimitInformation,
    SetInformationJobObject, JOBOBJECT_EXTENDED_LIMIT_INFORMATION, JOB_OBJECT_LIMIT,
    JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE,
};
use windows::Win32::System::LibraryLoader::{GetModuleHandleW, GetProcAddress, LoadLibraryW};
use windows::Win32::System::Memory::{
    CreateFileMappingW, MapViewOfFile, UnmapViewOfFile, FILE_MAP_ALL_ACCESS,
    MEMORY_MAPPED_VIEW_ADDRESS, PAGE_READWRITE,
};
use windows::Win32::System::Threading::{
    CreateProcessW, GetProcessTimes, OpenProcess, QueryFullProcessImageNameW, ResumeThread,
    TerminateProcess, WaitForInputIdle, WaitForSingleObject, CREATE_SUSPENDED,
    PROCESS_CREATION_FLAGS, PROCESS_INFORMATION, PROCESS_NAME_WIN32,
    PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_SET_QUOTA, PROCESS_SYNCHRONIZE, PROCESS_TERMINATE,
    STARTF_USEPOSITION, STARTF_USESHOWWINDOW, STARTF_USESIZE, STARTUPINFOW,
};
use windows::Win32::UI::Accessibility::{SetWinEventHook, UnhookWinEvent, HWINEVENTHOOK};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, IsWindowEnabled};
use windows::Win32::UI::WindowsAndMessaging::{
    AllowSetForegroundWindow, BringWindowToTop, CreateWindowExW, DefWindowProcW, DispatchMessageW,
    EnumWindows, GetClassNameW, GetClientRect, GetForegroundWindow, GetLayeredWindowAttributes,
    GetSystemMetrics, GetWindow, GetWindowLongPtrW, GetWindowPlacement, GetWindowRect,
    GetWindowTextLengthW, GetWindowThreadProcessId, IsIconic, IsWindow, IsWindowVisible, IsZoomed,
    MsgWaitForMultipleObjectsEx, PeekMessageW, PostMessageW, RegisterClassExW,
    RegisterWindowMessageW, SendMessageTimeoutW, SetForegroundWindow, SetLayeredWindowAttributes,
    SetWindowLongPtrW, SetWindowPlacement, SetWindowPos, SetWindowsHookExW, ShowWindow,
    TranslateMessage, UnhookWindowsHookEx, CHILDID_SELF, EVENT_OBJECT_CREATE,
    EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_SHOW, GWL_EXSTYLE, GWL_STYLE, GW_OWNER, HHOOK,
    HOOKPROC, HWND_BOTTOM, HWND_NOTOPMOST, HWND_TOPMOST, LAYERED_WINDOW_ATTRIBUTES_FLAGS,
    LWA_ALPHA, MSG, MWMO_INPUTAVAILABLE, OBJID_WINDOW, PM_REMOVE, QS_ALLINPUT, SMTO_ABORTIFHUNG,
    SMTO_NORMAL, SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN,
    SWP_ASYNCWINDOWPOS, SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOOWNERZORDER,
    SWP_NOSIZE, SWP_NOZORDER, SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE, WH_CALLWNDPROC, WH_CBT,
    WINDOWPLACEMENT, WINDOWPLACEMENT_FLAGS, WINDOW_EX_STYLE, WINDOW_STYLE, WINEVENT_OUTOFCONTEXT,
    WINEVENT_SKIPOWNPROCESS, WM_APP, WM_CLOSE, WM_ERASEBKGND, WM_NCDESTROY, WM_NULL, WNDCLASSEXW,
    WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN, WS_EX_LAYERED, WS_EX_NOACTIVATE,
    WS_EX_NOREDIRECTIONBITMAP, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_MAXIMIZEBOX,
    WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU, WS_THICKFRAME,
};

use super::audio_session;

use super::dimensions::{design_dimensions, picture_dimensions, vector_dimensions};
use super::displays::{dpi_at, work_area_at};
use super::hover_facts::HoverFacts;
use super::html_is_engine_drawn;
use super::layout::{
    compute_keyboard_layout, compute_mouse_layout, KeyboardPlacement, PreviewLayout,
};
use super::load_dispatch::HoverPlacement;
use super::media_data::ProbedGeometry;
use super::pointer::{AvoidRegion, ScreenRegion};
use super::preview_scale::{animated_scale_for, fit_reduced, HoverScales};
use super::video_probe::cached_video_geometry;
use crate::app::engine_processes;
use crate::config::config::{
    custom_preset_for, CustomCaptureTechnique, CustomPreset, CustomPromotionTechnique,
    CustomScaling, CustomWindowMode, PreviewScale, PreviewType, DEFAULT_AUDIO_SCALE,
};
use crate::readers::{pdf_preview, svg_preview, video_player};
use crate::CONFIG;

/// The class a custom preview's host is made from.
const HOST_CLASS: windows::core::PCWSTR = w!("RustHoverPreviewCustomPreviewHost");

/// The host's own message: the program's window is drawing, so the host may be shown.
const WM_THE_PICTURE_IS_READY: u32 = WM_APP + 41;
/// The host's own message: the program made no window to draw, or the hover is gone, so the host
/// leaves the screen.
const WM_THERE_IS_NOTHING_TO_SHOW: u32 = WM_APP + 42;

/// How long the program is given to make its window, and how often it is asked.
///
/// A program that takes longer than this is one whose preview never appears — a slow launch is
/// not a preview, and the hover has moved on by then. The wait runs on the worker thread, so
/// nothing of this app's is held by it. The cadence is the one a window is *hidden* at: a program
/// shows its own window the moment it has one, and every poll is a poll a window can be hidden on.
const WINDOW_WAIT: Duration = Duration::from_secs(6);

/// How often the launched process's own tree is read again while its window is looked for: a
/// program that hands its file to a process of its own is a window under that process's id.
const TREE_POLL: Duration = Duration::from_millis(200);

/// How small a window of the program's may be and still be the picture: a helper window is a
/// handle a program's thread is woken on, not a place a document is drawn.
///
/// Measured on PhotoCraft 0.23.1: its `Window Class` window is 2580x1059 and titled `PhotoCraft`,
/// and its `Winit Thread Event Target` is 16 pixels square at 0,0 with no title at all — and of
/// the two it is the event target that is visible at birth, because the launch hides the other.
const THE_SMALLEST_PICTURE: i64 = 64 * 64;

/// How often the program's window is looked at again, and kept at the box if the program moved
/// itself: a window shown again by its own program, or restored by it, is one the Shell would
/// compose at the place the program chose rather than at the hover's.
const KEEP_POLL: Duration = Duration::from_millis(40);

/// How long a program outlives the hover that started it.
///
/// **This is what makes a hover that flickers a hover rather than a launch loop.** The program's
/// own window is on the desktop for the moment between its first `ShowWindow` and this app's own
/// hide of it, and a window under the hand is a hover taken down: the hover comes back a moment
/// later, and it is this long a moment. Long enough to cover it, short enough that a program the
/// user has stopped hovering is not left running.
const LINGER: Duration = Duration::from_secs(2);

/// Neutral dark backdrop painted on the host where the target window does not fill the box.
const HOST_BACKDROP: COLORREF = COLORREF(0x0020_2020);

/// Grace period before allowing image-name fallback for brokered/packaged apps whose launcher
/// process has not exited yet.
const BROKER_FALLBACK_GRACE: Duration = Duration::from_millis(150);

/// Clock slack (2s in 100ns FILETIME ticks) when validating brokered process creation time.
const PROCESS_TIME_SLACK_100NS: u64 = 2 * 10_000_000;

/// Fallback off-screen parking coordinate when positive off-screen X is clamped by the target app.
/// Kept well inside i16 (-32768..32767) so WM_MOVE lParam never wraps.
const FALLBACK_PARKING: i32 = -20000;

/// Off-screen staging coordinate for `CustomCaptureTechnique::OffscreenStaged`.
const STAGE_COORD: i32 = -10_000;

/// 1/255 alpha: keeps DWM compositing the source for `DwmRegisterThumbnail` while invisible on screen.
const INVISIBLE_ALPHA: u8 = 1;
const DIRECT_EX_STYLE: u32 = WS_EX_TRANSPARENT.0 | WS_EX_TOOLWINDOW.0 | WS_EX_NOACTIVATE.0;
const INVISIBLE_EX_STYLE: u32 = WS_EX_LAYERED.0 | DIRECT_EX_STYLE;

/// What a hover of a custom preview is asked to be drawn at: the pointer's own placement, or the
/// focused item's.
///
/// It is the whole of what the hover knows about *where* the preview goes, and it is read on the
/// worker thread rather than here because the box cannot be worked out before the program's
/// window exists: the size the box is cut from is the file's own where it declares one, and the
/// window's only where it does not (see the module note).
#[derive(Clone, Copy)]
pub(super) enum Where {
    /// A pointer's hover: where the hand is, and the region the `Avoid` setting keeps a preview
    /// off.
    Pointer {
        cursor: (i32, i32),
        avoid: Option<AvoidRegion>,
    },
    /// A keyboard hover: the focused item's box, the region kept off, and whether its text is
    /// drawn as a row of its view (see `KeyboardPlacement`).
    Item {
        rect: ScreenRegion,
        avoid: Option<ScreenRegion>,
        columns: bool,
    },
}

/// Where the hover on screen is to be drawn. `None` is no hover.
struct Want {
    at: Where,
}

/// Per-preview Job Object with `KILL_ON_JOB_CLOSE` while previewing; `detach()` clears the flag
/// before closing so promoted windows outlive `rust-hover-preview.exe`.
struct PreviewJob {
    handle: HANDLE,
}

// SAFETY: Win32 Job Object handles are process-wide and safe to transfer across threads.
unsafe impl Send for PreviewJob {}

impl PreviewJob {
    fn create_and_assign(pid: u32) -> Option<Self> {
        unsafe {
            let job = CreateJobObjectW(None, PCWSTR::null()).ok()?;
            let mut limits = JOBOBJECT_EXTENDED_LIMIT_INFORMATION::default();
            limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
            if SetInformationJobObject(
                job,
                JobObjectExtendedLimitInformation,
                std::ptr::addr_of!(limits).cast(),
                std::mem::size_of::<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>() as u32,
            )
            .is_err()
            {
                let _ = CloseHandle(job);
                return None;
            }
            let Ok(proc_h) = OpenProcess(PROCESS_SET_QUOTA | PROCESS_TERMINATE, false, pid) else {
                let _ = CloseHandle(job);
                return None;
            };
            let assigned = AssignProcessToJobObject(job, proc_h).is_ok();
            let _ = CloseHandle(proc_h);
            if !assigned {
                let _ = CloseHandle(job);
                return None;
            }
            Some(Self { handle: job })
        }
    }

    /// Clear `KILL_ON_JOB_CLOSE` before closing so the process survives independently on promote.
    fn detach(mut self) -> bool {
        if self.handle.is_invalid() || self.handle.0.is_null() {
            return false;
        }
        let limits = JOBOBJECT_EXTENDED_LIMIT_INFORMATION {
            BasicLimitInformation:
                windows::Win32::System::JobObjects::JOBOBJECT_BASIC_LIMIT_INFORMATION {
                    LimitFlags: JOB_OBJECT_LIMIT(0),
                    ..Default::default()
                },
            ..Default::default()
        };
        let cleared = unsafe {
            SetInformationJobObject(
                self.handle,
                JobObjectExtendedLimitInformation,
                std::ptr::addr_of!(limits).cast(),
                std::mem::size_of::<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>() as u32,
            )
        }
        .is_ok();
        let _ = unsafe { CloseHandle(self.handle) };
        self.handle = HANDLE::default();
        cleared
    }
}

impl Drop for PreviewJob {
    fn drop(&mut self) {
        if !self.handle.is_invalid() && !self.handle.0.is_null() {
            let _ = unsafe { CloseHandle(self.handle) };
            self.handle = HANDLE::default();
        }
    }
}

/// The program this app started, while it is up: the file it was begun for, the program itself,
/// and the process it is.
///
/// It is not the hover's: it outlives one by [`LINGER`], and it is what a hover that comes back
/// finds rather than a second launch (see the module note).
struct Program {
    path: PathBuf,
    exe: String,
    process: Arc<OwnedHandle>,
    pid: u32,
    job: Option<PreviewJob>,
}

/// What one launch is: the process it began, the image that program runs under, and the windows
/// and processes that were on the machine before it — which is what tells the program's own
/// window from everything else, and the process it may end from the ones that are the user's.
struct Launch {
    /// The file the program was begun for: the size the box is cut to is the file's own, and it is
    /// read on the worker rather than here (see `base_size`).
    path: PathBuf,
    /// Intrinsic file dimensions pre-read before launch when known, avoiding a duplicate header read.
    initial_own: Option<(u32, u32)>,
    /// Provisional target rect `(x, y, w, h)` known at launch, used for pre-positioning hidden windows.
    provisional: Option<(i32, i32, i32, i32)>,
    /// True when launched exe resides under a standard Windows `WindowsApps` MSIX/alias directory.
    packaged_broker: bool,
    /// Retained handle to the spawned child process for `WaitForInputIdle`.
    process: Arc<OwnedHandle>,
    /// The process this app began.
    pid: u32,
    /// The preset program's own file name, lowercased: what a process the program hands its file
    /// to is running, however it was begun.
    exe: String,
    /// Foreground window before launch (e.g. Explorer), restored if the spawned app steals focus.
    previous_foreground: isize,
    /// Every top-level window there was before the launch, so that the program's own are the ones
    /// that were not.
    windows_before: Vec<isize>,
    /// Every process there was before the launch, so that one the program started is the only one
    /// this app may end (see `end_the_program`).
    processes_before: Vec<u32>,
    /// System FILETIME (100ns ticks) captured immediately before `CreateProcessW`.
    launch_filetime: u64,
    /// Monotonic timestamp captured immediately before `CreateProcessW`.
    launched_at: Instant,
    /// Which launch this is, so that a worker begun for an earlier one stands down.
    id: u64,
}

/// The host window, made once and kept off the screen between hovers.
static HOST: AtomicIsize = AtomicIsize::new(0);
/// Whether the host is on the screen.
static SHOWN: AtomicBool = AtomicBool::new(false);
/// True while `try_detach_active_preview` is promoting `SOURCE` behind `HOST`; suppresses premature `hide_the_host()`.
static DETACHING: AtomicBool = AtomicBool::new(false);
/// The program's window registered into it, while one is.
static SOURCE: AtomicIsize = AtomicIsize::new(0);
/// The process that window belongs to, which is not the launched one where the program is
/// packaged.
static SOURCE_PID: AtomicU32 = AtomicU32::new(0);
/// Whether that process is one the launch made, which is the only kind this app ends.
static SOURCE_IS_NEW: AtomicBool = AtomicBool::new(false);
/// True while active `SOURCE` is in `PlacementMode::InvisibleLayered` (1/255 alpha on-monitor).
static SOURCE_INVISIBLE_LAYERED: AtomicBool = AtomicBool::new(false);
/// Saved on-screen rect for `OffscreenStaged` reused non-owned windows so teardown never leaves them at `STAGE_COORD`.
static SOURCE_RESTORE_RECT: Mutex<Option<RECT>> = Mutex::new(None);
/// Active `InProcessHook` shared-memory control handle so `hide()` cloaks synchronously before hiding `host`.
static ACTIVE_HOOK_CONTROL: Mutex<Option<Arc<HookSharedView>>> = Mutex::new(None);
/// Last screen rect `(x, y, w, h)` of the active preview host, used by `open_fully` to promote in-place.
static LAST_HOST_RECT: Mutex<Option<(i32, i32, i32, i32)>> = Mutex::new(None);
/// Saved `GWL_STYLE` & `GWL_EXSTYLE` of `SOURCE` before preview styling so promote restores a normal window.
static SOURCE_ORIGINAL_STYLE: Mutex<Option<isize>> = Mutex::new(None);
static SOURCE_ORIGINAL_EX_STYLE: Mutex<Option<isize>> = Mutex::new(None);
static ACTIVE_CAPTURE_TECHNIQUE: Mutex<CustomCaptureTechnique> =
    Mutex::new(CustomCaptureTechnique::InProcessHook);
static ACTIVE_WINDOW_MODE: Mutex<CustomWindowMode> = Mutex::new(CustomWindowMode::Normal);
/// The thumbnail registration the Shell is drawing through, while one stands.
static THUMBNAIL: Mutex<isize> = Mutex::new(0);
/// The hover on screen, while there is one.
static WANT: Mutex<Option<Want>> = Mutex::new(None);
/// The program this app started for it, which outlives the hover by [`LINGER`].
static PROGRAM: Mutex<Option<Program>> = Mutex::new(None);
/// How many programs have been begun: a worker that was begun for an earlier one is one whose
/// program has been ended, and it neither places nor shows anything.
static LAUNCH: AtomicU64 = AtomicU64::new(0);
static CLASS_REGISTERED: Once = Once::new();

/// Whether a custom preview is on screen.
pub(super) fn is_showing() -> bool {
    WANT.lock().map(|want| want.is_some()).unwrap_or(false)
}

/// Active hovered custom preview path (`WANT` is `Some` and `PROGRAM` is held).
pub(super) fn active_path() -> Option<PathBuf> {
    let wanting = WANT.lock().ok().is_some_and(|want| want.is_some());
    if !wanting {
        return None;
    }
    PROGRAM
        .lock()
        .ok()
        .and_then(|held| held.as_ref().map(|p| p.path.clone()))
}

/// Whether a hovered custom preview is currently active.
#[allow(dead_code)]
pub(super) fn is_active() -> bool {
    active_path().is_some()
}

/// Pin message for an active hovered custom preview, routing through `PinCold` -> `open_fully`.
pub(super) fn pin_request_if_active() -> Option<super::PreviewMessage> {
    active_path().map(|path| super::PreviewMessage::PinCold { path })
}

/// Show `path` in the program its preset names, at the box that preset's scale and the tray's
/// placement ask for.
///
/// It answers whether a custom preview is what this hover is: `false` means no preset claims the
/// file and the caller goes on to lay out and load it the way it always would.
pub(super) fn show(path: &Path, at: Where) -> bool {
    let Some(preset) = preset_for(path) else {
        // A file no preset claims is not this road's, and one a preset no longer claims — a
        // preset deleted while its preview was up — takes the preview down with it.
        if is_showing() {
            hide();
        }
        return false;
    };

    let Some(exe) = preset.exe_path().and_then(resolved) else {
        // A preset whose program is not there is one no preview can be made of, and the hover is
        // left to whatever else this app would have done with the file.
        hide();
        return false;
    };

    // **The same program is not begun twice.** A hover of this file that this app's program is
    // already drawing is put back up wherever the hand now is — which is the whole of what keeps
    // the program's own first window from being a launch loop (see the module note).
    let the_same_program = PROGRAM
        .lock()
        .ok()
        .and_then(|held| {
            held.as_ref()
                .map(|program| (program.path.clone(), program.exe.clone()))
        })
        .is_some_and(|(path_of, exe_of)| {
            path_of.as_path() == path && exe_of.eq_ignore_ascii_case(&exe)
        });
    if the_same_program {
        if let Ok(mut want) = WANT.lock() {
            *want = Some(Want { at });
        }
        return true;
    }

    // A different file: the program that drew the last hover is not the one this one is, and it
    // goes now rather than at the end of its linger.
    end_the_program();

    // **The desktop is read before the program is begun**, so that the windows it puts up are the
    // ones that were not there, and the process it hands the file to is one this app may end.
    let previous_foreground = unsafe { GetForegroundWindow() }.0 as isize;
    let windows_before = top_level_windows();
    let processes_before = engine_processes::running().unwrap_or_default();

    let scaling = preset.scaling;
    let window_mode = preset.window_mode;
    let capture_technique = preset.capture_technique;
    let volume = preset.volume_enabled.then_some(preset.volume);
    let initial_own = base_size(path);
    let provisional = initial_own.and_then(|own| {
        let scale = resolve_scale(path, scaling);
        provisional_rect_for(at, own, scale, follow_cursor(), bypass_placement())
    });

    let launch_filetime = current_filetime_u64();
    let launched_at = Instant::now();
    let (process, pid, hook_session, capture_technique) = if capture_technique
        == CustomCaptureTechnique::InProcessHook
    {
        let Some((proc_handle, pid, session)) = launch_in_process_hook(&exe, path, provisional)
        else {
            return false;
        };
        let effective = if session.is_some() {
            CustomCaptureTechnique::InProcessHook
        } else {
            CustomCaptureTechnique::InvisibleLayered
        };
        (proc_handle, pid, session, effective)
    } else {
        let Some((proc_handle, pid)) = launch(&exe, path, provisional, capture_technique) else {
            return false;
        };
        (proc_handle, pid, None, capture_technique)
    };
    let process = Arc::new(process);
    let Some(host) = (unsafe { ensure_host() }) else {
        drop(hook_session);
        let _ = unsafe { TerminateProcess(HANDLE(process.as_raw_handle()), 1) };
        return false;
    };

    let exe_file = Path::new(&exe)
        .file_name()
        .map(|name| name.to_string_lossy().to_lowercase())
        .unwrap_or_default();
    let packaged_broker = is_packaged_broker_path(Path::new(&exe));
    // Per-preview detachable Job Object: crash-reaps while previewing, detaches cleanly on promote.
    let job = PreviewJob::create_and_assign(pid);
    engine_processes::record_player(&exe, pid);
    if let Ok(mut held) = PROGRAM.lock() {
        *held = Some(Program {
            path: path.to_path_buf(),
            exe: exe.clone(),
            process: Arc::clone(&process),
            pid,
            job,
        });
    }
    if let Ok(mut want) = WANT.lock() {
        *want = Some(Want { at });
    }
    if let Ok(mut slot) = ACTIVE_CAPTURE_TECHNIQUE.lock() {
        *slot = capture_technique;
    }
    if let Ok(mut slot) = ACTIVE_WINDOW_MODE.lock() {
        *slot = window_mode;
    }

    // Read after the launch is written down, so a worker begun for it is current exactly when it
    // is the program on screen.
    let id = LAUNCH.fetch_add(1, Ordering::SeqCst) + 1;
    let launch = Launch {
        path: path.to_path_buf(),
        initial_own,
        provisional,
        packaged_broker,
        process,
        pid,
        exe: exe_file,
        previous_foreground,
        windows_before,
        processes_before,
        launch_filetime,
        launched_at,
        id,
    };
    // The window is carried as a number: `HWND` is a pointer and a pointer is not `Send`, and
    // what the worker does with it is a placement the Shell answers for either way.
    let host = host.0 as isize;
    std::thread::spawn(move || {
        run_it(
            HWND(host as *mut _),
            launch,
            scaling,
            window_mode,
            capture_technique,
            volume,
            hook_session,
        )
    });
    true
}

/// Take the hover down: the host leaves the screen, and the program is kept a moment in case the
/// hover comes straight back (see [`LINGER`]).
///
/// **Nothing is killed here.** A hide that killed the program would make every flicker of the
/// hover a launch, and the program's own window is what makes it flicker.
pub(super) fn hide() {
    if let Ok(mut want) = WANT.lock() {
        *want = None;
    }
    let hook_ctrl = ACTIVE_HOOK_CONTROL.lock().ok().and_then(|g| g.clone());
    if let Some(ctrl) = hook_ctrl {
        let source = SOURCE.load(Ordering::SeqCst);
        ctrl.stage_offscreen();
        let cloaked = source != 0 && ctrl.cloak_sync(HWND(source as *mut _));
        if !cloaked {
            if SOURCE_INVISIBLE_LAYERED.load(Ordering::SeqCst) {
                keep_source_invisible(source);
            } else if source != 0 {
                let hwnd = HWND(source as *mut _);
                if unsafe { IsWindow(hwnd) }.as_bool() {
                    let _ = unsafe {
                        SetWindowPos(
                            hwnd,
                            HWND_BOTTOM,
                            STAGE_COORD,
                            STAGE_COORD,
                            0,
                            0,
                            SWP_NOSIZE | SWP_NOACTIVATE,
                        )
                    };
                }
            }
        }
    } else if SOURCE_INVISIBLE_LAYERED.load(Ordering::SeqCst) {
        keep_source_invisible(SOURCE.load(Ordering::SeqCst));
    } else {
        // Push on-monitor source off-screen before hiding host so it never flashes on desktop during LINGER.
        park_source_offscreen_now();
    }
    hide_the_host();
}

/// Open the file with the program its preset names, as the program's own window — the whole of
/// what the pin key means for a custom preview.
///
/// **Nothing of this app's stands in front of it.** A pin is a window this app draws and walks a
/// folder with; a file a preset claims is not this app's to draw, so a pin of one is the program
/// itself, started the way the listing would start it: its own frame, its own place, its own
/// taskbar button. The preview goes away with it.
///
/// It answers whether a preset claimed the file, so that the caller knows the pin was answered
/// here rather than by a window of this app's.
/// Clamp promoted window `(x, y, w, h)` to monitor work area so it never lands half off-screen.
fn clamp_rect_to_work_area(rect: (i32, i32, i32, i32), work: RECT) -> (i32, i32, i32, i32) {
    let work_w = (work.right - work.left).max(1);
    let work_h = (work.bottom - work.top).max(1);
    let w = rect.2.max(1).min(work_w);
    let h = rect.3.max(1).min(work_h);
    let max_x = (work.right - w).max(work.left);
    let max_y = (work.bottom - h).max(work.top);
    let x = rect.0.clamp(work.left, max_x);
    let y = rect.1.clamp(work.top, max_y);
    (x, y, w, h)
}

fn monitor_rect_for(rect: (i32, i32, i32, i32)) -> RECT {
    let screen = RECT {
        left: rect.0,
        top: rect.1,
        right: rect.0 + rect.2.max(1),
        bottom: rect.1 + rect.3.max(1),
    };
    let hmon = unsafe { MonitorFromRect(&screen, MONITOR_DEFAULTTONEAREST) };
    let mut info = MONITORINFO {
        cbSize: std::mem::size_of::<MONITORINFO>() as u32,
        ..Default::default()
    };
    if !hmon.is_invalid() && unsafe { GetMonitorInfoW(hmon, &mut info) }.as_bool() {
        info.rcMonitor
    } else {
        let bounds = work_area_at(rect.0, rect.1);
        RECT {
            left: bounds.left,
            top: bounds.top,
            right: bounds.right,
            bottom: bounds.bottom,
        }
    }
}

type WindowBox = (i32, i32, i32, i32);

/// Clamp visible `host_rect` to `monitor` (not `rcWork`) BEFORE adding invisible DWM border insets
/// so edge-touching / fit-to-screen previews never shrink by taskbar height or DWM border width.
fn promoted_window_rects(
    host_rect: WindowBox,
    insets: WindowBox,
    monitor: RECT,
) -> (WindowBox, WindowBox) {
    let visible = clamp_rect_to_work_area(host_rect, monitor);
    let outer = behind_host_window_rect(visible, insets);
    (visible, outer)
}

fn visible_rect_matches_target(actual_visible: RECT, target: (i32, i32, i32, i32)) -> bool {
    let aw = actual_visible.right - actual_visible.left;
    let ah = actual_visible.bottom - actual_visible.top;
    (actual_visible.left - target.0).abs() <= 1
        && (actual_visible.top - target.1).abs() <= 1
        && (aw - target.2).abs() <= 1
        && (ah - target.3).abs() <= 1
}

/// Shift `(dx, dy)` to keep `actual_visible` inside `bounds` without shrinking or clipping DWM borders.
fn visible_clamp_delta(actual_visible: RECT, bounds: RECT) -> (i32, i32) {
    let w = (actual_visible.right - actual_visible.left).max(1);
    let h = (actual_visible.bottom - actual_visible.top).max(1);
    let (fx, fy, _, _) =
        clamp_rect_to_work_area((actual_visible.left, actual_visible.top, w, h), bounds);
    (fx - actual_visible.left, fy - actual_visible.top)
}

struct DetachGuard;
impl Drop for DetachGuard {
    fn drop(&mut self) {
        DETACHING.store(false, Ordering::SeqCst);
    }
}

/// Raise `source` to top of normal Z-order (strictly above Explorer, strictly below `host` if visible).
/// Uses TOPMOST -> NOTOPMOST so Win32 lifts a HWND_BOTTOM window above Explorer even when foreground lock rejects SetForegroundWindow.
unsafe fn raise_source_above_explorer(
    source_hwnd: HWND,
    host_hwnd: Option<HWND>,
    x: i32,
    y: i32,
    w: i32,
    h: i32,
    extra_flags: windows::Win32::UI::WindowsAndMessaging::SET_WINDOW_POS_FLAGS,
) {
    let insert_after = match host_hwnd {
        Some(host) if IsWindowVisible(host).as_bool() => {
            let _ = SetWindowPos(
                host,
                HWND_TOPMOST,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );
            host
        }
        _ => HWND_TOPMOST,
    };
    let _ = SetWindowPos(
        source_hwnd,
        insert_after,
        x,
        y,
        w,
        h,
        extra_flags | SWP_SHOWWINDOW | SWP_NOACTIVATE,
    );
    let _ = SetWindowPos(
        source_hwnd,
        HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE,
    );
}

/// Detach & promote an active live custom preview for `path` in-place; returns `true` if detached.
fn try_detach_active_preview(path: &Path) -> bool {
    let matches_active = PROGRAM
        .lock()
        .ok()
        .is_some_and(|held| held.as_ref().is_some_and(|p| p.path.as_path() == path));
    let source_raw = SOURCE.load(Ordering::SeqCst);
    if !matches_active || source_raw == 0 {
        return false;
    }
    let source_hwnd = HWND(source_raw as *mut _);
    if !unsafe { IsWindow(source_hwnd) }.as_bool() {
        return false;
    }

    // Hold DETACHING until `source` is positioned, raised, and DWM-flushed behind `host` so neither
    // `run_it()` nor `hide()` can hide `host` early and cause a flash.
    DETACHING.store(true, Ordering::SeqCst);
    let detach_guard = DetachGuard;

    let host_hwnd = {
        let raw = HOST.load(Ordering::SeqCst);
        let hwnd = HWND(raw as *mut _);
        (raw != 0 && unsafe { IsWindow(hwnd) }.as_bool()).then_some(hwnd)
    };

    // Bump LAUNCH first so `run_it` stops touching `source` and drops `InProcessHookSession`.
    LAUNCH.fetch_add(1, Ordering::SeqCst);
    if let Ok(mut want) = WANT.lock() {
        *want = None;
    }
    if let Some(mut program) = PROGRAM.lock().ok().and_then(|mut held| held.take()) {
        if let Some(job) = program.job.take() {
            let _ = job.detach();
        }
        engine_processes::forget(program.pid);
    }

    SOURCE.store(0, Ordering::SeqCst);
    let mut source_pid = SOURCE_PID.swap(0, Ordering::SeqCst);
    if source_pid == 0 {
        let _ = unsafe { GetWindowThreadProcessId(source_hwnd, Some(&mut source_pid)) };
    }
    SOURCE_IS_NEW.store(false, Ordering::SeqCst);
    let was_invisible_layered = SOURCE_INVISIBLE_LAYERED.swap(false, Ordering::SeqCst);
    if let Ok(mut r) = SOURCE_RESTORE_RECT.lock() {
        *r = None;
    }
    let hook_ctrl = ACTIVE_HOOK_CONTROL.lock().ok().and_then(|mut g| g.take());
    let host_rect = LAST_HOST_RECT.lock().ok().and_then(|mut g| g.take());
    let orig_style = SOURCE_ORIGINAL_STYLE.lock().ok().and_then(|mut g| g.take());
    let orig_ex_style = SOURCE_ORIGINAL_EX_STYLE
        .lock()
        .ok()
        .and_then(|mut g| g.take());
    let capture_technique = ACTIVE_CAPTURE_TECHNIQUE
        .lock()
        .map(|g| *g)
        .unwrap_or(CustomCaptureTechnique::InProcessHook);
    let window_mode = ACTIVE_WINDOW_MODE
        .lock()
        .map(|g| *g)
        .unwrap_or(CustomWindowMode::Normal);

    // Detach hook subclass BEFORE SetWindowPos so rhp_subclass_proc never overrides Z-order or rect.
    if let Some(ctrl) = hook_ctrl {
        let _ = ctrl.detach_sync(source_hwnd);
    }
    if window_mode == CustomWindowMode::Borderless {
        if let Some(style) = orig_style {
            let _ = unsafe { SetWindowLongPtrW(source_hwnd, GWL_STYLE, style) };
        }
    }
    if was_invisible_layered {
        let _ = unsafe { SetLayeredWindowAttributes(source_hwnd, COLORREF(0), 255, LWA_ALPHA) };
        let target_ex = orig_ex_style.unwrap_or_else(|| {
            let cur = unsafe { GetWindowLongPtrW(source_hwnd, GWL_EXSTYLE) } as u32;
            (cur & !INVISIBLE_EX_STYLE) as isize
        });
        let _ = unsafe { SetWindowLongPtrW(source_hwnd, GWL_EXSTYLE, target_ex) };
        let _ = unsafe {
            SetWindowPos(
                source_hwnd,
                HWND::default(),
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        };
    } else if capture_technique == CustomCaptureTechnique::DirectWindow
        || window_mode == CustomWindowMode::Borderless
    {
        let target_ex = orig_ex_style.unwrap_or_else(|| {
            let cur = unsafe { GetWindowLongPtrW(source_hwnd, GWL_EXSTYLE) } as u32;
            (cur & !DIRECT_EX_STYLE) as isize
        });
        let _ = unsafe { SetWindowLongPtrW(source_hwnd, GWL_EXSTYLE, target_ex) };
        let _ = unsafe {
            SetWindowPos(
                source_hwnd,
                HWND::default(),
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        };
    }

    if let Some(hr) = host_rect {
        let monitor = monitor_rect_for(hr);
        let insets = dwm_frame_insets(source_hwnd);
        let (visible_target, (ax, ay, aw, ah)) = promoted_window_rects(hr, insets, monitor);
        unsafe {
            raise_source_above_explorer(
                source_hwnd,
                host_hwnd,
                ax,
                ay,
                aw,
                ah,
                windows::Win32::UI::WindowsAndMessaging::SET_WINDOW_POS_FLAGS(0),
            );
        }
        // Re-apply once on target monitor if WM_DPICHANGED resized `source` or DWM insets changed
        // after moving off STAGE_COORD / restoring frame styles.
        let insets_after = dwm_frame_insets(source_hwnd);
        let needs_reapply = insets_after != insets
            || visible_frame_rect(source_hwnd)
                .is_some_and(|vis| !visible_rect_matches_target(vis, visible_target));
        if needs_reapply {
            let (ax2, ay2, aw2, ah2) = behind_host_window_rect(visible_target, insets_after);
            let _ = unsafe {
                SetWindowPos(
                    source_hwnd,
                    HWND::default(),
                    ax2,
                    ay2,
                    aw2,
                    ah2,
                    SWP_NOZORDER | SWP_NOACTIVATE,
                )
            };
        }
        // Slide top-left only if app's WM_GETMINMAXINFO pushed visible bounds past monitor edges.
        if let (Some(vis), Some(actual)) =
            (visible_frame_rect(source_hwnd), window_rect(source_hwnd))
        {
            let (dx, dy) = visible_clamp_delta(vis, monitor);
            if dx != 0 || dy != 0 {
                let _ = unsafe {
                    SetWindowPos(
                        source_hwnd,
                        HWND::default(),
                        actual.left + dx,
                        actual.top + dy,
                        0,
                        0,
                        SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE,
                    )
                };
            }
        }
    } else if is_window_offscreen(source_hwnd) {
        let bounds = work_area_at(0, 0);
        let work = RECT {
            left: bounds.left,
            top: bounds.top,
            right: bounds.right,
            bottom: bounds.bottom,
        };
        let (aw, ah) = window_rect(source_hwnd)
            .map(|r| ((r.right - r.left).max(1), (r.bottom - r.top).max(1)))
            .unwrap_or((800, 600));
        let (fx, fy, _, _) = clamp_rect_to_work_area((work.left, work.top, aw, ah), work);
        unsafe {
            raise_source_above_explorer(source_hwnd, host_hwnd, fx, fy, 0, 0, SWP_NOSIZE);
        }
    } else {
        unsafe {
            raise_source_above_explorer(
                source_hwnd,
                host_hwnd,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE,
            );
        }
    }

    // Flush paint & DWM composition while `host` still covers `source` for a zero-flicker reveal.
    let _ = unsafe { InvalidateRect(source_hwnd, None, false) };
    let _ = unsafe { UpdateWindow(source_hwnd) };
    let _ = unsafe { DwmFlush() };

    drop(detach_guard);
    hide_the_host();
    if let Some(host) = host_hwnd {
        let _ = unsafe {
            SendMessageTimeoutW(
                host,
                WM_THERE_IS_NOTHING_TO_SHOW,
                WPARAM(0),
                LPARAM(0),
                SMTO_ABORTIFHUNG,
                50,
                None,
            )
        };
    }
    let thumbnail = THUMBNAIL
        .lock()
        .map(|mut held| std::mem::replace(&mut *held, 0))
        .unwrap_or(0);
    if thumbnail != 0 {
        let _ = unsafe { DwmUnregisterThumbnail(thumbnail) };
    }

    if source_pid != 0 {
        let _ = unsafe { AllowSetForegroundWindow(source_pid) };
    }
    let _ = unsafe { BringWindowToTop(source_hwnd) };
    let _ = unsafe { SetForegroundWindow(source_hwnd) };
    true
}

fn clear_want() {
    if let Ok(mut want) = WANT.lock() {
        *want = None;
    }
}

/// Spawn a completely untouched, detached process for `exe` + `path` (no job, no hooks, no STARTUPINFO).
fn spawn_detached_normal(exe: &str, path: &Path) {
    let mut cmd = Command::new(exe);
    cmd.arg(path);
    if let Some(dir) = Path::new(exe)
        .parent()
        .filter(|d| !d.as_os_str().is_empty() && d.is_dir())
    {
        cmd.current_dir(dir);
    }
    let _ = cmd.spawn();
}

fn dispatch_open_fully<D, S>(
    technique: CustomPromotionTechnique,
    path: &Path,
    exe: &str,
    try_detach: D,
    spawn_new: S,
) -> bool
where
    D: FnOnce(&Path) -> bool,
    S: FnOnce(&str, &Path),
{
    match technique {
        CustomPromotionTechnique::DirectPromotion => {
            if try_detach(path) {
                return true;
            }
            end_the_program();
            spawn_new(exe, path);
            true
        }
        CustomPromotionTechnique::OpenAsNewProcess => {
            clear_want();
            hide_the_host();
            end_the_program();
            spawn_new(exe, path);
            true
        }
    }
}

pub(super) fn open_fully(path: &Path) -> bool {
    let Some(preset) = preset_for(path) else {
        return false;
    };
    let Some(exe) = preset.exe_path().map(str::to_string) else {
        return false;
    };

    dispatch_open_fully(
        preset.promotion_technique,
        path,
        &exe,
        try_detach_active_preview,
        spawn_detached_normal,
    )
}

/// End the program, its window's registration, and the hover it was drawing.
fn end_the_program() {
    // First, so that a worker still waiting on the program's window stops placing and showing.
    LAUNCH.fetch_add(1, Ordering::SeqCst);

    if let Ok(mut want) = WANT.lock() {
        *want = None;
    }

    let program = PROGRAM.lock().ok().and_then(|mut held| held.take());
    let launched_pid = program.as_ref().map(|program| program.pid).unwrap_or(0);
    if let Some(program) = program {
        let _ = unsafe { TerminateProcess(HANDLE(program.process.as_raw_handle()), 1) };
        engine_processes::forget(program.pid);
    }

    // **The window this app asked for is closed, and the process that opened it is ended only
    // where the launch is what made it.** A packaged program's window belongs to a process this
    // app never started, and that process may be one the user had open before the hover: its
    // window is the document this app asked for, so the window goes and the process stays.
    let source = SOURCE.swap(0, Ordering::SeqCst);
    let source_pid = SOURCE_PID.swap(0, Ordering::SeqCst);
    let source_is_new = SOURCE_IS_NEW.swap(false, Ordering::SeqCst);
    SOURCE_INVISIBLE_LAYERED.store(false, Ordering::SeqCst);
    if let Ok(mut ctrl) = ACTIVE_HOOK_CONTROL.lock() {
        *ctrl = None;
    }
    if let Ok(mut r) = LAST_HOST_RECT.lock() {
        *r = None;
    }
    if let Ok(mut s) = SOURCE_ORIGINAL_STYLE.lock() {
        *s = None;
    }
    if let Ok(mut ex) = SOURCE_ORIGINAL_EX_STYLE.lock() {
        *ex = None;
    }
    let restore_rect = SOURCE_RESTORE_RECT.lock().ok().and_then(|mut r| r.take());
    if source != 0 {
        let source = HWND(source as *mut _);
        if unsafe { IsWindow(source) }.as_bool() {
            if let Some(target) = restore_rect {
                if window_rect(source).is_some_and(|cur| {
                    should_restore_on_teardown(source_is_new, cur, virtual_screen())
                }) {
                    let _ = unsafe {
                        SetWindowPos(
                            source,
                            HWND_BOTTOM,
                            target.left,
                            target.top,
                            (target.right - target.left).max(1),
                            (target.bottom - target.top).max(1),
                            SWP_NOACTIVATE | SWP_NOZORDER,
                        )
                    };
                }
            }
            let _ = unsafe { PostMessageW(source, WM_CLOSE, WPARAM(0), LPARAM(0)) };
        }
    }
    if source_is_new && source_pid != 0 && source_pid != launched_pid {
        end_process(source_pid);
    }

    let thumbnail = THUMBNAIL
        .lock()
        .map(|mut held| std::mem::replace(&mut *held, 0))
        .unwrap_or(0);
    if thumbnail != 0 {
        let _ = unsafe { DwmUnregisterThumbnail(thumbnail) };
    }
    hide_the_host();
}

/// End a process this app did not begin: a program that hands its file to a process of its own is
/// one whose window outlives the launcher, and it is the window that is the preview.
fn end_process(pid: u32) {
    let Ok(process) = (unsafe { OpenProcess(PROCESS_TERMINATE, false, pid) }) else {
        return;
    };
    let _ = unsafe { TerminateProcess(process, 1) };
    let _ = unsafe { CloseHandle(process) };
}

/// Take the host off the screen, on the thread that made it.
fn hide_the_host() {
    if DETACHING.load(Ordering::SeqCst) {
        return;
    }
    SHOWN.store(false, Ordering::SeqCst);

    let host = HOST.load(Ordering::SeqCst);
    if host == 0 {
        return;
    }
    let host = HWND(host as *mut _);
    if unsafe { IsWindow(host) }.as_bool() {
        let _ = unsafe { PostMessageW(host, WM_THERE_IS_NOTHING_TO_SHOW, WPARAM(0), LPARAM(0)) };
    }
}

/// The preset that claims `path`, read from the running configuration.
fn preset_for(path: &Path) -> Option<CustomPreset> {
    CONFIG
        .lock()
        .ok()
        .and_then(|config| custom_preset_for(&config, path).cloned())
}

/// Whether the tray's `Placement → Position` follows the pointer rather than choosing the
/// roomiest side of it, which is the one setting a placement is made by.
fn follow_cursor() -> bool {
    CONFIG
        .lock()
        .map(|config| config.follow_cursor)
        .unwrap_or(true)
}

/// Whether the tray's `Scaling → Bypass Placement` is on: the box then takes the whole room on
/// the cursor's free side, whatever the preset's own scale asks, as every other preview's does.
fn bypass_placement() -> bool {
    CONFIG
        .lock()
        .map(|config| config.bypass_placement)
        .unwrap_or(false)
}

/// The program as the desktop's own loader wants it: **a full path**.
///
/// `CreateProcessW` is handed the program as `lpApplicationName`, and a name there is a place and
/// never a search — so a preset that says `mspaint.exe` is a launch the desktop refuses, which is
/// a preview with nothing on screen. The search `Command` would have done is done here instead,
/// so a preset that names a program is a preset that runs.
fn resolved(exe: &str) -> Option<String> {
    // **Trimmed of the quotes a paste brings.** The field is written as it was pasted — Ctrl+Shift+C
    // hands over `"C:\Program Files\IrfanView\i_view64.exe"` — and nothing compared here is a path
    // while the quotes are still around it.
    let exe = exe.trim().trim_matches('"');
    let named = Path::new(exe);
    if named.is_file() {
        return Some(exe.to_string());
    }
    // A path with a folder in it is a path: one that is not there is not a program to look for.
    if named.components().count() > 1 {
        return None;
    }

    let file = if named.extension().is_some() {
        exe.to_string()
    } else {
        format!("{exe}.exe")
    };
    let path = std::env::var_os("PATH")?;
    std::env::split_paths(&path)
        .map(|folder| folder.join(&file))
        .find(|candidate| candidate.is_file())
        .map(|candidate| candidate.to_string_lossy().into_owned())
}

fn wide_null(s: &OsStr) -> Vec<u16> {
    s.encode_wide().chain(std::iter::once(0)).collect()
}

/// Quote one argument per Win32 CRT / `CommandLineToArgvW` rules (always wrapped in `"`).
fn quote_win32_arg(arg: &OsStr, out: &mut Vec<u16>) {
    const QUOTE: u16 = b'"' as u16;
    const BACKSLASH: u16 = b'\\' as u16;

    out.push(QUOTE);
    let mut backslashes = 0usize;
    for unit in arg.encode_wide() {
        if unit == BACKSLASH {
            backslashes += 1;
        } else if unit == QUOTE {
            // Double preceding backslashes + 1 to escape the embedded quote.
            for _ in 0..(backslashes * 2 + 1) {
                out.push(BACKSLASH);
            }
            out.push(QUOTE);
            backslashes = 0;
        } else {
            for _ in 0..backslashes {
                out.push(BACKSLASH);
            }
            backslashes = 0;
            out.push(unit);
        }
    }
    // Escape trailing backslashes so they do not escape the closing quote.
    for _ in 0..(backslashes * 2) {
        out.push(BACKSLASH);
    }
    out.push(QUOTE);
}

/// Build NUL-terminated UTF-16 command line `"exe" "path"` for `CreateProcessW`.
fn build_command_line(exe: &OsStr, path: &OsStr) -> Vec<u16> {
    let mut line = Vec::new();
    quote_win32_arg(exe, &mut line);
    line.push(b' ' as u16);
    quote_win32_arg(path, &mut line);
    line.push(0);
    line
}

fn provisional_rect_for(
    at: Where,
    own: (u32, u32),
    scale: PreviewScale,
    follow_cursor: bool,
    bypass_placement: bool,
) -> Option<(i32, i32, i32, i32)> {
    let layout = layout_for(at, own, scale, follow_cursor, bypass_placement)?;
    Some((
        layout.pos_x,
        layout.pos_y,
        layout.preview_w as i32,
        layout.preview_h as i32,
    ))
}

/// `STARTUPINFOW` hinting `SW_HIDE` and (when known) provisional preview rect.
fn launch_startup_info(provisional: Option<(i32, i32, i32, i32)>) -> STARTUPINFOW {
    let mut si = STARTUPINFOW {
        cb: std::mem::size_of::<STARTUPINFOW>() as u32,
        dwFlags: STARTF_USESHOWWINDOW,
        wShowWindow: SW_HIDE.0 as u16,
        ..Default::default()
    };
    if let Some((x, y, w, h)) = provisional {
        si.dwFlags |= STARTF_USEPOSITION | STARTF_USESIZE;
        si.dwX = x.max(0) as u32;
        si.dwY = y.max(0) as u32;
        si.dwXSize = w.max(1) as u32;
        si.dwYSize = h.max(1) as u32;
    }
    si
}

/// Safe positive off-screen `STARTUPINFOW` origin (`dwX`/`dwY` are `u32`).
fn offscreen_staged_startup_origin(vscreen: RECT) -> (i32, i32) {
    ((vscreen.right + 2000).max(0), vscreen.top.max(0))
}

/// `STARTUPINFOW` tailored for `capture_technique`; existing 4 techniques delegate unchanged.
fn launch_startup_info_for(
    provisional: Option<(i32, i32, i32, i32)>,
    capture_technique: CustomCaptureTechnique,
    vscreen: RECT,
) -> STARTUPINFOW {
    if !matches!(
        capture_technique,
        CustomCaptureTechnique::OffscreenStaged | CustomCaptureTechnique::InProcessHook
    ) {
        return launch_startup_info(provisional);
    }
    let (sx, sy) = offscreen_staged_startup_origin(vscreen);
    let (w, h) = provisional
        .map(|(_, _, w, h)| (w.max(1), h.max(1)))
        .unwrap_or((800, 600));
    STARTUPINFOW {
        cb: std::mem::size_of::<STARTUPINFOW>() as u32,
        dwFlags: STARTF_USESHOWWINDOW | STARTF_USEPOSITION | STARTF_USESIZE,
        wShowWindow: SW_HIDE.0 as u16,
        dwX: sx as u32,
        dwY: sy as u32,
        dwXSize: w as u32,
        dwYSize: h as u32,
        ..Default::default()
    }
}

/// Spawn target normally via `CreateProcessW` (no `CREATE_NO_WINDOW`, no stdio redirection) with
/// its exe folder as cwd and `STARTUPINFOW` `SW_HIDE` + provisional position/size hints.
fn launch(
    exe: &str,
    path: &Path,
    provisional: Option<(i32, i32, i32, i32)>,
    capture_technique: CustomCaptureTechnique,
) -> Option<(OwnedHandle, u32)> {
    let application = wide_null(OsStr::new(exe));
    let mut cmdline = build_command_line(OsStr::new(exe), path.as_os_str());
    let cwd_wide = Path::new(exe)
        .parent()
        .filter(|dir| !dir.as_os_str().is_empty())
        .map(|dir| wide_null(dir.as_os_str()));
    let cwd_ptr = cwd_wide
        .as_ref()
        .map_or(PCWSTR::null(), |buf| PCWSTR(buf.as_ptr()));
    let si = launch_startup_info_for(provisional, capture_technique, virtual_screen());
    let mut pi = PROCESS_INFORMATION::default();

    unsafe {
        CreateProcessW(
            PCWSTR(application.as_ptr()),
            PWSTR(cmdline.as_mut_ptr()),
            None,
            None,
            BOOL(0),
            PROCESS_CREATION_FLAGS(0),
            None,
            cwd_ptr,
            &si,
            &mut pi,
        )
    }
    .ok()?;
    let _ = unsafe { CloseHandle(pi.hThread) };
    // SAFETY: `pi.hProcess` is uniquely owned by this call and closed on `OwnedHandle` drop.
    let process = unsafe { OwnedHandle::from_raw_handle(pi.hProcess.0) };
    Some((process, pi.dwProcessId))
}

unsafe extern "system" fn host_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_ERASEBKGND => {
            let hdc = windows::Win32::Graphics::Gdi::HDC(wparam.0 as *mut _);
            let mut client = RECT::default();
            let _ = unsafe { GetClientRect(hwnd, &mut client) };
            let brush = unsafe { CreateSolidBrush(HOST_BACKDROP) };
            unsafe {
                FillRect(hdc, &client, brush);
                let _ = DeleteObject(HGDIOBJ(brush.0));
            }
            LRESULT(1)
        }
        WM_THE_PICTURE_IS_READY => {
            let _ = unsafe { ShowWindow(hwnd, SW_SHOWNOACTIVATE) };
            LRESULT(0)
        }
        WM_THERE_IS_NOTHING_TO_SHOW => {
            let _ = unsafe { ShowWindow(hwnd, SW_HIDE) };
            LRESULT(0)
        }
        WM_NCDESTROY => {
            HOST.store(0, Ordering::SeqCst);
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

/// The host, made once and kept off the screen between hovers.
///
/// Safety: makes a window of this process's and asks the desktop about it; the handle is kept in
/// `HOST` for the life of the run.
unsafe fn ensure_host() -> Option<HWND> {
    let existing = HOST.load(Ordering::SeqCst);
    if existing != 0 {
        let host = HWND(existing as *mut _);
        if unsafe { IsWindow(host) }.as_bool() {
            return Some(host);
        }
        HOST.store(0, Ordering::SeqCst);
    }

    CLASS_REGISTERED.call_once(|| {
        let Ok(hinstance) = (unsafe { GetModuleHandleW(None) }) else {
            return;
        };
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(host_proc),
            hInstance: hinstance.into(),
            hbrBackground: HBRUSH(unsafe { GetStockObject(NULL_BRUSH) }.0),
            lpszClassName: HOST_CLASS,
            ..Default::default()
        };
        unsafe { RegisterClassExW(&class) };
    });

    let hinstance = unsafe { GetModuleHandleW(None) }.ok()?;
    // Borderless, never activated, kept above the listing, and skipped by a press
    // (`WindowFromPoint`) because it is disabled: a press on the picture is the hover's press
    // and belongs to the file behind it (see the module doc).
    let host = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE(
                WS_EX_TOOLWINDOW.0 | WS_EX_NOACTIVATE.0 | WS_EX_TOPMOST.0 | WS_EX_TRANSPARENT.0,
            ),
            HOST_CLASS,
            w!("custom preview"),
            WINDOW_STYLE(WS_POPUP.0 | WS_CLIPCHILDREN.0),
            0,
            0,
            1,
            1,
            None,
            None,
            hinstance,
            None,
        )
    }
    .ok()?;

    let disable_transitions: u32 = 1;
    let _ = unsafe {
        DwmSetWindowAttribute(
            host,
            DWMWA_TRANSITIONS_FORCEDISABLED,
            &disable_transitions as *const u32 as *const _,
            std::mem::size_of::<u32>() as u32,
        )
    };
    let _ = unsafe { EnableWindow(host, false) };
    HOST.store(host.0 as isize, Ordering::SeqCst);
    Some(host)
}

/// Whether the program this worker was begun for is still the one on screen.
fn is_current(id: u64) -> bool {
    LAUNCH.load(Ordering::SeqCst) == id
}

/// Find the program's window, keep the box and the picture in step with the hover, and keep the
/// window parked until the hover has been gone for [`LINGER`].
///
/// **One worker per program**, not per hover: the hover can come and go while the program is
/// starting, and what that must cost is a hidden host and a kept window rather than a launch.
fn run_it(
    host: HWND,
    launch: Launch,
    scaling: CustomScaling,
    window_mode: CustomWindowMode,
    capture_technique: CustomCaptureTechnique,
    volume: Option<u32>,
    hook_session: Option<InProcessHookSession>,
) {
    let Some(source) = (if let Some(ref session) = hook_session {
        wait_for_a_window_in_process_hook(&launch, session)
    } else if capture_technique == CustomCaptureTechnique::OffscreenStaged {
        wait_for_a_window_offscreen_staged(&launch)
    } else {
        wait_for_a_window(&launch)
    }) else {
        // A program that made no window is a preview with nothing to show: it goes, and the host
        // is already off the screen.
        if is_current(launch.id) {
            end_the_program();
        }
        return;
    };
    if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
        return;
    }

    if let Some(ref session) = hook_session {
        session.view.set_captured_hwnd(source);
        if let Ok(mut ctrl) = ACTIVE_HOOK_CONTROL.lock() {
            *ctrl = Some(Arc::clone(&session.view));
        }
    }

    let source_pid = pid_of(source);
    let source_is_new = source_pid != 0 && !launch.processes_before.contains(&source_pid);
    SOURCE.store(source.0 as isize, Ordering::SeqCst);
    SOURCE_PID.store(source_pid, Ordering::SeqCst);
    SOURCE_IS_NEW.store(source_is_new, Ordering::SeqCst);
    SOURCE_INVISIBLE_LAYERED.store(false, Ordering::SeqCst);
    if let Ok(mut s) = SOURCE_ORIGINAL_STYLE.lock() {
        *s = Some(unsafe { GetWindowLongPtrW(source, GWL_STYLE) });
    }
    if let Ok(mut ex) = SOURCE_ORIGINAL_EX_STYLE.lock() {
        *ex = Some(unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) });
    }

    // Lock ownership strictly to our launched tree (+ adopted brokered tree if newly spawned).
    let mut processes = owned_pids_for(launch.pid, source_pid, source_is_new);
    if let Some(vol) = volume {
        audio_session::apply_custom_preview_volume(processes.clone(), vol);
    }
    let mut last_volume_apply = Instant::now();

    // Visible app (ignored `SW_HIDE`): make 1/255 invisible BEFORE any style/size/move changes
    // when technique is `InvisibleLayered`.
    let was_hidden = !unsafe { IsWindowVisible(source) }.as_bool();
    let mut invisible_layered = if !was_hidden && tries_invisible_layered(capture_technique) {
        let ok = make_source_invisible(source.0 as isize);
        SOURCE_INVISIBLE_LAYERED.store(ok, Ordering::SeqCst);
        restore_foreground(launch.previous_foreground, &processes);
        ok
    } else {
        false
    };

    // Read intrinsic shape once before any placement mutates `rcNormalPosition`.
    let own = launch.initial_own.unwrap_or_else(|| own_size(source));
    let scale = resolve_scale(&launch.path, scaling);

    // Normal mode keeps target chrome/styles untouched; Borderless strips frame chrome.
    if window_mode == CustomWindowMode::Borderless {
        strip_the_frame(source);
        if invisible_layered {
            keep_source_invisible(source.0 as isize);
        }
    }
    if capture_technique == CustomCaptureTechnique::DirectWindow {
        apply_direct_window_style(source);
    }

    let mut placed_mode: Option<PlacementMode> = None;
    let mut placed_host_rect: Option<(i32, i32, i32, i32)> = None;
    let mut expected_rect: Option<RECT> = None;
    let mut the_ask: Option<(i32, i32)> = None;

    let current_target_rect = || -> Option<(i32, i32, i32, i32)> {
        let at = WANT
            .lock()
            .ok()
            .and_then(|want| want.as_ref().map(|want| want.at))?;
        provisional_rect_for(at, own, scale, follow_cursor(), bypass_placement())
    };

    let mut tree_read = Instant::now();
    let mut images: HashMap<u32, Option<String>> = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();

    let mut box_of_the_host: Option<(i32, i32, i32, i32)> = None;
    let mut registered: isize = 0;

    if let Some(ref session) = hook_session {
        let mut staged = OffscreenStagedState {
            invisible_layered: &mut invisible_layered,
            placed_mode: &mut placed_mode,
            placed_host_rect: &mut placed_host_rect,
            expected_rect: &mut expected_rect,
            the_ask: &mut the_ask,
            box_of_the_host: &mut box_of_the_host,
            registered: &mut registered,
        };
        if !run_in_process_hook_settle_and_reveal(
            host,
            source,
            &launch,
            session,
            &processes,
            &mut images,
            &mut creation_times,
            was_hidden,
            window_mode,
            &current_target_rect,
            &mut staged,
        ) {
            return;
        }
    } else if capture_technique == CustomCaptureTechnique::OffscreenStaged {
        let mut staged = OffscreenStagedState {
            invisible_layered: &mut invisible_layered,
            placed_mode: &mut placed_mode,
            placed_host_rect: &mut placed_host_rect,
            expected_rect: &mut expected_rect,
            the_ask: &mut the_ask,
            box_of_the_host: &mut box_of_the_host,
            registered: &mut registered,
        };
        if !run_offscreen_staged_settle_and_reveal(
            host,
            source,
            &launch,
            &processes,
            &mut images,
            &mut creation_times,
            was_hidden,
            own,
            window_mode,
            &current_target_rect,
            &mut staged,
        ) {
            return;
        }
    } else {
        // Scoped WinEvent hook during settle: re-snaps `source` in <1ms on LOCATIONCHANGE and
        // suppresses owned splash/companion windows on CREATE/SHOW without blocking on WaitForInputIdle.
        WINEVENT_EVENTS.with(|q| q.borrow_mut().clear());
        let settle_hook = unsafe {
            SetWinEventHook(
                EVENT_OBJECT_CREATE,
                EVENT_OBJECT_LOCATIONCHANGE,
                None,
                Some(win_event_proc),
                0,
                0,
                WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
            )
        };
        let settle_guard = WinEventGuard(settle_hook);

        if was_hidden {
            // App honored `SW_HIDE`: atomically position + reveal via `SetWindowPlacement` so it never
            // flashes at its old saved coordinates.
            if let Some(target_rect) = current_target_rect() {
                let mode = match capture_technique {
                    CustomCaptureTechnique::OffScreen => {
                        let vscreen = virtual_screen();
                        let (px, py) = parking_origin(vscreen);
                        let placement = WINDOWPLACEMENT {
                            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
                            flags: WINDOWPLACEMENT_FLAGS(0),
                            showCmd: SW_SHOWNOACTIVATE.0 as u32,
                            rcNormalPosition: RECT {
                                left: px,
                                top: py,
                                right: px + target_rect.2,
                                bottom: py + target_rect.3,
                            },
                            ..Default::default()
                        };
                        let _ = unsafe { SetWindowPlacement(source, &placement) };
                        let ask = the_ask_for(
                            (target_rect.2, target_rect.3),
                            source_size(source, window_mode),
                        );
                        the_ask = Some(ask);
                        park_offscreen(source, ask.0, ask.1);
                        restore_foreground(launch.previous_foreground, &processes);
                        PlacementMode::OffScreen
                    }
                    CustomCaptureTechnique::DirectWindow => {
                        let insets = if window_mode == CustomWindowMode::Normal {
                            dwm_frame_insets(source)
                        } else {
                            (0, 0, 0, 0)
                        };
                        let (wx, wy, ww, wh) = behind_host_window_rect(target_rect, insets);
                        let placement = WINDOWPLACEMENT {
                            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
                            flags: WINDOWPLACEMENT_FLAGS(0),
                            showCmd: SW_SHOWNOACTIVATE.0 as u32,
                            rcNormalPosition: workspace_rect_for_screen(RECT {
                                left: wx,
                                top: wy,
                                right: wx + ww,
                                bottom: wy + wh,
                            }),
                            ..Default::default()
                        };
                        let _ = unsafe { SetWindowPlacement(source, &placement) };
                        place_direct_window(source, target_rect, window_mode);
                        restore_foreground(launch.previous_foreground, &processes);
                        PlacementMode::DirectWindow
                    }
                    CustomCaptureTechnique::InvisibleLayered
                    | CustomCaptureTechnique::BehindHost
                    | CustomCaptureTechnique::OffscreenStaged
                    | CustomCaptureTechnique::InProcessHook => {
                        let insets = if window_mode == CustomWindowMode::Normal {
                            dwm_frame_insets(source)
                        } else {
                            (0, 0, 0, 0)
                        };
                        let (wx, wy, ww, wh) = behind_host_window_rect(target_rect, insets);
                        let workspace_rect = workspace_rect_for_screen(RECT {
                            left: wx,
                            top: wy,
                            right: wx + ww,
                            bottom: wy + wh,
                        });
                        let placement = WINDOWPLACEMENT {
                            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
                            flags: WINDOWPLACEMENT_FLAGS(0),
                            showCmd: SW_SHOWNOACTIVATE.0 as u32,
                            rcNormalPosition: workspace_rect,
                            ..Default::default()
                        };
                        let _ = unsafe { SetWindowPlacement(source, &placement) };

                        if tries_invisible_layered(capture_technique) {
                            invisible_layered = make_source_invisible(source.0 as isize);
                            SOURCE_INVISIBLE_LAYERED.store(invisible_layered, Ordering::SeqCst);
                        }
                        restore_foreground(launch.previous_foreground, &processes);

                        if invisible_layered {
                            place_invisible_layered(source, target_rect, window_mode, &mut the_ask);
                            PlacementMode::InvisibleLayered
                        } else {
                            // Re-read DWM invisible border insets now that `source` is visible for exact alignment.
                            let insets2 = if window_mode == CustomWindowMode::Normal {
                                dwm_frame_insets(source)
                            } else {
                                (0, 0, 0, 0)
                            };
                            let (wx2, wy2, ww2, wh2) =
                                behind_host_window_rect(target_rect, insets2);
                            let _ = unsafe {
                                SetWindowPos(
                                    source,
                                    HWND_BOTTOM,
                                    wx2,
                                    wy2,
                                    ww2,
                                    wh2,
                                    SWP_NOACTIVATE,
                                )
                            };
                            restore_foreground(launch.previous_foreground, &processes);
                            let _ = unsafe { InvalidateRect(source, None, false) };
                            let _ = unsafe { UpdateWindow(source) };

                            if aligned_behind_host(source, target_rect) {
                                PlacementMode::BehindHost
                            } else {
                                let ask = the_ask_for(
                                    (target_rect.2, target_rect.3),
                                    source_size(source, window_mode),
                                );
                                the_ask = Some(ask);
                                park_offscreen(source, ask.0, ask.1);
                                PlacementMode::OffScreen
                            }
                        }
                    }
                };
                placed_mode = Some(mode);
                placed_host_rect = Some(target_rect);
                let _ = drain_settle_events(source, &launch, &processes);
                expected_rect = window_rect(source);
            }
        } else if let Some(target_rect) = current_target_rect() {
            // App ignored `SW_HIDE`: apply the selected capture technique immediately.
            placed_mode = Some(if invisible_layered {
                place_invisible_layered(source, target_rect, window_mode, &mut the_ask);
                restore_foreground(launch.previous_foreground, &processes);
                PlacementMode::InvisibleLayered
            } else {
                snap_source(
                    source,
                    capture_technique,
                    target_rect,
                    window_mode,
                    launch.previous_foreground,
                    &processes,
                    &mut the_ask,
                    placed_mode.zip(placed_host_rect),
                )
            });
            placed_host_rect = Some(target_rect);
            let _ = drain_settle_events(source, &launch, &processes);
            expected_rect = window_rect(source);
        }

        // Immediately suppress any already-visible companion windows strictly in `processes`.
        census(
            &launch,
            &processes,
            &mut images,
            &mut creation_times,
            source.0 as isize,
            true,
            OwnershipMode::OwnedOnly,
        )
        .take();

        let source_proc = (source_pid != 0 && source_pid != launch.pid)
            .then(|| open_process_for_idle(source_pid))
            .flatten();
        let settle_start = Instant::now();
        let mut last_activity = settle_start;
        let settle_deadline = settle_start
            + if was_hidden {
                Duration::from_millis(250)
            } else {
                Duration::from_millis(500)
            };

        while Instant::now() < settle_deadline {
            if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
                return;
            }
            let Some(target_rect) = current_target_rect() else {
                break;
            };

            if invisible_layered {
                keep_source_invisible(source.0 as isize);
            } else if capture_technique == CustomCaptureTechnique::DirectWindow {
                apply_direct_window_style(source);
            }

            // Ignore self-triggered LOCATIONCHANGE from our own SetWindowPos.
            let raw_source_moved = drain_settle_events(source, &launch, &processes);
            let source_moved = raw_source_moved && window_rect(source) != expected_rect;

            let needs_resnap = source_moved
                || placed_host_rect != Some(target_rect)
                || match placed_mode {
                    Some(PlacementMode::InvisibleLayered | PlacementMode::DirectWindow) => false,
                    Some(PlacementMode::BehindHost) => !aligned_behind_host(source, target_rect),
                    Some(PlacementMode::OffScreen) => !parked(source),
                    None => true,
                };
            if needs_resnap {
                if placed_host_rect.is_some_and(|p| p.2 != target_rect.2 || p.3 != target_rect.3) {
                    the_ask = None;
                }
                placed_mode = Some(if invisible_layered {
                    place_invisible_layered(source, target_rect, window_mode, &mut the_ask);
                    restore_foreground(launch.previous_foreground, &processes);
                    PlacementMode::InvisibleLayered
                } else {
                    snap_source(
                        source,
                        capture_technique,
                        target_rect,
                        window_mode,
                        launch.previous_foreground,
                        &processes,
                        &mut the_ask,
                        placed_mode.zip(placed_host_rect),
                    )
                });
                placed_host_rect = Some(target_rect);
                let _ = drain_settle_events(source, &launch, &processes);
                expected_rect = window_rect(source);
                last_activity = Instant::now();
            }

            let launch_idle = is_input_idle_now(HANDLE(launch.process.as_raw_handle()));
            let source_idle = source_proc
                .as_ref()
                .is_none_or(|h| is_input_idle_now(HANDLE(h.as_raw_handle())));
            // Exit settle promptly once BehindHost is aligned so host covers source right away.
            let behind_host_ready = placed_mode == Some(PlacementMode::BehindHost)
                && aligned_behind_host(source, target_rect);
            // Visible InvisibleLayered apps get a short quiet window so startup WindowFit finishes while 1/255 invisible.
            let invisible_quiet = placed_mode == Some(PlacementMode::InvisibleLayered)
                && (was_hidden || last_activity.elapsed() >= Duration::from_millis(60));
            let ready_to_probe = if placed_mode == Some(PlacementMode::InvisibleLayered) {
                invisible_quiet
                    && (was_hidden
                        || (launch_idle && source_idle)
                        || settle_start.elapsed() >= Duration::from_millis(150))
            } else {
                behind_host_ready
                    || placed_mode == Some(PlacementMode::DirectWindow)
                    || was_hidden
                    || (launch_idle && source_idle)
                    || settle_start.elapsed() >= Duration::from_millis(150)
            };

            if ready_to_probe && !source_moved {
                let answered = unsafe {
                    SendMessageTimeoutW(
                        source,
                        WM_NULL,
                        WPARAM(0),
                        LPARAM(0),
                        SMTO_ABORTIFHUNG | SMTO_NORMAL,
                        50,
                        None,
                    )
                }
                .0 != 0;
                if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
                    return;
                }
                if answered {
                    let no_drift = window_rect(source) == expected_rect;
                    let aligned = no_drift
                        && match placed_mode {
                            Some(PlacementMode::InvisibleLayered) => {
                                keep_source_invisible(source.0 as isize);
                                true
                            }
                            Some(PlacementMode::DirectWindow) => true,
                            Some(PlacementMode::BehindHost) => {
                                aligned_behind_host(source, target_rect)
                            }
                            Some(PlacementMode::OffScreen) => parked(source),
                            None => false,
                        };
                    if aligned {
                        break;
                    }
                }
            }

            let slice_ms = settle_deadline
                .saturating_duration_since(Instant::now())
                .as_millis()
                .min(10) as u32;
            if slice_ms > 0 {
                let _ = unsafe {
                    MsgWaitForMultipleObjectsEx(None, slice_ms, QS_ALLINPUT, MWMO_INPUTAVAILABLE)
                };
            }
        }
        drop(settle_guard);
        drop(source_proc);
    }

    let mut lingering: Option<Instant> = None;

    loop {
        if !is_current(launch.id) {
            return;
        }
        if !unsafe { IsWindow(source) }.as_bool() {
            // A program that closed its own window has nothing left to draw, and nothing left to
            // keep: it is ended here rather than left running behind a hover that shows nothing.
            if is_current(launch.id) {
                end_the_program();
            }
            return;
        }

        let at = WANT
            .lock()
            .ok()
            .and_then(|want| want.as_ref().map(|want| want.at));
        let Some(at) = at else {
            // Re-check after reading WANT: open_fully/end_the_program bump LAUNCH before clearing WANT.
            if !is_current(launch.id) {
                return;
            }
            // Hover gone (LINGER): never run hide_others census so user-opened instances stay untouched.
            let since = *lingering.get_or_insert_with(Instant::now);
            placed_host_rect = None;
            if let Some(ref session) = hook_session {
                session.view.stage_offscreen();
                if !session.view.cloak_sync(source) {
                    if invisible_layered {
                        keep_source_invisible(source.0 as isize);
                    } else {
                        placed_mode = None;
                        let _ = unsafe {
                            SetWindowPos(
                                source,
                                HWND_BOTTOM,
                                STAGE_COORD,
                                STAGE_COORD,
                                0,
                                0,
                                SWP_NOSIZE | SWP_NOACTIVATE,
                            )
                        };
                    }
                }
            } else if invisible_layered {
                keep_source_invisible(source.0 as isize);
            } else {
                placed_mode = None;
                if !is_window_offscreen(source) {
                    push_offscreen(source);
                }
            }
            hide_the_host();
            if since.elapsed() >= LINGER {
                end_the_program();
                return;
            }
            std::thread::sleep(KEEP_POLL);
            continue;
        };
        lingering = None;

        // Hide secondary splash/dialog windows belonging strictly to `processes` (never by exe name).
        census(
            &launch,
            &processes,
            &mut images,
            &mut creation_times,
            source.0 as isize,
            true,
            OwnershipMode::OwnedOnly,
        )
        .take();
        if tree_read.elapsed() >= Duration::from_secs(1) {
            processes = owned_pids_for(launch.pid, source_pid, source_is_new);
            tree_read = Instant::now();
        }
        if let Some(vol) = volume {
            if last_volume_apply.elapsed() >= Duration::from_millis(500) {
                audio_session::apply_custom_preview_volume(processes.clone(), vol);
                last_volume_apply = Instant::now();
            }
        }

        let Some(layout) = layout_for(at, own, scale, follow_cursor(), bypass_placement()) else {
            if !is_current(launch.id) {
                return;
            }
            placed_host_rect = None;
            if let Some(ref session) = hook_session {
                session.view.stage_offscreen();
                if !session.view.cloak_sync(source) {
                    if invisible_layered {
                        keep_source_invisible(source.0 as isize);
                    } else {
                        placed_mode = None;
                        let _ = unsafe {
                            SetWindowPos(
                                source,
                                HWND_BOTTOM,
                                STAGE_COORD,
                                STAGE_COORD,
                                0,
                                0,
                                SWP_NOSIZE | SWP_NOACTIVATE,
                            )
                        };
                    }
                }
            } else if invisible_layered {
                keep_source_invisible(source.0 as isize);
            } else {
                placed_mode = None;
                if !is_window_offscreen(source) {
                    push_offscreen(source);
                }
            }
            hide_the_host();
            std::thread::sleep(KEEP_POLL);
            continue;
        };
        let (x, y) = (layout.pos_x, layout.pos_y);
        let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);
        let host_rect = (x, y, width, height);
        if let Ok(mut slot) = LAST_HOST_RECT.lock() {
            *slot = Some(host_rect);
        }

        if let Some(ref session) = hook_session {
            let needs_reposition = session.view.is_cloaked()
                || placed_mode.is_none()
                || placed_host_rect != Some(host_rect);
            if needs_reposition {
                if placed_host_rect.is_some_and(|p| p.2 != host_rect.2 || p.3 != host_rect.3) {
                    the_ask = None;
                }
                let mut staged = OffscreenStagedState {
                    invisible_layered: &mut invisible_layered,
                    placed_mode: &mut placed_mode,
                    placed_host_rect: &mut placed_host_rect,
                    expected_rect: &mut expected_rect,
                    the_ask: &mut the_ask,
                    box_of_the_host: &mut box_of_the_host,
                    registered: &mut registered,
                };
                reveal_in_process_hook(
                    host,
                    source,
                    &launch,
                    session,
                    &processes,
                    host_rect,
                    window_mode,
                    &mut staged,
                );
                std::thread::sleep(KEEP_POLL);
                continue;
            }
        } else if placed_mode.is_none() || placed_host_rect != Some(host_rect) {
            if placed_host_rect.is_some_and(|p| p.2 != host_rect.2 || p.3 != host_rect.3) {
                the_ask = None;
            }
            placed_mode = Some(if invisible_layered {
                place_invisible_layered(source, host_rect, window_mode, &mut the_ask);
                restore_foreground(launch.previous_foreground, &processes);
                PlacementMode::InvisibleLayered
            } else {
                snap_source(
                    source,
                    capture_technique,
                    host_rect,
                    window_mode,
                    launch.previous_foreground,
                    &processes,
                    &mut the_ask,
                    placed_mode.zip(placed_host_rect),
                )
            });
            placed_host_rect = Some(host_rect);
            expected_rect = window_rect(source);
        } else {
            match placed_mode {
                Some(PlacementMode::InvisibleLayered) => {
                    keep_source_invisible(source.0 as isize);
                    if window_rect(source) != expected_rect {
                        place_invisible_layered(source, host_rect, window_mode, &mut the_ask);
                        restore_foreground(launch.previous_foreground, &processes);
                        expected_rect = window_rect(source);
                    }
                }
                Some(PlacementMode::DirectWindow) => {
                    apply_direct_window_style(source);
                    if window_rect(source) != expected_rect {
                        place_direct_window(source, host_rect, window_mode);
                        restore_foreground(launch.previous_foreground, &processes);
                        expected_rect = window_rect(source);
                    }
                }
                Some(PlacementMode::BehindHost) if !aligned_behind_host(source, host_rect) => {
                    if !place_behind_host(source, host_rect, window_mode) {
                        let ask = the_ask_for((width, height), source_size(source, window_mode));
                        the_ask = Some(ask);
                        park_offscreen(source, ask.0, ask.1);
                        placed_mode = Some(PlacementMode::OffScreen);
                    }
                    restore_foreground(launch.previous_foreground, &processes);
                }
                Some(PlacementMode::OffScreen) if !parked(source) => {
                    let ask = the_ask.unwrap_or((width, height));
                    park_offscreen(source, ask.0, ask.1);
                    restore_foreground(launch.previous_foreground, &processes);
                }
                _ => {}
            }
        }

        if !uses_dwm_host(capture_technique) {
            // DirectWindow shows `source` directly as HWND_TOPMOST; never register thumbnail or show host.
            std::thread::sleep(KEEP_POLL);
            continue;
        }

        let placement_ok = match placed_mode {
            Some(PlacementMode::InvisibleLayered) => {
                keep_source_invisible(source.0 as isize);
                true
            }
            Some(PlacementMode::DirectWindow) => true,
            Some(PlacementMode::BehindHost) => aligned_behind_host(source, host_rect),
            Some(PlacementMode::OffScreen) => parked(source),
            None => false,
        };
        if !placement_ok {
            if !is_current(launch.id) {
                return;
            }
            hide_the_host();
            std::thread::sleep(KEEP_POLL);
            continue;
        }

        // Only after `source` is settled and aligned: size `host`, bind thumbnail, and show `host`.
        if box_of_the_host != Some(host_rect) {
            let _ = unsafe {
                SetWindowPos(
                    host,
                    HWND::default(),
                    x,
                    y,
                    width,
                    height,
                    SWP_NOZORDER | SWP_NOACTIVATE,
                )
            };
            box_of_the_host = Some(host_rect);
        }

        if registered == 0 {
            if let Ok(registration) = unsafe { DwmRegisterThumbnail(host, source) } {
                registered = registration;
                if let Ok(mut held) = THUMBNAIL.lock() {
                    *held = registration;
                }
            }
        }
        if registered != 0 {
            let mut client = RECT::default();
            let _ = unsafe { GetClientRect(host, &mut client) };
            // Prefer DWM's own reported source size in Normal mode (matches exact composited frame
            // without Win11 invisible-border mismatch); fall back to window/client rect.
            let picture = if window_mode == CustomWindowMode::Normal {
                unsafe { DwmQueryThumbnailSourceSize(registered) }
                    .ok()
                    .filter(|s| s.cx > 0 && s.cy > 0)
                    .map(|s| (s.cx, s.cy))
                    .unwrap_or_else(|| {
                        let (w, h) = source_size(source, window_mode);
                        (w as i32, h as i32)
                    })
            } else {
                let (w, h) = source_size(source, window_mode);
                (w as i32, h as i32)
            };
            let properties =
                thumbnail_properties(client.right, client.bottom, picture, window_mode);
            if unsafe { DwmUpdateThumbnailProperties(registered, &properties) }.is_ok()
                && !SHOWN.load(Ordering::SeqCst)
            {
                SHOWN.store(true, Ordering::SeqCst);
                let _ =
                    unsafe { PostMessageW(host, WM_THE_PICTURE_IS_READY, WPARAM(0), LPARAM(0)) };
            }
        }

        std::thread::sleep(KEEP_POLL);
    }
}

fn thumbnail_properties(
    client_width: i32,
    client_height: i32,
    picture: (i32, i32),
    window_mode: CustomWindowMode,
) -> DWM_THUMBNAIL_PROPERTIES {
    DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION
            | DWM_TNP_VISIBLE
            | DWM_TNP_SOURCECLIENTAREAONLY
            | DWM_TNP_OPACITY,
        rcDestination: the_picture_in(client_width, client_height, picture),
        rcSource: RECT::default(),
        opacity: 255,
        fVisible: true.into(),
        fSourceClientAreaOnly: (window_mode == CustomWindowMode::Borderless).into(),
    }
}

fn resolve_scale(path: &Path, scaling: CustomScaling) -> PreviewScale {
    match scaling {
        // Explicit preset overrides keep "share of room" semantics (Screen 100%/75%/50%, Fit, Custom %).
        CustomScaling::Override(scale) => fit_reduced(scale),
        CustomScaling::FollowTray => tray_scale_for(path),
    }
}

fn tray_scale_for(path: &Path) -> PreviewScale {
    let hover = HoverFacts::read(path);
    let audio_scale = CONFIG
        .lock()
        .map(|cfg| cfg.audio_scale)
        .unwrap_or(DEFAULT_AUDIO_SCALE);
    let kind = match hover.route.content {
        crate::formats::content_type::Content::Kind(kind) => kind,
        _ => hover.route.named.unwrap_or(PreviewType::Images),
    };
    tray_scale_for_kind(kind, path, hover.scales, audio_scale)
}

fn tray_scale_for_kind(
    kind: PreviewType,
    path: &Path,
    scales: HoverScales,
    audio_scale: PreviewScale,
) -> PreviewScale {
    match kind {
        PreviewType::Images => animated_scale_for(path, scales).unwrap_or(scales.picture),
        PreviewType::Magick => scales.picture,
        PreviewType::Videos => scales.video,
        PreviewType::Audio => fit_reduced(audio_scale),
        PreviewType::Ebook | PreviewType::Calibre => fit_reduced(scales.ebook),
        PreviewType::Document | PreviewType::Libre => fit_reduced(scales.document),
        PreviewType::Fonts => fit_reduced(scales.font),
        PreviewType::Design => fit_reduced(scales.design),
        PreviewType::Vector => fit_reduced(scales.vector),
        PreviewType::Models => fit_reduced(scales.model),
        PreviewType::Text if html_is_engine_drawn(path) => fit_reduced(scales.document),
        PreviewType::Text | PreviewType::Archives | PreviewType::Peazip => {
            PreviewScale::Percent(100)
        }
    }
}

fn layout_for(
    at: Where,
    own: (u32, u32),
    scale: PreviewScale,
    follow_cursor: bool,
    bypass_placement: bool,
) -> Option<PreviewLayout> {
    match at {
        Where::Pointer { cursor, avoid } => {
            let bounds = work_area_at(cursor.0, cursor.1);
            let dpi = dpi_at(cursor.0, cursor.1);
            compute_mouse_layout(
                cursor.0,
                cursor.1,
                HoverPlacement {
                    orig_dims: own,
                    avoid,
                    follow_cursor,
                    bypass_placement,
                    preview_scale: scale,
                    at_the_pointer_corner: false,
                },
                bounds,
                dpi,
            )
        }
        Where::Item {
            rect,
            avoid,
            columns,
        } => {
            let (left, top, right, bottom) = rect;
            let center = ((left + right) / 2, (top + bottom) / 2);
            let bounds = work_area_at(center.0, center.1);
            let dpi = dpi_at(center.0, center.1);
            compute_keyboard_layout(
                KeyboardPlacement {
                    item_rect: rect,
                    avoid,
                    columns,
                    orig_dims: own,
                    follow_cursor,
                    bypass_placement,
                    preview_scale: scale,
                },
                bounds,
                dpi,
            )
        }
    }
}

/// The size the program's window is asked for: the box's own shape, raised to the smallest size the
/// window will take at that shape.
///
/// **A window that will not take the box is asked for a bigger box of the same shape.** The box is
/// the room cut to the file's shape, so a tall picture's box is a narrow one — narrower than a
/// program is obliged to draw itself: Windows 11's Paint takes 704x500 as its floor and answers a
/// smaller ask with that, a shape of its own rather than the box's. Composed into the box as it
/// stands the picture was letterboxed in the host's background, and grown to the size it took the
/// hover was the window's shape rather than the file's. So the window is asked for the box's shape
/// at the smallest size that clears the least it will take (see `the_size_it_took`), and it is
/// composed into the box at exactly that shape (see `the_picture_in`): the hover is the size the
/// setting says it is and the program's picture fills it, the window being off every display where
/// it stands.
fn the_ask_for(box_: (i32, i32), taken: (u32, u32)) -> (i32, i32) {
    let box_ = (box_.0.max(1), box_.1.max(1));
    let shape = box_.0 as f32 / box_.1 as f32;

    let width = (taken.0 as i32)
        .max((taken.1 as f32 * shape).ceil() as i32)
        .max(box_.0);
    let height = ((width as f32 / shape).ceil() as i32).max(1);

    (width, height)
}

/// Fast intrinsic dimensions of `path` (SVG, PDF, image, video, design, vector) without spawning
/// external tools; returns `None` when the file declares no dimensions so `own_size` is used.
fn base_size(path: &Path) -> Option<(u32, u32)> {
    let valid = |dims: (u32, u32)| dims.0 > 0 && dims.1 > 0;
    if svg_preview::is_svg_file(path) {
        if let Some(dims) = svg_preview::measure(path).filter(|&d| valid(d)) {
            return Some(dims);
        }
    }
    if pdf_preview::is_pdf_preview(path) {
        if let Some(dims) = pdf_preview::page_dimensions(path).filter(|&d| valid(d)) {
            return Some(dims);
        }
    }
    if let Some(dims) = picture_dimensions(path).filter(|&d| valid(d)) {
        return Some(dims);
    }
    if let Some(ProbedGeometry::Measured(geom)) = cached_video_geometry(path) {
        if valid((geom.width, geom.height)) {
            return Some((geom.width, geom.height));
        }
    }
    if let Some(dims) = video_player::header(path)
        .map(|h| video_player::displayed_shape(h.frame_width, h.frame_height, h.aspect))
        .filter(|&d| valid(d))
    {
        return Some(dims);
    }
    design_dimensions(path)
        .or_else(|| vector_dimensions(path))
        .filter(|&d| valid(d))
}

/// The size the program's own window opens at: the shape of a file that declares none of its own.
///
/// It is the fallback [`base_size`] leaves, and it is a shape of the program's rather than of the
/// file's, which is the whole reason it is the fallback.
///
/// It is the window's *restored* box and not the one it is showing at, because a program that
/// remembers being maximized is showing at the whole display — and the display is what the fit is
/// for, not what the shape means. A window that never had a restored box is measured where it
/// stands.
fn own_size(window: HWND) -> (u32, u32) {
    let mut placement = WINDOWPLACEMENT {
        length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
        ..Default::default()
    };
    if unsafe { GetWindowPlacement(window, &mut placement) }.is_ok() {
        let rect = placement.rcNormalPosition;
        let (width, height) = (rect.right - rect.left, rect.bottom - rect.top);
        if width >= 64 && height >= 64 {
            return (width as u32, height as u32);
        }
    }

    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_ok() {
        let (width, height) = (
            (rect.right - rect.left).max(1),
            (rect.bottom - rect.top).max(1),
        );
        return (width as u32, height as u32);
    }

    (800, 600)
}

/// Captured size of `window`: full window rect in `Normal` mode, client rect in `Borderless` mode.
fn source_size(window: HWND, window_mode: CustomWindowMode) -> (u32, u32) {
    let mut rect = RECT::default();
    let ok = match window_mode {
        CustomWindowMode::Normal => unsafe { GetWindowRect(window, &mut rect) }.is_ok(),
        CustomWindowMode::Borderless => unsafe { GetClientRect(window, &mut rect) }.is_ok(),
    };
    if !ok {
        return (0, 0);
    }
    (
        (rect.right - rect.left).max(0) as u32,
        (rect.bottom - rect.top).max(0) as u32,
    )
}

/// The rectangle in the host's client the program's window is drawn in: the window's own client at
/// the size it has where that fits, and its own shape fitted into the client where it does not.
///
/// **The box and the window are the same shape normally**, because the window is asked for the box's
/// own shape at the smallest size it will take (see `the_ask_for`), and this is then the host's whole
/// client with the program's picture filling it. They part where the program would not take that ask
/// either — its floor larger than the display, the room beside the text a hover keeps off — and a
/// preview that shows the program at its own shape inside the box says more than one that shows
/// nothing. The window is never stretched: the shape it is drawn at is its own, and `background` is
/// what the host paints in the letterbox.
fn the_picture_in(client_width: i32, client_height: i32, picture: (i32, i32)) -> RECT {
    let client_width = client_width.max(1);
    let client_height = client_height.max(1);
    let picture_width = picture.0.max(1);
    let picture_height = picture.1.max(1);

    let (width, height) = if picture_width <= client_width && picture_height <= client_height {
        // The window is no bigger than the box: it is drawn at its own size, which is the case
        // every program that takes the box is in — and a window smaller than the box is a picture
        // at its own size with the background around it rather than one stretched to fill it.
        (picture_width, picture_height)
    } else {
        let scale = (client_width as f32 / picture_width as f32)
            .min(client_height as f32 / picture_height as f32);
        (
            ((picture_width as f32 * scale).round() as i32).clamp(1, client_width),
            ((picture_height as f32 * scale).round() as i32).clamp(1, client_height),
        )
    };

    let left = (client_width - width) / 2;
    let top = (client_height - height) / 2;

    RECT {
        left,
        top,
        right: left + width,
        bottom: top + height,
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum PlacementMode {
    InvisibleLayered,
    BehindHost,
    OffScreen,
    DirectWindow,
}

fn tries_invisible_layered(technique: CustomCaptureTechnique) -> bool {
    technique == CustomCaptureTechnique::InvisibleLayered
}

fn uses_dwm_host(technique: CustomCaptureTechnique) -> bool {
    technique != CustomCaptureTechnique::DirectWindow
}

/// Ensure click-through, no-taskbar, non-activating extended styles for `DirectWindow` mode.
fn apply_direct_window_style(window: HWND) {
    let ex = unsafe { GetWindowLongPtrW(window, GWL_EXSTYLE) };
    if (ex as u32 & DIRECT_EX_STYLE) != DIRECT_EX_STYLE {
        let _ = unsafe { SetWindowLongPtrW(window, GWL_EXSTYLE, ex | (DIRECT_EX_STYLE as isize)) };
        let _ = unsafe {
            SetWindowPos(
                window,
                HWND::default(),
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        };
    }
}

/// Place `window` directly at `host_rect` as `HWND_TOPMOST` click-through (no DWM host/thumbnail).
fn place_direct_window(
    window: HWND,
    host_rect: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
) {
    apply_direct_window_style(window);
    let insets = if window_mode == CustomWindowMode::Normal {
        dwm_frame_insets(window)
    } else {
        (0, 0, 0, 0)
    };
    let (wx, wy, ww, wh) = behind_host_window_rect(host_rect, insets);

    if unsafe { IsZoomed(window) }.as_bool() || unsafe { IsIconic(window) }.as_bool() {
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags: WINDOWPLACEMENT_FLAGS(0),
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: workspace_rect_for_screen(RECT {
                left: wx,
                top: wy,
                right: wx + ww,
                bottom: wy + wh,
            }),
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(window, &placement) };
    }

    let _ = unsafe { SetWindowPos(window, HWND_TOPMOST, wx, wy, ww, wh, SWP_NOACTIVATE) };
    if !unsafe { IsWindowVisible(window) }.as_bool() {
        let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
    }
    let _ = unsafe { InvalidateRect(window, None, false) };
    let _ = unsafe { UpdateWindow(window) };
}

/// True when `WS_EX_LAYERED` + `LWA_ALPHA = 1` is safe for a window with `ex_style`.
/// Rejects DirectComposition/WinUI 3 (`WS_EX_NOREDIRECTIONBITMAP`) and `UpdateLayeredWindow` windows.
fn can_use_invisible_layered(ex_style: u32, layered_attr_ok: Option<bool>) -> bool {
    if (ex_style & WS_EX_NOREDIRECTIONBITMAP.0) != 0 {
        return false;
    }
    if (ex_style & WS_EX_LAYERED.0) != 0 && layered_attr_ok == Some(false) {
        return false;
    }
    true
}

fn layered_alpha_flags(source: HWND) -> Option<(u8, LAYERED_WINDOW_ATTRIBUTES_FLAGS)> {
    let mut key = COLORREF(0);
    let mut alpha = 0u8;
    let mut flags = LAYERED_WINDOW_ATTRIBUTES_FLAGS(0);
    unsafe {
        GetLayeredWindowAttributes(source, Some(&mut key), Some(&mut alpha), Some(&mut flags))
    }
    .is_ok()
    .then_some((alpha, flags))
}

/// Apply `WS_EX_LAYERED` + `LWA_ALPHA = 1` on a visible source; reverts immediately if rejected.
fn make_source_invisible(source: isize) -> bool {
    if source == 0 {
        return false;
    }
    let hwnd = HWND(source as *mut _);
    // Must be visible first: applying LWA_ALPHA before first show can be ignored by Windows.
    if !unsafe { IsWindow(hwnd) }.as_bool() || !unsafe { IsWindowVisible(hwnd) }.as_bool() {
        return false;
    }
    let original_ex = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    let ex_u32 = original_ex as u32;
    let layered_attr_ok =
        ((ex_u32 & WS_EX_LAYERED.0) != 0).then(|| layered_alpha_flags(hwnd).is_some());
    if !can_use_invisible_layered(ex_u32, layered_attr_ok) {
        return false;
    }

    let _ = unsafe {
        SetWindowLongPtrW(
            hwnd,
            GWL_EXSTYLE,
            original_ex | (INVISIBLE_EX_STYLE as isize),
        )
    };
    let _ = unsafe {
        SetWindowPos(
            hwnd,
            HWND_BOTTOM,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };
    let applied =
        unsafe { SetLayeredWindowAttributes(hwnd, COLORREF(0), INVISIBLE_ALPHA, LWA_ALPHA) };
    let verified = applied.is_ok()
        && layered_alpha_flags(hwnd)
            .is_some_and(|(alpha, flags)| (flags.0 & LWA_ALPHA.0) != 0 && alpha == INVISIBLE_ALPHA);
    if !verified {
        // Revert immediately to prevent black-band artifact on windows that reject LWA_ALPHA.
        let _ = unsafe { SetWindowLongPtrW(hwnd, GWL_EXSTYLE, original_ex) };
        let _ = unsafe {
            SetWindowPos(
                hwnd,
                HWND::default(),
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        };
        return false;
    }
    true
}

/// Re-assert `INVISIBLE_EX_STYLE` + `LWA_ALPHA = 1` + `HWND_BOTTOM` if the app stripped or reset them.
fn keep_source_invisible(source: isize) {
    if source == 0 {
        return;
    }
    let hwnd = HWND(source as *mut _);
    if !unsafe { IsWindow(hwnd) }.as_bool() {
        return;
    }
    let ex = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    let missing_style = (ex as u32 & INVISIBLE_EX_STYLE) != INVISIBLE_EX_STYLE;
    let alpha_ok = !missing_style
        && layered_alpha_flags(hwnd)
            .is_some_and(|(alpha, flags)| (flags.0 & LWA_ALPHA.0) != 0 && alpha == INVISIBLE_ALPHA);
    if alpha_ok {
        return;
    }
    if missing_style {
        let _ = unsafe { SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ex | (INVISIBLE_EX_STYLE as isize)) };
        let _ = unsafe {
            SetWindowPos(
                hwnd,
                HWND_BOTTOM,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        };
    } else {
        let _ = unsafe {
            SetWindowPos(
                hwnd,
                HWND_BOTTOM,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
            )
        };
    }
    let _ = unsafe { SetLayeredWindowAttributes(hwnd, COLORREF(0), INVISIBLE_ALPHA, LWA_ALPHA) };
}

/// Position `window` on-monitor at `(host_x, host_y, ask_w, ask_h)` in `InvisibleLayered` mode.
/// Never flings off-screen even if min-size floor exceeds `host_rect`.
fn place_invisible_layered(
    window: HWND,
    host_rect: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
    the_ask: &mut Option<(i32, i32)>,
) {
    keep_source_invisible(window.0 as isize);
    let (hx, hy, hw, hh) = host_rect;
    let (tw, th) = the_ask.unwrap_or((hw, hh));
    let insets = if window_mode == CustomWindowMode::Normal {
        dwm_frame_insets(window)
    } else {
        (0, 0, 0, 0)
    };
    let (wx, wy, ww, wh) = behind_host_window_rect((hx, hy, tw, th), insets);

    if unsafe { IsZoomed(window) }.as_bool() || unsafe { IsIconic(window) }.as_bool() {
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags: WINDOWPLACEMENT_FLAGS(0),
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: workspace_rect_for_screen(RECT {
                left: wx,
                top: wy,
                right: wx + ww,
                bottom: wy + wh,
            }),
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(window, &placement) };
    }

    let _ = unsafe { SetWindowPos(window, HWND_BOTTOM, wx, wy, ww, wh, SWP_NOACTIVATE) };

    // If app's min-size floor refused `(tw, th)`, resize on-monitor to `the_ask_for` keeping aspect ratio.
    let taken = source_size(window, window_mode);
    let ask = the_ask_for((hw, hh), taken);
    if ask != (tw, th) {
        *the_ask = (ask != (hw, hh)).then_some(ask);
        let (wx2, wy2, ww2, wh2) = behind_host_window_rect((hx, hy, ask.0, ask.1), insets);
        let _ = unsafe { SetWindowPos(window, HWND_BOTTOM, wx2, wy2, ww2, wh2, SWP_NOACTIVATE) };
    }

    if !unsafe { IsWindowVisible(window) }.as_bool() {
        let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
    }
    keep_source_invisible(window.0 as isize);
    let _ = unsafe { InvalidateRect(window, None, false) };
    let _ = unsafe { UpdateWindow(window) };
}

fn virtual_screen() -> RECT {
    let left = unsafe { GetSystemMetrics(SM_XVIRTUALSCREEN) };
    let top = unsafe { GetSystemMetrics(SM_YVIRTUALSCREEN) };
    let width = unsafe { GetSystemMetrics(SM_CXVIRTUALSCREEN) }.max(1);
    let height = unsafe { GetSystemMetrics(SM_CYVIRTUALSCREEN) }.max(1);
    RECT {
        left,
        top,
        right: left + width,
        bottom: top + height,
    }
}

/// Invisible Win11 DWM border insets `(left, top, right, bottom)` between `GetWindowRect` and
/// `DWMWA_EXTENDED_FRAME_BOUNDS` (zeros if unavailable).
fn dwm_frame_insets(window: HWND) -> (i32, i32, i32, i32) {
    let mut wr = RECT::default();
    if unsafe { GetWindowRect(window, &mut wr) }.is_err() {
        return (0, 0, 0, 0);
    }
    let mut ef = RECT::default();
    let ok = unsafe {
        DwmGetWindowAttribute(
            window,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            &mut ef as *mut RECT as *mut _,
            std::mem::size_of::<RECT>() as u32,
        )
    }
    .is_ok();
    if !ok || ef.right <= ef.left || ef.bottom <= ef.top {
        return (0, 0, 0, 0);
    }
    (
        (ef.left - wr.left).max(0),
        (ef.top - wr.top).max(0),
        (wr.right - ef.right).max(0),
        (wr.bottom - ef.bottom).max(0),
    )
}

/// Visible frame rect via `DWMWA_EXTENDED_FRAME_BOUNDS`, falling back to `GetWindowRect`.
fn visible_frame_rect(window: HWND) -> Option<RECT> {
    let mut ef = RECT::default();
    if unsafe {
        DwmGetWindowAttribute(
            window,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            &mut ef as *mut RECT as *mut _,
            std::mem::size_of::<RECT>() as u32,
        )
    }
    .is_ok()
        && ef.right > ef.left
        && ef.bottom > ef.top
    {
        return Some(ef);
    }
    let mut wr = RECT::default();
    unsafe { GetWindowRect(window, &mut wr) }.ok().map(|()| wr)
}

/// Outer window rect `(x, y, w, h)` so visible frame aligns pixel-for-pixel inside `host_rect`.
fn behind_host_window_rect(
    host_rect: (i32, i32, i32, i32),
    insets: (i32, i32, i32, i32),
) -> (i32, i32, i32, i32) {
    let (hx, hy, hw, hh) = host_rect;
    let (il, it, ir, ib) = insets;
    (
        hx - il,
        hy - it,
        (hw + il + ir).max(1),
        (hh + it + ib).max(1),
    )
}

fn window_rect(window: HWND) -> Option<RECT> {
    let mut rect = RECT::default();
    unsafe { GetWindowRect(window, &mut rect) }
        .ok()
        .map(|()| rect)
}

/// Drain queued WinEvents; suppress owned companion windows and report `source` location changes.
fn drain_settle_events(source: HWND, launch: &Launch, processes: &[u32]) -> bool {
    let mut msg = MSG::default();
    while unsafe { PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE) }.as_bool() {
        let _ = unsafe { TranslateMessage(&msg) };
        unsafe { DispatchMessageW(&msg) };
    }
    let mut source_moved = false;
    WINEVENT_EVENTS.with(|q| {
        for (event, raw) in q.borrow_mut().drain(..) {
            if raw == source.0 as isize {
                if event == EVENT_OBJECT_LOCATIONCHANGE {
                    source_moved = true;
                }
            } else if (event == EVENT_OBJECT_CREATE || event == EVENT_OBJECT_SHOW)
                && !launch.windows_before.contains(&raw)
            {
                let hwnd = HWND(raw as *mut _);
                if processes.contains(&pid_of(hwnd)) {
                    suppress_companion_window(hwnd, source.0 as isize);
                }
            }
        }
    });
    source_moved
}

const ALIGN_TOLERANCE: i32 = 4;

/// True when `actual_visible` fits inside `host_rect` (+4px DPI rounding tolerance).
fn is_aligned_behind_host(actual_visible: RECT, host_rect: (i32, i32, i32, i32)) -> bool {
    let (hx, hy, hw, hh) = host_rect;
    let aw = actual_visible.right - actual_visible.left;
    let ah = actual_visible.bottom - actual_visible.top;
    aw > 0
        && ah > 0
        && actual_visible.left >= hx - ALIGN_TOLERANCE
        && actual_visible.top >= hy - ALIGN_TOLERANCE
        && actual_visible.right <= hx + hw + ALIGN_TOLERANCE
        && actual_visible.bottom <= hy + hh + ALIGN_TOLERANCE
}

fn aligned_behind_host(window: HWND, host_rect: (i32, i32, i32, i32)) -> bool {
    visible_frame_rect(window).is_some_and(|r| is_aligned_behind_host(r, host_rect))
        && unsafe { IsWindowVisible(window) }.as_bool()
        && !unsafe { IsZoomed(window) }.as_bool()
        && !unsafe { IsIconic(window) }.as_bool()
}

/// Convert screen rect to `WINDOWPLACEMENT::rcNormalPosition` workspace coordinates.
fn screen_to_workspace_rect(screen: RECT, monitor: RECT, work: RECT) -> RECT {
    let dx = work.left - monitor.left;
    let dy = work.top - monitor.top;
    RECT {
        left: screen.left - dx,
        top: screen.top - dy,
        right: screen.right - dx,
        bottom: screen.bottom - dy,
    }
}

fn workspace_rect_for_screen(screen: RECT) -> RECT {
    let hmon = unsafe { MonitorFromRect(&screen, MONITOR_DEFAULTTONEAREST) };
    let mut info = MONITORINFO {
        cbSize: std::mem::size_of::<MONITORINFO>() as u32,
        ..Default::default()
    };
    if !hmon.is_invalid() && unsafe { GetMonitorInfoW(hmon, &mut info) }.as_bool() {
        screen_to_workspace_rect(screen, info.rcMonitor, info.rcWork)
    } else {
        screen
    }
}

/// Place `window` 100% on-monitor directly behind `host` at `HWND_BOTTOM` without `SWP_FRAMECHANGED`
/// or `RDW_ERASE` so classic GDI apps (IrfanView) repaint natively. Returns true if it fits inside `host`.
fn place_behind_host(
    window: HWND,
    host_rect: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
) -> bool {
    let insets = if window_mode == CustomWindowMode::Normal {
        dwm_frame_insets(window)
    } else {
        (0, 0, 0, 0)
    };
    let (wx, wy, ww, wh) = behind_host_window_rect(host_rect, insets);

    if unsafe { IsZoomed(window) }.as_bool() || unsafe { IsIconic(window) }.as_bool() {
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags: WINDOWPLACEMENT_FLAGS(0),
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: workspace_rect_for_screen(RECT {
                left: wx,
                top: wy,
                right: wx + ww,
                bottom: wy + wh,
            }),
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(window, &placement) };
    }

    let _ = unsafe { SetWindowPos(window, HWND_BOTTOM, wx, wy, ww, wh, SWP_NOACTIVATE) };

    if !unsafe { IsWindowVisible(window) }.as_bool() {
        let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
    }

    let _ = unsafe { InvalidateRect(window, None, false) };
    let _ = unsafe { UpdateWindow(window) };

    aligned_behind_host(window, host_rect)
}

/// Off-screen parking origin past virtual screen's right edge (clamped to i16 range).
fn parking_origin(vscreen: RECT) -> (i32, i32) {
    let x = (vscreen.right + 1000).clamp(-30000, 30000);
    let y = vscreen.top.clamp(-30000, 30000);
    (x, y)
}

/// True when `rect` is completely outside `vscreen`.
fn is_off_screen(rect: &RECT, vscreen: RECT) -> bool {
    rect.left >= vscreen.right
        || rect.right <= vscreen.left
        || rect.top >= vscreen.bottom
        || rect.bottom <= vscreen.top
}

fn is_window_offscreen(window: HWND) -> bool {
    let mut rect = RECT::default();
    unsafe { GetWindowRect(window, &mut rect) }.is_ok() && is_off_screen(&rect, virtual_screen())
}

fn parked(window: HWND) -> bool {
    is_window_offscreen(window)
        && unsafe { IsWindowVisible(window) }.as_bool()
        && !unsafe { IsZoomed(window) }.as_bool()
        && !unsafe { IsIconic(window) }.as_bool()
}

/// True only when primary parking `SetWindowPos` failed or primary coords overlap `vscreen`.
/// Never fling to `FALLBACK_PARKING` when the app clamped the window onto the monitor.
fn should_try_fallback_parking(
    primary_swp_ok: bool,
    primary_origin: (i32, i32),
    actual_rect: Option<RECT>,
    vscreen: RECT,
) -> bool {
    if !primary_swp_ok {
        return true;
    }
    let primary_overlaps_vscreen = primary_origin.0 >= vscreen.left
        && primary_origin.0 < vscreen.right
        && primary_origin.1 >= vscreen.top
        && primary_origin.1 < vscreen.bottom;
    if primary_overlaps_vscreen {
        return true;
    }
    let Some(rect) = actual_rect else {
        return true;
    };
    if is_off_screen(&rect, vscreen) {
        return false;
    }
    let clamped_onto_monitor = rect.left < vscreen.right && rect.right > vscreen.left;
    !clamped_onto_monitor
}

/// Fast off-screen move preserving size (used on `hide()` / `LINGER` so `source` never stays visible).
fn push_offscreen(window: HWND) {
    let vscreen = virtual_screen();
    let (px, py) = parking_origin(vscreen);
    let swp_ok = unsafe {
        SetWindowPos(
            window,
            HWND_BOTTOM,
            px,
            py,
            0,
            0,
            SWP_NOSIZE | SWP_NOACTIVATE,
        )
    }
    .is_ok();
    if should_try_fallback_parking(swp_ok, (px, py), window_rect(window), vscreen) {
        let _ = unsafe {
            SetWindowPos(
                window,
                HWND_BOTTOM,
                FALLBACK_PARKING,
                FALLBACK_PARKING,
                0,
                0,
                SWP_NOSIZE | SWP_NOACTIVATE,
            )
        };
    }
}

/// Synchronously push active `SOURCE` off-screen before hiding `host`.
fn park_source_offscreen_now() {
    let source = SOURCE.load(Ordering::SeqCst);
    if source == 0 {
        return;
    }
    let hwnd = HWND(source as *mut _);
    if unsafe { IsWindow(hwnd) }.as_bool() && !is_window_offscreen(hwnd) {
        push_offscreen(hwnd);
    }
}

/// Fallback off-screen park when `window`'s minimum size exceeds `host`.
fn park_offscreen(window: HWND, width: i32, height: i32) {
    let vscreen = virtual_screen();
    let (px, py) = parking_origin(vscreen);
    let target = RECT {
        left: px,
        top: py,
        right: px + width,
        bottom: py + height,
    };

    if unsafe { IsZoomed(window) }.as_bool() || unsafe { IsIconic(window) }.as_bool() {
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: target,
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(window, &placement) };
    }

    let swp_ok =
        unsafe { SetWindowPos(window, HWND_BOTTOM, px, py, width, height, SWP_NOACTIVATE) }.is_ok();

    if should_try_fallback_parking(swp_ok, (px, py), window_rect(window), vscreen) {
        let _ = unsafe {
            SetWindowPos(
                window,
                HWND_BOTTOM,
                FALLBACK_PARKING,
                FALLBACK_PARKING,
                0,
                0,
                SWP_NOSIZE | SWP_NOACTIVATE,
            )
        };
    }

    if !unsafe { IsWindowVisible(window) }.as_bool() {
        let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
    }

    let _ = unsafe { InvalidateRect(window, None, false) };
    let _ = unsafe { UpdateWindow(window) };
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SnapAction {
    KeepInvisibleLayered,
    PlaceDirectWindow,
    KeepOffScreen,
    ReparkOffScreen,
    TryBehindHost,
}

/// Decide placement action for `technique`; never bounce `InvisibleLayered` or `OffScreen` through `BehindHost`.
fn decide_snap_action(
    technique: CustomCaptureTechnique,
    current_mode: Option<PlacementMode>,
    is_parked: bool,
    prev_host_rect: Option<(i32, i32, i32, i32)>,
    target_rect: (i32, i32, i32, i32),
) -> SnapAction {
    match technique {
        CustomCaptureTechnique::DirectWindow => SnapAction::PlaceDirectWindow,
        CustomCaptureTechnique::OffScreen => {
            if is_parked {
                SnapAction::KeepOffScreen
            } else {
                SnapAction::ReparkOffScreen
            }
        }
        CustomCaptureTechnique::InvisibleLayered
        | CustomCaptureTechnique::BehindHost
        | CustomCaptureTechnique::OffscreenStaged
        | CustomCaptureTechnique::InProcessHook => {
            if current_mode == Some(PlacementMode::InvisibleLayered) {
                return SnapAction::KeepInvisibleLayered;
            }
            let host_grew = prev_host_rect
                .is_some_and(|(_, _, pw, ph)| target_rect.2 > pw || target_rect.3 > ph);
            if current_mode == Some(PlacementMode::OffScreen) && !host_grew {
                if is_parked {
                    SnapAction::KeepOffScreen
                } else {
                    SnapAction::ReparkOffScreen
                }
            } else {
                SnapAction::TryBehindHost
            }
        }
    }
}

/// Snap `source` according to `technique` and restore focus.
#[allow(clippy::too_many_arguments)]
fn snap_source(
    source: HWND,
    technique: CustomCaptureTechnique,
    target_rect: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
    previous_foreground: isize,
    processes: &[u32],
    the_ask: &mut Option<(i32, i32)>,
    prior: Option<(PlacementMode, (i32, i32, i32, i32))>,
) -> PlacementMode {
    let (current_mode, prev_host_rect) = prior.unzip();
    match decide_snap_action(
        technique,
        current_mode,
        parked(source),
        prev_host_rect,
        target_rect,
    ) {
        SnapAction::KeepInvisibleLayered => {
            place_invisible_layered(source, target_rect, window_mode, the_ask);
            restore_foreground(previous_foreground, processes);
            PlacementMode::InvisibleLayered
        }
        SnapAction::PlaceDirectWindow => {
            place_direct_window(source, target_rect, window_mode);
            restore_foreground(previous_foreground, processes);
            PlacementMode::DirectWindow
        }
        SnapAction::KeepOffScreen => PlacementMode::OffScreen,
        SnapAction::ReparkOffScreen => {
            let ask = the_ask.unwrap_or_else(|| {
                the_ask_for(
                    (target_rect.2, target_rect.3),
                    source_size(source, window_mode),
                )
            });
            *the_ask = Some(ask);
            park_offscreen(source, ask.0, ask.1);
            restore_foreground(previous_foreground, processes);
            PlacementMode::OffScreen
        }
        SnapAction::TryBehindHost => {
            let mode = if place_behind_host(source, target_rect, window_mode) {
                *the_ask = None;
                PlacementMode::BehindHost
            } else {
                let ask = the_ask_for(
                    (target_rect.2, target_rect.3),
                    source_size(source, window_mode),
                );
                *the_ask = Some(ask);
                park_offscreen(source, ask.0, ask.1);
                PlacementMode::OffScreen
            };
            restore_foreground(previous_foreground, processes);
            mode
        }
    }
}

fn wait_for_input_idle_handle(process: HANDLE, timeout_ms: u32) {
    let _ = unsafe { WaitForInputIdle(process, timeout_ms) };
}

fn is_input_idle_now(process: HANDLE) -> bool {
    // `WAIT_TIMEOUT` means still starting; 0 (idle) or `WAIT_FAILED` (exited/no GUI) means done.
    let status = unsafe { WaitForInputIdle(process, 0) };
    status != WAIT_TIMEOUT.0
}

fn open_process_for_idle(pid: u32) -> Option<OwnedHandle> {
    let process = unsafe {
        OpenProcess(
            PROCESS_SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            pid,
        )
    }
    .ok()?;
    // SAFETY: `process` is uniquely owned and closed on `OwnedHandle` drop.
    Some(unsafe { OwnedHandle::from_raw_handle(process.0) })
}

/// Restore previous foreground window (e.g. Explorer) if the spawned app stole focus.
fn restore_foreground(previous: isize, processes: &[u32]) {
    if previous == 0 {
        return;
    }
    let prev = HWND(previous as *mut _);
    if !unsafe { IsWindow(prev) }.as_bool()
        || !unsafe { IsWindowVisible(prev) }.as_bool()
        || !unsafe { IsWindowEnabled(prev) }.as_bool()
    {
        return;
    }
    let current = unsafe { GetForegroundWindow() };
    if current.0.is_null() || current == prev {
        return;
    }
    if processes.contains(&pid_of(current)) {
        let _ = unsafe { SetForegroundWindow(prev) };
    }
}

/// Take the frame off the window, and keep it out of the taskbar and out of the focus: what the
/// user asked to see is the program's picture, not its title bar and not its button.
fn strip_the_frame(window: HWND) {
    let style = unsafe { GetWindowLongPtrW(window, GWL_STYLE) } as u32;
    let stripped = style
        & !(WS_CAPTION.0
            | WS_THICKFRAME.0
            | WS_SYSMENU.0
            | WS_MINIMIZEBOX.0
            | WS_MAXIMIZEBOX.0
            | WS_CHILD.0);
    let _ = unsafe { SetWindowLongPtrW(window, GWL_STYLE, stripped as isize) };

    let ex_style = unsafe { GetWindowLongPtrW(window, GWL_EXSTYLE) } as u32;
    let stripped = ex_style | WS_EX_TOOLWINDOW.0 | WS_EX_NOACTIVATE.0;
    let _ = unsafe { SetWindowLongPtrW(window, GWL_EXSTYLE, stripped as isize) };
    let _ = unsafe {
        SetWindowPos(
            window,
            HWND::default(),
            0,
            0,
            0,
            0,
            SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE,
        )
    };
}

/// True when `exe` lives under a standard Windows MSIX / AppExecutionAlias `WindowsApps` folder.
fn is_packaged_broker_path(exe: &Path) -> bool {
    exe.parent().is_some_and(|dir| {
        dir.components().any(|c| match c {
            std::path::Component::Normal(seg) => {
                seg.to_string_lossy().eq_ignore_ascii_case("WindowsApps")
            }
            _ => false,
        })
    })
}

/// Suppress a visible secondary/companion window of `processes`; owned `#32770` dialogs get `WM_CLOSE`.
fn suppress_companion_window(hwnd: HWND, keep: isize) {
    if hwnd.0 as isize == keep || hwnd.0.is_null() {
        return;
    }
    let owner_hwnd = unsafe { GetWindow(hwnd, GW_OWNER) }
        .ok()
        .filter(|o| !o.0.is_null());
    // Never SW_HIDE an owned #32770 modal dialog: hiding leaves `source` WS_DISABLED
    // (white/frozen + Windows ding). Dismiss via WM_CLOSE instead.
    let owned_by_keep = owner_hwnd.is_some_and(|o| o.0 as isize == keep);
    if owned_by_keep && is_dialog_class(hwnd) {
        let _ = unsafe { PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) };
    } else if !owned_by_keep && unsafe { IsWindowVisible(hwnd) }.as_bool() {
        let _ = unsafe { ShowWindow(hwnd, SW_HIDE) };
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OwnershipMode {
    /// Discovery phase: match `owned_pids` first; allow image-name fallback only for brokered apps.
    Discover { allow_image_fallback: bool },
    /// Post-adoption phase: match strictly `owned_pids`, never by image name.
    OwnedOnly,
}

/// Pure ownership predicate deciding if `pid` belongs to the active preview.
#[allow(clippy::too_many_arguments)]
fn belongs_to_preview(
    pid: u32,
    owned_pids: &[u32],
    mode: OwnershipMode,
    launch_exe: &str,
    launch_filetime: u64,
    processes_before: &[u32],
    image_name: Option<&str>,
    process_creation_time: Option<u64>,
) -> bool {
    if pid == 0 {
        return false;
    }
    if owned_pids.contains(&pid) {
        return true;
    }
    let OwnershipMode::Discover {
        allow_image_fallback: true,
    } = mode
    else {
        return false;
    };
    if launch_exe.is_empty() || image_name != Some(launch_exe) {
        return false;
    }
    if processes_before.contains(&pid) {
        return true;
    }
    let min_time = launch_filetime.saturating_sub(PROCESS_TIME_SLACK_100NS);
    process_creation_time.is_some_and(|ct| ct >= min_time)
}

/// Owned PID set after adoption: launched tree union adopted brokered tree (only if newly spawned).
fn owned_pids_for(launch_pid: u32, source_pid: u32, source_is_new: bool) -> Vec<u32> {
    let mut owned = process_tree(launch_pid);
    if source_is_new && source_pid != 0 {
        for pid in process_tree(source_pid) {
            if !owned.contains(&pid) {
                owned.push(pid);
            }
        }
    }
    owned
}

/// One look at the desktop: every window the program has put up since it was begun.
struct Census<'a> {
    launch: &'a Launch,
    processes: &'a [u32],
    images: &'a mut HashMap<u32, Option<String>>,
    creation_times: &'a mut HashMap<u32, Option<u64>>,
    /// The window this app has already adopted, which is the picture and not to be hidden.
    keep: isize,
    /// Hide visible non-`keep` windows only after `source` is adopted (never during discovery).
    hide_others: bool,
    mode: OwnershipMode,
    /// The program's document windows, by the desktop's order, with the area each one covers.
    found: Vec<(i64, HWND)>,
    /// The largest document window the program was showing, and how much of the display it
    /// covers: a splash is a window, and the document is the one that fills a frame.
    shown: Option<(i64, HWND)>,
}

/// Look at the desktop once, on the terms `Census` states.
fn census<'a>(
    launch: &'a Launch,
    processes: &'a [u32],
    images: &'a mut HashMap<u32, Option<String>>,
    creation_times: &'a mut HashMap<u32, Option<u64>>,
    keep: isize,
    hide_others: bool,
    mode: OwnershipMode,
) -> Census<'a> {
    unsafe extern "system" fn visit(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let census = unsafe { &mut *(lparam.0 as *mut Census<'_>) };
        census.visit(hwnd);
        BOOL(1)
    }

    let mut census = Census {
        launch,
        processes,
        images,
        creation_times,
        keep,
        hide_others,
        mode,
        found: Vec::new(),
        shown: None,
    };
    let _ = unsafe { EnumWindows(Some(visit), LPARAM(&mut census as *mut Census<'_> as isize)) };
    census
}

impl Census<'_> {
    fn visit(&mut self, hwnd: HWND) {
        if hwnd.0 as isize == self.keep || hwnd.0.is_null() {
            return;
        }
        let owner = pid_of(hwnd);
        if owner == 0 || !self.belongs(owner) {
            return;
        }
        // Not a window the program put up: it was on the desktop before the program was begun.
        if self.launch.windows_before.contains(&(hwnd.0 as isize)) {
            return;
        }

        let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
        let owner_hwnd = unsafe { GetWindow(hwnd, GW_OWNER) }
            .ok()
            .filter(|o| !o.0.is_null());
        let owned = owner_hwnd.is_some();
        let titled = unsafe { GetWindowTextLengthW(hwnd) } > 0;
        let mut rect = RECT::default();
        let area = if unsafe { GetWindowRect(hwnd, &mut rect) }.is_ok() {
            (rect.right - rect.left) as i64 * (rect.bottom - rect.top) as i64
        } else {
            0
        };
        // **A document window, and not every window a program owns** (see the module note): a
        // helper is untitled or a few pixels square, and is never the picture.
        let document =
            titled && !owned && ex_style & WS_EX_TOOLWINDOW.0 == 0 && area >= THE_SMALLEST_PICTURE;

        if unsafe { IsWindowVisible(hwnd) }.as_bool() {
            if document && self.shown.is_none_or(|(best, _)| area > best) {
                self.shown = Some((area, hwnd));
            }
            if self.hide_others {
                suppress_companion_window(hwnd, self.keep);
            }
        }

        if document {
            self.found.push((area, hwnd));
        }
    }

    fn belongs(&mut self, pid: u32) -> bool {
        if self.processes.contains(&pid) {
            return true;
        }
        let OwnershipMode::Discover {
            allow_image_fallback: true,
        } = self.mode
        else {
            return false;
        };
        let image = image_name_of(pid, self.images);
        let creation = if image.as_deref() == Some(self.launch.exe.as_str())
            && !self.launch.processes_before.contains(&pid)
        {
            creation_time_of(pid, self.creation_times)
        } else {
            None
        };
        belongs_to_preview(
            pid,
            self.processes,
            self.mode,
            &self.launch.exe,
            self.launch.launch_filetime,
            &self.launch.processes_before,
            image.as_deref(),
            creation,
        )
    }

    /// Nothing is read out of the census: what it did is the hiding, and what it found is what the
    /// caller asked for (see `wait_for_a_window`).
    fn take(self) {}
}

thread_local! {
    /// Top-level `(event, HWND)` pairs reported by WinEvent hooks on the worker thread.
    static WINEVENT_EVENTS: RefCell<Vec<(u32, isize)>> = const { RefCell::new(Vec::new()) };
}

/// RAII guard ensuring `UnhookWinEvent` always runs when a WinEvent hook scope exits.
struct WinEventGuard(HWINEVENTHOOK);

impl Drop for WinEventGuard {
    fn drop(&mut self) {
        if !self.0.is_invalid() {
            let _ = unsafe { UnhookWinEvent(self.0) };
        }
        WINEVENT_EVENTS.with(|q| {
            if let Ok(mut list) = q.try_borrow_mut() {
                list.clear();
            }
        });
    }
}

unsafe extern "system" fn win_event_proc(
    _hook: HWINEVENTHOOK,
    event: u32,
    hwnd: HWND,
    id_object: i32,
    id_child: i32,
    _event_thread: u32,
    _event_time: u32,
) {
    if matches!(
        event,
        EVENT_OBJECT_CREATE | EVENT_OBJECT_SHOW | EVENT_OBJECT_LOCATIONCHANGE
    ) && id_object == OBJID_WINDOW.0
        && id_child == CHILDID_SELF as i32
        && !hwnd.0.is_null()
    {
        WINEVENT_EVENTS.with(|q| {
            if let Ok(mut list) = q.try_borrow_mut() {
                list.push((event, hwnd.0 as isize));
            }
        });
    }
}

/// Pre-position a newly created hidden top-level candidate at `provisional` while still hidden so
/// apps that ignore `SW_HIDE` and call `ShowWindow(SW_SHOW)` appear directly at target rect.
fn preposition_hidden_candidate(hwnd: HWND, provisional: (i32, i32, i32, i32)) {
    if !unsafe { IsWindow(hwnd) }.as_bool() || unsafe { IsWindowVisible(hwnd) }.as_bool() {
        return;
    }
    let owned = unsafe { GetWindow(hwnd, GW_OWNER) }
        .ok()
        .is_some_and(|o| !o.0.is_null());
    if owned {
        return;
    }
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    if ex_style & WS_EX_TOOLWINDOW.0 != 0 {
        return;
    }
    let (x, y, w, h) = provisional;
    let _ = unsafe {
        SetWindowPos(
            hwnd,
            HWND_BOTTOM,
            x,
            y,
            w,
            h,
            SWP_NOACTIVATE | SWP_NOOWNERZORDER,
        )
    };
}

fn stage_rect(w: i32, h: i32) -> RECT {
    RECT {
        left: STAGE_COORD,
        top: STAGE_COORD,
        right: STAGE_COORD + w.max(1),
        bottom: STAGE_COORD + h.max(1),
    }
}

fn stage_workspace_rect(w: i32, h: i32) -> RECT {
    workspace_rect_for_screen(stage_rect(w, h))
}

fn is_at_stage_coord(rect: RECT) -> bool {
    (rect.left - STAGE_COORD).abs() <= 64 && (rect.top - STAGE_COORD).abs() <= 64
}

fn should_restore_on_teardown(source_is_new: bool, rect: RECT, vscreen: RECT) -> bool {
    !source_is_new && (is_off_screen(&rect, vscreen) || is_at_stage_coord(rect))
}

/// Immediately stage a newly created or shown top-level candidate at `(STAGE_COORD, STAGE_COORD)`.
fn stage_candidate_offscreen(hwnd: HWND, w: i32, h: i32) {
    if !unsafe { IsWindow(hwnd) }.as_bool() {
        return;
    }
    let owned = unsafe { GetWindow(hwnd, GW_OWNER) }
        .ok()
        .is_some_and(|o| !o.0.is_null());
    if owned {
        return;
    }
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    if ex_style & WS_EX_TOOLWINDOW.0 != 0 {
        return;
    }
    let _ = unsafe {
        SetWindowPos(
            hwnd,
            HWND_BOTTOM,
            STAGE_COORD,
            STAGE_COORD,
            w.max(1),
            h.max(1),
            SWP_NOACTIVATE | SWP_NOOWNERZORDER,
        )
    };
}

/// Non-blocking window discovery for `OffscreenStaged`: never calls `WaitForInputIdle`, stages
/// candidates at `(STAGE_COORD, STAGE_COORD)` on CREATE/SHOW, adopts hidden candidates after 40ms.
fn wait_for_a_window_offscreen_staged(launch: &Launch) -> Option<HWND> {
    let deadline = Instant::now() + WINDOW_WAIT;
    let mut processes = process_tree(launch.pid);
    let mut read = Instant::now();
    let mut images: HashMap<u32, Option<String>> = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();
    let mut first_hidden: Option<(Instant, HWND)> = None;
    let (stage_w, stage_h) = launch
        .provisional
        .map(|(_, _, w, h)| (w.max(1), h.max(1)))
        .unwrap_or((800, 600));

    WINEVENT_EVENTS.with(|q| q.borrow_mut().clear());
    let hook = unsafe {
        SetWinEventHook(
            EVENT_OBJECT_CREATE,
            EVENT_OBJECT_LOCATIONCHANGE,
            None,
            Some(win_event_proc),
            0,
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
        )
    };
    let _guard = WinEventGuard(hook);

    let broker_grace = if launch.packaged_broker {
        Duration::ZERO
    } else {
        BROKER_FALLBACK_GRACE
    };
    let fallback_allowed = |launch: &Launch| -> bool {
        process_has_exited(HANDLE(launch.process.as_raw_handle()))
            || launch.launched_at.elapsed() >= broker_grace
    };

    let check_now = |processes: &[u32],
                     images: &mut HashMap<u32, Option<String>>,
                     creation_times: &mut HashMap<u32, Option<u64>>,
                     first_hidden: &mut Option<(Instant, HWND)>|
     -> Option<HWND> {
        let tree_only = OwnershipMode::Discover {
            allow_image_fallback: false,
        };
        let look = census(
            launch,
            processes,
            images,
            creation_times,
            0,
            false,
            tree_only,
        );
        if let Some((_, window)) = look.shown {
            return Some(window);
        }
        if let Some((_, hidden)) = look.found.iter().copied().max_by_key(|(area, _)| *area) {
            let entry = first_hidden.get_or_insert_with(|| (Instant::now(), hidden));
            entry.1 = hidden;
            if entry.0.elapsed() >= Duration::from_millis(40) {
                return Some(hidden);
            }
            return None;
        }

        if !fallback_allowed(launch) {
            return None;
        }
        let with_fallback = OwnershipMode::Discover {
            allow_image_fallback: true,
        };
        let look = census(
            launch,
            processes,
            images,
            creation_times,
            0,
            false,
            with_fallback,
        );
        if let Some((_, window)) = look.shown {
            return Some(window);
        }
        if let Some((_, hidden)) = look.found.iter().copied().max_by_key(|(area, _)| *area) {
            let entry = first_hidden.get_or_insert_with(|| (Instant::now(), hidden));
            entry.1 = hidden;
            if entry.0.elapsed() >= Duration::from_millis(40) {
                return Some(hidden);
            }
        }
        None
    };

    while Instant::now() < deadline {
        if !is_current(launch.id) {
            return None;
        }
        if let Some(window) = check_now(
            &processes,
            &mut images,
            &mut creation_times,
            &mut first_hidden,
        ) {
            return Some(window);
        }

        if read.elapsed() >= TREE_POLL {
            processes = process_tree(launch.pid);
            read = Instant::now();
        }

        let remaining_ms = deadline
            .saturating_duration_since(Instant::now())
            .as_millis()
            .min(5) as u32;
        if remaining_ms > 0 {
            let _ = unsafe {
                MsgWaitForMultipleObjectsEx(None, remaining_ms, QS_ALLINPUT, MWMO_INPUTAVAILABLE)
            };
        }

        let mut msg = MSG::default();
        while unsafe { PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE) }.as_bool() {
            let _ = unsafe { TranslateMessage(&msg) };
            unsafe { DispatchMessageW(&msg) };

            let allow_fallback = fallback_allowed(launch);
            let mode = OwnershipMode::Discover {
                allow_image_fallback: allow_fallback,
            };
            let mut matched = false;
            WINEVENT_EVENTS.with(|q| {
                for (event, raw) in q.borrow_mut().drain(..) {
                    if launch.windows_before.contains(&raw) {
                        continue;
                    }
                    let hwnd = HWND(raw as *mut _);
                    let pid = pid_of(hwnd);
                    if pid == 0 {
                        continue;
                    }
                    let belongs = if processes.contains(&pid) {
                        true
                    } else if allow_fallback {
                        let image = image_name_of(pid, &mut images);
                        let creation = if image.as_deref() == Some(launch.exe.as_str())
                            && !launch.processes_before.contains(&pid)
                        {
                            creation_time_of(pid, &mut creation_times)
                        } else {
                            None
                        };
                        belongs_to_preview(
                            pid,
                            &processes,
                            mode,
                            &launch.exe,
                            launch.launch_filetime,
                            &launch.processes_before,
                            image.as_deref(),
                            creation,
                        )
                    } else {
                        false
                    };
                    if belongs {
                        if event == EVENT_OBJECT_CREATE || event == EVENT_OBJECT_SHOW {
                            stage_candidate_offscreen(hwnd, stage_w, stage_h);
                            restore_foreground(launch.previous_foreground, &processes);
                        }
                        matched = true;
                    }
                }
            });
            if matched {
                if let Some(window) = check_now(
                    &processes,
                    &mut images,
                    &mut creation_times,
                    &mut first_hidden,
                ) {
                    return Some(window);
                }
            }
        }
    }
    first_hidden.map(|(_, hwnd)| hwnd)
}

struct OffscreenStagedState<'a> {
    invisible_layered: &'a mut bool,
    placed_mode: &'a mut Option<PlacementMode>,
    placed_host_rect: &'a mut Option<(i32, i32, i32, i32)>,
    expected_rect: &'a mut Option<RECT>,
    the_ask: &'a mut Option<(i32, i32)>,
    box_of_the_host: &'a mut Option<(i32, i32, i32, i32)>,
    registered: &'a mut isize,
}

/// Strictly isolated stage -> settle at `(STAGE_COORD, STAGE_COORD)` -> reveal phase for `OffscreenStaged`.
/// Returns `false` if launch was superseded or `source` was destroyed.
#[allow(clippy::too_many_arguments)]
fn run_offscreen_staged_settle_and_reveal(
    host: HWND,
    source: HWND,
    launch: &Launch,
    processes: &[u32],
    images: &mut HashMap<u32, Option<String>>,
    creation_times: &mut HashMap<u32, Option<u64>>,
    was_hidden: bool,
    own: (u32, u32),
    window_mode: CustomWindowMode,
    current_target_rect: &dyn Fn() -> Option<(i32, i32, i32, i32)>,
    state: &mut OffscreenStagedState<'_>,
) -> bool {
    let vscreen = virtual_screen();
    let initial_target = current_target_rect();
    // Save an on-screen restore rect for reused non-owned windows so teardown never leaves them at STAGE_COORD.
    let saved_restore = window_rect(source)
        .filter(|r| !is_off_screen(r, vscreen) && !is_at_stage_coord(*r))
        .or_else(|| {
            initial_target.map(|(x, y, w, h)| RECT {
                left: x,
                top: y,
                right: x + w.max(1),
                bottom: y + h.max(1),
            })
        })
        .unwrap_or(RECT {
            left: 100,
            top: 100,
            right: 900,
            bottom: 700,
        });
    if let Ok(mut slot) = SOURCE_RESTORE_RECT.lock() {
        *slot = Some(saved_restore);
    }

    let (stage_w, stage_h) = initial_target
        .map(|(_, _, w, h)| (w.max(1), h.max(1)))
        .unwrap_or((own.0.max(1) as i32, own.1.max(1) as i32));

    WINEVENT_EVENTS.with(|q| q.borrow_mut().clear());
    let settle_hook = unsafe {
        SetWinEventHook(
            EVENT_OBJECT_CREATE,
            EVENT_OBJECT_LOCATIONCHANGE,
            None,
            Some(win_event_proc),
            0,
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
        )
    };
    let settle_guard = WinEventGuard(settle_hook);

    if was_hidden {
        // Reveal atomically at `(STAGE_COORD, STAGE_COORD)` so zero pixels ever appear on-screen.
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags: WINDOWPLACEMENT_FLAGS(0),
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: stage_workspace_rect(stage_w, stage_h),
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(source, &placement) };
        let _ = unsafe {
            SetWindowPos(
                source,
                HWND_BOTTOM,
                STAGE_COORD,
                STAGE_COORD,
                stage_w,
                stage_h,
                SWP_NOACTIVATE,
            )
        };
        restore_foreground(launch.previous_foreground, processes);
    } else {
        if unsafe { IsZoomed(source) }.as_bool() || unsafe { IsIconic(source) }.as_bool() {
            let placement = WINDOWPLACEMENT {
                length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
                flags: WINDOWPLACEMENT_FLAGS(0),
                showCmd: SW_SHOWNOACTIVATE.0 as u32,
                rcNormalPosition: stage_workspace_rect(stage_w, stage_h),
                ..Default::default()
            };
            let _ = unsafe { SetWindowPlacement(source, &placement) };
        }
        let _ = unsafe {
            SetWindowPos(
                source,
                HWND_BOTTOM,
                STAGE_COORD,
                STAGE_COORD,
                stage_w,
                stage_h,
                SWP_NOACTIVATE,
            )
        };
        restore_foreground(launch.previous_foreground, processes);
    }

    // Suppress any owned companion windows strictly in `processes`.
    census(
        launch,
        processes,
        images,
        creation_times,
        source.0 as isize,
        true,
        OwnershipMode::OwnedOnly,
    )
    .take();

    let _ = drain_settle_events(source, launch, processes);
    let mut staged_expected = window_rect(source);
    // If app immediately clamped our first move back onto the monitor, don't fight it in a loop.
    let stage_refused = staged_expected.is_some_and(|r| !is_off_screen(&r, vscreen));

    let settle_start = Instant::now();
    let mut last_activity = settle_start;
    let settle_deadline = settle_start + Duration::from_millis(550);

    while !stage_refused && Instant::now() < settle_deadline {
        if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
            return false;
        }
        if current_target_rect().is_none() {
            break;
        }

        let raw_moved = drain_settle_events(source, launch, processes);
        let cur_rect = window_rect(source);
        let drifted = (raw_moved && cur_rect != staged_expected)
            || cur_rect.is_some_and(|r| !is_off_screen(&r, vscreen));
        if drifted {
            let _ = unsafe {
                SetWindowPos(
                    source,
                    HWND_BOTTOM,
                    STAGE_COORD,
                    STAGE_COORD,
                    stage_w,
                    stage_h,
                    SWP_NOACTIVATE,
                )
            };
            restore_foreground(launch.previous_foreground, processes);
            let _ = drain_settle_events(source, launch, processes);
            staged_expected = window_rect(source);
            if staged_expected.is_some_and(|r| !is_off_screen(&r, vscreen)) {
                break;
            }
            last_activity = Instant::now();
        } else if last_activity.elapsed() >= Duration::from_millis(100) {
            let answered = unsafe {
                SendMessageTimeoutW(
                    source,
                    WM_NULL,
                    WPARAM(0),
                    LPARAM(0),
                    SMTO_ABORTIFHUNG | SMTO_NORMAL,
                    50,
                    None,
                )
            }
            .0 != 0;
            if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
                return false;
            }
            if answered && window_rect(source) == staged_expected {
                break;
            }
        }

        let slice_ms = settle_deadline
            .saturating_duration_since(Instant::now())
            .as_millis()
            .min(5) as u32;
        if slice_ms > 0 {
            let _ = unsafe {
                MsgWaitForMultipleObjectsEx(None, slice_ms, QS_ALLINPUT, MWMO_INPUTAVAILABLE)
            };
        }
    }
    drop(settle_guard);

    if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
        return false;
    }

    let Some(host_rect) = current_target_rect() else {
        return true;
    };
    if let Ok(mut slot) = LAST_HOST_RECT.lock() {
        *slot = Some(host_rect);
    }
    let (hx, hy, hw, hh) = host_rect;

    // Branch A: try `WS_EX_LAYERED` + `LWA_ALPHA = 1` WHILE STILL AT `(STAGE_COORD, STAGE_COORD)`.
    if make_source_invisible(source.0 as isize) {
        *state.invisible_layered = true;
        SOURCE_INVISIBLE_LAYERED.store(true, Ordering::SeqCst);
        place_invisible_layered(source, host_rect, window_mode, state.the_ask);
        if window_rect(source).is_some_and(|r| is_off_screen(&r, vscreen)) {
            let (aw, ah) = state.the_ask.unwrap_or((hw, hh));
            let _ = unsafe { SetWindowPos(source, HWND_BOTTOM, hx, hy, aw, ah, SWP_NOACTIVATE) };
            let _ = unsafe { InvalidateRect(source, None, false) };
            let _ = unsafe { UpdateWindow(source) };
        }
        keep_source_invisible(source.0 as isize);
        restore_foreground(launch.previous_foreground, processes);
        *state.placed_mode = Some(PlacementMode::InvisibleLayered);
        *state.placed_host_rect = Some(host_rect);
        *state.expected_rect = window_rect(source);
        return true;
    }

    // Branch B: `WS_EX_NOREDIRECTIONBITMAP` GPU/Avalonia/WinUI apps.
    // 1. Resize while still at `(STAGE_COORD, STAGE_COORD)`.
    let insets = if window_mode == CustomWindowMode::Normal {
        dwm_frame_insets(source)
    } else {
        (0, 0, 0, 0)
    };
    let (_, _, ww, wh) = behind_host_window_rect(host_rect, insets);
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND_BOTTOM,
            STAGE_COORD,
            STAGE_COORD,
            ww,
            wh,
            SWP_NOACTIVATE,
        )
    };

    // 2. Size `host`, bind DWM thumbnail, update properties, and SHOW `host` FIRST synchronously
    // so our opaque topmost host already covers `(hx, hy, hw, hh)` before `source` moves on-screen.
    let _ = unsafe {
        SetWindowPos(
            host,
            HWND::default(),
            hx,
            hy,
            hw,
            hh,
            SWP_NOZORDER | SWP_NOACTIVATE,
        )
    };
    *state.box_of_the_host = Some(host_rect);

    if *state.registered == 0 {
        if let Ok(registration) = unsafe { DwmRegisterThumbnail(host, source) } {
            *state.registered = registration;
            if let Ok(mut held) = THUMBNAIL.lock() {
                *held = registration;
            }
        }
    }
    if *state.registered != 0 {
        let mut client = RECT::default();
        let _ = unsafe { GetClientRect(host, &mut client) };
        let picture = if window_mode == CustomWindowMode::Normal {
            unsafe { DwmQueryThumbnailSourceSize(*state.registered) }
                .ok()
                .filter(|s| s.cx > 0 && s.cy > 0)
                .map(|s| (s.cx, s.cy))
                .unwrap_or_else(|| {
                    let (w, h) = source_size(source, window_mode);
                    (w as i32, h as i32)
                })
        } else {
            let (w, h) = source_size(source, window_mode);
            (w as i32, h as i32)
        };
        let properties = thumbnail_properties(client.right, client.bottom, picture, window_mode);
        if unsafe { DwmUpdateThumbnailProperties(*state.registered, &properties) }.is_ok() {
            SHOWN.store(true, Ordering::SeqCst);
            let _ = unsafe {
                SendMessageTimeoutW(
                    host,
                    WM_THE_PICTURE_IS_READY,
                    WPARAM(0),
                    LPARAM(0),
                    SMTO_ABORTIFHUNG | SMTO_NORMAL,
                    100,
                    None,
                )
            };
        }
    }

    // 3. Move `source` straight onto real screen behind the already-visible `host`.
    let mode = if place_behind_host(source, host_rect, window_mode) {
        PlacementMode::BehindHost
    } else {
        let ask = the_ask_for((hw, hh), source_size(source, window_mode));
        *state.the_ask = Some(ask);
        let _ = unsafe {
            SetWindowPos(
                source,
                HWND_BOTTOM,
                STAGE_COORD,
                STAGE_COORD,
                ask.0,
                ask.1,
                SWP_NOACTIVATE,
            )
        };
        let _ = unsafe { InvalidateRect(source, None, false) };
        let _ = unsafe { UpdateWindow(source) };
        PlacementMode::OffScreen
    };
    restore_foreground(launch.previous_foreground, processes);
    *state.placed_mode = Some(mode);
    *state.placed_host_rect = Some(host_rect);
    *state.expected_rect = window_rect(source);
    true
}

// --- CustomCaptureTechnique::InProcessHook ---

const HOOK_MAGIC: u32 = 0x5248_5048;
const HOOK_VERSION: u32 = 1;
const CMD_UNCLOAK: usize = 1;
const CMD_CLOAK: usize = 2;
const CMD_DETACH: usize = 3;
const HOOK_CONTROL_MSG_NAME: PCWSTR = w!("RhpCustomPreviewHookControl");

#[repr(C)]
struct HookSharedState {
    magic: u32,
    version: u32,
    target_pid: u32,
    uncloak_msg: u32,
    cloak_enabled: AtomicU32,
    rect_locked: AtomicU32,
    x: AtomicI32,
    y: AtomicI32,
    w: AtomicI32,
    h: AtomicI32,
    captured_hwnd: AtomicIsize,
    hook_hits: AtomicU32,
}

/// Shared memory mapping for `InProcessHook`; safe to clone across `hide()` and worker thread.
struct HookSharedView {
    _mapping: OwnedHandle,
    view: *mut HookSharedState,
    uncloak_msg: u32,
}

// SAFETY: `view` points to a process-wide file mapping whose fields are atomics / immutable header.
unsafe impl Send for HookSharedView {}
unsafe impl Sync for HookSharedView {}

impl Drop for HookSharedView {
    fn drop(&mut self) {
        if !self.view.is_null() {
            let _ = unsafe {
                UnmapViewOfFile(MEMORY_MAPPED_VIEW_ADDRESS {
                    Value: self.view as *mut _,
                })
            };
        }
    }
}

impl HookSharedView {
    fn create(
        pid: u32,
        provisional: Option<(i32, i32, i32, i32)>,
        uncloak_msg: u32,
    ) -> Option<Self> {
        let name = wide_null(OsStr::new(&format!("Local\\RhpCustomPreviewHook_{pid}")));
        let size = std::mem::size_of::<HookSharedState>() as u32;
        let mapping = unsafe {
            CreateFileMappingW(
                INVALID_HANDLE_VALUE,
                None,
                PAGE_READWRITE,
                0,
                size,
                PCWSTR(name.as_ptr()),
            )
        }
        .ok()?;
        let mapping = unsafe { OwnedHandle::from_raw_handle(mapping.0) };
        let mapped = unsafe {
            MapViewOfFile(
                HANDLE(mapping.as_raw_handle()),
                FILE_MAP_ALL_ACCESS,
                0,
                0,
                size as usize,
            )
        };
        if mapped.Value.is_null() {
            return None;
        }
        let ptr = mapped.Value as *mut HookSharedState;
        let (w, h) = provisional.map_or((0, 0), |(_, _, w, h)| (w.max(1), h.max(1)));
        unsafe {
            ptr.write(HookSharedState {
                magic: HOOK_MAGIC,
                version: HOOK_VERSION,
                target_pid: pid,
                uncloak_msg,
                cloak_enabled: AtomicU32::new(1),
                rect_locked: AtomicU32::new(1),
                x: AtomicI32::new(STAGE_COORD),
                y: AtomicI32::new(STAGE_COORD),
                w: AtomicI32::new(w),
                h: AtomicI32::new(h),
                captured_hwnd: AtomicIsize::new(0),
                hook_hits: AtomicU32::new(0),
            });
        }
        Some(Self {
            _mapping: mapping,
            view: ptr,
            uncloak_msg,
        })
    }

    #[inline]
    fn state(&self) -> &HookSharedState {
        unsafe { &*self.view }
    }

    fn set_rect(&self, x: i32, y: i32, w: i32, h: i32) {
        let s = self.state();
        s.x.store(x, Ordering::Release);
        s.y.store(y, Ordering::Release);
        s.w.store(w, Ordering::Release);
        s.h.store(h, Ordering::Release);
    }

    fn stage_offscreen(&self) {
        let s = self.state();
        s.x.store(STAGE_COORD, Ordering::Release);
        s.y.store(STAGE_COORD, Ordering::Release);
    }

    fn set_captured_hwnd(&self, hwnd: HWND) {
        self.state()
            .captured_hwnd
            .store(hwnd.0 as isize, Ordering::Release);
    }

    fn captured_hwnd(&self) -> Option<HWND> {
        let raw = self.state().captured_hwnd.load(Ordering::Acquire);
        (raw != 0).then_some(HWND(raw as *mut _))
    }

    fn is_cloaked(&self) -> bool {
        self.state().cloak_enabled.load(Ordering::Acquire) == 1
    }

    fn hook_hits(&self) -> u32 {
        self.state().hook_hits.load(Ordering::Acquire)
    }

    fn ensure_cloaked_and_locked(&self) {
        let s = self.state();
        s.cloak_enabled.store(1, Ordering::Release);
        s.rect_locked.store(1, Ordering::Release);
    }

    /// Synchronously cloak via `SendMessageTimeoutW`; `WH_CALLWNDPROC` never sees `PostMessageW`.
    /// Returns `true` only if `rhp_callwnd_proc` actually handled `CMD_CLOAK` on `hwnd`'s thread.
    fn cloak_sync(&self, hwnd: HWND) -> bool {
        self.state().cloak_enabled.store(1, Ordering::Release);
        if hwnd.0.is_null() || !unsafe { IsWindow(hwnd) }.as_bool() {
            return false;
        }
        let hits_before = self.hook_hits();
        let sent = unsafe {
            SendMessageTimeoutW(
                hwnd,
                self.uncloak_msg,
                WPARAM(CMD_CLOAK),
                LPARAM(0),
                SMTO_ABORTIFHUNG | SMTO_NORMAL,
                50,
                None,
            )
        }
        .0 != 0;
        sent && self.hook_hits() > hits_before
    }

    fn uncloak_sync(&self, hwnd: HWND) {
        self.state().cloak_enabled.store(0, Ordering::Release);
        if !hwnd.0.is_null() && unsafe { IsWindow(hwnd) }.as_bool() {
            let _ = unsafe {
                SendMessageTimeoutW(
                    hwnd,
                    self.uncloak_msg,
                    WPARAM(CMD_UNCLOAK),
                    LPARAM(0),
                    SMTO_ABORTIFHUNG | SMTO_NORMAL,
                    150,
                    None,
                )
            };
        }
    }

    /// Clear cloak + rect lock and synchronously remove the in-process subclass hook (`CMD_DETACH`).
    fn detach_sync(&self, hwnd: HWND) -> bool {
        let s = self.state();
        s.cloak_enabled.store(0, Ordering::Release);
        s.rect_locked.store(0, Ordering::Release);
        if hwnd.0.is_null() || !unsafe { IsWindow(hwnd) }.as_bool() {
            return false;
        }
        let hits_before = self.hook_hits();
        let sent = unsafe {
            SendMessageTimeoutW(
                hwnd,
                self.uncloak_msg,
                WPARAM(CMD_DETACH),
                LPARAM(0),
                SMTO_ABORTIFHUNG | SMTO_NORMAL,
                100,
                None,
            )
        }
        .0 != 0;
        sent && self.hook_hits() > hits_before
    }
}

/// Active hook handles + DLL module + shared memory for `CustomCaptureTechnique::InProcessHook`.
struct InProcessHookSession {
    view: Arc<HookSharedView>,
    cbt_hook: HHOOK,
    callwnd_hook: HHOOK,
    dll_module: HMODULE,
}

// SAFETY: `HHOOK` and `HMODULE` are process-wide Win32 handles; `UnhookWindowsHookEx` / `FreeLibrary`
// are safe to call from the worker thread on teardown.
unsafe impl Send for InProcessHookSession {}

impl Drop for InProcessHookSession {
    fn drop(&mut self) {
        if let Ok(mut active) = ACTIVE_HOOK_CONTROL.lock() {
            if active
                .as_ref()
                .is_some_and(|ctrl| Arc::ptr_eq(ctrl, &self.view))
            {
                *active = None;
            }
        }
        if !self.cbt_hook.0.is_null() {
            let _ = unsafe { UnhookWindowsHookEx(self.cbt_hook) };
        }
        if !self.callwnd_hook.0.is_null() {
            let _ = unsafe { UnhookWindowsHookEx(self.callwnd_hook) };
        }
        if !self.dll_module.0.is_null() {
            let _ = unsafe { FreeLibrary(self.dll_module) };
        }
    }
}

/// Beside the executable first, so a portable zip carries its own copy; then the per-user folder,
/// which is where a lone executable's copy has to be put by hand. Tests run from
/// `target/<profile>/deps/`, and the DLL that build made is in `target/<profile>/` above them.
fn find_hook_dll() -> Option<PathBuf> {
    let dir = std::env::current_exe().ok()?.parent()?.to_path_buf();
    hook_dll_candidates(&dir, user_hook_folder())
        .into_iter()
        .find(|path| path.is_file())
}

/// The order `find_hook_dll` searches, as paths: beside the executable, one level above it, then
/// the per-user folder.
fn hook_dll_candidates(dir: &Path, user_folder: Option<PathBuf>) -> Vec<PathBuf> {
    const NAME: &str = "custom_preview_hook.dll";
    let mut candidates = vec![dir.join(NAME)];
    if let Some(parent) = dir.parent() {
        candidates.push(parent.join(NAME));
    }
    candidates.extend(user_folder.map(|folder| folder.join(NAME)));
    candidates
}

/// `<local appdata>\rust-hover-preview` — local and not the roaming folder `config.ini` lives in,
/// which a domain profile syncs and no binary belongs in.
fn user_hook_folder() -> Option<PathBuf> {
    if let Some(portable) = crate::paths::data_root() {
        return Some(portable.to_path_buf());
    }
    directories::BaseDirs::new().map(|base| base.data_local_dir().join("rust-hover-preview"))
}

/// Whether a launch would find the hook DLL, read once for the preset editor's Capture Technique
/// row. The DLL does not move while this app runs, and a hover searches for it again anyway.
pub(crate) fn hook_dll_present() -> bool {
    *HOOK_DLL_PRESENT.get_or_init(|| find_hook_dll().is_some())
}

static HOOK_DLL_PRESENT: std::sync::OnceLock<bool> = std::sync::OnceLock::new();

fn load_hook_dll() -> Option<(HMODULE, HOOKPROC, HOOKPROC)> {
    let path = find_hook_dll()?;
    let wide = wide_null(path.as_os_str());
    let hmod = unsafe { LoadLibraryW(PCWSTR(wide.as_ptr())) }.ok()?;
    let cbt_raw = unsafe { GetProcAddress(hmod, PCSTR(c"rhp_cbt_proc".as_ptr().cast())) };
    let cwp_raw = unsafe { GetProcAddress(hmod, PCSTR(c"rhp_callwnd_proc".as_ptr().cast())) };
    match (cbt_raw, cwp_raw) {
        (Some(cbt_fn), Some(cwp_fn)) => {
            // SAFETY: Both exported functions match Win32 `HOOKPROC` (`extern "system" fn(i32, WPARAM, LPARAM) -> LRESULT`).
            let cbt_proc: HOOKPROC = Some(unsafe {
                std::mem::transmute::<
                    unsafe extern "system" fn() -> isize,
                    unsafe extern "system" fn(i32, WPARAM, LPARAM) -> LRESULT,
                >(cbt_fn)
            });
            let cwp_proc: HOOKPROC = Some(unsafe {
                std::mem::transmute::<
                    unsafe extern "system" fn() -> isize,
                    unsafe extern "system" fn(i32, WPARAM, LPARAM) -> LRESULT,
                >(cwp_fn)
            });
            Some((hmod, cbt_proc, cwp_proc))
        }
        _ => {
            let _ = unsafe { FreeLibrary(hmod) };
            None
        }
    }
}

fn try_attach_thread_hooks(
    hmod: HMODULE,
    cbt_proc: HOOKPROC,
    cwp_proc: HOOKPROC,
    tid: u32,
) -> Option<(HHOOK, HHOOK)> {
    let hinst = HINSTANCE(hmod.0);
    let cbt = unsafe { SetWindowsHookExW(WH_CBT, cbt_proc, hinst, tid) }.ok()?;
    match unsafe { SetWindowsHookExW(WH_CALLWNDPROC, cwp_proc, hinst, tid) } {
        Ok(cwp) => Some((cbt, cwp)),
        Err(_) => {
            let _ = unsafe { UnhookWindowsHookEx(cbt) };
            None
        }
    }
}

const IN_PROCESS_HOOK_SHOW_CMD: u16 = SW_HIDE.0 as u16;

/// Spawn target with `CREATE_SUSPENDED`, initialize per-PID shared memory, and attach `WH_CBT` +
/// `WH_CALLWNDPROC` before/immediately after `ResumeThread`. Falls back to `None` session
/// (caller degrades cleanly to `InvisibleLayered`) if DLL is missing or hook cannot attach.
fn launch_in_process_hook(
    exe: &str,
    path: &Path,
    provisional: Option<(i32, i32, i32, i32)>,
) -> Option<(OwnedHandle, u32, Option<InProcessHookSession>)> {
    let Some((hmod, cbt_proc, cwp_proc)) = load_hook_dll() else {
        let (proc_handle, pid) = launch(
            exe,
            path,
            provisional,
            CustomCaptureTechnique::InvisibleLayered,
        )?;
        return Some((proc_handle, pid, None));
    };

    let application = wide_null(OsStr::new(exe));
    let mut cmdline = build_command_line(OsStr::new(exe), path.as_os_str());
    let cwd_wide = Path::new(exe)
        .parent()
        .filter(|dir| !dir.as_os_str().is_empty())
        .map(|dir| wide_null(dir.as_os_str()));
    let cwd_ptr = cwd_wide
        .as_ref()
        .map_or(PCWSTR::null(), |buf| PCWSTR(buf.as_ptr()));

    // Off-screen positive STARTUPINFOW + SW_HIDE (u32-safe); hook locks real creation to STAGE_COORD.
    let mut si = launch_startup_info_for(
        provisional,
        CustomCaptureTechnique::InProcessHook,
        virtual_screen(),
    );
    si.wShowWindow = IN_PROCESS_HOOK_SHOW_CMD;

    let mut pi = PROCESS_INFORMATION::default();
    if unsafe {
        CreateProcessW(
            PCWSTR(application.as_ptr()),
            PWSTR(cmdline.as_mut_ptr()),
            None,
            None,
            BOOL(0),
            CREATE_SUSPENDED,
            None,
            cwd_ptr,
            &si,
            &mut pi,
        )
    }
    .is_err()
    {
        let _ = unsafe { FreeLibrary(hmod) };
        return None;
    }

    // SAFETY: `pi.hProcess` is uniquely owned and closed on `OwnedHandle` drop.
    let process = unsafe { OwnedHandle::from_raw_handle(pi.hProcess.0) };
    let pid = pi.dwProcessId;
    let tid = pi.dwThreadId;
    let uncloak_msg = unsafe { RegisterWindowMessageW(HOOK_CONTROL_MSG_NAME) };
    let view = HookSharedView::create(pid, provisional, uncloak_msg).map(Arc::new);

    // Attempt hook install before `ResumeThread`; if Win11 rejects a suspended thread before
    // `W32THREAD` init (`ERROR_INVALID_THREAD_ID`), spin-yield 30ms (sub-0.1ms attach; avoids
    // 15.6ms timer-tick sleep) then sleep(1ms) up to 200ms.
    let mut hooks = if view.is_some() {
        try_attach_thread_hooks(hmod, cbt_proc, cwp_proc, tid)
    } else {
        None
    };

    let _ = unsafe { ResumeThread(pi.hThread) };
    let _ = unsafe { CloseHandle(pi.hThread) };

    if view.is_some() && hooks.is_none() {
        let retry_start = Instant::now();
        let spin_deadline = retry_start + Duration::from_millis(30);
        let retry_deadline = retry_start + Duration::from_millis(200);
        while Instant::now() < retry_deadline {
            if process_has_exited(HANDLE(process.as_raw_handle())) {
                break;
            }
            if let Some(pair) = try_attach_thread_hooks(hmod, cbt_proc, cwp_proc, tid) {
                hooks = Some(pair);
                break;
            }
            if Instant::now() < spin_deadline {
                std::thread::yield_now();
            } else {
                std::thread::sleep(Duration::from_millis(1));
            }
        }
    }

    let session = match (view, hooks) {
        (Some(view), Some((cbt_hook, callwnd_hook))) => Some(InProcessHookSession {
            view,
            cbt_hook,
            callwnd_hook,
            dll_module: hmod,
        }),
        _ => {
            let _ = unsafe { FreeLibrary(hmod) };
            None
        }
    };
    Some((process, pid, session))
}

/// Validate a `captured_hwnd` candidate: live, unowned, non-toolwindow, and titled or >= 64x64.
fn validate_captured_hwnd(hwnd: HWND) -> Option<(i64, bool, HWND)> {
    if !unsafe { IsWindow(hwnd) }.as_bool() {
        return None;
    }
    if unsafe { GetWindow(hwnd, GW_OWNER) }
        .ok()
        .is_some_and(|owner| !owner.0.is_null())
    {
        return None;
    }
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    if (ex_style & WS_EX_TOOLWINDOW.0) != 0 {
        return None;
    }
    let mut rect = RECT::default();
    let area = if unsafe { GetWindowRect(hwnd, &mut rect) }.is_ok() {
        (rect.right - rect.left).max(0) as i64 * (rect.bottom - rect.top).max(0) as i64
    } else {
        0
    };
    let titled = unsafe { GetWindowTextLengthW(hwnd) } > 0;
    if titled || area >= THE_SMALLEST_PICTURE {
        Some((area, titled, hwnd))
    } else {
        None
    }
}

/// Fast discovery for `InProcessHook`: validates `session.captured_hwnd` alongside `census` and
/// prefers a larger titled main window in `processes` when one exists.
fn wait_for_a_window_in_process_hook(
    launch: &Launch,
    session: &InProcessHookSession,
) -> Option<HWND> {
    let deadline = Instant::now() + WINDOW_WAIT;
    let mut processes = process_tree(launch.pid);
    let mut read = Instant::now();
    let mut images: HashMap<u32, Option<String>> = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();
    let mut first_candidate: Option<(Instant, HWND)> = None;

    while Instant::now() < deadline {
        if !is_current(launch.id) {
            return None;
        }
        let valid_captured = session
            .view
            .captured_hwnd()
            .and_then(validate_captured_hwnd);

        let look = census(
            launch,
            &processes,
            &mut images,
            &mut creation_times,
            0,
            false,
            OwnershipMode::Discover {
                allow_image_fallback: false,
            },
        );
        if let Some((shown_area, shown_hwnd)) = look.shown {
            if let Some((cap_area, true, cap_hwnd)) = valid_captured {
                if cap_area > shown_area {
                    return Some(cap_hwnd);
                }
            }
            return Some(shown_hwnd);
        }

        let best_hidden = look.found.iter().copied().max_by_key(|(area, _)| *area);
        if let Some((cap_area, true, cap_hwnd)) = valid_captured {
            if let Some((hidden_area, hidden_hwnd)) = best_hidden {
                if hidden_area > cap_area {
                    return Some(hidden_hwnd);
                }
            }
            if cap_area >= THE_SMALLEST_PICTURE {
                return Some(cap_hwnd);
            }
        }
        if let Some((_, hidden)) = best_hidden {
            let entry = first_candidate.get_or_insert_with(|| (Instant::now(), hidden));
            entry.1 = hidden;
            if entry.0.elapsed() >= Duration::from_millis(40) {
                return Some(hidden);
            }
        } else if let Some((_, _, cap_hwnd)) = valid_captured {
            first_candidate.get_or_insert_with(|| (Instant::now(), cap_hwnd));
        }

        if read.elapsed() >= TREE_POLL {
            processes = process_tree(launch.pid);
            read = Instant::now();
        }
        std::thread::sleep(Duration::from_millis(5));
    }
    first_candidate
        .map(|(_, hwnd)| hwnd)
        .or_else(|| wait_for_a_window(launch))
}

#[allow(clippy::too_many_arguments)]
fn reveal_in_process_hook(
    host: HWND,
    source: HWND,
    launch: &Launch,
    session: &InProcessHookSession,
    processes: &[u32],
    host_rect: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
    state: &mut OffscreenStagedState<'_>,
) {
    if let Ok(mut slot) = LAST_HOST_RECT.lock() {
        *slot = Some(host_rect);
    }
    // Late-hooked fallback (`hook_hits == 0`): source is not cloaked; use InvisibleLayered if active.
    if session.view.hook_hits() == 0 && *state.invisible_layered {
        session
            .view
            .state()
            .cloak_enabled
            .store(0, Ordering::Release);
        place_invisible_layered(source, host_rect, window_mode, state.the_ask);
        restore_foreground(launch.previous_foreground, processes);
        *state.placed_mode = Some(PlacementMode::InvisibleLayered);
        *state.placed_host_rect = Some(host_rect);
        *state.expected_rect = window_rect(source);
        return;
    }

    let (hx, hy, hw, hh) = host_rect;
    let insets = if window_mode == CustomWindowMode::Normal {
        dwm_frame_insets(source)
    } else {
        (0, 0, 0, 0)
    };
    let (wx, wy, ww, wh) = behind_host_window_rect(host_rect, insets);

    // Stage & size at `(STAGE_COORD, STAGE_COORD)` while cloaked before moving anything on-screen.
    session.view.set_rect(STAGE_COORD, STAGE_COORD, ww, wh);
    if !session.view.is_cloaked() {
        let _ = session.view.cloak_sync(source);
    }
    if unsafe { IsZoomed(source) }.as_bool() || unsafe { IsIconic(source) }.as_bool() {
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags: WINDOWPLACEMENT_FLAGS(0),
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: stage_workspace_rect(ww, wh),
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(source, &placement) };
    }
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND_BOTTOM,
            STAGE_COORD,
            STAGE_COORD,
            ww,
            wh,
            SWP_NOACTIVATE,
        )
    };
    if !unsafe { IsWindowVisible(source) }.as_bool() {
        let _ = unsafe { ShowWindow(source, SW_SHOWNOACTIVATE) };
    }

    // Check fit using off-screen dimensions so `source` never touches the monitor before `host`.
    let fits_host = visible_frame_rect(source).is_some_and(|r| {
        let aw = r.right - r.left;
        let ah = r.bottom - r.top;
        aw > 0 && ah > 0 && aw <= hw + ALIGN_TOLERANCE * 2 && ah <= hh + ALIGN_TOLERANCE * 2
    }) && unsafe { IsWindowVisible(source) }.as_bool()
        && !unsafe { IsZoomed(source) }.as_bool()
        && !unsafe { IsIconic(source) }.as_bool();

    if !fits_host {
        let ask = the_ask_for((hw, hh), source_size(source, window_mode));
        *state.the_ask = Some(ask);
        session
            .view
            .set_rect(STAGE_COORD, STAGE_COORD, ask.0, ask.1);
        let _ = unsafe {
            SetWindowPos(
                source,
                HWND_BOTTOM,
                STAGE_COORD,
                STAGE_COORD,
                ask.0,
                ask.1,
                SWP_NOACTIVATE,
            )
        };
    }

    // Size `host`, register thumbnail, and show opaque topmost `host` FIRST.
    if *state.box_of_the_host != Some(host_rect) {
        let _ = unsafe {
            SetWindowPos(
                host,
                HWND::default(),
                hx,
                hy,
                hw,
                hh,
                SWP_NOZORDER | SWP_NOACTIVATE,
            )
        };
        *state.box_of_the_host = Some(host_rect);
    }
    if *state.registered == 0 {
        if let Ok(registration) = unsafe { DwmRegisterThumbnail(host, source) } {
            *state.registered = registration;
            if let Ok(mut held) = THUMBNAIL.lock() {
                *held = registration;
            }
        }
    }
    if *state.registered != 0 {
        let mut client = RECT::default();
        let _ = unsafe { GetClientRect(host, &mut client) };
        let picture = if window_mode == CustomWindowMode::Normal {
            unsafe { DwmQueryThumbnailSourceSize(*state.registered) }
                .ok()
                .filter(|s| s.cx > 0 && s.cy > 0)
                .map(|s| (s.cx, s.cy))
                .unwrap_or_else(|| {
                    let (w, h) = source_size(source, window_mode);
                    (w as i32, h as i32)
                })
        } else {
            let (w, h) = source_size(source, window_mode);
            (w as i32, h as i32)
        };
        let properties = thumbnail_properties(client.right, client.bottom, picture, window_mode);
        if unsafe { DwmUpdateThumbnailProperties(*state.registered, &properties) }.is_ok() {
            SHOWN.store(true, Ordering::SeqCst);
            let _ = unsafe {
                SendMessageTimeoutW(
                    host,
                    WM_THE_PICTURE_IS_READY,
                    WPARAM(0),
                    LPARAM(0),
                    SMTO_ABORTIFHUNG | SMTO_NORMAL,
                    100,
                    None,
                )
            };
        }
    }

    // Wait for DWM to physically composite opaque `host` on screen before moving/uncloaking `source`.
    let _ = unsafe { DwmFlush() };

    let mode = if fits_host {
        session.view.set_rect(wx, wy, ww, wh);
        let _ = unsafe { SetWindowPos(source, HWND_BOTTOM, wx, wy, ww, wh, SWP_NOACTIVATE) };
        session.view.uncloak_sync(source);
        let _ = unsafe { InvalidateRect(source, None, false) };
        let _ = unsafe { UpdateWindow(source) };
        PlacementMode::BehindHost
    } else {
        session.view.uncloak_sync(source);
        let _ = unsafe { InvalidateRect(source, None, false) };
        let _ = unsafe { UpdateWindow(source) };
        PlacementMode::OffScreen
    };

    restore_foreground(launch.previous_foreground, processes);
    *state.placed_mode = Some(mode);
    *state.placed_host_rect = Some(host_rect);
    *state.expected_rect = window_rect(source);
}

#[allow(clippy::too_many_arguments)]
fn run_in_process_hook_settle_and_reveal(
    host: HWND,
    source: HWND,
    launch: &Launch,
    session: &InProcessHookSession,
    processes: &[u32],
    images: &mut HashMap<u32, Option<String>>,
    creation_times: &mut HashMap<u32, Option<u64>>,
    was_hidden: bool,
    window_mode: CustomWindowMode,
    current_target_rect: &dyn Fn() -> Option<(i32, i32, i32, i32)>,
    state: &mut OffscreenStagedState<'_>,
) -> bool {
    let cloak_or_park = |invisible_layered: bool| {
        session.view.stage_offscreen();
        if !session.view.cloak_sync(source) {
            if invisible_layered {
                keep_source_invisible(source.0 as isize);
            } else {
                let _ = unsafe {
                    SetWindowPos(
                        source,
                        HWND_BOTTOM,
                        STAGE_COORD,
                        STAGE_COORD,
                        0,
                        0,
                        SWP_NOSIZE | SWP_NOACTIVATE,
                    )
                };
            }
        }
    };

    if was_hidden {
        let Some(target_rect) = current_target_rect() else {
            // Hover already left while window is still hidden; never show it.
            return true;
        };
        session.view.ensure_cloaked_and_locked();
        let insets = if window_mode == CustomWindowMode::Normal {
            dwm_frame_insets(source)
        } else {
            (0, 0, 0, 0)
        };
        let (_, _, ww, wh) = behind_host_window_rect(target_rect, insets);
        session.view.set_rect(STAGE_COORD, STAGE_COORD, ww, wh);
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags: WINDOWPLACEMENT_FLAGS(0),
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: stage_workspace_rect(ww, wh),
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(source, &placement) };
    }

    // Confirm hook actually ran on `source`'s thread; if late-hooked / secondary-thread UI,
    // attempt one synchronous cloak, else degrade immediately to InvisibleLayered or OffScreen.
    if session.view.hook_hits() == 0 && !session.view.cloak_sync(source) {
        let ok = make_source_invisible(source.0 as isize);
        *state.invisible_layered = ok;
        SOURCE_INVISIBLE_LAYERED.store(ok, Ordering::SeqCst);
        if !ok && !is_window_offscreen(source) {
            push_offscreen(source);
        }
    }

    census(
        launch,
        processes,
        images,
        creation_times,
        source.0 as isize,
        true,
        OwnershipMode::OwnedOnly,
    )
    .take();

    // Settle while `source` is cloaked (`DWMWA_CLOAK = 1`) and rect-locked at `(STAGE_COORD, STAGE_COORD)`.
    let settle_start = Instant::now();
    let settle_deadline = settle_start + Duration::from_millis(500);
    let mut last_rect = window_rect(source);
    let mut last_change = settle_start;

    while Instant::now() < settle_deadline {
        if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
            return false;
        }
        let Some(target_rect) = current_target_rect() else {
            if !is_current(launch.id) {
                return false;
            }
            cloak_or_park(*state.invisible_layered);
            return true;
        };
        let insets = if window_mode == CustomWindowMode::Normal {
            dwm_frame_insets(source)
        } else {
            (0, 0, 0, 0)
        };
        let (_, _, ww, wh) = behind_host_window_rect(target_rect, insets);
        session.view.set_rect(STAGE_COORD, STAGE_COORD, ww, wh);

        let cur = window_rect(source);
        if cur != last_rect {
            last_rect = cur;
            last_change = Instant::now();
        } else if last_change.elapsed() >= Duration::from_millis(100) {
            let answered = unsafe {
                SendMessageTimeoutW(
                    source,
                    WM_NULL,
                    WPARAM(0),
                    LPARAM(0),
                    SMTO_ABORTIFHUNG | SMTO_NORMAL,
                    50,
                    None,
                )
            }
            .0 != 0;
            if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
                return false;
            }
            if answered && window_rect(source) == last_rect {
                break;
            }
        }
        std::thread::sleep(Duration::from_millis(10));
    }

    if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
        return false;
    }
    let Some(host_rect) = current_target_rect() else {
        cloak_or_park(*state.invisible_layered);
        return true;
    };
    reveal_in_process_hook(
        host,
        source,
        launch,
        session,
        processes,
        host_rect,
        window_mode,
        state,
    );
    true
}

/// The program's window, waited for within [`WINDOW_WAIT`] via out-of-context WinEvent hook
/// (`EVENT_OBJECT_CREATE`..`EVENT_OBJECT_SHOW`) with a 5ms `MsgWaitForMultipleObjectsEx` fallback.
fn wait_for_a_window(launch: &Launch) -> Option<HWND> {
    let deadline = Instant::now() + WINDOW_WAIT;
    let mut processes = process_tree(launch.pid);
    let mut read = Instant::now();
    let mut images: HashMap<u32, Option<String>> = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();

    WINEVENT_EVENTS.with(|q| q.borrow_mut().clear());
    let hook = unsafe {
        SetWinEventHook(
            EVENT_OBJECT_CREATE,
            EVENT_OBJECT_SHOW,
            None,
            Some(win_event_proc),
            0,
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
        )
    };
    let _guard = WinEventGuard(hook);

    let broker_grace = if launch.packaged_broker {
        Duration::ZERO
    } else {
        BROKER_FALLBACK_GRACE
    };
    let fallback_allowed = |launch: &Launch| -> bool {
        process_has_exited(HANDLE(launch.process.as_raw_handle()))
            || launch.launched_at.elapsed() >= broker_grace
    };

    let check_now = |processes: &[u32],
                     images: &mut HashMap<u32, Option<String>>,
                     creation_times: &mut HashMap<u32, Option<u64>>|
     -> Option<HWND> {
        // Pass 1: ALWAYS check process_tree(launch.pid) first; never adopt sibling PIDs when tree has a window.
        let tree_only = OwnershipMode::Discover {
            allow_image_fallback: false,
        };
        let look = census(
            launch,
            processes,
            images,
            creation_times,
            0,
            false,
            tree_only,
        );
        if let Some((_, window)) = look.shown {
            return Some(window);
        }
        if let Some((_, hidden)) = look.found.iter().copied().max_by_key(|(area, _)| *area) {
            wait_for_input_idle_handle(HANDLE(launch.process.as_raw_handle()), 250);
            if !is_current(launch.id) {
                return None;
            }
            let look2 = census(
                launch,
                processes,
                images,
                creation_times,
                0,
                false,
                tree_only,
            );
            if let Some((_, window)) = look2.shown {
                return Some(window);
            }
            return look2
                .found
                .iter()
                .copied()
                .max_by_key(|(area, _)| *area)
                .map(|(_, window)| window)
                .or(Some(hidden));
        }

        // Pass 2: Brokered/packaged app fallback ONLY when process_tree has no window AND
        // launcher exited (broker stub) or grace elapsed.
        if !fallback_allowed(launch) {
            return None;
        }
        let with_fallback = OwnershipMode::Discover {
            allow_image_fallback: true,
        };
        let look = census(
            launch,
            processes,
            images,
            creation_times,
            0,
            false,
            with_fallback,
        );
        if let Some((_, window)) = look.shown {
            return Some(window);
        }
        if let Some((_, hidden)) = look.found.iter().copied().max_by_key(|(area, _)| *area) {
            wait_for_input_idle_handle(HANDLE(launch.process.as_raw_handle()), 250);
            if !is_current(launch.id) {
                return None;
            }
            let look2 = census(
                launch,
                processes,
                images,
                creation_times,
                0,
                false,
                with_fallback,
            );
            if let Some((_, window)) = look2.shown {
                return Some(window);
            }
            return look2
                .found
                .iter()
                .copied()
                .max_by_key(|(area, _)| *area)
                .map(|(_, window)| window)
                .or(Some(hidden));
        }
        None
    };

    while Instant::now() < deadline {
        if !is_current(launch.id) {
            return None;
        }

        if let Some(window) = check_now(&processes, &mut images, &mut creation_times) {
            return Some(window);
        }

        if read.elapsed() >= TREE_POLL {
            processes = process_tree(launch.pid);
            read = Instant::now();
        }

        let remaining_ms = deadline
            .saturating_duration_since(Instant::now())
            .as_millis()
            .min(5) as u32;
        if remaining_ms > 0 {
            let _ = unsafe {
                MsgWaitForMultipleObjectsEx(None, remaining_ms, QS_ALLINPUT, MWMO_INPUTAVAILABLE)
            };
        }

        let mut msg = MSG::default();
        while unsafe { PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE) }.as_bool() {
            let _ = unsafe { TranslateMessage(&msg) };
            unsafe { DispatchMessageW(&msg) };

            let allow_fallback = fallback_allowed(launch);
            let mode = OwnershipMode::Discover {
                allow_image_fallback: allow_fallback,
            };
            let mut matched = false;
            WINEVENT_EVENTS.with(|q| {
                for (event, raw) in q.borrow_mut().drain(..) {
                    if launch.windows_before.contains(&raw) {
                        continue;
                    }
                    let hwnd = HWND(raw as *mut _);
                    let pid = pid_of(hwnd);
                    if pid == 0 {
                        continue;
                    }
                    let belongs = if processes.contains(&pid) {
                        true
                    } else if allow_fallback {
                        let image = image_name_of(pid, &mut images);
                        let creation = if image.as_deref() == Some(launch.exe.as_str())
                            && !launch.processes_before.contains(&pid)
                        {
                            creation_time_of(pid, &mut creation_times)
                        } else {
                            None
                        };
                        belongs_to_preview(
                            pid,
                            &processes,
                            mode,
                            &launch.exe,
                            launch.launch_filetime,
                            &launch.processes_before,
                            image.as_deref(),
                            creation,
                        )
                    } else {
                        false
                    };
                    if belongs {
                        if event == EVENT_OBJECT_CREATE {
                            if let Some(provisional) = launch.provisional {
                                preposition_hidden_candidate(hwnd, provisional);
                            }
                        }
                        matched = true;
                    }
                }
            });
            if matched {
                if let Some(window) = check_now(&processes, &mut images, &mut creation_times) {
                    return Some(window);
                }
            }
        }
    }
    None
}

fn filetime_to_u64(ft: FILETIME) -> u64 {
    ((ft.dwHighDateTime as u64) << 32) | (ft.dwLowDateTime as u64)
}

fn current_filetime_u64() -> u64 {
    // Windows FILETIME epoch (1601-01-01) to Unix epoch (1970-01-01) in 100ns ticks.
    const WINDOWS_EPOCH_OFFSET_100NS: u64 = 116_444_736_000_000_000;
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| {
            WINDOWS_EPOCH_OFFSET_100NS
                .saturating_add(d.as_secs().saturating_mul(10_000_000))
                .saturating_add(u64::from(d.subsec_nanos() / 100))
        })
        .unwrap_or(0)
}

fn process_has_exited(process: HANDLE) -> bool {
    unsafe { WaitForSingleObject(process, 0) == WAIT_OBJECT_0 }
}

fn creation_time_of(pid: u32, cache: &mut HashMap<u32, Option<u64>>) -> Option<u64> {
    if let Some(known) = cache.get(&pid) {
        return *known;
    }
    let time = unsafe { process_creation_time(pid) };
    cache.insert(pid, time);
    time
}

unsafe fn process_creation_time(pid: u32) -> Option<u64> {
    let process = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) }.ok()?;
    let mut created = FILETIME::default();
    let mut exited = FILETIME::default();
    let mut kernel = FILETIME::default();
    let mut user = FILETIME::default();
    let ok = unsafe { GetProcessTimes(process, &mut created, &mut exited, &mut kernel, &mut user) }
        .is_ok();
    let _ = unsafe { CloseHandle(process) };
    ok.then(|| filetime_to_u64(created))
}

/// The process and everything it started, as deep as the machine's own parent links go.
///
/// A program that hands its file to a process of its own — a launcher and the window it puts up,
/// a viewer and its renderer — is a program whose window is not the one the launched id owns, and
/// the tree is what finds it.
fn process_tree(pid: u32) -> Vec<u32> {
    /// How many processes are looked up: a tree is a handful, and a bound keeps a snapshot walk
    /// that cannot end a walk that does.
    const THE_LOOKUPS: usize = 8;

    let mut found = vec![pid];
    let mut at = 0;
    let mut lookups = 0;
    while at < found.len() && lookups < THE_LOOKUPS {
        let parent = found[at];
        at += 1;
        lookups += 1;
        for child in engine_processes::children_of(parent) {
            if !found.contains(&child) {
                found.push(child);
            }
        }
    }
    found
}

/// True when `hwnd` is a standard Win32 dialog (`#32770`).
fn is_dialog_class(hwnd: HWND) -> bool {
    let mut buf = [0u16; 32];
    let len = unsafe { GetClassNameW(hwnd, &mut buf) } as usize;
    len == 6 && String::from_utf16_lossy(&buf[..len]) == "#32770"
}

/// The window's process.
fn pid_of(window: HWND) -> u32 {
    let mut pid = 0u32;
    unsafe { GetWindowThreadProcessId(window, Some(&mut pid)) };
    pid
}

/// Every top-level window on the desktop, as handles a `Send` fact can hold.
fn top_level_windows() -> Vec<isize> {
    unsafe extern "system" fn visit(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let found = unsafe { &mut *(lparam.0 as *mut Vec<isize>) };
        found.push(hwnd.0 as isize);
        BOOL(1)
    }

    let mut found: Vec<isize> = Vec::new();
    let _ = unsafe { EnumWindows(Some(visit), LPARAM(&mut found as *mut Vec<isize> as isize)) };
    found
}

/// The image a process is running, as its file name and lowercased, or nothing where the machine
/// will not say.
///
/// Asked once per process per launch: the answer does not change while the process lives, and
/// every window on the desktop is asked about on every look.
fn image_name_of(pid: u32, images: &mut HashMap<u32, Option<String>>) -> Option<String> {
    if let Some(known) = images.get(&pid) {
        return known.clone();
    }
    let name = unsafe { process_image_name(pid) };
    images.insert(pid, name.clone());
    name
}

/// Safety: opens the process for the query alone and closes it again; nothing is dereferenced.
unsafe fn process_image_name(pid: u32) -> Option<String> {
    let process = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) }.ok()?;
    let mut buffer = [0u16; 1024];
    let mut length = buffer.len() as u32;
    let queried = unsafe {
        QueryFullProcessImageNameW(
            process,
            PROCESS_NAME_WIN32,
            windows::core::PWSTR(buffer.as_mut_ptr()),
            &mut length,
        )
    };
    let _ = unsafe { CloseHandle(process) };
    queried.ok()?;

    let path = String::from_utf16_lossy(&buffer[..length as usize]);
    Some(
        path.rsplit(['\\', '/'])
            .next()
            .unwrap_or(&path)
            .to_lowercase(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn pointer(cursor: (i32, i32)) -> Where {
        Where::Pointer {
            cursor,
            avoid: None,
        }
    }

    #[test]
    fn a_hundred_percent_of_a_program_window_is_the_room_cut_to_its_shape() {
        let room = work_area_at(0, 0);
        let scale = resolve_scale(
            Path::new("sample.png"),
            CustomScaling::Override(PreviewScale::Percent(100)),
        );
        let layout = layout_for(pointer((0, 0)), (600, 400), scale, true, false)
            .expect("a display has room for a preview");

        assert!(
            layout.preview_w as i32 > 600 || layout.preview_h as i32 > 400,
            "the share is of the room rather than of the size the program's window opens at"
        );
        assert!(
            layout.pos_x >= room.left
                && layout.pos_x + layout.preview_w as i32 <= room.right
                && layout.pos_y >= room.top
                && layout.pos_y + layout.preview_h as i32 <= room.bottom,
            "and the box it comes out at is one the display holds"
        );
        let shape = layout.preview_w as f32 / layout.preview_h as f32;
        assert!(
            (shape - 600.0 / 400.0).abs() < 0.01,
            "cut to the window's own shape, which is {shape} rather than 1.5"
        );
    }

    #[test]
    fn a_percentage_is_that_share_of_the_room() {
        let whole_scale = resolve_scale(
            Path::new("sample.png"),
            CustomScaling::Override(PreviewScale::Percent(100)),
        );
        let half_scale = resolve_scale(
            Path::new("sample.png"),
            CustomScaling::Override(PreviewScale::Percent(50)),
        );
        let whole = layout_for(pointer((0, 0)), (600, 400), whole_scale, true, false)
            .expect("a display has room for a preview");
        let half = layout_for(pointer((0, 0)), (600, 400), half_scale, true, false)
            .expect("half of a room is a preview a display can hold");

        assert!(
            (half.preview_w as f32 - whole.preview_w as f32 / 2.0).abs() <= 2.0
                && (half.preview_h as f32 - whole.preview_h as f32 / 2.0).abs() <= 2.0,
            "50% is half the box 100% is: {}x{} against {}x{}",
            half.preview_w,
            half.preview_h,
            whole.preview_w,
            whole.preview_h
        );
    }

    #[test]
    fn the_fit_takes_the_room_and_the_box_stays_on_the_display() {
        let room = work_area_at(0, 0);
        let layout = layout_for(
            pointer((0, 0)),
            (600, 400),
            PreviewScale::FitToScreen,
            false,
            false,
        )
        .expect("a display has room for a preview");
        assert!(
            layout.preview_w as i32 > 600 || layout.preview_h as i32 > 400,
            "a fit is the room rather than the program's own size"
        );
        assert!(layout.pos_x >= room.left);
        assert!(layout.pos_x + layout.preview_w as i32 <= room.right);
        assert!(layout.pos_y >= room.top);
        assert!(layout.pos_y + layout.preview_h as i32 <= room.bottom);
    }

    #[test]
    fn a_picture_is_the_room_cut_to_its_own_shape() {
        let room = work_area_at(0, 0);
        let at_a_hundred = layout_for(
            pointer((0, 0)),
            (600, 400),
            resolve_scale(
                Path::new("sample.png"),
                CustomScaling::Override(PreviewScale::Percent(100)),
            ),
            true,
            false,
        )
        .expect("a display has room for a preview");
        assert!(
            at_a_hundred.preview_w > 600 && at_a_hundred.preview_h > 400,
            "100% is the whole of the room rather than the picture at its own pixel size"
        );
        let shape = at_a_hundred.preview_w as f32 / at_a_hundred.preview_h as f32;
        assert!(
            (shape - 600.0 / 400.0).abs() < 0.01,
            "cut to the picture's own shape, which is {shape} rather than 1.5"
        );
        assert!(
            at_a_hundred.pos_x >= room.left
                && at_a_hundred.pos_x + at_a_hundred.preview_w as i32 <= room.right
                && at_a_hundred.pos_y >= room.top
                && at_a_hundred.pos_y + at_a_hundred.preview_h as i32 <= room.bottom,
            "and the box it comes out at is one the display holds"
        );

        let fitted = layout_for(
            pointer((0, 0)),
            (600, 400),
            PreviewScale::FitToScreen,
            true,
            false,
        )
        .expect("a display has room for a preview");
        assert_eq!(
            (fitted.preview_w, fitted.preview_h),
            (at_a_hundred.preview_w, at_a_hundred.preview_h),
            "a fit is the same whole of the room"
        );

        let half = layout_for(
            pointer((0, 0)),
            (600, 400),
            resolve_scale(
                Path::new("sample.png"),
                CustomScaling::Override(PreviewScale::Percent(50)),
            ),
            true,
            false,
        )
        .expect("half of a room is a preview a display can hold");
        assert!(
            (half.preview_w as f32 - at_a_hundred.preview_w as f32 / 2.0).abs() <= 2.0
                && (half.preview_h as f32 - at_a_hundred.preview_h as f32 / 2.0).abs() <= 2.0,
            "a percentage below the fit is that share of the room: {}x{} against {}x{}",
            half.preview_w,
            half.preview_h,
            at_a_hundred.preview_w,
            at_a_hundred.preview_h
        );
    }

    #[test]
    fn follow_tray_scale_maps_every_file_kind_to_its_tray_setting() {
        let scales = HoverScales {
            picture: PreviewScale::Percent(80),
            video: PreviewScale::FitToScreenReduced(75),
            animated: PreviewScale::Percent(60),
            ebook: PreviewScale::Percent(70),
            document: PreviewScale::Percent(65),
            font: PreviewScale::Percent(55),
            design: PreviewScale::Percent(85),
            vector: PreviewScale::Percent(90),
            model: PreviewScale::Percent(50),
        };
        let audio = PreviewScale::Percent(40);
        let p = Path::new("sample.bin");

        assert_eq!(
            tray_scale_for_kind(PreviewType::Images, p, scales, audio),
            PreviewScale::Percent(80)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Videos, p, scales, audio),
            PreviewScale::FitToScreenReduced(75)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Audio, p, scales, audio),
            PreviewScale::FitToScreenReduced(40)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Ebook, p, scales, audio),
            PreviewScale::FitToScreenReduced(70)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Document, p, scales, audio),
            PreviewScale::FitToScreenReduced(65)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Fonts, p, scales, audio),
            PreviewScale::FitToScreenReduced(55)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Design, p, scales, audio),
            PreviewScale::FitToScreenReduced(85)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Vector, p, scales, audio),
            PreviewScale::FitToScreenReduced(90)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Models, p, scales, audio),
            PreviewScale::FitToScreenReduced(50)
        );
        assert_eq!(
            tray_scale_for_kind(PreviewType::Text, Path::new("notes.txt"), scales, audio),
            PreviewScale::Percent(100)
        );
    }

    #[test]
    fn thumbnail_properties_sets_source_client_area_and_opacity_flags() {
        let normal = thumbnail_properties(400, 300, (400, 300), CustomWindowMode::Normal);
        assert_ne!(normal.dwFlags & DWM_TNP_SOURCECLIENTAREAONLY, 0);
        assert_ne!(normal.dwFlags & DWM_TNP_OPACITY, 0);
        assert_eq!(normal.opacity, 255);
        assert!(!normal.fSourceClientAreaOnly.as_bool());

        let borderless = thumbnail_properties(400, 300, (400, 300), CustomWindowMode::Borderless);
        assert_ne!(borderless.dwFlags & DWM_TNP_SOURCECLIENTAREAONLY, 0);
        assert_ne!(borderless.dwFlags & DWM_TNP_OPACITY, 0);
        assert_eq!(borderless.opacity, 255);
        assert!(borderless.fSourceClientAreaOnly.as_bool());
    }

    #[test]
    fn behind_host_placement_aligns_visible_frame_and_detects_overflow() {
        let host = (200, 120, 640, 480);
        let insets = (7, 0, 7, 7);
        assert_eq!(
            behind_host_window_rect(host, insets),
            (193, 120, 654, 487),
            "outer window expands by invisible DWM insets so visible frame matches host"
        );

        let exact_visible = RECT {
            left: 200,
            top: 120,
            right: 840,
            bottom: 600,
        };
        assert!(is_aligned_behind_host(exact_visible, host));

        // Smaller contained window and up to 4px DPI rounding tolerance still aligned.
        let smaller_contained = RECT {
            left: 220,
            top: 140,
            right: 800,
            bottom: 560,
        };
        assert!(is_aligned_behind_host(smaller_contained, host));

        let dpi_rounded = RECT {
            left: 196,
            top: 116,
            right: 844,
            bottom: 604,
        };
        assert!(is_aligned_behind_host(dpi_rounded, host));

        // Window refusing to shrink (e.g. Paint min size 704x500 > host 640x480) triggers fallback.
        let too_big = RECT {
            left: 200,
            top: 120,
            right: 904,
            bottom: 620,
        };
        assert!(!is_aligned_behind_host(too_big, host));

        // Drifted window (recentered by app) is detected as unaligned.
        let drifted = RECT {
            left: 300,
            top: 120,
            right: 940,
            bottom: 600,
        };
        assert!(!is_aligned_behind_host(drifted, host));
    }

    #[test]
    fn offscreen_mode_never_bounces_through_behind_host_unless_host_grew() {
        let host = (200, 120, 640, 480);
        // Already parked OffScreen -> no-op.
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::BehindHost,
                Some(PlacementMode::OffScreen),
                true,
                Some(host),
                host
            ),
            SnapAction::KeepOffScreen
        );
        // OffScreen drifted on-screen -> re-park directly, never bounce through BehindHost.
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::BehindHost,
                Some(PlacementMode::OffScreen),
                false,
                Some(host),
                host
            ),
            SnapAction::ReparkOffScreen
        );
        // Host moved without growing -> stay OffScreen.
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::BehindHost,
                Some(PlacementMode::OffScreen),
                true,
                Some(host),
                (300, 180, 640, 480),
            ),
            SnapAction::KeepOffScreen
        );
        // Host grew larger -> retry BehindHost.
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::BehindHost,
                Some(PlacementMode::OffScreen),
                true,
                Some(host),
                (200, 120, 800, 480),
            ),
            SnapAction::TryBehindHost
        );
        // Initial or BehindHost -> try BehindHost.
        assert_eq!(
            decide_snap_action(CustomCaptureTechnique::BehindHost, None, false, None, host),
            SnapAction::TryBehindHost
        );
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::BehindHost,
                Some(PlacementMode::BehindHost),
                false,
                Some(host),
                host
            ),
            SnapAction::TryBehindHost
        );
    }

    #[test]
    fn fallback_parking_skipped_when_app_clamped_primary_move_onto_monitor() {
        let vscreen = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let primary = parking_origin(vscreen);
        // App clamped right-edge off-screen move back onto right edge of monitor -> do NOT fling to -20000.
        let clamped_right = RECT {
            left: 1120,
            top: 0,
            right: 1920,
            bottom: 600,
        };
        assert!(!should_try_fallback_parking(
            true,
            primary,
            Some(clamped_right),
            vscreen
        ));
        // Already off-screen -> no fallback needed.
        let offscreen = RECT {
            left: primary.0,
            top: primary.1,
            right: primary.0 + 800,
            bottom: primary.1 + 600,
        };
        assert!(!should_try_fallback_parking(
            true,
            primary,
            Some(offscreen),
            vscreen
        ));
        // Primary SetWindowPos failed -> try fallback.
        assert!(should_try_fallback_parking(
            false,
            primary,
            Some(clamped_right),
            vscreen
        ));
        // Primary origin overlaps virtual screen -> try fallback.
        assert!(should_try_fallback_parking(
            true,
            (500, 100),
            Some(clamped_right),
            vscreen
        ));
    }

    #[test]
    fn parking_origin_is_completely_offscreen_within_i16_range() {
        let vscreen = RECT {
            left: -1920,
            top: 0,
            right: 3840,
            bottom: 2160,
        };
        let (x, y) = parking_origin(vscreen);
        assert_eq!((x, y), (vscreen.right + 1000, vscreen.top));
        assert!(x >= i16::MIN as i32 && x <= i16::MAX as i32);
        assert!(y >= i16::MIN as i32 && y <= i16::MAX as i32);
        let parked_rect = RECT {
            left: x,
            top: y,
            right: x + 800,
            bottom: y + 600,
        };
        assert!(is_off_screen(&parked_rect, vscreen));
        let fallback_rect = RECT {
            left: FALLBACK_PARKING,
            top: FALLBACK_PARKING,
            right: FALLBACK_PARKING + 800,
            bottom: FALLBACK_PARKING + 600,
        };
        assert!(is_off_screen(&fallback_rect, vscreen));
        let on_screen = RECT {
            left: 100,
            top: 100,
            right: 900,
            bottom: 700,
        };
        assert!(!is_off_screen(&on_screen, vscreen));
    }

    #[test]
    fn a_window_that_keeps_a_size_of_its_own_is_composed_into_the_box() {
        let whole = the_picture_in(704, 500, (704, 500));
        assert_eq!(
            (whole.left, whole.top, whole.right, whole.bottom),
            (0, 0, 704, 500),
            "a window that took the box is drawn in the whole of it, at its own size"
        );

        let smaller = the_picture_in(300, 200, (100, 60));
        assert_eq!(
            (smaller.left, smaller.top, smaller.right, smaller.bottom),
            (100, 70, 200, 130),
            "and a window smaller than the box is centred in it rather than stretched to fill it"
        );

        let fitted = the_picture_in(200, 150, (704, 500));
        assert_eq!(
            (fitted.left, fitted.top, fitted.right, fitted.bottom),
            (0, 4, 200, 146),
            "a window too big for the box is its own shape fitted into it, and the letterbox is \
             what the background shows"
        );
        assert!(
            fitted.top == 150 - fitted.bottom,
            "the window stands in the middle of the box rather than against one edge"
        );
        assert!(
            fitted.right - fitted.left <= 200 && fitted.bottom - fitted.top <= 150,
            "and never past the box it is drawn in"
        );
    }

    #[test]
    fn a_window_is_asked_for_the_shape_of_the_box_it_will_not_take() {
        // The box is cut to the picture's 4:3 shape and Paint will not be smaller than 704x500, so
        // it is asked for the box's shape at the smallest size that clears the floor it answered
        // with.
        let asked = the_ask_for((80, 60), (704, 500));
        assert_eq!(
            asked,
            (704, 528),
            "the ask keeps the box's own shape, raised to the least the window will take"
        );

        let picture = the_picture_in(80, 60, asked);
        assert_eq!(
            (picture.left, picture.top, picture.right, picture.bottom),
            (0, 0, 80, 60),
            "and the window fills that box, with none of the background around it"
        );

        // A tall box is asked for at its own shape too, tall past the floor's own height. This is
        // the reported hover: a 2135x6012 picture's box refused its width was a window in a band of
        // background above and below the picture.
        let asked = the_ask_for((490, 1380), (704, 1380));
        assert_eq!(
            asked,
            (704, 1983),
            "a box narrower than the window is asked for taller, keeping its own shape"
        );
        let picture = the_picture_in(490, 1380, asked);
        assert_eq!(
            (picture.left, picture.top, picture.right, picture.bottom),
            (0, 0, 490, 1380),
            "so the picture fills the box rather than standing letterboxed in it"
        );

        // A window that took the box is asked for nothing more, and one that will not say what it
        // took leaves the box it was asked for.
        assert_eq!(the_ask_for((600, 400), (600, 400)), (600, 400));
        assert_eq!(the_ask_for((600, 400), (0, 0)), (600, 400));
    }

    #[test]
    fn a_tall_picture_is_a_tall_box_whatever_the_program_last_opened_at() {
        let folder = std::env::temp_dir().join("rust-hover-preview-custom-shape");
        let _ = std::fs::create_dir_all(&folder);
        let tall = folder.join("tall.png");
        image::RgbaImage::from_pixel(40, 90, image::Rgba([10, 20, 30, 255]))
            .save(&tall)
            .expect("a picture to write");

        assert_eq!(
            base_size(&tall),
            Some((40, 90)),
            "the shape is the picture's own rather than the window the program opens at"
        );

        let svg = folder.join("box.svg");
        std::fs::write(
            &svg,
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="320" height="180"></svg>"#,
        )
        .expect("an svg to write");
        assert_eq!(
            base_size(&svg),
            Some((320, 180)),
            "an SVG declares its own intrinsic dimensions"
        );

        let words = folder.join("notes.txt");
        std::fs::write(&words, "a file with no size of its own").expect("a file to write");
        assert!(
            base_size(&words).is_none(),
            "a file that is not a picture leaves the program's window as the shape"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    #[test]
    fn a_program_named_without_a_folder_is_looked_for_where_the_machine_looks() {
        let found = resolved("cmd.exe").expect("a machine has a command processor");
        assert!(
            Path::new(&found).is_file(),
            "a name is answered as the full path the desktop's own loader takes"
        );
        assert!(
            resolved("a program that is not installed.exe").is_none(),
            "a name nothing answers to is no program"
        );
        assert!(
            resolved(r"C:\no such folder\no such program.exe").is_none(),
            "and a path that is not there is not a name to search for"
        );
        assert!(
            resolved("\"cmd.exe\"").is_some(),
            "the quotes `Copy as path` brings around a program are not part of it"
        );
    }

    #[test]
    fn command_line_quoting_handles_spaces_quotes_and_trailing_backslashes() {
        let decode = |exe: &str, path: &str| {
            let wide = build_command_line(OsStr::new(exe), OsStr::new(path));
            assert_eq!(wide.last(), Some(&0), "must be NUL-terminated");
            String::from_utf16_lossy(&wide[..wide.len() - 1])
        };

        assert_eq!(
            decode("viewer.exe", "photo.png"),
            r#""viewer.exe" "photo.png""#
        );
        assert_eq!(
            decode(r"Program Files\My Viewer\app.exe", r"my folder\pic 1.png"),
            r#""Program Files\My Viewer\app.exe" "my folder\pic 1.png""#
        );
        assert_eq!(
            decode("app.exe", r#"say "hi".txt"#),
            r#""app.exe" "say \"hi\".txt""#
        );
        assert_eq!(
            decode("app.exe", r#"dir\"quoted".png"#),
            r#""app.exe" "dir\\\"quoted\".png""#
        );
        assert_eq!(
            decode("app.exe", r"trailing\folder\\"),
            r#""app.exe" "trailing\folder\\\\""#
        );
    }

    #[test]
    fn startup_info_sets_hide_and_clamps_provisional_rect() {
        let none = launch_startup_info(None);
        assert_eq!(none.cb, std::mem::size_of::<STARTUPINFOW>() as u32);
        assert_eq!(none.dwFlags, STARTF_USESHOWWINDOW);
        assert_eq!(none.wShowWindow, SW_HIDE.0 as u16);

        let rect = provisional_rect_for(
            pointer((0, 0)),
            (800, 600),
            PreviewScale::Percent(75),
            true,
            false,
        )
        .expect("provisional rect computed for known dims");
        assert!(rect.2 > 0 && rect.3 > 0);

        let with_rect = launch_startup_info(Some(rect));
        assert_eq!(
            with_rect.dwFlags,
            STARTF_USESHOWWINDOW | STARTF_USEPOSITION | STARTF_USESIZE
        );
        assert_eq!(with_rect.wShowWindow, SW_HIDE.0 as u16);
        assert_eq!(with_rect.dwX, rect.0.max(0) as u32);
        assert_eq!(with_rect.dwY, rect.1.max(0) as u32);
        assert_eq!(with_rect.dwXSize, rect.2 as u32);
        assert_eq!(with_rect.dwYSize, rect.3 as u32);

        // Negative multi-monitor coordinates and zero/negative sizes clamp safely.
        let clamped = launch_startup_info(Some((-120, -40, 0, -10)));
        assert_eq!(clamped.dwX, 0);
        assert_eq!(clamped.dwY, 0);
        assert_eq!(clamped.dwXSize, 1);
        assert_eq!(clamped.dwYSize, 1);
    }

    #[test]
    fn screen_to_workspace_rect_subtracts_taskbar_offset_on_primary_and_secondary() {
        let screen = RECT {
            left: 300,
            top: 150,
            right: 940,
            bottom: 630,
        };
        // Primary monitor with bottom taskbar: work origin equals monitor origin.
        let primary_mon = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let bottom_bar = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1040,
        };
        assert_eq!(
            screen_to_workspace_rect(screen, primary_mon, bottom_bar),
            screen
        );

        // Primary monitor with top taskbar (48px) and left dock (64px): shifts by (-64, -48).
        let top_left_bar = RECT {
            left: 64,
            top: 48,
            right: 1920,
            bottom: 1080,
        };
        assert_eq!(
            screen_to_workspace_rect(screen, primary_mon, top_left_bar),
            RECT {
                left: 236,
                top: 102,
                right: 876,
                bottom: 582,
            }
        );

        // Secondary monitor to the left (-1920..0) with 40px top taskbar.
        let secondary_mon = RECT {
            left: -1920,
            top: 0,
            right: 0,
            bottom: 1080,
        };
        let secondary_work = RECT {
            left: -1920,
            top: 40,
            right: 0,
            bottom: 1080,
        };
        let sec_screen = RECT {
            left: -1500,
            top: 200,
            right: -700,
            bottom: 800,
        };
        assert_eq!(
            screen_to_workspace_rect(sec_screen, secondary_mon, secondary_work),
            RECT {
                left: -1500,
                top: 160,
                right: -700,
                bottom: 760,
            }
        );
    }

    #[test]
    fn ownership_never_matches_manually_launched_sibling_exe_once_owned() {
        let launch_ft = 100_000_000_000u64;
        let owned = [1001u32, 1002];
        let before = [500u32];

        // Post-adoption (`OwnedOnly`): manually-launched sibling PID with same exe is NEVER matched.
        assert!(belongs_to_preview(
            1001,
            &owned,
            OwnershipMode::OwnedOnly,
            "i_view64.exe",
            launch_ft,
            &before,
            Some("i_view64.exe"),
            Some(launch_ft + 5_000_000),
        ));
        assert!(!belongs_to_preview(
            2005,
            &owned,
            OwnershipMode::OwnedOnly,
            "i_view64.exe",
            launch_ft,
            &before,
            Some("i_view64.exe"),
            Some(launch_ft + 5_000_000),
        ));

        // Discovery pass 1 (`allow_image_fallback: false`): sibling PID is NEVER matched.
        assert!(!belongs_to_preview(
            2005,
            &owned,
            OwnershipMode::Discover {
                allow_image_fallback: false,
            },
            "i_view64.exe",
            launch_ft,
            &before,
            Some("i_view64.exe"),
            Some(launch_ft + 5_000_000),
        ));

        // Discovery pass 2 (`allow_image_fallback: true` for brokered apps):
        // - Newly spawned brokered process (>= launch_ft - 2s) matches.
        assert!(belongs_to_preview(
            3001,
            &owned,
            OwnershipMode::Discover {
                allow_image_fallback: true,
            },
            "mspaint.exe",
            launch_ft,
            &before,
            Some("mspaint.exe"),
            Some(launch_ft - 5_000_000),
        ));
        // - Unrelated process started well before launch (> 2s earlier, not in processes_before) rejected.
        assert!(!belongs_to_preview(
            3002,
            &owned,
            OwnershipMode::Discover {
                allow_image_fallback: true,
            },
            "mspaint.exe",
            launch_ft,
            &before,
            Some("mspaint.exe"),
            Some(launch_ft - PROCESS_TIME_SLACK_100NS - 1),
        ));
        // - Existing single-instance packaged host in `processes_before` matches (gated by windows_before).
        assert!(belongs_to_preview(
            500,
            &owned,
            OwnershipMode::Discover {
                allow_image_fallback: true,
            },
            "mspaint.exe",
            launch_ft,
            &before,
            Some("mspaint.exe"),
            None,
        ));
        // - Wrong exe name rejected.
        assert!(!belongs_to_preview(
            3003,
            &owned,
            OwnershipMode::Discover {
                allow_image_fallback: true,
            },
            "mspaint.exe",
            launch_ft,
            &before,
            Some("notepad.exe"),
            Some(launch_ft + 1000),
        ));
    }

    #[test]
    fn packaged_broker_path_matches_windowsapps_segments_only() {
        assert!(is_packaged_broker_path(Path::new(
            r"Program Files\WindowsApps\Vendor.App_1.0.0_x64__abc\app.exe"
        )));
        assert!(is_packaged_broker_path(Path::new(
            r"AppData\Local\Microsoft\WindowsApps\viewer.exe"
        )));
        assert!(is_packaged_broker_path(Path::new(
            r"AppData\Local\Microsoft\windowsapps\viewer.exe"
        )));
        assert!(!is_packaged_broker_path(Path::new(
            r"Program Files\Viewer\viewer.exe"
        )));
        assert!(!is_packaged_broker_path(Path::new(r"tools\WindowsApps")));
        assert!(!is_packaged_broker_path(Path::new("viewer.exe")));
    }

    #[test]
    fn can_use_invisible_layered_rejects_noredirectionbitmap_and_update_layered_window() {
        // DirectComposition / WinUI 3 (WS_EX_NOREDIRECTIONBITMAP) must be rejected.
        assert!(!can_use_invisible_layered(
            WS_EX_NOREDIRECTIONBITMAP.0,
            None
        ));
        assert!(!can_use_invisible_layered(
            WS_EX_NOREDIRECTIONBITMAP.0 | WS_EX_LAYERED.0,
            Some(true)
        ));
        // Already WS_EX_LAYERED via UpdateLayeredWindow (GetLayeredWindowAttributes fails) rejected.
        assert!(!can_use_invisible_layered(WS_EX_LAYERED.0, Some(false)));
        // Normal window and LWA_ALPHA-compatible layered window accepted.
        assert!(can_use_invisible_layered(0, None));
        assert!(can_use_invisible_layered(WS_EX_TOOLWINDOW.0, None));
        assert!(can_use_invisible_layered(WS_EX_LAYERED.0, Some(true)));
    }

    #[test]
    fn invisible_layered_mode_stays_on_monitor_across_host_moves_and_resizes() {
        let host = (200, 120, 640, 480);
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::InvisibleLayered,
                Some(PlacementMode::InvisibleLayered),
                false,
                Some(host),
                host
            ),
            SnapAction::KeepInvisibleLayered
        );
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::InvisibleLayered,
                Some(PlacementMode::InvisibleLayered),
                false,
                Some(host),
                (300, 200, 900, 700)
            ),
            SnapAction::KeepInvisibleLayered
        );
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::InvisibleLayered,
                Some(PlacementMode::InvisibleLayered),
                true,
                Some(host),
                (100, 80, 320, 240)
            ),
            SnapAction::KeepInvisibleLayered
        );
    }

    #[test]
    fn capture_technique_dispatches_host_and_snap_actions_accurately() {
        let host = (200, 120, 640, 480);
        // Only InvisibleLayered attempts WS_EX_LAYERED; only DirectWindow skips DWM host.
        assert!(tries_invisible_layered(
            CustomCaptureTechnique::InvisibleLayered
        ));
        assert!(!tries_invisible_layered(CustomCaptureTechnique::BehindHost));
        assert!(!tries_invisible_layered(CustomCaptureTechnique::OffScreen));
        assert!(!tries_invisible_layered(
            CustomCaptureTechnique::DirectWindow
        ));

        assert!(uses_dwm_host(CustomCaptureTechnique::InvisibleLayered));
        assert!(uses_dwm_host(CustomCaptureTechnique::BehindHost));
        assert!(uses_dwm_host(CustomCaptureTechnique::OffScreen));
        assert!(!uses_dwm_host(CustomCaptureTechnique::DirectWindow));

        // OffScreen technique always parks/keeps off-screen, even if host grew.
        assert_eq!(
            decide_snap_action(CustomCaptureTechnique::OffScreen, None, false, None, host),
            SnapAction::ReparkOffScreen
        );
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::OffScreen,
                Some(PlacementMode::OffScreen),
                true,
                Some(host),
                (200, 120, 900, 700)
            ),
            SnapAction::KeepOffScreen
        );

        // DirectWindow always places directly without BehindHost/OffScreen.
        assert_eq!(
            decide_snap_action(
                CustomCaptureTechnique::DirectWindow,
                None,
                false,
                None,
                host
            ),
            SnapAction::PlaceDirectWindow
        );
        assert!(uses_dwm_host(CustomCaptureTechnique::OffscreenStaged));
        assert!(!tries_invisible_layered(
            CustomCaptureTechnique::OffscreenStaged
        ));
        assert!(uses_dwm_host(CustomCaptureTechnique::InProcessHook));
        assert!(!tries_invisible_layered(
            CustomCaptureTechnique::InProcessHook
        ));
    }

    #[test]
    fn offscreen_staged_startup_stage_rect_and_teardown_safeguard_helpers() {
        let vscreen = RECT {
            left: -1920,
            top: 0,
            right: 3840,
            bottom: 2160,
        };
        let si = launch_startup_info_for(
            Some((200, 150, 900, 700)),
            CustomCaptureTechnique::OffscreenStaged,
            vscreen,
        );
        assert_eq!(
            si.dwFlags,
            STARTF_USESHOWWINDOW | STARTF_USEPOSITION | STARTF_USESIZE
        );
        assert_eq!(si.wShowWindow, SW_HIDE.0 as u16);
        assert_eq!(si.dwX, (vscreen.right + 2000) as u32);
        assert_eq!(si.dwY, 0);
        assert_eq!(si.dwXSize, 900);
        assert_eq!(si.dwYSize, 700);

        // Existing techniques remain delegated unchanged to launch_startup_info.
        let legacy = launch_startup_info_for(
            Some((200, 150, 900, 700)),
            CustomCaptureTechnique::InvisibleLayered,
            vscreen,
        );
        assert_eq!(legacy.dwX, 200);
        assert_eq!(legacy.dwY, 150);

        // Stage rect sits at (-10000, -10000) and is completely off-screen.
        let st = stage_rect(640, 480);
        assert_eq!(
            (st.left, st.top, st.right, st.bottom),
            (-10_000, -10_000, -9360, -9520)
        );
        assert!(is_off_screen(&st, vscreen));
        assert!(is_at_stage_coord(st));

        // Anti-stuck teardown guard restores reused non-owned windows when off-screen/staged,
        // and never touches newly spawned windows (terminated) or already on-screen windows.
        assert!(should_restore_on_teardown(false, st, vscreen));
        assert!(!should_restore_on_teardown(true, st, vscreen));
        let on_screen = RECT {
            left: 100,
            top: 100,
            right: 900,
            bottom: 700,
        };
        assert!(!should_restore_on_teardown(false, on_screen, vscreen));
    }

    #[test]
    fn the_hook_dll_is_sought_beside_the_executable_before_the_user_folder() {
        // The portable zip and the installer both carry the DLL beside the exe, so that copy wins
        // over whatever a lone executable's folder holds.
        let dir = Path::new(r"C:\apps\rhp");
        let user = PathBuf::from(r"C:\users\u\AppData\Local\rust-hover-preview");
        let candidates = hook_dll_candidates(dir, Some(user));
        assert_eq!(
            candidates,
            vec![
                dir.join("custom_preview_hook.dll"),
                Path::new(r"C:\apps").join("custom_preview_hook.dll"),
                PathBuf::from(r"C:\users\u\AppData\Local\rust-hover-preview")
                    .join("custom_preview_hook.dll"),
            ]
        );
        // No local-appdata folder on the machine is two candidates rather than a panic.
        assert_eq!(hook_dll_candidates(dir, None).len(), 2);
    }

    #[test]
    fn in_process_hook_dll_exports_shm_layout_and_thread_hook_attachment() {
        // 1. Companion DLL must exist beside the test binary (or parent target dir) and export both hooks.
        let (hmod, cbt_proc, cwp_proc) = load_hook_dll().expect(
            "custom_preview_hook.dll must be built and export rhp_cbt_proc + rhp_callwnd_proc",
        );

        // 2. Verify shared memory creation and field layout.
        let pid = std::process::id();
        let msg = unsafe { RegisterWindowMessageW(HOOK_CONTROL_MSG_NAME) };
        assert_ne!(msg, 0);
        let view = HookSharedView::create(pid, Some((120, 80, 640, 480)), msg)
            .expect("HookSharedView::create must succeed");
        assert_eq!(std::mem::size_of::<HookSharedState>(), 56);
        assert_eq!(view.state().magic, HOOK_MAGIC);
        assert_eq!(view.state().version, HOOK_VERSION);
        assert_eq!(view.state().target_pid, pid);
        assert!(view.is_cloaked());
        assert_eq!(view.state().rect_locked.load(Ordering::Acquire), 1);
        assert_eq!(view.state().x.load(Ordering::Acquire), STAGE_COORD);
        assert_eq!(view.state().y.load(Ordering::Acquire), STAGE_COORD);
        assert_eq!(view.state().w.load(Ordering::Acquire), 640);
        assert_eq!(view.state().h.load(Ordering::Acquire), 480);

        // Off-screen positive STARTUPINFOW for InProcessHook (u32-safe).
        let vscreen = RECT {
            left: -1920,
            top: 0,
            right: 3840,
            bottom: 2160,
        };
        let hook_si = launch_startup_info_for(
            Some((120, 80, 640, 480)),
            CustomCaptureTechnique::InProcessHook,
            vscreen,
        );
        assert_eq!(hook_si.wShowWindow, SW_HIDE.0 as u16);
        assert_eq!(hook_si.dwX, (vscreen.right + 2000) as u32);
        assert_eq!(hook_si.dwY, 0);
        assert_eq!(hook_si.dwXSize, 640);
        assert_eq!(hook_si.dwYSize, 480);

        // 3. Empirical Win11 check: spawn a benign suspended process (`cmd.exe /c exit 0`) and test
        //    whether `SetWindowsHookExW(WH_CBT | WH_CALLWNDPROC, ..., tid)` succeeds before vs after `ResumeThread`.
        let cmd_exe = std::env::var_os("ComSpec").expect("ComSpec must be set on Windows");
        let app_w = wide_null(OsStr::new(&cmd_exe));
        let mut cmdline = wide_null(OsStr::new("cmd.exe /c exit 0"));
        let si = STARTUPINFOW {
            cb: std::mem::size_of::<STARTUPINFOW>() as u32,
            dwFlags: STARTF_USESHOWWINDOW,
            wShowWindow: SW_HIDE.0 as u16,
            ..Default::default()
        };
        let mut pi = PROCESS_INFORMATION::default();
        unsafe {
            CreateProcessW(
                PCWSTR(app_w.as_ptr()),
                PWSTR(cmdline.as_mut_ptr()),
                None,
                None,
                BOOL(0),
                CREATE_SUSPENDED,
                None,
                PCWSTR::null(),
                &si,
                &mut pi,
            )
            .expect("CreateProcessW(CREATE_SUSPENDED) failed");
        }
        let attached_suspended = try_attach_thread_hooks(hmod, cbt_proc, cwp_proc, pi.dwThreadId);
        let pre_resume_ok = attached_suspended.is_some();
        if let Some((cbt, cwp)) = attached_suspended {
            let _ = unsafe { UnhookWindowsHookEx(cbt) };
            let _ = unsafe { UnhookWindowsHookEx(cwp) };
        }

        // Also verify attaching to a running thread (current thread, guaranteed W32THREAD-ready),
        // SW_HIDE launch default, cloak_sync delivery through WH_CALLWNDPROC, and captured_hwnd
        // ignoring 0x0 hidden helper windows while latching a real titled/captioned window.
        assert_eq!(IN_PROCESS_HOOK_SHOW_CMD, SW_HIDE.0 as u16);

        let current_tid = unsafe { windows::Win32::System::Threading::GetCurrentThreadId() };
        let current_hooks = try_attach_thread_hooks(hmod, cbt_proc, cwp_proc, current_tid)
            .expect("SetWindowsHookExW on active thread must succeed");

        let helper_hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!(""),
                WS_POPUP,
                0,
                0,
                0,
                0,
                None,
                None,
                None,
                None,
            )
            .expect("create 0x0 hidden helper window")
        };
        assert_eq!(
            view.captured_hwnd(),
            None,
            "0x0 hidden helper window must not be latched into captured_hwnd"
        );
        assert!(
            view.hook_hits() > 0,
            "hook must still intercept and cloak hidden helper windows"
        );

        let hits_before_cloak = view.hook_hits();
        assert!(
            view.cloak_sync(helper_hwnd),
            "cloak_sync must deliver CMD_CLOAK synchronously through WH_CALLWNDPROC"
        );
        assert!(
            view.hook_hits() > hits_before_cloak,
            "CMD_CLOAK must increment hook_hits inside rhp_callwnd_proc"
        );

        let main_hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!("InProcessHook Main Candidate"),
                WS_CAPTION,
                120,
                80,
                320,
                240,
                None,
                None,
                None,
                None,
            )
            .expect("create hidden main candidate window")
        };
        let _ = unsafe { SetWindowPos(main_hwnd, HWND_BOTTOM, 120, 80, 320, 240, SWP_NOACTIVATE) };
        let staged_rect = window_rect(main_hwnd).expect("main_hwnd rect");
        assert!(
            is_at_stage_coord(staged_rect),
            "hook must keep newly created window at STAGE_COORD before reveal, got {staged_rect:?}"
        );
        assert_eq!(
            view.captured_hwnd(),
            Some(main_hwnd),
            "real captioned window >= 64x64 must be latched into captured_hwnd"
        );

        // Reveal: update shared memory rect to target and move cleanly on-screen.
        view.set_rect(120, 80, 320, 240);
        let _ = unsafe { SetWindowPos(main_hwnd, HWND_BOTTOM, 120, 80, 320, 240, SWP_NOACTIVATE) };
        let revealed_rect = window_rect(main_hwnd).expect("revealed main_hwnd rect");
        assert_eq!(
            (revealed_rect.left, revealed_rect.top),
            (120, 80),
            "reveal must move window cleanly to updated shared-memory rect"
        );
        view.stage_offscreen();
        assert_eq!(view.state().x.load(Ordering::Acquire), STAGE_COORD);
        assert_eq!(view.state().y.load(Ordering::Acquire), STAGE_COORD);

        // CMD_DETACH clears cloak_enabled + rect_locked and removes the in-process subclass hook.
        assert!(
            view.detach_sync(main_hwnd),
            "detach_sync must deliver CMD_DETACH synchronously and increment hook_hits"
        );
        assert_eq!(view.state().cloak_enabled.load(Ordering::Acquire), 0);
        assert_eq!(view.state().rect_locked.load(Ordering::Acquire), 0);
        let _ = unsafe {
            SetWindowPos(
                main_hwnd,
                HWND_NOTOPMOST,
                210,
                140,
                320,
                240,
                SWP_NOACTIVATE,
            )
        };
        let detached_rect = window_rect(main_hwnd).expect("detached main_hwnd rect");
        assert_eq!(
            (detached_rect.left, detached_rect.top),
            (210, 140),
            "detached window must move freely with zero leftover subclass rect-lock"
        );

        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(main_hwnd) };
        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(helper_hwnd) };
        let _ = unsafe { UnhookWindowsHookEx(current_hooks.0) };
        let _ = unsafe { UnhookWindowsHookEx(current_hooks.1) };

        let _ = unsafe { TerminateProcess(pi.hProcess, 0) };
        let _ = unsafe { CloseHandle(pi.hThread) };
        let _ = unsafe { CloseHandle(pi.hProcess) };
        let _ = unsafe { FreeLibrary(hmod) };

        eprintln!(
            "[empirical] SetWindowsHookExW(WH_CBT+WH_CALLWNDPROC) on CREATE_SUSPENDED thread: {}",
            if pre_resume_ok {
                "ACCEPTED_BEFORE_RESUME"
            } else {
                "REJECTED_BEFORE_RESUME (retry-after-resume path active)"
            }
        );
    }

    #[test]
    fn custom_preview_volume_params_clamps_and_mutes_at_zero() {
        assert_eq!(audio_session::custom_preview_volume_params(0), (0.0, true));
        assert_eq!(
            audio_session::custom_preview_volume_params(50),
            (0.5, false)
        );
        assert_eq!(
            audio_session::custom_preview_volume_params(100),
            (1.0, false)
        );
        assert_eq!(
            audio_session::custom_preview_volume_params(250),
            (1.0, false)
        );
    }

    #[test]
    fn preview_job_detach_clears_kill_on_close_and_preserves_child() {
        let mut child = engine_processes::hidden_command("cmd.exe")
            .args(["/C", "ping -n 6 127.0.0.1 >nul"])
            .spawn()
            .expect("spawn hidden helper child");
        let pid = child.id();
        let job = PreviewJob::create_and_assign(pid).expect("create per-preview job");
        assert!(job.detach(), "detach must clear KILL_ON_JOB_CLOSE");
        // Child must still be alive after detached job is closed.
        std::thread::sleep(Duration::from_millis(50));
        assert!(
            child.try_wait().expect("try_wait").is_none(),
            "detached child must survive job close"
        );
        let _ = child.kill();
        let _ = child.wait();
    }

    static TEST_STATE_LOCK: Mutex<()> = Mutex::new(());

    fn dummy_program(path: PathBuf) -> Program {
        let handle = unsafe {
            OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, std::process::id())
                .expect("open current process query handle")
        };
        Program {
            path,
            exe: "test.exe".to_string(),
            process: Arc::new(unsafe { OwnedHandle::from_raw_handle(handle.0) }),
            pid: 0,
            job: None,
        }
    }

    #[test]
    fn active_path_tracks_want_and_program_and_clears_on_hide() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        let path = PathBuf::from("sample.custom");

        if let Ok(mut w) = WANT.lock() {
            *w = None;
        }
        if let Ok(mut p) = PROGRAM.lock() {
            *p = None;
        }
        assert_eq!(active_path(), None);
        assert!(!is_active());

        if let Ok(mut p) = PROGRAM.lock() {
            *p = Some(dummy_program(path.clone()));
        }
        // Program lingering without WANT must not report active.
        assert_eq!(active_path(), None);
        assert!(!is_active());

        if let Ok(mut w) = WANT.lock() {
            *w = Some(Want {
                at: pointer((100, 100)),
            });
        }
        assert_eq!(active_path().as_deref(), Some(path.as_path()));
        assert!(is_active());

        // Pin-press routing must emit PinCold for the active custom preview path.
        assert!(matches!(
            pin_request_if_active(),
            Some(super::super::PreviewMessage::PinCold { path: routed }) if routed == path
        ));

        hide();
        assert_eq!(active_path(), None);
        assert!(!is_active());
        assert!(pin_request_if_active().is_none());

        if let Ok(mut p) = PROGRAM.lock() {
            *p = None;
        }
    }

    #[test]
    fn clamp_rect_to_work_area_keeps_promoted_window_on_monitor() {
        let work = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1040,
        };
        // Off-screen staged coordinates clamp cleanly onto work area.
        assert_eq!(
            clamp_rect_to_work_area((STAGE_COORD, STAGE_COORD, 400, 300), work),
            (0, 0, 400, 300)
        );
        // Right/bottom overflow slides back inside work area.
        assert_eq!(
            clamp_rect_to_work_area((1700, 900, 500, 400), work),
            (1420, 640, 500, 400)
        );
        // Oversized dimensions clamp to work area bounds.
        assert_eq!(
            clamp_rect_to_work_area((100, 100, 3000, 2000), work),
            (0, 0, 1920, 1040)
        );
    }

    #[test]
    fn promoted_window_rects_preserves_fit_to_screen_size_with_dwm_insets() {
        let monitor = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let insets = (7, 0, 7, 7);

        // Fit-to-screen preview touching top and bottom of work area (1040px) or full monitor (1080px):
        // visible rect must NOT shrink, and outer rect must include invisible DWM border overhang.
        let (vis_work, outer_work) = promoted_window_rects((300, 0, 1000, 1040), insets, monitor);
        assert_eq!(vis_work, (300, 0, 1000, 1040));
        assert_eq!(outer_work, (293, 0, 1014, 1047));

        let (vis_full, outer_full) = promoted_window_rects((0, 0, 1920, 1080), insets, monitor);
        assert_eq!(vis_full, (0, 0, 1920, 1080));
        assert_eq!(outer_full, (-7, 0, 1934, 1087));

        // Edge-aligned visible frame needs zero clamp shift; overflowing min-size frame slides cleanly.
        assert_eq!(
            visible_clamp_delta(
                RECT {
                    left: 0,
                    top: 0,
                    right: 1920,
                    bottom: 1080
                },
                monitor
            ),
            (0, 0)
        );
        assert_eq!(
            visible_clamp_delta(
                RECT {
                    left: 1600,
                    top: 700,
                    right: 2100,
                    bottom: 1120
                },
                monitor
            ),
            (-180, -40)
        );

        // Post-move size verification catches WM_DPICHANGED / stale-inset shrink.
        assert!(visible_rect_matches_target(
            RECT {
                left: 300,
                top: 0,
                right: 1300,
                bottom: 1040
            },
            vis_work
        ));
        assert!(!visible_rect_matches_target(
            RECT {
                left: 300,
                top: 0,
                right: 1300,
                bottom: 992
            },
            vis_work
        ));
    }

    #[test]
    fn open_fully_detach_clears_state_so_subsequent_hide_is_noop() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        let path = PathBuf::from("promote.custom");
        let bounds = work_area_at(0, 0);
        let target_x = bounds.left + 120;
        let target_y = bounds.top + 90;

        let hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(INVISIBLE_EX_STYLE),
                w!("STATIC"),
                w!("Promote Test"),
                WS_CAPTION,
                STAGE_COORD,
                STAGE_COORD,
                320,
                240,
                None,
                None,
                None,
                None,
            )
            .expect("create test window")
        };
        let _ = unsafe { SetLayeredWindowAttributes(hwnd, COLORREF(0), 1, LWA_ALPHA) };

        if let Ok(mut p) = PROGRAM.lock() {
            *p = Some(dummy_program(path.clone()));
        }
        if let Ok(mut w) = WANT.lock() {
            *w = Some(Want {
                at: pointer((target_x, target_y)),
            });
        }
        SOURCE.store(hwnd.0 as isize, Ordering::SeqCst);
        SOURCE_INVISIBLE_LAYERED.store(true, Ordering::SeqCst);
        if let Ok(mut r) = LAST_HOST_RECT.lock() {
            *r = Some((target_x, target_y, 320, 240));
        }
        if let Ok(mut ex) = SOURCE_ORIGINAL_EX_STYLE.lock() {
            *ex = Some(0);
        }

        let launch_before = LAUNCH.load(Ordering::SeqCst);
        assert!(try_detach_active_preview(&path));
        assert!(LAUNCH.load(Ordering::SeqCst) > launch_before);
        assert!(WANT.lock().unwrap().is_none());
        assert!(PROGRAM.lock().unwrap().is_none());
        assert_eq!(SOURCE.load(Ordering::SeqCst), 0);
        assert!(!SOURCE_INVISIBLE_LAYERED.load(Ordering::SeqCst));
        assert!(ACTIVE_HOOK_CONTROL.lock().unwrap().is_none());
        assert!(LAST_HOST_RECT.lock().unwrap().is_none());
        assert_eq!(active_path(), None);

        let ex_after = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
        assert_eq!(
            ex_after & INVISIBLE_EX_STYLE,
            0,
            "promote must strip invisible/preview extended styles"
        );
        let promoted_rect = window_rect(hwnd).expect("promoted rect");
        assert!(
            !is_at_stage_coord(promoted_rect),
            "promoted window must move off STAGE_COORD onto monitor, got {promoted_rect:?}"
        );

        // Subsequent hide() on hover-away must be a 100% no-op on the promoted window.
        hide();
        let after_hide_rect = window_rect(hwnd).expect("rect after hide");
        assert_eq!(
            after_hide_rect, promoted_rect,
            "hide() after promote must not move or park the promoted window"
        );

        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(hwnd) };
    }

    #[test]
    fn promote_lifts_bottom_window_above_sibling_without_staying_topmost() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        let path = PathBuf::from("zorder.custom");
        let bounds = work_area_at(0, 0);
        let x = bounds.left + 100;
        let y = bounds.top + 100;

        let sibling = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!("Explorer Sibling"),
                WS_CAPTION,
                x,
                y,
                320,
                240,
                None,
                None,
                None,
                None,
            )
            .expect("create sibling window")
        };
        let _ = unsafe { ShowWindow(sibling, SW_SHOWNOACTIVATE) };

        let source = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!("Preview Source"),
                WS_CAPTION,
                x,
                y,
                320,
                240,
                None,
                None,
                None,
                None,
            )
            .expect("create source window")
        };
        let _ = unsafe {
            SetWindowPos(
                source,
                HWND_BOTTOM,
                x,
                y,
                320,
                240,
                SWP_SHOWWINDOW | SWP_NOACTIVATE,
            )
        };

        if let Ok(mut p) = PROGRAM.lock() {
            *p = Some(dummy_program(path.clone()));
        }
        SOURCE.store(source.0 as isize, Ordering::SeqCst);
        if let Ok(mut r) = LAST_HOST_RECT.lock() {
            *r = Some((x, y, 320, 240));
        }

        // Verify hide_the_host() is suppressed while DETACHING is set.
        SHOWN.store(true, Ordering::SeqCst);
        DETACHING.store(true, Ordering::SeqCst);
        hide_the_host();
        assert!(
            SHOWN.load(Ordering::SeqCst),
            "hide_the_host must be suppressed during DETACHING"
        );
        DETACHING.store(false, Ordering::SeqCst);

        assert!(try_detach_active_preview(&path));
        assert!(!DETACHING.load(Ordering::SeqCst));

        let ex_after = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) } as u32;
        assert_eq!(
            ex_after & WS_EX_TOPMOST.0,
            0,
            "promoted window must be demoted to HWND_NOTOPMOST"
        );

        // Walk downward from `source` via GW_HWNDNEXT; `sibling` must appear strictly below `source`.
        let mut cursor = source;
        let mut sibling_below = false;
        while let Ok(next) =
            unsafe { GetWindow(cursor, windows::Win32::UI::WindowsAndMessaging::GW_HWNDNEXT) }
        {
            if next == sibling {
                sibling_below = true;
                break;
            }
            if next.0.is_null() {
                break;
            }
            cursor = next;
        }
        assert!(
            sibling_below,
            "promoted window must sit strictly above sibling normal window in Z-order"
        );

        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(source) };
        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(sibling) };
    }

    #[test]
    fn open_fully_dispatch_respects_promotion_technique() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        let path = PathBuf::from("dispatch.custom");

        // DirectPromotion: uses try_detach first and skips spawning when active detach succeeds.
        let mut detached = false;
        let mut spawned = false;
        assert!(dispatch_open_fully(
            CustomPromotionTechnique::DirectPromotion,
            &path,
            "viewer.exe",
            |p| {
                assert_eq!(p, path.as_path());
                detached = true;
                true
            },
            |_, _| spawned = true,
        ));
        assert!(detached);
        assert!(!spawned);

        // DirectPromotion fallback when cold/inactive: falls through to normal detached spawn.
        let mut detached = false;
        let mut spawned = false;
        assert!(dispatch_open_fully(
            CustomPromotionTechnique::DirectPromotion,
            &path,
            "viewer.exe",
            |_| {
                detached = true;
                false
            },
            |exe, p| {
                assert_eq!(exe, "viewer.exe");
                assert_eq!(p, path.as_path());
                spawned = true;
            },
        ));
        assert!(detached);
        assert!(spawned);

        // OpenAsNewProcess: never detaches; clears WANT/SHOWN/SOURCE and spawns a fresh process.
        if let Ok(mut w) = WANT.lock() {
            *w = Some(Want {
                at: Where::Pointer {
                    cursor: (0, 0),
                    avoid: None,
                },
            });
        }
        SHOWN.store(true, Ordering::SeqCst);
        let mut detached = false;
        let mut spawned = false;
        assert!(dispatch_open_fully(
            CustomPromotionTechnique::OpenAsNewProcess,
            &path,
            "viewer.exe",
            |_| {
                detached = true;
                true
            },
            |exe, p| {
                assert_eq!(exe, "viewer.exe");
                assert_eq!(p, path.as_path());
                spawned = true;
            },
        ));
        assert!(!detached, "OpenAsNewProcess must not detach in-place");
        assert!(spawned, "OpenAsNewProcess must spawn a fresh process");
        assert!(!is_showing(), "OpenAsNewProcess must clear WANT");
        assert!(
            !SHOWN.load(Ordering::SeqCst),
            "OpenAsNewProcess must hide host"
        );
    }
}

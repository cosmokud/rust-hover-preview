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
//! **And the box is a client, while the picture is the whole frame around it.** A program draws its
//! document in its *client*; the caption above it and the resize border round it are the frame, and
//! a preview composed from the whole window gives the document only the box *minus* that frame — a
//! player handed a client 4 % wider than the film's own shape fills the height and draws the rest as
//! black bars either side. So the hover is **the box grown by the program's own frame** (see
//! `the_host_for`), the window is placed so that its *visible frame* — the client plus that frame — is
//! that grown rect (see `the_ask_for_host` and `behind_host_window_rect`), and the host is that rect
//! too: the client lands on the box, the caption keeps the strip the hover grew by, and the picture
//! is composed into the whole of the host rather than fitted into part of it. The frame is read off
//! the window on every poll (`chrome_now`), so a program that restyles itself is followed, and a
//! `Borderless` preset — whose frame is stripped — has none and is exactly the box it always was.
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
use std::collections::{HashMap, HashSet};
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::os::windows::io::{AsRawHandle, FromRawHandle, OwnedHandle};
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{
    AtomicBool, AtomicI32, AtomicIsize, AtomicU32, AtomicU64, AtomicU8, Ordering,
};
use std::sync::{Arc, Mutex, Once};
use std::time::{Duration, Instant};

use windows::core::{w, PCSTR, PCWSTR, PWSTR};
use windows::Win32::Foundation::{
    CloseHandle, FreeLibrary, GetLastError, SetHandleInformation, BOOL, COLORREF,
    ERROR_ALREADY_EXISTS, ERROR_MORE_DATA, FILETIME, HANDLE, HANDLE_FLAG_INHERIT, HINSTANCE,
    HMODULE, HWND, INVALID_HANDLE_VALUE, LPARAM, LRESULT, POINT, RECT, WAIT_OBJECT_0, WAIT_TIMEOUT,
    WPARAM,
};
use windows::Win32::Graphics::Dwm::{
    DwmFlush, DwmGetWindowAttribute, DwmQueryThumbnailSourceSize, DwmRegisterThumbnail,
    DwmSetWindowAttribute, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties, DWMWA_CLOAK,
    DWMWA_CLOAKED, DWMWA_EXTENDED_FRAME_BOUNDS, DWMWA_TRANSITIONS_FORCEDISABLED,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_OPACITY, DWM_TNP_RECTDESTINATION,
    DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{
    ClientToScreen, CreateSolidBrush, DeleteObject, FillRect, GetMonitorInfoW, GetStockObject,
    InvalidateRect, MonitorFromRect, UpdateWindow, HBRUSH, HGDIOBJ, MONITORINFO,
    MONITOR_DEFAULTTONEAREST, NULL_BRUSH,
};
use windows::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, Thread32First, Thread32Next, TH32CS_SNAPTHREAD, THREADENTRY32,
};
use windows::Win32::System::JobObjects::{
    AssignProcessToJobObject, CreateJobObjectW, JobObjectBasicProcessIdList,
    JobObjectExtendedLimitInformation, QueryInformationJobObject, SetInformationJobObject,
    JOBOBJECT_BASIC_PROCESS_ID_LIST, JOBOBJECT_EXTENDED_LIMIT_INFORMATION, JOB_OBJECT_LIMIT,
    JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE,
};
use windows::Win32::System::LibraryLoader::{GetModuleHandleW, GetProcAddress, LoadLibraryW};
use windows::Win32::System::Memory::{
    CreateFileMappingW, MapViewOfFile, UnmapViewOfFile, FILE_MAP_ALL_ACCESS,
    MEMORY_MAPPED_VIEW_ADDRESS, PAGE_READWRITE,
};
use windows::Win32::System::Threading::{
    CreateProcessW, GetProcessTimes, GetThreadId, IsWow64Process, OpenProcess,
    QueryFullProcessImageNameW, ResumeThread, TerminateProcess, WaitForInputIdle,
    WaitForSingleObject, CREATE_NO_WINDOW, CREATE_SUSPENDED, PROCESS_CREATION_FLAGS,
    PROCESS_INFORMATION, PROCESS_NAME_WIN32, PROCESS_QUERY_INFORMATION,
    PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_SET_QUOTA, PROCESS_SYNCHRONIZE, PROCESS_TERMINATE,
    STARTF_USEPOSITION, STARTF_USESHOWWINDOW, STARTF_USESIZE, STARTF_USESTDHANDLES, STARTUPINFOW,
    THREAD_QUERY_LIMITED_INFORMATION,
};
use windows::Win32::UI::Accessibility::{SetWinEventHook, UnhookWinEvent, HWINEVENTHOOK};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, IsWindowEnabled};
use windows::Win32::UI::WindowsAndMessaging::{
    AllowSetForegroundWindow, BringWindowToTop, CreateWindowExW, DefWindowProcW, DispatchMessageW,
    EnumWindows, GetAncestor, GetClassNameW, GetClientRect, GetForegroundWindow,
    GetLayeredWindowAttributes, GetSystemMetrics, GetWindow, GetWindowLongPtrW, GetWindowPlacement,
    GetWindowRect, GetWindowTextLengthW, GetWindowThreadProcessId, IsIconic, IsWindow,
    IsWindowVisible, IsZoomed, MoveWindow, MsgWaitForMultipleObjectsEx, PeekMessageW, PostMessageW,
    RegisterClassExW, RegisterWindowMessageW, SendMessageTimeoutW, SetForegroundWindow,
    SetLayeredWindowAttributes, SetWindowLongPtrW, SetWindowPlacement, SetWindowPos,
    SetWindowsHookExW, ShowWindow, TranslateMessage, UnhookWindowsHookEx, WindowFromPoint,
    CHILDID_SELF, EVENT_OBJECT_CREATE, EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_SHOW, GA_ROOT,
    GWL_EXSTYLE, GWL_STYLE, GW_OWNER, HHOOK, HOOKPROC, HWND_BOTTOM, HWND_NOTOPMOST, HWND_TOPMOST,
    LAYERED_WINDOW_ATTRIBUTES_FLAGS, LWA_ALPHA, MSG, MWMO_INPUTAVAILABLE, OBJID_WINDOW, PM_REMOVE,
    QS_ALLINPUT, SET_WINDOW_POS_FLAGS, SMTO_ABORTIFHUNG, SMTO_NORMAL, SM_CXVIRTUALSCREEN,
    SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN, SWP_ASYNCWINDOWPOS, SWP_FRAMECHANGED,
    SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOOWNERZORDER, SWP_NOSENDCHANGING, SWP_NOSIZE, SWP_NOZORDER,
    SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE, WH_CALLWNDPROC, WH_CBT, WINDOWPLACEMENT,
    WINDOWPLACEMENT_FLAGS, WINDOW_EX_STYLE, WINDOW_STYLE, WINEVENT_OUTOFCONTEXT,
    WINEVENT_SKIPOWNPROCESS, WM_APP, WM_CLOSE, WM_ERASEBKGND, WM_NCDESTROY, WM_NULL, WNDCLASSEXW,
    WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN, WS_EX_LAYERED, WS_EX_NOACTIVATE,
    WS_EX_NOREDIRECTIONBITMAP, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_MAXIMIZE,
    WS_MAXIMIZEBOX, WS_MINIMIZE, WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU, WS_THICKFRAME,
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
    build_launch_args_with, custom_preset_for, CustomCaptureTechnique, CustomChildDepth,
    CustomChildScope, CustomPathMethod, CustomPictureWindow, CustomPlacement, CustomPreset,
    CustomPromotionTechnique, CustomScaling, CustomSiblingWindows, CustomWindowMode, PreviewScale,
    PreviewType, COMPARABLE_AREA_PERCENT, DEFAULT_AUDIO_SCALE,
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

/// The smallest a window may be, in either direction, and still be a picture **when it has never been
/// seen visible**.
///
/// A hidden candidate is a guess: nothing of it has been composited, so its rectangle is all there is
/// to judge it by, and an area alone cannot tell a document from a program's own furniture — **VLC
/// puts up a top-level `Static` window titled `VLC ghk 3.0.24` at 136x39 before its media window
/// exists**, and at 5304 px² that cleared 64x64 and was adopted as the picture: the hover became a
/// white box while the real 1936x1192 window arrived 150 ms later, which is exactly the empty box a
/// VLC preset showed. Its 39-pixel height is what says it is a strip of text rather than a place a
/// picture is drawn, and the floor leaves the windows that already worked alone: mpv's first hidden
/// window is 136x100, and 100 clears this.
const THE_SMALLEST_SIDE: i32 = 64;

/// What a window's rectangle is worth to the ranking: the area it is drawn in, and whether either of
/// its sides is too thin for a picture (see `THE_SMALLEST_SIDE`).
///
/// Carried rather than recomputed because `classify_candidate` has to answer from the shape as well as
/// the size, and a caller holding a `window_rect` has both already.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct WindowSize {
    area: i64,
    /// True where a side is thinner than a picture's: a tooltip, a status strip, an event target.
    a_strip: bool,
}

impl WindowSize {
    fn of(rect: RECT) -> Self {
        let width = (rect.right - rect.left).max(0);
        let height = (rect.bottom - rect.top).max(0);
        Self {
            area: i64::from(width) * i64::from(height),
            a_strip: width.min(height) < THE_SMALLEST_SIDE,
        }
    }
}

/// How often the program's window is looked at again, and kept at the box if the program moved
/// itself: a window shown again by its own program, or restored by it, is one the Shell would
/// compose at the place the program chose rather than at the hover's.
const KEEP_POLL: Duration = Duration::from_millis(40);
const FORCED_POLL: Duration = Duration::from_millis(8);
const FORCED_STEADY_POLL: Duration = Duration::from_millis(16);
const FORCED_FAST_POLL_WINDOW: Duration = Duration::from_millis(1500);
const FORCED_SETTLE_STABLE: Duration = Duration::from_millis(200);
const FORCED_SETTLE_CAP: Duration = Duration::from_millis(1200);
const FORCED_CURSOR_STANDOFF_PX: i32 = 12;
const HOOK_FALLBACK_OWN_SIZE: (u32, u32) = (1280, 800);

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
    placement: CustomPlacement,
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

    /// Query live process IDs assigned to this Job Object, retrying once on `ERROR_MORE_DATA`.
    fn process_ids(&self) -> Vec<u32> {
        if self.handle.is_invalid() || self.handle.0.is_null() {
            return Vec::new();
        }
        let my_pid = std::process::id();
        let header_size = std::mem::size_of::<JOBOBJECT_BASIC_PROCESS_ID_LIST>();
        let ptr_size = std::mem::size_of::<usize>();
        let header_usizes = header_size.div_ceil(ptr_size);
        let mut cap: usize = 64;

        for attempt in 0..2 {
            let total_usizes = header_usizes + cap;
            let total_bytes = (total_usizes * ptr_size) as u32;
            let mut buf = vec![0usize; total_usizes];
            let res = unsafe {
                QueryInformationJobObject(
                    self.handle,
                    JobObjectBasicProcessIdList,
                    buf.as_mut_ptr().cast(),
                    total_bytes,
                    None,
                )
            };
            if let Err(err) = res {
                let is_more_data = err.code() == ERROR_MORE_DATA.to_hresult()
                    || unsafe { GetLastError() } == ERROR_MORE_DATA;
                if attempt == 0 && is_more_data {
                    let list =
                        unsafe { &*(buf.as_ptr() as *const JOBOBJECT_BASIC_PROCESS_ID_LIST) };
                    let needed = list.NumberOfAssignedProcesses as usize;
                    cap = needed.max(cap * 2).min(4096);
                    continue;
                }
                return Vec::new();
            }

            return parse_job_process_id_list(&buf, my_pid);
        }
        Vec::new()
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
struct Program {
    path: PathBuf,
    exe: String,
    arguments: String,
    previous_foreground: isize,
    process: Arc<OwnedHandle>,
    pid: u32,
    job: Option<PreviewJob>,
    processes_before: Vec<u32>,
    /// The preset's `Child Processes` row as it stood when this program was begun: how far the
    /// teardown sweep may walk to end what it started.
    child_scope: CustomChildScope,
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
    /// Initial primary thread id of the spawned process (`0` if unavailable).
    tid: u32,
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
    /// Preset placement mode (`FollowTray`, `Forced`, or `Hook`).
    placement: CustomPlacement,
    /// Capture technique originally requested by the preset (before 32-bit hook fallback).
    requested_capture: CustomCaptureTechnique,
    /// How far past `pid` discovery may look — the preset's `Child Processes` row, read once at
    /// launch so a config change mid-hover cannot half-move a window that is already adopted.
    child_scope: CustomChildScope,
    /// Whether a brokered/stub-handoff process may be adopted by image name or install folder.
    broker_fallback: bool,
    /// What happens to the program's other windows while one of them is the picture.
    sibling_windows: CustomSiblingWindows,
    /// Which of the program's windows is the picture, read once at launch for the same reason the
    /// rows beside it are: a preset edited mid-hover cannot half-change what is already adopted.
    picture_window: CustomPictureWindow,
    /// Dedicated install directory of the launched executable when safe for stub fallback matching.
    install_dir: Option<PathBuf>,
    /// Which launch this is, so that a worker begun for an earlier one stands down.
    id: u64,
}

fn parse_job_process_id_list(buf: &[usize], my_pid: u32) -> Vec<u32> {
    let ptr_size = std::mem::size_of::<usize>();
    let total_bytes = std::mem::size_of_val(buf);
    if total_bytes < std::mem::size_of::<JOBOBJECT_BASIC_PROCESS_ID_LIST>() {
        return Vec::new();
    }
    let list = unsafe { &*(buf.as_ptr() as *const JOBOBJECT_BASIC_PROCESS_ID_LIST) };
    let list_ptr = std::ptr::addr_of!(list.ProcessIdList).cast::<usize>();
    let offset_bytes = (list_ptr as usize).saturating_sub(buf.as_ptr() as usize);
    let max_slots = total_bytes.saturating_sub(offset_bytes) / ptr_size;
    let count = (list.NumberOfProcessIdsInList as usize).min(max_slots);
    let raw_pids = unsafe { std::slice::from_raw_parts(list_ptr, count) };
    let mut out = Vec::with_capacity(count);
    for &raw in raw_pids {
        if let Ok(pid) = u32::try_from(raw) {
            if pid != 0 && pid != my_pid && !out.contains(&pid) {
                out.push(pid);
            }
        }
    }
    out
}

fn active_job_pids(launch_id: u64) -> Vec<u32> {
    if !is_current(launch_id) {
        return Vec::new();
    }
    let Ok(held) = PROGRAM.lock() else {
        return Vec::new();
    };
    if !is_current(launch_id) {
        return Vec::new();
    }
    held.as_ref()
        .and_then(|program| program.job.as_ref())
        .map(PreviewJob::process_ids)
        .unwrap_or_default()
}

/// True only when `dir` is a dedicated application folder rather than a shallow or shared system/user container.
fn is_safe_app_install_dir(dir: &Path) -> bool {
    let mut count = 0usize;
    let mut last = None;
    for comp in dir.components() {
        match comp {
            std::path::Component::Normal(s) => {
                count += 1;
                last = Some(s);
            }
            std::path::Component::ParentDir | std::path::Component::CurDir => return false,
            std::path::Component::Prefix(_) | std::path::Component::RootDir => {}
        }
    }
    if count < 3 {
        return false;
    }
    let Some(last) = last else {
        return false;
    };
    let last_lower = last.to_string_lossy().to_lowercase();
    !matches!(
        last_lower.as_str(),
        "windows"
            | "system32"
            | "syswow64"
            | "program files"
            | "program files (x86)"
            | "programdata"
            | "local"
            | "locallow"
            | "roaming"
            | "programs"
            | "desktop"
            | "downloads"
            | "documents"
            | "temp"
            | "tmp"
            | "users"
            | "bin"
            | "winapps"
            | "windowsapps"
    )
}

fn normalized_path_components(path: &Path) -> Option<Vec<String>> {
    let mut out = Vec::new();
    for comp in path.components() {
        match comp {
            std::path::Component::Prefix(prefix) => match prefix.kind() {
                std::path::Prefix::Disk(d) | std::path::Prefix::VerbatimDisk(d) => {
                    out.push(format!("disk:{}", d.to_ascii_lowercase() as char));
                }
                std::path::Prefix::UNC(server, share)
                | std::path::Prefix::VerbatimUNC(server, share) => {
                    out.push(format!(
                        "unc:{}:{}",
                        server.to_string_lossy().to_lowercase(),
                        share.to_string_lossy().to_lowercase()
                    ));
                }
                _ => {
                    out.push(prefix.as_os_str().to_string_lossy().to_lowercase());
                }
            },
            std::path::Component::RootDir => {
                out.push("/".to_owned());
            }
            std::path::Component::CurDir => {}
            std::path::Component::ParentDir => return None,
            std::path::Component::Normal(seg) => {
                out.push(seg.to_string_lossy().to_lowercase());
            }
        }
    }
    (!out.is_empty()).then_some(out)
}

/// Case-insensitive path-component prefix check (`exe_path` strictly inside `install_dir`).
fn exe_is_inside_install_dir(exe_path: &Path, install_dir: &Path) -> bool {
    let (Some(exe_comps), Some(dir_comps)) = (
        normalized_path_components(exe_path),
        normalized_path_components(install_dir),
    ) else {
        return false;
    };
    exe_comps.len() > dir_comps.len() && exe_comps.starts_with(&dir_comps)
}

#[inline]
fn enforces_placement(placement: CustomPlacement) -> bool {
    matches!(placement, CustomPlacement::Forced | CustomPlacement::Hook)
}

#[inline]
fn uses_hook_enforcement(placement: CustomPlacement) -> bool {
    placement == CustomPlacement::Hook
}

#[inline]
fn sh_forced_mode(placement: CustomPlacement) -> u8 {
    match placement {
        CustomPlacement::FollowTray => 0,
        CustomPlacement::Forced => 1,
        CustomPlacement::Hook => 2,
    }
}

#[inline]
fn z_order_mode_for(capture_technique: CustomCaptureTechnique) -> u8 {
    match capture_technique {
        CustomCaptureTechnique::DirectWindow => ZORDER_TOPMOST,
        _ => ZORDER_BOTTOM,
    }
}

/// Concrete `Tray -> Placement` target rect for hook initialization (`Hook` mode uses fallback size
/// when the file declares no intrinsic dimensions so `HCBT_CREATEWND` always receives `w, h > 0`).
fn hook_provisional_rect(
    at: Where,
    path: &Path,
    scaling: CustomScaling,
    provisional: Option<(i32, i32, i32, i32)>,
) -> Option<(i32, i32, i32, i32)> {
    provisional.or_else(|| {
        let scale = resolve_scale(path, scaling);
        provisional_rect_for(
            at,
            HOOK_FALLBACK_OWN_SIZE,
            scale,
            follow_cursor(),
            bypass_placement(),
        )
    })
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
/// The program this app started for it, active while the hover is live.
static PROGRAM: Mutex<Option<Program>> = Mutex::new(None);
/// Lock-free flag indicating a custom preview launch/hover is live (set in `show`, cleared in
/// `end_the_program` / `try_detach_active_preview`).
static LIVE: AtomicBool = AtomicBool::new(false);
/// Lock-free flag indicating a `CustomPlacement::Forced` launch is in its initial settle/snap window.
static FORCED_SETTLING: AtomicBool = AtomicBool::new(false);
/// Launched child PID for the currently active custom preview.
static LAUNCHED_PID: AtomicU32 = AtomicU32::new(0);
/// Newly-created PIDs owned by the active custom preview (updated by `run_it`, cleared on teardown/detach).
static OWNED_PIDS: Mutex<Vec<u32>> = Mutex::new(Vec::new());
/// How many programs have been begun: a worker that was begun for an earlier one is one whose
/// program has been ended, and it neither places nor shows anything.
static LAUNCH: AtomicU64 = AtomicU64::new(0);
static CLASS_REGISTERED: Once = Once::new();

/// Whether a custom preview is on screen.
pub(super) fn is_showing() -> bool {
    WANT.lock().map(|want| want.is_some()).unwrap_or(false)
}

/// Lock-free check whether a custom preview is currently live (and not being promoted/detached).
pub(super) fn has_live_preview() -> bool {
    !DETACHING.load(Ordering::Acquire) && LIVE.load(Ordering::Acquire)
}

/// Lock-free check whether a `Forced`-placement custom preview is currently settling/snapping.
pub(super) fn is_forced_settling() -> bool {
    !DETACHING.load(Ordering::Acquire) && FORCED_SETTLING.load(Ordering::Acquire)
}

fn point_in_rect(pt: POINT, rect: RECT) -> bool {
    rect.right > rect.left
        && rect.bottom > rect.top
        && pt.x >= rect.left
        && pt.x < rect.right
        && pt.y >= rect.top
        && pt.y < rect.bottom
}

fn point_in_xywh(pt: POINT, rect: (i32, i32, i32, i32)) -> bool {
    rect.2 > 0
        && rect.3 > 0
        && pt.x >= rect.0
        && pt.x < rect.0 + rect.2
        && pt.y >= rect.1
        && pt.y < rect.1 + rect.3
}

fn is_rect_visible_on_screen(rect: RECT, vscreen: RECT) -> bool {
    rect.right > rect.left
        && rect.bottom > rect.top
        && !is_at_stage_coord(rect)
        && !is_off_screen(&rect, vscreen)
}

/// True when `hwnd` is visible, uncloaked, non-zero-alpha, and positioned on the virtual screen
/// (excludes hidden, cloaked, invisible-layered, and offscreen-staged helper/preview windows).
fn window_is_visible_on_screen(hwnd: HWND) -> bool {
    if hwnd.0.is_null() || !unsafe { IsWindow(hwnd) }.as_bool() {
        return false;
    }
    if !unsafe { IsWindowVisible(hwnd) }.as_bool() || unsafe { IsIconic(hwnd) }.as_bool() {
        return false;
    }

    let mut cloaked: u32 = 0;
    if unsafe {
        DwmGetWindowAttribute(
            hwnd,
            DWMWA_CLOAKED,
            &mut cloaked as *mut u32 as *mut _,
            std::mem::size_of::<u32>() as u32,
        )
    }
    .is_ok()
        && cloaked != 0
    {
        return false;
    }

    let source_raw = SOURCE.load(Ordering::Acquire);
    if hwnd.0 as isize == source_raw && SOURCE_INVISIBLE_LAYERED.load(Ordering::Acquire) {
        return false;
    }

    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    if (ex_style & WS_EX_LAYERED.0) != 0
        && layered_alpha_flags(hwnd)
            .is_some_and(|(alpha, flags)| (flags.0 & LWA_ALPHA.0) != 0 && alpha == 0)
    {
        return false;
    }

    window_rect(hwnd).is_some_and(|r| is_rect_visible_on_screen(r, virtual_screen()))
}

fn is_active_owned_pid(pid: u32, my_pid: u32) -> bool {
    if pid == 0 || pid == my_pid {
        return false;
    }
    if LAUNCHED_PID.load(Ordering::Acquire) == pid {
        return true;
    }
    if SOURCE_IS_NEW.load(Ordering::Acquire) && SOURCE_PID.load(Ordering::Acquire) == pid {
        return true;
    }
    OWNED_PIDS
        .lock()
        .ok()
        .is_some_and(|pids| pids.contains(&pid))
}

/// True when cursor `pt` hits the visible custom preview `HOST` rect, visible uncloaked `SOURCE`
/// rect (handles `WS_EX_TRANSPARENT` / `DirectWindow`), or any visible uncloaked window owned by
/// the spawned custom `.exe` / its newly created PIDs.
pub(super) fn cursor_is_over_preview(pt: POINT) -> bool {
    if !has_live_preview() {
        return false;
    }

    // Check 1: Visible HOST rect (handles WS_EX_TRANSPARENT where WindowFromPoint skips HOST).
    if SHOWN.load(Ordering::Acquire) {
        if let Ok(guard) = LAST_HOST_RECT.lock() {
            if guard.is_some_and(|r| point_in_xywh(pt, r)) {
                return true;
            }
        }
        let host_raw = HOST.load(Ordering::Acquire);
        if host_raw != 0 {
            let host_hwnd = HWND(host_raw as *mut _);
            if window_is_visible_on_screen(host_hwnd)
                && window_rect(host_hwnd).is_some_and(|r| point_in_rect(pt, r))
            {
                return true;
            }
        }
    }

    // Check 2 & 3 (SOURCE / OWNED_PIDS) are deferred while Forced placement is settling its window off cursor.
    if FORCED_SETTLING.load(Ordering::Acquire) {
        return false;
    }

    // Check 2: Visible uncloaked on-screen SOURCE rect (catches DirectWindow and startup flashes).
    let source_raw = SOURCE.load(Ordering::Acquire);
    if source_raw != 0 {
        let source_hwnd = HWND(source_raw as *mut _);
        if window_is_visible_on_screen(source_hwnd)
            && visible_frame_rect(source_hwnd).is_some_and(|r| point_in_rect(pt, r))
        {
            return true;
        }
    }

    // Check 3: WindowFromPoint and its GA_ROOT ancestor against HOST, SOURCE, and owned PIDs.
    let wfp = unsafe { WindowFromPoint(pt) };
    if !wfp.is_invalid() && !wfp.0.is_null() {
        let root = unsafe { GetAncestor(wfp, GA_ROOT) };
        let root = if root.is_invalid() || root.0.is_null() {
            wfp
        } else {
            root
        };
        let wfp_raw = wfp.0 as isize;
        let root_raw = root.0 as isize;

        let host_raw = HOST.load(Ordering::Acquire);
        if host_raw != 0
            && SHOWN.load(Ordering::Acquire)
            && (wfp_raw == host_raw || root_raw == host_raw)
            && window_is_visible_on_screen(HWND(host_raw as *mut _))
        {
            return true;
        }

        if source_raw != 0
            && (wfp_raw == source_raw || root_raw == source_raw)
            && window_is_visible_on_screen(HWND(source_raw as *mut _))
        {
            return true;
        }

        let my_pid = std::process::id();
        let wfp_pid = pid_of(wfp);
        let root_pid = pid_of(root);
        if (is_active_owned_pid(wfp_pid, my_pid) || is_active_owned_pid(root_pid, my_pid))
            && window_is_visible_on_screen(root)
        {
            return true;
        }
    }

    false
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

/// Env-gated diagnostic trace (`RHP_CUSTOM_PREVIEW_TRACE`) writing to `%TEMP%\rhp-custom-preview-trace.log`.
/// Truncates once per process via `STARTED: Once` and incurs only an `AtomicBool` load when unset.
mod trace {
    use std::ffi::OsStr;
    use std::fs::{File, OpenOptions};
    use std::io::Write;
    use std::path::{Path, PathBuf};
    use std::sync::atomic::{AtomicBool, Ordering};
    use std::sync::{Mutex, Once};
    use std::time::{SystemTime, UNIX_EPOCH};

    static READ: Once = Once::new();
    static ASKED: AtomicBool = AtomicBool::new(false);
    static STARTED: Once = Once::new();
    static FILE_LOCK: Mutex<()> = Mutex::new(());

    pub(super) fn env_flag_enabled(val: Option<&OsStr>) -> bool {
        let Some(val) = val else {
            return false;
        };
        let s = val.to_string_lossy();
        let trimmed = s.trim();
        !trimmed.is_empty() && trimmed != "0" && !trimmed.eq_ignore_ascii_case("false")
    }

    pub(super) fn enabled() -> bool {
        READ.call_once(|| {
            let on = env_flag_enabled(std::env::var_os("RHP_CUSTOM_PREVIEW_TRACE").as_deref());
            ASKED.store(on, Ordering::Relaxed);
        });
        ASKED.load(Ordering::Relaxed)
    }

    pub(super) fn log_path() -> PathBuf {
        std::env::temp_dir().join("rhp-custom-preview-trace.log")
    }

    pub(super) fn write_with_once(once: &Once, path: &Path, line: &str) {
        let _guard = FILE_LOCK.lock().ok();
        once.call_once(|| {
            let _ = File::create(path);
        });
        let since_epoch = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default();
        let secs = since_epoch.as_secs();
        let millis = since_epoch.subsec_millis();
        if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(path) {
            let _ = writeln!(file, "{secs}.{millis:03} {line}");
            let _ = file.flush();
        }
    }

    pub(super) fn record(line: &str) {
        if !enabled() {
            return;
        }
        write_with_once(&STARTED, &log_path(), line);
    }
}

macro_rules! cp_trace {
    ($($arg:tt)*) => {
        if trace::enabled() {
            trace::record(&format!($($arg)*));
        }
    };
}

fn win_err() -> u32 {
    unsafe { GetLastError().0 }
}

fn path_basename(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| "<none>".to_owned())
}

fn hwnd_class_name(hwnd: HWND) -> String {
    let mut buf = [0u16; 128];
    let len = unsafe { GetClassNameW(hwnd, &mut buf) } as usize;
    String::from_utf16_lossy(&buf[..len.min(buf.len())])
}

fn process_is_wow64(handle: HANDLE) -> Option<bool> {
    let mut wow64 = BOOL(0);
    unsafe { IsWow64Process(handle, &mut wow64) }
        .is_ok()
        .then_some(wow64.as_bool())
}

fn pid_is_wow64(pid: u32) -> Option<bool> {
    if pid == 0 {
        return None;
    }
    let handle = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) }.ok()?;
    let owned = unsafe { OwnedHandle::from_raw_handle(handle.0) };
    process_is_wow64(HANDLE(owned.as_raw_handle()))
}

#[allow(clippy::too_many_arguments)]
fn trace_window_diag(
    label: &str,
    hwnd: HWND,
    source_pid: u32,
    source_tid: u32,
    source_is_new: bool,
    hook_hits: Option<u32>,
    captured_hwnd: Option<isize>,
    extra: &str,
) {
    if !trace::enabled() {
        return;
    }
    let owner = unsafe { GetWindow(hwnd, GW_OWNER) }.map_or(0usize, |h| h.0 as usize);
    let style = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) } as u32;
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    let visible = unsafe { IsWindowVisible(hwnd) }.as_bool();
    let zoomed = unsafe { IsZoomed(hwnd) }.as_bool();
    let iconic = unsafe { IsIconic(hwnd) }.as_bool();
    let class = hwnd_class_name(hwnd);
    let wow64 = pid_is_wow64(source_pid);
    let win_rect = window_rect(hwnd);
    let vis_rect = visible_frame_rect(hwnd);
    cp_trace!(
        "{label}: hwnd=0x{:x} pid={} tid={} is_new={} wow64={:?} class={:?} style=0x{:x} ex=0x{:x} owner=0x{:x} vis={} zoom={} min={} win={:?} vis_rect={:?} hook_hits={:?} cap_hwnd={:?}{extra}",
        hwnd.0 as usize,
        source_pid,
        source_tid,
        source_is_new,
        wow64,
        class,
        style,
        ex_style,
        owner,
        visible,
        zoomed,
        iconic,
        win_rect,
        vis_rect,
        hook_hits,
        captured_hwnd,
    );
}

type RectTuple = (i32, i32, i32, i32);
type DriftSig = (Option<RectTuple>, RectTuple);

#[allow(clippy::too_many_arguments)]
fn trace_drift(
    count: &mut u32,
    last_sig: &mut Option<DriftSig>,
    id: u64,
    reason: &str,
    hwnd: HWND,
    expected: Option<RECT>,
    host_rect: RectTuple,
    hook_hits: Option<u32>,
) {
    if !trace::enabled() {
        return;
    }
    let cur = window_rect(hwnd);
    let cur_tuple = cur.map(|r| (r.left, r.top, r.right - r.left, r.bottom - r.top));
    let sig = (cur_tuple, host_rect);
    if *count < 12 || *last_sig != Some(sig) {
        cp_trace!(
            "drift[{}]: id={} reason={} hwnd=0x{:x} cur={:?} vis={:?} expected={:?} host={:?} zoom={} hook_hits={:?}",
            *count,
            id,
            reason,
            hwnd.0 as usize,
            cur,
            visible_frame_rect(hwnd),
            expected,
            host_rect,
            unsafe { IsZoomed(hwnd) }.as_bool(),
            hook_hits
        );
        *last_sig = Some(sig);
    }
    *count = count.saturating_add(1);
}

const TOOLWINDOW_FALLBACK_MIN_AREA: i64 = 160 * 120;

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum CandidateTier {
    ToolWindow = 1,
    UntitledVisible = 2,
    HiddenOwner = 3,
    Main = 4,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct CandidateRank {
    tier: CandidateTier,
    visible: bool,
    area: i64,
}

fn owner_rect_is_blocking(
    owner_present: bool,
    owner_visible: bool,
    owner_wh: Option<(i32, i32)>,
) -> bool {
    if !owner_present || !owner_visible {
        return false;
    }
    owner_wh.is_some_and(|(w, h)| w > 1 && h > 1)
}

fn window_is_blocking_owner(owner: HWND) -> bool {
    if owner.0.is_null() {
        return false;
    }
    let owner_visible = unsafe { IsWindowVisible(owner) }.as_bool();
    if !owner_visible {
        return false;
    }
    let mut rect = RECT::default();
    let owner_wh = unsafe { GetWindowRect(owner, &mut rect) }
        .is_ok()
        .then_some((rect.right - rect.left, rect.bottom - rect.top));
    owner_rect_is_blocking(true, owner_visible, owner_wh)
}

/// True only when `hwnd` has an owner that is both visible and non-tiny (`w > 1 && h > 1`).
/// Hidden/1x1 helper owners such as Delphi `TApplication` or WinForms owners are non-blocking.
fn owner_is_blocking(hwnd: HWND) -> bool {
    unsafe { GetWindow(hwnd, GW_OWNER) }
        .ok()
        .is_some_and(window_is_blocking_owner)
}

fn classify_candidate(
    titled: bool,
    strictly_owned: bool,
    owner_blocking: bool,
    tool_window: bool,
    visible: bool,
    size: WindowSize,
) -> Option<CandidateTier> {
    // **A window never seen visible is held to a picture's *shape* as well as its size**, because its
    // rectangle is all that is known about it: a program's tooltip is a titled, unowned, non-tool
    // window too, and VLC's is 136x39 (see `THE_SMALLEST_SIDE`). One that has been seen is judged as
    // before, so a small program's window is a small picture rather than no picture at all.
    let a_strip = !visible && size.a_strip;
    let area = size.area;
    if titled && !strictly_owned && !tool_window && !a_strip && area >= THE_SMALLEST_PICTURE {
        Some(CandidateTier::Main)
    } else if titled && !owner_blocking && !tool_window && !a_strip && area >= THE_SMALLEST_PICTURE
    {
        Some(CandidateTier::HiddenOwner)
    } else if visible && !owner_blocking && !tool_window && area >= THE_SMALLEST_PICTURE {
        Some(CandidateTier::UntitledVisible)
    } else if visible && !owner_blocking && tool_window && area >= TOOLWINDOW_FALLBACK_MIN_AREA {
        Some(CandidateTier::ToolWindow)
    } else {
        None
    }
}

fn select_best_candidate<T: Copy>(candidates: &[(CandidateRank, T)]) -> Option<(CandidateRank, T)> {
    let mut best: Option<(CandidateRank, T)> = None;
    for &(rank, item) in candidates {
        if best.is_none_or(|(best_rank, _)| rank > best_rank) {
            best = Some((rank, item));
        }
    }
    best
}

thread_local! {
    /// What this launch's discovery has watched appear, on the thread doing the watching.
    ///
    /// **One worker thread per launch**, so a thread-local is a per-launch record without a parameter
    /// threaded through every census call: `show` spawns the worker that adopts the window, and it is
    /// that thread which walks the desktop (see `run_it`).
    static SIGHTINGS: RefCell<Sightings> = RefCell::new(Sightings::default());
}

/// Forget every sighting, at the start of a discovery wait.
fn forget_sightings() {
    let _ = SIGHTINGS.try_with(|held| held.borrow_mut().clear());
}

/// How long the program is watched for a *new* candidate before a hidden winner is taken, where the
/// preset's `Picture Window` row asks for the window made last **and more than one candidate is of
/// the same size class**.
///
/// **The wait is the point of the row, not a cost of it.** A program that puts up a helper first and
/// its document second is a program that is still making windows while the helper is the largest one
/// there is; taking the last window made means knowing the program has stopped making them. Measured
/// on MPV Player's `mpv.exe`: the `mpv smtc` helper is born at 25–50 ms and the film's window at
/// 195–280 ms, so a settle shorter than that adopts the helper anyway and the row does nothing.
///
/// **What it costs a program that is not MPV Player**: two thirds of a second before a preview that
/// appears after a hidden window has stood still, because until the program stops making windows
/// there is no telling the last one from the next one. A window the program *shows* skips the wait.
const NEWEST_HIDDEN_SETTLE: Duration = Duration::from_millis(350);

/// What the census has watched appear, so that "the window the program made last" is a fact about
/// the launch rather than a guess about the z-order.
///
/// One per discovery wait, owned by the wait and handed to every census it takes: a census is one
/// look at the desktop, and which window arrived when is the one thing a look cannot answer.
#[derive(Default)]
struct Sightings {
    /// When each candidate window was first seen.
    first_seen: HashMap<isize, Instant>,
    /// The instant a candidate was last seen for the first time: the program's own last word on the
    /// subject, and the clock the hidden settle is measured from.
    last_new: Option<Instant>,
}

impl Sightings {
    fn note(&mut self, hwnd: HWND) {
        let raw = hwnd.0 as isize;
        if self.first_seen.contains_key(&raw) {
            return;
        }
        let now = Instant::now();
        self.first_seen.insert(raw, now);
        self.last_new = Some(now);
    }

    fn seen_at(&self, hwnd: HWND) -> Option<Instant> {
        self.first_seen.get(&(hwnd.0 as isize)).copied()
    }

    fn clear(&mut self) {
        self.first_seen.clear();
        self.last_new = None;
    }

    /// Whether the program has stopped making candidate windows for long enough that the last one it
    /// made can be taken as its last word.
    fn has_settled(&self) -> bool {
        self.last_new
            .is_none_or(|at| at.elapsed() >= NEWEST_HIDDEN_SETTLE)
    }
}

/// The picture among `candidates`, on the terms the preset's `Picture Window` row names.
///
/// `Largest` is the ranking this app has always had. `NewestComparable` fixes the best tier and
/// visibility first — a window the program *shows* still beats one it keeps hidden — and then, of the
/// windows whose area is within [`COMPARABLE_AREA_PERCENT`] of the largest there, prefers the one made
/// last. The margin is what keeps a toast from being mistaken for a document; the settle is what keeps
/// the first of a program's own two windows from being mistaken for its last (see
/// `NEWEST_HIDDEN_SETTLE`).
fn select_the_picture(
    candidates: &[(CandidateRank, HWND)],
    sightings: &Sightings,
    mode: CustomPictureWindow,
) -> Option<(CandidateRank, HWND)> {
    let largest = select_best_candidate(candidates)?;
    if !mode.prefers_the_newest() {
        return Some(largest);
    }
    let same_class: Vec<(CandidateRank, HWND)> = candidates
        .iter()
        .copied()
        .filter(|(rank, _)| rank.tier == largest.0.tier && rank.visible == largest.0.visible)
        .collect();
    let biggest = same_class.iter().map(|(rank, _)| rank.area).max()?;
    let comparable_area = biggest.saturating_mul(COMPARABLE_AREA_PERCENT) / 100;
    let mut best: Option<(CandidateRank, HWND, Instant)> = None;
    for (rank, item) in same_class {
        if rank.area < comparable_area {
            continue;
        }
        let seen_at = sightings.seen_at(item).unwrap_or_else(Instant::now);
        if best.is_none_or(|(_, _, at)| seen_at >= at) {
            best = Some((rank, item, seen_at));
        }
    }
    let (rank, item, _) = best?;
    // **A hidden winner waits for the program to stop making windows, whatever it looks like.** A
    // window that is up is the program's own answer and is taken at once; a window nobody has seen
    // is a guess, and the guess this row exists to correct is exactly the case where *one* bigger
    // window is already there and the right one is still to come — MPV Player's helper is alone on
    // the desktop from 25 ms to 250 ms, so "one candidate is nothing to choose" is what adopted it.
    // Measured: skipping this wait for a single hidden candidate adopts `mpv smtc` at 26 ms, and
    // waiting adopts the film at 578 ms.
    if !rank.visible && !sightings.has_settled() {
        return None;
    }
    if item != largest.1 {
        cp_trace!(
            "picture: newest_comparable chose hwnd=0x{:x} area={} visible={} over hwnd=0x{:x} \
             area={} of {} candidates",
            item.0 as usize,
            rank.area,
            rank.visible,
            largest.1 .0 as usize,
            largest.0.area,
            candidates.len()
        );
    }
    Some((rank, item))
}

static LAST_DUMPED_300MS_ID: AtomicU64 = AtomicU64::new(0);

type ProcessImages = HashMap<u32, Option<PathBuf>>;

fn maybe_trace_300ms_no_match(
    tag: &str,
    launch: &Launch,
    started: Instant,
    dumped: &mut bool,
    processes: &[u32],
    images: &mut ProcessImages,
) {
    if !trace::enabled() || *dumped || started.elapsed() < Duration::from_millis(300) {
        return;
    }
    *dumped = true;
    if LAST_DUMPED_300MS_ID.swap(launch.id, Ordering::AcqRel) == launch.id {
        return;
    }

    struct DumpState<'a> {
        launch: &'a Launch,
        tree: &'a [u32],
        job_pids: &'a [u32],
        images: &'a mut ProcessImages,
        my_pid: u32,
        rows: Vec<(HWND, u32, u32, bool, bool)>,
        extra_visible: Vec<(HWND, u32)>,
    }

    unsafe extern "system" fn visit_dump(hwnd: HWND, lparam: LPARAM) -> BOOL {
        if hwnd.0.is_null() {
            return BOOL(1);
        }
        let state = unsafe { &mut *(lparam.0 as *mut DumpState<'_>) };
        let mut pid = 0u32;
        let tid = unsafe { GetWindowThreadProcessId(hwnd, Some(&mut pid)) };
        if pid == 0 {
            return BOOL(1);
        }
        let in_tree = state.tree.contains(&pid);
        let in_job = state.job_pids.contains(&pid);
        let full_path = image_path_of(pid, state.images);
        let matches_exe = !state.launch.exe.is_empty()
            && full_path
                .as_deref()
                .and_then(|p| p.file_name())
                .map(|n| n.to_string_lossy().to_lowercase())
                .as_deref()
                == Some(state.launch.exe.as_str());
        let in_install_dir = !state.launch.processes_before.contains(&pid)
            && state
                .launch
                .install_dir
                .as_deref()
                .zip(full_path.as_deref())
                .is_some_and(|(dir, exe)| exe_is_inside_install_dir(exe, dir));
        if in_tree || in_job || matches_exe || in_install_dir {
            state.rows.push((hwnd, pid, tid, in_tree, in_job));
        } else if state.extra_visible.len() < 10
            && pid != state.my_pid
            && !state.launch.windows_before.contains(&(hwnd.0 as isize))
            && unsafe { IsWindowVisible(hwnd) }.as_bool()
        {
            state.extra_visible.push((hwnd, pid));
        }
        BOOL(1)
    }

    let tree = process_tree(launch.pid);
    let _ = processes;
    let job_pids = active_job_pids(launch.id);
    let mut state = DumpState {
        launch,
        tree: &tree,
        job_pids: &job_pids,
        images,
        my_pid: std::process::id(),
        rows: Vec::new(),
        extra_visible: Vec::new(),
    };
    let _ = unsafe {
        EnumWindows(
            Some(visit_dump),
            LPARAM(&mut state as *mut DumpState<'_> as isize),
        )
    };
    let row_count = state.rows.len();
    cp_trace!(
        "{tag}: 300ms_no_match_dump id={} exe={} launch_pid={} tree={:?} job_pids={:?} count={}",
        launch.id,
        launch.exe,
        launch.pid,
        tree,
        job_pids,
        row_count
    );
    for (hwnd, pid, tid, in_tree, in_job) in state.rows {
        let class = hwnd_class_name(hwnd);
        let title_len = unsafe { GetWindowTextLengthW(hwnd) };
        let style = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) } as u32;
        let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
        let owner_hwnd = unsafe { GetWindow(hwnd, GW_OWNER) }
            .ok()
            .filter(|o| !o.0.is_null());
        let owner_raw = owner_hwnd.map_or(0usize, |o| o.0 as usize);
        let owner_vis = owner_hwnd.is_some_and(|o| unsafe { IsWindowVisible(o) }.as_bool());
        let vis = unsafe { IsWindowVisible(hwnd) }.as_bool();
        let rect = window_rect(hwnd);
        let in_before = launch.windows_before.contains(&(hwnd.0 as isize));
        cp_trace!(
            "{tag}_win: id={} hwnd=0x{:x} pid={} tid={} in_tree={} in_job={} class={:?} title_len={} style=0x{:x} ex=0x{:x} owner=0x{:x} owner_vis={} vis={} rect={:?} in_before={}",
            launch.id,
            hwnd.0 as usize,
            pid,
            tid,
            in_tree,
            in_job,
            class,
            title_len,
            style,
            ex_style,
            owner_raw,
            owner_vis,
            vis,
            rect,
            in_before
        );
    }
    if row_count == 0 {
        for (hwnd, pid) in state.extra_visible {
            let exe = image_name_of(pid, state.images).unwrap_or_else(|| "<unknown>".to_owned());
            let class = hwnd_class_name(hwnd);
            let rect = window_rect(hwnd);
            cp_trace!(
                "{tag}_desktop_new: id={} hwnd=0x{:x} pid={} exe={:?} class={:?} rect={:?}",
                launch.id,
                hwnd.0 as usize,
                pid,
                exe,
                class,
                rect
            );
        }
    }
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
        cp_trace!("show: exe_missing path={}", path_basename(path));
        // A preset whose program is not there is one no preview can be made of, and the hover is
        // left to whatever else this app would have done with the file.
        hide();
        return false;
    };

    let placement = preset.placement;
    let arguments = preset.arguments.clone();
    let path_method = preset.path_method;

    // **The same program is not begun twice.** A hover of this file that this app's program is
    // already drawing is put back up wherever the hand now is — which is the whole of what keeps
    // the program's own first window from being a launch loop (see the module note).
    let the_same_program = PROGRAM
        .lock()
        .ok()
        .and_then(|held| {
            held.as_ref().map(|program| {
                (
                    program.path.clone(),
                    program.exe.clone(),
                    program.arguments.clone(),
                )
            })
        })
        .is_some_and(|(path_of, exe_of, args_of)| {
            path_of.as_path() == path && exe_of.eq_ignore_ascii_case(&exe) && args_of == arguments
        });
    if the_same_program {
        if let Ok(mut want) = WANT.lock() {
            *want = Some(Want { at, placement });
        }
        LIVE.store(true, Ordering::Release);
        cp_trace!(
            "show: reuse path={} exe={} arguments={:?} placement={:?} cursor={:?}",
            path_basename(path),
            path_basename(Path::new(&exe)),
            arguments,
            placement,
            where_cursor(at)
        );
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
    let requested_capture = preset.capture_technique;
    let child_scope = preset.child_scope;
    let broker_fallback = preset.broker_fallback;
    let sibling_windows = preset.sibling_windows;
    let picture_window = preset.picture_window;
    let volume = preset.volume_enabled.then_some(preset.volume);
    let initial_own = base_size(path);
    let provisional = initial_own.and_then(|own| {
        let scale = resolve_scale(path, scaling);
        provisional_rect_for(at, own, scale, follow_cursor(), bypass_placement())
    });
    let hook_provisional = if uses_hook_enforcement(placement) {
        hook_provisional_rect(at, path, scaling, provisional)
    } else {
        provisional
    };

    let launch_filetime = current_filetime_u64();
    let launched_at = Instant::now();
    let (process, pid, tid, hook_session, capture_technique) = if requested_capture
        == CustomCaptureTechnique::InProcessHook
        || uses_hook_enforcement(placement)
    {
        let init = hook_init_for(requested_capture, placement, hook_provisional);
        let Some((proc_handle, pid, tid, session)) = launch_with_hook(
            &exe,
            &arguments,
            path,
            path_method,
            hook_provisional,
            requested_capture,
            init,
        ) else {
            cp_trace!(
                "show: launch_with_hook_failed path={} exe={} req_cap={:?} placement={:?}",
                path_basename(path),
                path_basename(Path::new(&exe)),
                requested_capture,
                placement
            );
            return false;
        };
        let effective =
            if requested_capture == CustomCaptureTechnique::InProcessHook && session.is_none() {
                CustomCaptureTechnique::InvisibleLayered
            } else {
                requested_capture
            };
        (proc_handle, pid, tid, session, effective)
    } else {
        let Some((proc_handle, pid)) = launch(
            &exe,
            &arguments,
            path,
            path_method,
            provisional,
            requested_capture,
        ) else {
            cp_trace!(
                "show: launch_failed path={} exe={} req_cap={:?} placement={:?}",
                path_basename(path),
                path_basename(Path::new(&exe)),
                requested_capture,
                placement
            );
            return false;
        };
        (proc_handle, pid, 0, None, requested_capture)
    };
    let process = Arc::new(process);
    let Some(host) = (unsafe { ensure_host() }) else {
        cp_trace!(
            "show: ensure_host_failed path={} exe={} pid={}",
            path_basename(path),
            path_basename(Path::new(&exe)),
            pid
        );
        drop(hook_session);
        let _ = unsafe { TerminateProcess(HANDLE(process.as_raw_handle()), 1) };
        return false;
    };

    let exe_path = Path::new(&exe);
    let exe_file = exe_path
        .file_name()
        .map(|name| name.to_string_lossy().to_lowercase())
        .unwrap_or_default();
    let packaged_broker = is_packaged_broker_path(exe_path);
    let install_dir = exe_path
        .parent()
        .filter(|dir| is_safe_app_install_dir(dir))
        .map(Path::to_path_buf);
    // Per-preview detachable Job Object: crash-reaps while previewing, detaches cleanly on promote.
    let job = PreviewJob::create_and_assign(pid);
    engine_processes::record_player(&exe, pid);
    if let Ok(mut held) = PROGRAM.lock() {
        *held = Some(Program {
            path: path.to_path_buf(),
            exe: exe.clone(),
            arguments: arguments.clone(),
            previous_foreground,
            process: Arc::clone(&process),
            pid,
            job,
            processes_before: processes_before.clone(),
            child_scope,
        });
    }
    if let Ok(mut want) = WANT.lock() {
        *want = Some(Want { at, placement });
    }
    LAUNCHED_PID.store(pid, Ordering::Release);
    if let Ok(mut owned) = OWNED_PIDS.lock() {
        *owned = (pid != 0 && pid != std::process::id())
            .then_some(pid)
            .into_iter()
            .collect();
    }
    FORCED_SETTLING.store(enforces_placement(placement), Ordering::Release);
    LIVE.store(true, Ordering::Release);
    if let Ok(mut slot) = ACTIVE_CAPTURE_TECHNIQUE.lock() {
        *slot = capture_technique;
    }
    if let Ok(mut slot) = ACTIVE_WINDOW_MODE.lock() {
        *slot = window_mode;
    }

    // Read after the launch is written down, so a worker begun for it is current exactly when it
    // is the program on screen.
    let id = LAUNCH.fetch_add(1, Ordering::SeqCst) + 1;
    let prev_fg_class = if trace::enabled() && previous_foreground != 0 {
        hwnd_class_name(HWND(previous_foreground as *mut _))
    } else {
        String::new()
    };
    cp_trace!(
        "show: launch id={} path={} exe={} arguments={:?} req_cap={:?} cap={:?} placement={:?} win_mode={:?} scaling={:?} own={:?} prov={:?} hook_prov={:?} pid={} tid={} hook_session={} packaged={} has_install_dir={} prev_fg=0x{:x} prev_fg_class={:?} cursor={:?}",
        id,
        path_basename(path),
        exe_file,
        arguments,
        requested_capture,
        capture_technique,
        placement,
        window_mode,
        scaling,
        initial_own,
        provisional,
        hook_provisional,
        pid,
        tid,
        hook_session.is_some(),
        packaged_broker,
        install_dir.is_some(),
        previous_foreground as usize,
        prev_fg_class,
        where_cursor(at)
    );
    let launch = Launch {
        path: path.to_path_buf(),
        initial_own,
        provisional: hook_provisional,
        packaged_broker,
        process,
        pid,
        tid,
        exe: exe_file,
        previous_foreground,
        windows_before,
        processes_before,
        launch_filetime,
        launched_at,
        placement,
        requested_capture,
        child_scope,
        broker_fallback,
        sibling_windows,
        picture_window,
        install_dir,
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

/// Take the hover down immediately: tears down the host and spawned preview program unless a
/// promote/detach is currently in flight.
pub(super) fn hide() {
    let detaching = DETACHING.load(Ordering::Acquire);
    let showing = is_showing();
    if detaching || showing {
        cp_trace!("hide: detaching={} showing={}", detaching, showing);
    }
    if detaching {
        return;
    }
    end_the_program();
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
    // Clear hook shared-memory locks immediately before promotion moves/restores `source_hwnd`.
    LAUNCH.fetch_add(1, Ordering::SeqCst);
    FORCED_SETTLING.store(false, Ordering::Release);
    let hook_ctrl = ACTIVE_HOOK_CONTROL.lock().ok().and_then(|mut g| g.take());
    if let Some(ref ctrl) = hook_ctrl {
        ctrl.clear_locks_for_detach();
    }
    LIVE.store(false, Ordering::Release);
    LAUNCHED_PID.store(0, Ordering::Release);
    if let Ok(mut owned) = OWNED_PIDS.lock() {
        owned.clear();
    }
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
    let _ = set_window_cloak(source_hwnd, false);
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
        engine_processes::forget(source_pid);
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

/// Spawn a completely untouched, detached process for `exe` + `arguments` + `path`
/// (with `CREATE_NO_WINDOW`, null stdio handles, no job, no hooks, no STARTUPINFO).
fn spawn_detached_normal(exe: &str, arguments: &str, path: &Path, method: CustomPathMethod) {
    let mut cmd = Command::new(exe);
    cmd.args(build_launch_args_with(arguments, path, method));
    cmd.stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null());
    cmd.creation_flags(CREATE_NO_WINDOW.0);
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
    arguments: &str,
    method: CustomPathMethod,
    try_detach: D,
    spawn_new: S,
) -> bool
where
    D: FnOnce(&Path) -> bool,
    S: FnOnce(&str, &str, &Path, CustomPathMethod),
{
    match technique {
        CustomPromotionTechnique::DirectPromotion => {
            if try_detach(path) {
                return true;
            }
            end_the_program();
            spawn_new(exe, arguments, path, method);
            true
        }
        CustomPromotionTechnique::OpenAsNewProcess => {
            clear_want();
            hide_the_host();
            end_the_program();
            spawn_new(exe, arguments, path, method);
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
        &preset.arguments,
        preset.path_method,
        try_detach_active_preview,
        spawn_detached_normal,
    )
}

struct OwnedPidInputs<'a> {
    launched_pid: u32,
    launched_tree: &'a [u32],
    job_pids: &'a [u32],
    source_pid: u32,
    source_is_new: bool,
    source_tree: &'a [u32],
    tracked_owned: &'a [u32],
    processes_before: &'a [u32],
    my_pid: u32,
}

fn merge_owned_pids_to_end(inputs: OwnedPidInputs<'_>) -> Vec<u32> {
    let OwnedPidInputs {
        launched_pid,
        launched_tree,
        job_pids,
        source_pid,
        source_is_new,
        source_tree,
        tracked_owned,
        processes_before,
        my_pid,
    } = inputs;
    let mut pids = Vec::new();
    let mut push_if_owned = |pid: u32, allow_if_launched: bool| {
        if pid == 0 || pid == my_pid {
            return;
        }
        if !allow_if_launched && processes_before.contains(&pid) {
            return;
        }
        if !pids.contains(&pid) {
            pids.push(pid);
        }
    };

    if launched_pid != 0 {
        push_if_owned(launched_pid, true);
        for &pid in launched_tree {
            push_if_owned(pid, pid == launched_pid);
        }
    }
    for &pid in job_pids {
        push_if_owned(pid, pid == launched_pid && launched_pid != 0);
    }
    if source_is_new && source_pid != 0 {
        push_if_owned(source_pid, false);
        for &pid in source_tree {
            push_if_owned(pid, false);
        }
    }
    for &pid in tracked_owned {
        push_if_owned(pid, pid == launched_pid && launched_pid != 0);
    }
    pids
}

/// End the program, its window's registration, and the hover it was drawing.
fn end_the_program() {
    // First, so that a worker still waiting on the program's window stops placing and showing.
    LAUNCH.fetch_add(1, Ordering::SeqCst);
    FORCED_SETTLING.store(false, Ordering::Release);
    LIVE.store(false, Ordering::Release);
    LAUNCHED_PID.store(0, Ordering::Release);
    SHOWN.store(false, Ordering::SeqCst);

    if let Ok(mut want) = WANT.lock() {
        *want = None;
    }

    let program = PROGRAM.lock().ok().and_then(|mut held| held.take());
    let tracked_owned = OWNED_PIDS
        .lock()
        .ok()
        .map(|mut held| std::mem::take(&mut *held))
        .unwrap_or_default();

    // **The window this app asked for is closed, and the process that opened it is ended only
    // where the launch is what made it.** A packaged program's window belongs to a process this
    // app never started, and that process may be one the user had open before the hover: its
    // window is the document this app asked for, so the window goes and the process stays.
    let source = SOURCE.swap(0, Ordering::SeqCst);
    let source_pid = SOURCE_PID.swap(0, Ordering::SeqCst);
    let source_is_new = SOURCE_IS_NEW.swap(false, Ordering::SeqCst);
    SOURCE_INVISIBLE_LAYERED.store(false, Ordering::SeqCst);
    let hook_ctrl = ACTIVE_HOOK_CONTROL
        .lock()
        .ok()
        .and_then(|mut ctrl| ctrl.take());
    if let Some(ref ctrl) = hook_ctrl {
        ctrl.clear_locks_for_detach();
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
    if program.is_some() || source != 0 || !tracked_owned.is_empty() {
        cp_trace!(
            "end_the_program: has_program={} source=0x{:x} source_pid={} source_is_new={} owned_pids={:?} restore_rect={:?}",
            program.is_some(),
            source as usize,
            source_pid,
            source_is_new,
            tracked_owned,
            restore_rect
        );
    }
    if source != 0 {
        let source = HWND(source as *mut _);
        if unsafe { IsWindow(source) }.as_bool() {
            if !source_is_new {
                if let Some(ref ctrl) = hook_ctrl {
                    let _ = ctrl.detach_sync(source);
                }
                let _ = set_window_cloak(source, false);
            }
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

    let fg_before = unsafe { GetForegroundWindow() };
    let previous_foreground = program
        .as_ref()
        .map(|program| program.previous_foreground)
        .unwrap_or(0);
    let my_pid = std::process::id();
    let launched_pid = program.as_ref().map(|program| program.pid).unwrap_or(0);
    let processes_before = program
        .as_ref()
        .map(|program| program.processes_before.as_slice())
        .unwrap_or(&[]);
    let launched_tree = if launched_pid != 0 && launched_pid != my_pid {
        match program.as_ref() {
            Some(program) => process_tree_to(launched_pid, program.child_scope.depth()),
            None => Vec::new(),
        }
    } else {
        Vec::new()
    };
    let job_pids = program
        .as_ref()
        .filter(|program| program.child_scope != CustomChildScope::Disabled)
        .and_then(|program| program.job.as_ref())
        .map(PreviewJob::process_ids)
        .unwrap_or_default();
    let source_tree = if source_is_new
        && source_pid != 0
        && source_pid != my_pid
        && program
            .as_ref()
            .is_none_or(|program| program.child_scope != CustomChildScope::Disabled)
    {
        process_tree_to(
            source_pid,
            program
                .as_ref()
                .map_or(CustomChildDepth::WholeTree, |program| {
                    program.child_scope.depth()
                }),
        )
    } else {
        Vec::new()
    };
    let pids_to_end = merge_owned_pids_to_end(OwnedPidInputs {
        launched_pid,
        launched_tree: &launched_tree,
        job_pids: &job_pids,
        source_pid,
        source_is_new,
        source_tree: &source_tree,
        tracked_owned: &tracked_owned,
        processes_before,
        my_pid,
    });

    if let Some(program) = program {
        if program.pid != 0 && program.pid != my_pid {
            let _ = unsafe { TerminateProcess(HANDLE(program.process.as_raw_handle()), 1) };
            engine_processes::forget(program.pid);
        }
        // Dropping `program` drops `PreviewJob` (triggering `KILL_ON_JOB_CLOSE`).
    }
    for &pid in &pids_to_end {
        end_process(pid);
    }

    let thumbnail = THUMBNAIL
        .lock()
        .map(|mut held| std::mem::replace(&mut *held, 0))
        .unwrap_or(0);
    if thumbnail != 0 {
        let _ = unsafe { DwmUnregisterThumbnail(thumbnail) };
    }
    hide_the_host();
    restore_foreground_on_teardown(previous_foreground, fg_before, &pids_to_end);
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
pub(crate) fn resolved(exe: &str) -> Option<String> {
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

/// Build NUL-terminated UTF-16 command line `"exe" ...args...` for `CreateProcessW`.
fn build_command_line(
    exe: &Path,
    arguments: &str,
    path: &Path,
    method: CustomPathMethod,
) -> Vec<u16> {
    let mut line = Vec::new();
    quote_win32_arg(exe.as_os_str(), &mut line);
    for arg in build_launch_args_with(arguments, path, method) {
        line.push(b' ' as u16);
        quote_win32_arg(OsStr::new(&arg), &mut line);
    }
    line.push(0);
    line
}

/// Process creation flags used when launching custom-preview processes (always silent: `CREATE_NO_WINDOW`).
fn preview_creation_flags(suspended: bool) -> PROCESS_CREATION_FLAGS {
    if suspended {
        CREATE_SUSPENDED | CREATE_NO_WINDOW
    } else {
        CREATE_NO_WINDOW
    }
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

/// Open `NUL` read+write with `HANDLE_FLAG_INHERIT` so console-subsystem preview targets (e.g.
/// `mpv.exe`) never inherit/touch an open terminal/console handle.
fn open_inheritable_nul() -> Option<std::fs::File> {
    let nul = std::fs::OpenOptions::new()
        .read(true)
        .write(true)
        .open("NUL")
        .ok()?;
    let handle = HANDLE(nul.as_raw_handle());
    unsafe { SetHandleInformation(handle, HANDLE_FLAG_INHERIT.0, HANDLE_FLAG_INHERIT) }.ok()?;
    Some(nul)
}

/// Point `hStdInput`, `hStdOutput`, and `hStdError` at `nul_handle` and set `STARTF_USESTDHANDLES`.
fn attach_nul_stdio(si: &mut STARTUPINFOW, nul_handle: HANDLE) {
    si.dwFlags |= STARTF_USESTDHANDLES;
    si.hStdInput = nul_handle;
    si.hStdOutput = nul_handle;
    si.hStdError = nul_handle;
}

/// Spawn target via `CreateProcessW` with `CREATE_NO_WINDOW` and `NUL` stdio redirection
/// (`STARTF_USESTDHANDLES`), its exe folder as cwd, and `STARTUPINFOW` `SW_HIDE` + provisional position/size hints.
fn launch(
    exe: &str,
    arguments: &str,
    path: &Path,
    method: CustomPathMethod,
    provisional: Option<(i32, i32, i32, i32)>,
    capture_technique: CustomCaptureTechnique,
) -> Option<(OwnedHandle, u32)> {
    let application = wide_null(OsStr::new(exe));
    let mut cmdline = build_command_line(Path::new(exe), arguments, path, method);
    let cwd_wide = Path::new(exe)
        .parent()
        .filter(|dir| !dir.as_os_str().is_empty())
        .map(|dir| wide_null(dir.as_os_str()));
    let cwd_ptr = cwd_wide
        .as_ref()
        .map_or(PCWSTR::null(), |buf| PCWSTR(buf.as_ptr()));
    let mut si = launch_startup_info_for(provisional, capture_technique, virtual_screen());
    let nul = open_inheritable_nul();
    let inherit_handles = if let Some(ref file) = nul {
        attach_nul_stdio(&mut si, HANDLE(file.as_raw_handle()));
        BOOL(1)
    } else {
        BOOL(0)
    };
    let mut pi = PROCESS_INFORMATION::default();

    let created = unsafe {
        CreateProcessW(
            PCWSTR(application.as_ptr()),
            PWSTR(cmdline.as_mut_ptr()),
            None,
            None,
            inherit_handles,
            preview_creation_flags(false),
            None,
            cwd_ptr,
            &si,
            &mut pi,
        )
    };
    drop(nul);
    created.ok()?;
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

fn launch_still_wanted(launch: &Launch) -> bool {
    is_current(launch.id) && is_showing()
}

fn update_owned_pids_if_current(launch: &Launch, pids: &[u32]) {
    let my_pid = std::process::id();
    let filtered: Vec<u32> = pids
        .iter()
        .copied()
        .filter(|&pid| {
            pid != 0
                && pid != my_pid
                && (pid == launch.pid || !launch.processes_before.contains(&pid))
        })
        .collect();
    if let Ok(mut owned) = OWNED_PIDS.lock() {
        if is_current(launch.id) {
            *owned = filtered;
        }
    }
}

/// End `launch.pid` and any newly spawned descendants or active job processes, excluding pre-launch and self PIDs.
///
/// **Depth is the preset's `Child Processes` row here too**, so a preset that asked for the
/// launched pid alone leaves the processes it started running — which is what a user who gated the
/// walk is asking for, and why the gate is the same one discovery reads rather than a second switch.
fn sweep_launch_tree(launch: &Launch) {
    let my_pid = std::process::id();
    let mut to_end = Vec::new();
    if launch.pid != 0 && launch.pid != my_pid {
        to_end.extend(discovery_tree(launch));
    }
    for pid in discovery_job_pids(launch) {
        if !to_end.contains(&pid) {
            to_end.push(pid);
        }
    }
    for pid in to_end {
        if pid != 0
            && pid != my_pid
            && (pid == launch.pid || !launch.processes_before.contains(&pid))
        {
            end_process(pid);
        }
    }
}

/// Clean up a newly discovered `source` window (and its new process tree if newly spawned) when
/// a launch is aborted concurrently by `hide()` / `end_the_program()`.
fn cleanup_aborted_source(source: HWND, launch: &Launch) {
    let my_pid = std::process::id();
    let mut source_pid = 0u32;
    let is_window = !source.0.is_null() && unsafe { IsWindow(source) }.as_bool();
    if is_window {
        source_pid = pid_of(source);
    }
    let source_is_new =
        source_pid != 0 && source_pid != my_pid && !launch.processes_before.contains(&source_pid);
    if is_window {
        if !source_is_new {
            let _ = set_window_cloak(source, false);
        }
        let _ = unsafe { PostMessageW(source, WM_CLOSE, WPARAM(0), LPARAM(0)) };
    }
    if source_is_new && launch.child_scope != CustomChildScope::Disabled {
        for pid in process_tree_to(source_pid, launch.child_scope.depth()) {
            if pid != 0 && pid != my_pid && !launch.processes_before.contains(&pid) {
                end_process(pid);
            }
        }
    }
    sweep_launch_tree(launch);
}

/// Find the program's window and keep the box and the picture in step with the hover until the
/// hover ends.
fn run_it(
    host: HWND,
    launch: Launch,
    scaling: CustomScaling,
    window_mode: CustomWindowMode,
    mut capture_technique: CustomCaptureTechnique,
    volume: Option<u32>,
    mut hook_session: Option<InProcessHookSession>,
) {
    let Some(discovered) = (if capture_technique == CustomCaptureTechnique::InProcessHook {
        if let Some(ref session) = hook_session {
            wait_for_a_window_in_process_hook(&launch, session, capture_technique)
        } else {
            wait_for_a_window(&launch, capture_technique)
        }
    } else if capture_technique == CustomCaptureTechnique::OffscreenStaged {
        wait_for_a_window_offscreen_staged(&launch, capture_technique)
    } else {
        wait_for_a_window(&launch, capture_technique)
    }) else {
        cp_trace!(
            "run_it: no_window_found id={} pid={} cap={:?} hook_hits={:?}",
            launch.id,
            launch.pid,
            capture_technique,
            hook_session.as_ref().map(|s| s.view.hook_hits())
        );
        // A program that made no window is a preview with nothing to show: it goes, and the host
        // is already off the screen. Even if `hide()`/`end_the_program()` already invalidated this
        // launch concurrently, sweep `launch.pid` so children spawned during the final sleep cannot leak.
        if is_current(launch.id) {
            end_the_program();
        }
        sweep_launch_tree(&launch);
        return;
    };
    let source = discovered.hwnd;
    let mut pre_cloaked = discovered.pre_cloaked;
    if hook_session.is_none() {
        hook_session = discovered.hook_session;
    } else if let Some(extra) = discovered.hook_session {
        extra.view.clear_locks_for_detach();
    }
    if !launch_still_wanted(&launch) || !unsafe { IsWindow(source) }.as_bool() {
        cp_trace!(
            "run_it: aborted_after_find id={} hwnd=0x{:x} still_wanted={} is_window={}",
            launch.id,
            source.0 as usize,
            launch_still_wanted(&launch),
            unsafe { IsWindow(source) }.as_bool()
        );
        if let Some(ref session) = hook_session {
            session.view.clear_locks_for_detach();
            let _ = session.view.detach_sync(source);
        }
        drop(hook_session.take());
        cleanup_aborted_source(source, &launch);
        if is_current(launch.id) {
            end_the_program();
        }
        return;
    }

    let mut source_pid = 0u32;
    let source_tid = unsafe { GetWindowThreadProcessId(source, Some(&mut source_pid)) };
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

    // Lock ownership strictly to our launched tree (+ active job PIDs + adopted brokered tree if newly spawned).
    let mut processes = owned_pids_for(&launch, launch.pid, source_pid, source_is_new);
    update_owned_pids_if_current(&launch, &processes);
    if let Some(vol) = volume {
        audio_session::apply_custom_preview_volume(processes.clone(), vol);
    }
    let mut last_volume_apply = Instant::now();

    // If discovery pre-attached `custom_preview_hook.dll` directly onto the 64-bit grandchild,
    // promote `capture_technique` back to `InProcessHook` BEFORE checking `tries_invisible_layered`.
    if launch.requested_capture == CustomCaptureTechnique::InProcessHook
        && capture_technique != CustomCaptureTechnique::InProcessHook
        && hook_session
            .as_ref()
            .is_some_and(|session| session.pid == source_pid && !session.hooks.is_empty())
    {
        capture_technique = CustomCaptureTechnique::InProcessHook;
        if let Ok(mut slot) = ACTIVE_CAPTURE_TECHNIQUE.lock() {
            *slot = capture_technique;
        }
    }

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

    if trace::enabled() {
        let tree_pids = process_tree(launch.pid);
        let job_pids = active_job_pids(launch.id);
        let discover_via = if source_pid == launch.pid {
            "launch_pid"
        } else if tree_pids.contains(&source_pid) {
            "process_tree"
        } else if job_pids.contains(&source_pid) {
            "job_pid"
        } else if source_is_new
            && launch.install_dir.as_deref().is_some_and(|dir| {
                unsafe { process_image_path(source_pid) }
                    .is_some_and(|p| exe_is_inside_install_dir(&p, dir))
            })
        {
            "install_dir"
        } else {
            "image_name"
        };
        if matches!(discover_via, "job_pid" | "install_dir") {
            let exe_name = unsafe { process_image_path(source_pid) }
                .as_deref()
                .and_then(image_name_from_path)
                .unwrap_or_else(|| "<unknown>".to_owned());
            cp_trace!(
                "discover_fallback: id={} via={} hwnd=0x{:x} pid={} launch_pid={} exe={:?}",
                launch.id,
                discover_via,
                source.0 as usize,
                source_pid,
                launch.pid,
                exe_name
            );
        }
        let extra = format!(
            " id={} via={} launch_pid={} launch_tid={} was_hidden={} inv_layered={} own={:?} scale={:?}",
            launch.id, discover_via, launch.pid, launch.tid, was_hidden, invisible_layered, own, scale
        );
        trace_window_diag(
            "discovered",
            source,
            source_pid,
            source_tid,
            source_is_new,
            hook_session.as_ref().map(|s| s.view.hook_hits()),
            hook_session
                .as_ref()
                .and_then(|s| s.view.captured_hwnd().map(|h| h.0 as isize)),
            &extra,
        );
    }

    let current_target_rect = || -> Option<(i32, i32, i32, i32)> {
        let at = WANT
            .lock()
            .ok()
            .and_then(|want| want.as_ref().map(|want| want.at))?;
        provisional_rect_for(at, own, scale, follow_cursor(), bypass_placement())
    };

    // Ensure hook covers secondary UI threads and child/brokered PIDs (including 64-bit
    // grandchildren spawned by 32-bit stubs) BEFORE modifying window styles, and register the
    // active SHM handle for detach/hide.
    let wants_hook_session = uses_hook_enforcement(launch.placement)
        || launch.requested_capture == CustomCaptureTechnique::InProcessHook;
    if wants_hook_session {
        let target_hint = current_target_rect().or(launch.provisional);
        let desired_capture = if launch.requested_capture == CustomCaptureTechnique::InProcessHook
            && source_pid != launch.pid
            && pid_is_wow64(source_pid) != Some(true)
        {
            CustomCaptureTechnique::InProcessHook
        } else {
            capture_technique
        };
        let mut init = hook_init_for(desired_capture, launch.placement, target_hint);
        init.captured_hwnd = source.0 as isize;
        let same_pid_hooked = hook_session
            .as_ref()
            .is_some_and(|session| session.pid == source_pid);
        if same_pid_hooked {
            let mut hooked_ok = true;
            if let Some(ref mut session) = hook_session {
                if source_tid != 0 && !session.hooks.iter().any(|&(t, _, _)| t == source_tid) {
                    let attached = session.attach_thread(source_tid);
                    cp_trace!(
                        "run_it: attach_secondary_thread id={} pid={} tid={} launch_tid={} ok={}",
                        launch.id,
                        source_pid,
                        source_tid,
                        launch.tid,
                        attached
                    );
                    hooked_ok = !session.hooks.is_empty();
                } else {
                    hooked_ok = !session.hooks.is_empty();
                }
            }
            if !hooked_ok {
                if let Some(ref session) = hook_session {
                    session.view.clear_locks_for_detach();
                }
                drop(hook_session.take());
            }
        } else {
            let prev_pid = hook_session.as_ref().map(|s| s.pid);
            if let Some(ref session) = hook_session {
                session.view.clear_locks_for_detach();
            }
            drop(hook_session.take());
            let should_dynamic_attach = source_pid != 0
                && source_tid != 0
                && (source_pid != launch.pid || uses_hook_enforcement(launch.placement))
                && pid_is_wow64(source_pid) != Some(true);
            if should_dynamic_attach {
                let dynamic = attach_dynamic_hook_session(source_pid, source_tid, init);
                cp_trace!(
                    "run_it: attach_dynamic_session id={} prev_pid={:?} src_pid={} src_tid={} req_cap={:?} ok={}",
                    launch.id,
                    prev_pid,
                    source_pid,
                    source_tid,
                    launch.requested_capture,
                    dynamic.is_some()
                );
                if let Some(dynamic) = dynamic {
                    hook_session = Some(dynamic);
                    if capture_technique != desired_capture {
                        capture_technique = desired_capture;
                        if let Ok(mut slot) = ACTIVE_CAPTURE_TECHNIQUE.lock() {
                            *slot = capture_technique;
                        }
                    }
                }
            }
        }
    }
    if hook_session.is_none() && capture_technique == CustomCaptureTechnique::InProcessHook {
        capture_technique = CustomCaptureTechnique::InvisibleLayered;
        if let Ok(mut slot) = ACTIVE_CAPTURE_TECHNIQUE.lock() {
            *slot = capture_technique;
        }
        if !was_hidden && !invisible_layered {
            invisible_layered = make_source_invisible(source.0 as isize);
            SOURCE_INVISIBLE_LAYERED.store(invisible_layered, Ordering::SeqCst);
            restore_foreground(launch.previous_foreground, &processes);
        }
    }
    if let Some(ref session) = hook_session {
        session.view.set_captured_hwnd(source);
        if let Ok(mut ctrl) = ACTIVE_HOOK_CONTROL.lock() {
            *ctrl = Some(Arc::clone(&session.view));
        }
        if uses_hook_enforcement(launch.placement) {
            let init = hook_init_for(
                capture_technique,
                launch.placement,
                current_target_rect().or(launch.provisional),
            );
            session.view.set_forced_mode(init.forced_mode);
            session.view.set_z_order_mode(init.z_order_mode);
            session
                .view
                .set_rect(init.rect.0, init.rect.1, init.rect.2, init.rect.3);
            if capture_technique != CustomCaptureTechnique::InProcessHook {
                let enforced = session.view.enforce_now_sync(source);
                cp_trace!(
                    "run_it: initial_enforce_non_inprocess id={} cap={:?} ok={} init_rect={:?}",
                    launch.id,
                    capture_technique,
                    enforced,
                    init.rect
                );
            }
        }
        cp_trace!(
            "run_it: hook_ready id={} hook_pid={} hooks_len={} hits={} rect=({},{},{},{}) forced={} cloak={} z_mode={}",
            launch.id,
            session.pid,
            session.hooks.len(),
            session.view.hook_hits(),
            session.view.state().x.load(Ordering::Relaxed),
            session.view.state().y.load(Ordering::Relaxed),
            session.view.state().w.load(Ordering::Relaxed),
            session.view.state().h.load(Ordering::Relaxed),
            session.view.state().forced_placement.load(Ordering::Relaxed),
            session.view.is_cloaked(),
            session.view.state().z_order_mode.load(Ordering::Relaxed)
        );
    } else if uses_hook_enforcement(launch.placement) {
        cp_trace!(
            "run_it: no_hook_session id={} src_pid={} src_tid={} cap={:?}",
            launch.id,
            source_pid,
            source_tid,
            capture_technique
        );
    }

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

    let mut tree_read = Instant::now();
    let mut images: ProcessImages = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();

    let mut box_of_the_host: Option<(i32, i32, i32, i32)> = None;
    let mut registered: isize = 0;

    if capture_technique == CustomCaptureTechnique::InProcessHook {
        if let Some(ref session) = hook_session {
            let mut staged = OffscreenStagedState {
                invisible_layered: &mut invisible_layered,
                pre_cloaked: &mut pre_cloaked,
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
        }
    } else if capture_technique == CustomCaptureTechnique::OffscreenStaged {
        let mut staged = OffscreenStagedState {
            invisible_layered: &mut invisible_layered,
            pre_cloaked: &mut pre_cloaked,
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
            hook_session.as_ref().map(|session| &*session.view),
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

        let hook_view = hook_session.as_ref().map(|session| &*session.view);
        if was_hidden {
            // App honored `SW_HIDE`: atomically position + reveal via `SetWindowPlacement` so it never
            // flashes at its old saved coordinates.
            if let Some(box_) = current_target_rect() {
                let chrome = chrome_now(source, window_mode);
                let target_rect = the_host_for(box_, chrome);
                let ctx = placement_ctx_from_want(&launch, &processes, None, hook_view);
                let mode = match capture_technique {
                    CustomCaptureTechnique::OffScreen => {
                        let vscreen = virtual_screen();
                        let (px, py) = parking_origin(vscreen);
                        if let Some(view) = ctx.hook {
                            view.set_z_order_mode(ZORDER_BOTTOM);
                            view.set_rect(px, py, target_rect.2.max(1), target_rect.3.max(1));
                        }
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
                        let ask = the_ask_for_window(source, (box_.2, box_.3), window_mode);
                        the_ask = Some(ask);
                        park_offscreen(source, ask.0, ask.1, ctx);
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
                        if let Some(view) = ctx.hook {
                            view.set_z_order_mode(ZORDER_TOPMOST);
                            view.set_rect(wx, wy, ww, wh);
                        }
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
                        place_direct_window(source, target_rect, window_mode, ctx);
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
                        if let Some(view) = ctx.hook {
                            view.set_z_order_mode(ZORDER_BOTTOM);
                            view.set_rect(wx, wy, ww, wh);
                        }
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
                            place_invisible_layered(
                                source,
                                target_rect,
                                chrome,
                                window_mode,
                                &mut the_ask,
                                ctx,
                            );
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
                            if ctx.forced {
                                apply_hook_rect_now(
                                    ctx.hook,
                                    source,
                                    (wx2, wy2, ww2, wh2),
                                    ZORDER_BOTTOM,
                                );
                                force_place_window(
                                    source,
                                    (wx2, wy2, ww2, wh2),
                                    HWND_BOTTOM,
                                    SET_WINDOW_POS_FLAGS(0),
                                    PlacementContext {
                                        avoid_cursor: None,
                                        ..ctx
                                    },
                                );
                            } else {
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
                            }
                            let _ = unsafe { InvalidateRect(source, None, false) };
                            let _ = unsafe { UpdateWindow(source) };

                            if aligned_behind_host(source, target_rect) {
                                PlacementMode::BehindHost
                            } else {
                                let ask = the_ask_for_window(source, (box_.2, box_.3), window_mode);
                                the_ask = Some(ask);
                                park_offscreen(source, ask.0, ask.1, ctx);
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
        } else if let Some(box_) = current_target_rect() {
            let chrome = chrome_now(source, window_mode);
            let target_rect = the_host_for(box_, chrome);
            let ctx = placement_ctx_from_want(&launch, &processes, None, hook_view);
            // App ignored `SW_HIDE`: apply the selected capture technique immediately.
            placed_mode = Some(if invisible_layered {
                place_invisible_layered(
                    source,
                    target_rect,
                    chrome,
                    window_mode,
                    &mut the_ask,
                    ctx,
                );
                restore_foreground(launch.previous_foreground, &processes);
                PlacementMode::InvisibleLayered
            } else {
                snap_source(
                    source,
                    capture_technique,
                    target_rect,
                    chrome,
                    window_mode,
                    ctx,
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
            if !is_current(launch.id) {
                return;
            }
            if !unsafe { IsWindow(source) }.as_bool() {
                if is_current(launch.id) {
                    end_the_program();
                }
                return;
            }
            let Some(box_) = current_target_rect() else {
                if is_current(launch.id) {
                    end_the_program();
                }
                return;
            };
            // Everything that places `source` wants the frame its picture will be composed from — the
            // box grown by the program's own caption and border (see `the_host_for`) — while the ask
            // is a question about the box itself. One reading of the frame answers both.
            let chrome = chrome_now(source, window_mode);
            let target_rect = the_host_for(box_, chrome);

            if invisible_layered {
                keep_source_invisible(source.0 as isize);
            } else if capture_technique == CustomCaptureTechnique::DirectWindow {
                apply_direct_window_style(source);
            }

            // Ignore self-triggered LOCATIONCHANGE from our own SetWindowPos.
            let raw_source_moved = drain_settle_events(source, &launch, &processes);
            let source_moved = raw_source_moved && window_rect(source) != expected_rect;
            let ctx = placement_ctx_from_want(&launch, &processes, None, hook_view);

            let needs_resnap = source_moved
                || placed_host_rect != Some(target_rect)
                || (ctx.forced && unsafe { IsZoomed(source) }.as_bool())
                || match placed_mode {
                    Some(PlacementMode::InvisibleLayered | PlacementMode::DirectWindow) => false,
                    Some(PlacementMode::BehindHost) => !aligned_behind_host(source, target_rect),
                    Some(PlacementMode::OffScreen) => should_repark_offscreen(
                        source,
                        expected_rect,
                        ctx.forced,
                        ctx.avoid_cursor,
                        false,
                    ),
                    None => true,
                };
            if needs_resnap {
                if placed_host_rect.is_some_and(|p| p.2 != target_rect.2 || p.3 != target_rect.3) {
                    the_ask = None;
                }
                placed_mode = Some(if invisible_layered {
                    place_invisible_layered(
                        source,
                        target_rect,
                        chrome,
                        window_mode,
                        &mut the_ask,
                        ctx,
                    );
                    restore_foreground(launch.previous_foreground, &processes);
                    PlacementMode::InvisibleLayered
                } else {
                    snap_source(
                        source,
                        capture_technique,
                        target_rect,
                        chrome,
                        window_mode,
                        ctx,
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
                if !is_current(launch.id) {
                    return;
                }
                if !unsafe { IsWindow(source) }.as_bool() {
                    if is_current(launch.id) {
                        end_the_program();
                    }
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

    let mut revealed_at: Option<Instant> = None;
    let mut off_cursor_since: Option<Instant> = None;

    let forced_poll_duration =
        |forced: bool, launched_at: Instant, revealed: Option<Instant>| -> Duration {
            if !forced {
                return KEEP_POLL;
            }
            if launched_at.elapsed() < FORCED_FAST_POLL_WINDOW
                || revealed.is_some_and(|t| t.elapsed() < FORCED_FAST_POLL_WINDOW)
            {
                FORCED_POLL
            } else {
                FORCED_STEADY_POLL
            }
        };

    let update_forced_settling =
        |forced: bool,
         preview_shown: bool,
         cursor_pt: Option<POINT>,
         off_cursor_since: &mut Option<Instant>,
         revealed_at: &mut Option<Instant>| {
            if preview_shown && revealed_at.is_none() {
                *revealed_at = Some(Instant::now());
            }
            if !forced || !FORCED_SETTLING.load(Ordering::Acquire) {
                return;
            }
            if launch.launched_at.elapsed() >= FORCED_SETTLE_CAP {
                FORCED_SETTLING.store(false, Ordering::Release);
                cp_trace!(
                    "forced_settling: cleared id={} reason=cap elapsed_ms={}",
                    launch.id,
                    launch.launched_at.elapsed().as_millis()
                );
                return;
            }
            let source_off_cursor = visible_frame_rect(source)
                .or_else(|| window_rect(source))
                .is_some_and(|r| cursor_pt.is_none_or(|pt| !point_in_rect(pt, r)));
            if preview_shown && source_off_cursor {
                let started = *off_cursor_since.get_or_insert_with(Instant::now);
                if started.elapsed() >= FORCED_SETTLE_STABLE {
                    FORCED_SETTLING.store(false, Ordering::Release);
                    cp_trace!(
                        "forced_settling: cleared id={} reason=stable stable_ms={} total_ms={}",
                        launch.id,
                        started.elapsed().as_millis(),
                        launch.launched_at.elapsed().as_millis()
                    );
                }
            } else {
                *off_cursor_since = None;
            }
        };

    let mut drift_count: u32 = 0;
    let mut last_drift_sig: Option<DriftSig> = None;
    let mut was_layout_none = false;

    loop {
        if !is_current(launch.id) {
            cp_trace!("run_it: exit reason=superseded id={}", launch.id);
            return;
        }
        if !unsafe { IsWindow(source) }.as_bool() {
            cp_trace!(
                "run_it: exit reason=window_destroyed id={} hwnd=0x{:x}",
                launch.id,
                source.0 as usize
            );
            // A program that closed its own window has nothing left to draw, and nothing left to
            // keep: it is ended here rather than left running behind a hover that shows nothing.
            if is_current(launch.id) {
                end_the_program();
            }
            return;
        }

        let Some((at, placement)) = current_want_snapshot() else {
            cp_trace!(
                "run_it: exit reason=want_cleared id={} hwnd=0x{:x}",
                launch.id,
                source.0 as usize
            );
            // Re-check after reading WANT: open_fully/end_the_program bump LAUNCH before clearing WANT.
            if is_current(launch.id) {
                end_the_program();
            }
            return;
        };
        let forced = enforces_placement(placement);
        let hook_mode = uses_hook_enforcement(placement);
        let active_hook_view = hook_mode
            .then(|| hook_session.as_ref().map(|session| &*session.view))
            .flatten();
        let cursor_pt = where_cursor(at);

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
            processes = owned_pids_for(&launch, launch.pid, source_pid, source_is_new);
            update_owned_pids_if_current(&launch, &processes);
            tree_read = Instant::now();
        }
        if let Some(vol) = volume {
            if last_volume_apply.elapsed() >= Duration::from_millis(500) {
                audio_session::apply_custom_preview_volume(processes.clone(), vol);
                last_volume_apply = Instant::now();
            }
        }
        let ctx = PlacementContext {
            forced,
            hook: active_hook_view,
            avoid_cursor: cursor_pt,
            previous_foreground: launch.previous_foreground,
            processes: &processes,
        };

        let Some(layout) = layout_for(at, own, scale, follow_cursor(), bypass_placement()) else {
            if !is_current(launch.id) {
                return;
            }
            if !was_layout_none {
                was_layout_none = true;
                cp_trace!(
                    "run_it: layout_none id={} own={:?} scale={:?}",
                    launch.id,
                    own,
                    scale
                );
            }
            placed_host_rect = None;
            if capture_technique == CustomCaptureTechnique::InProcessHook {
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
                } else {
                    placed_mode = None;
                    if !is_window_offscreen(source) {
                        push_offscreen(source);
                    }
                }
            } else if invisible_layered {
                keep_source_invisible(source.0 as isize);
            } else {
                placed_mode = None;
                if let Some(view) = active_hook_view {
                    view.stage_offscreen();
                    let _ = view.enforce_now_sync(source);
                }
                if !is_window_offscreen(source) {
                    push_offscreen(source);
                }
            }
            hide_the_host();
            std::thread::sleep(forced_poll_duration(
                forced,
                launch.launched_at,
                revealed_at,
            ));
            continue;
        };
        was_layout_none = false;
        let (x, y) = (layout.pos_x, layout.pos_y);
        let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);
        // **The hover is the box plus the chrome the program draws around its picture.** The box is the
        // room cut to the file's shape, and the picture — the program's client — is what goes on it; a
        // player composed from its whole window instead gets the box *minus* its caption and border,
        // which is a client wider than the film and the black bars down the sides that follow from it
        // (see `chrome_of`). Read every poll, because a program that restyles itself restyles this.
        let chrome = chrome_now(source, window_mode);
        let host_rect = the_host_for((x, y, width, height), chrome);
        if !(forced
            && placed_mode == Some(PlacementMode::DirectWindow)
            && placed_host_rect == Some(host_rect))
        {
            if let Ok(mut slot) = LAST_HOST_RECT.lock() {
                *slot = Some(host_rect);
            }
        }

        if let Some(ref session) = hook_session {
            if !is_current(launch.id) || DETACHING.load(Ordering::Acquire) {
                return;
            }
            session.view.set_forced_mode(sh_forced_mode(placement));
        }

        if capture_technique == CustomCaptureTechnique::InProcessHook {
            let Some(ref session) = hook_session else {
                continue;
            };
            let needs_reposition = session.view.is_cloaked()
                || placed_mode.is_none()
                || placed_host_rect != Some(host_rect);
            if needs_reposition {
                cp_trace!(
                    "run_it: inprocess_reposition id={} cloaked={} mode={:?} prev_host={:?} host={:?}",
                    launch.id,
                    session.view.is_cloaked(),
                    placed_mode,
                    placed_host_rect,
                    host_rect
                );
                if placed_host_rect.is_some_and(|p| p.2 != host_rect.2 || p.3 != host_rect.3) {
                    the_ask = None;
                }
                let mut staged = OffscreenStagedState {
                    invisible_layered: &mut invisible_layered,
                    pre_cloaked: &mut pre_cloaked,
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
            } else if forced {
                match placed_mode {
                    Some(PlacementMode::BehindHost) => {
                        if !aligned_behind_host(source, host_rect)
                            || unsafe { IsZoomed(source) }.as_bool()
                            || window_rect(source) != expected_rect
                        {
                            trace_drift(
                                &mut drift_count,
                                &mut last_drift_sig,
                                launch.id,
                                "inprocess_behind_host_drift",
                                source,
                                expected_rect,
                                host_rect,
                                Some(session.view.hook_hits()),
                            );
                            let insets = if window_mode == CustomWindowMode::Normal {
                                dwm_frame_insets(source)
                            } else {
                                (0, 0, 0, 0)
                            };
                            let (wx, wy, ww, wh) = behind_host_window_rect(host_rect, insets);
                            let no_shift_ctx = PlacementContext {
                                avoid_cursor: None,
                                ..ctx
                            };
                            session.view.set_rect(wx, wy, ww, wh);
                            apply_hook_rect_now(
                                active_hook_view,
                                source,
                                (wx, wy, ww, wh),
                                ZORDER_BOTTOM,
                            );
                            force_place_window(
                                source,
                                (wx, wy, ww, wh),
                                HWND_BOTTOM,
                                SET_WINDOW_POS_FLAGS(0),
                                no_shift_ctx,
                            );
                            if aligned_behind_host(source, host_rect) {
                                expected_rect = window_rect(source);
                            } else {
                                let ask = the_ask_for_window(source, (width, height), window_mode);
                                cp_trace!(
                                    "run_it: inprocess_behind_host_fallback_offscreen id={} host={:?} ask={:?}",
                                    launch.id,
                                    host_rect,
                                    ask
                                );
                                the_ask = Some(ask);
                                session
                                    .view
                                    .set_rect(STAGE_COORD, STAGE_COORD, ask.0, ask.1);
                                apply_hook_rect_now(
                                    active_hook_view,
                                    source,
                                    (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                                    ZORDER_BOTTOM,
                                );
                                force_place_window(
                                    source,
                                    (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                                    HWND_BOTTOM,
                                    SET_WINDOW_POS_FLAGS(0),
                                    no_shift_ctx,
                                );
                                nudge_clamped_offscreen_off_cursor(
                                    source,
                                    Some(&session.view),
                                    cursor_pt,
                                    launch.previous_foreground,
                                    &processes,
                                );
                                placed_mode = Some(PlacementMode::OffScreen);
                                expected_rect = window_rect(source);
                            }
                        }
                    }
                    Some(PlacementMode::OffScreen) => {
                        if should_repark_offscreen(
                            source,
                            expected_rect,
                            forced,
                            cursor_pt,
                            session.view.is_cloaked(),
                        ) {
                            trace_drift(
                                &mut drift_count,
                                &mut last_drift_sig,
                                launch.id,
                                "inprocess_offscreen_repark",
                                source,
                                expected_rect,
                                host_rect,
                                Some(session.view.hook_hits()),
                            );
                            let ask = the_ask.unwrap_or_else(|| {
                                the_ask_for_window(source, (width, height), window_mode)
                            });
                            the_ask = Some(ask);
                            session
                                .view
                                .set_rect(STAGE_COORD, STAGE_COORD, ask.0, ask.1);
                            apply_hook_rect_now(
                                active_hook_view,
                                source,
                                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                                ZORDER_BOTTOM,
                            );
                            force_place_window(
                                source,
                                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                                HWND_BOTTOM,
                                SET_WINDOW_POS_FLAGS(0),
                                PlacementContext {
                                    avoid_cursor: None,
                                    ..ctx
                                },
                            );
                            nudge_clamped_offscreen_off_cursor(
                                source,
                                Some(&session.view),
                                cursor_pt,
                                launch.previous_foreground,
                                &processes,
                            );
                            expected_rect = window_rect(source);
                        } else if forced
                            && !is_window_offscreen(source)
                            && FORCED_SETTLING.load(Ordering::Acquire)
                        {
                            trace_drift(
                                &mut drift_count,
                                &mut last_drift_sig,
                                launch.id,
                                "inprocess_offscreen_nudge",
                                source,
                                expected_rect,
                                host_rect,
                                Some(session.view.hook_hits()),
                            );
                            nudge_clamped_offscreen_off_cursor(
                                source,
                                Some(&session.view),
                                cursor_pt,
                                launch.previous_foreground,
                                &processes,
                            );
                            expected_rect = window_rect(source);
                        }
                    }
                    Some(PlacementMode::InvisibleLayered) => {
                        keep_source_invisible(source.0 as isize);
                        if window_rect(source) != expected_rect
                            || unsafe { IsZoomed(source) }.as_bool()
                        {
                            trace_drift(
                                &mut drift_count,
                                &mut last_drift_sig,
                                launch.id,
                                "inprocess_inv_layered_drift",
                                source,
                                expected_rect,
                                host_rect,
                                Some(session.view.hook_hits()),
                            );
                            place_invisible_layered(
                                source,
                                host_rect,
                                chrome,
                                window_mode,
                                &mut the_ask,
                                ctx,
                            );
                            expected_rect = window_rect(source);
                        }
                    }
                    Some(PlacementMode::DirectWindow) | None => {}
                }
            }
            let preview_shown = SHOWN.load(Ordering::Acquire);
            update_forced_settling(
                forced,
                preview_shown,
                cursor_pt,
                &mut off_cursor_since,
                &mut revealed_at,
            );
            std::thread::sleep(forced_poll_duration(
                forced,
                launch.launched_at,
                revealed_at,
            ));
            continue;
        } else if placed_mode.is_none() || placed_host_rect != Some(host_rect) {
            if placed_host_rect.is_some_and(|p| p.2 != host_rect.2 || p.3 != host_rect.3) {
                the_ask = None;
            }
            placed_mode = Some(if invisible_layered {
                place_invisible_layered(source, host_rect, chrome, window_mode, &mut the_ask, ctx);
                restore_foreground(launch.previous_foreground, &processes);
                PlacementMode::InvisibleLayered
            } else {
                snap_source(
                    source,
                    capture_technique,
                    host_rect,
                    chrome,
                    window_mode,
                    ctx,
                    &mut the_ask,
                    placed_mode.zip(placed_host_rect),
                )
            });
            placed_host_rect = Some(host_rect);
            expected_rect = window_rect(source);
            cp_trace!(
                "run_it: std_placed id={} cap={:?} mode={:?} host={:?} expected={:?}",
                launch.id,
                capture_technique,
                placed_mode,
                host_rect,
                expected_rect
            );
        } else {
            match placed_mode {
                Some(PlacementMode::InvisibleLayered) => {
                    keep_source_invisible(source.0 as isize);
                    if window_rect(source) != expected_rect
                        || (forced && unsafe { IsZoomed(source) }.as_bool())
                    {
                        trace_drift(
                            &mut drift_count,
                            &mut last_drift_sig,
                            launch.id,
                            "std_inv_layered_drift",
                            source,
                            expected_rect,
                            host_rect,
                            active_hook_view.map(|v| v.hook_hits()),
                        );
                        place_invisible_layered(
                            source,
                            host_rect,
                            chrome,
                            window_mode,
                            &mut the_ask,
                            ctx,
                        );
                        restore_foreground(launch.previous_foreground, &processes);
                        expected_rect = window_rect(source);
                    }
                }
                Some(PlacementMode::DirectWindow) => {
                    apply_direct_window_style(source);
                    let cur_rect = window_rect(source);
                    let overlaps_cursor = forced
                        && cursor_pt.is_some_and(|c| cur_rect.is_some_and(|r| point_in_rect(c, r)));
                    if cur_rect != expected_rect
                        || (forced && unsafe { IsZoomed(source) }.as_bool())
                        || overlaps_cursor
                    {
                        trace_drift(
                            &mut drift_count,
                            &mut last_drift_sig,
                            launch.id,
                            "std_direct_drift",
                            source,
                            expected_rect,
                            host_rect,
                            active_hook_view.map(|v| v.hook_hits()),
                        );
                        place_direct_window(source, host_rect, window_mode, ctx);
                        restore_foreground(launch.previous_foreground, &processes);
                        expected_rect = window_rect(source);
                    }
                }
                Some(PlacementMode::BehindHost)
                    if !aligned_behind_host(source, host_rect)
                        || (forced && unsafe { IsZoomed(source) }.as_bool()) =>
                {
                    trace_drift(
                        &mut drift_count,
                        &mut last_drift_sig,
                        launch.id,
                        "std_behind_host_drift",
                        source,
                        expected_rect,
                        host_rect,
                        active_hook_view.map(|v| v.hook_hits()),
                    );
                    if !place_behind_host(source, host_rect, window_mode, ctx) {
                        let ask = the_ask_for_window(source, (width, height), window_mode);
                        the_ask = Some(ask);
                        park_offscreen(source, ask.0, ask.1, ctx);
                        placed_mode = Some(PlacementMode::OffScreen);
                    }
                    restore_foreground(launch.previous_foreground, &processes);
                    expected_rect = window_rect(source);
                }
                Some(PlacementMode::OffScreen) => {
                    if should_repark_offscreen(source, expected_rect, forced, cursor_pt, false) {
                        trace_drift(
                            &mut drift_count,
                            &mut last_drift_sig,
                            launch.id,
                            "std_offscreen_repark",
                            source,
                            expected_rect,
                            host_rect,
                            active_hook_view.map(|v| v.hook_hits()),
                        );
                        let ask = the_ask.unwrap_or((width, height));
                        park_offscreen(source, ask.0, ask.1, ctx);
                        restore_foreground(launch.previous_foreground, &processes);
                        expected_rect = window_rect(source);
                    } else if forced
                        && !is_window_offscreen(source)
                        && FORCED_SETTLING.load(Ordering::Acquire)
                    {
                        trace_drift(
                            &mut drift_count,
                            &mut last_drift_sig,
                            launch.id,
                            "std_offscreen_nudge",
                            source,
                            expected_rect,
                            host_rect,
                            active_hook_view.map(|v| v.hook_hits()),
                        );
                        nudge_clamped_offscreen_off_cursor(
                            source,
                            active_hook_view,
                            cursor_pt,
                            launch.previous_foreground,
                            &processes,
                        );
                        expected_rect = window_rect(source);
                    }
                }
                _ => {}
            }
        }

        if !uses_dwm_host(capture_technique) {
            // DirectWindow shows `source` directly as HWND_TOPMOST; never register thumbnail or show host.
            let direct_shown = placed_mode == Some(PlacementMode::DirectWindow);
            if direct_shown && pre_cloaked {
                let _ = set_window_cloak(source, false);
                pre_cloaked = false;
            }
            update_forced_settling(
                forced,
                direct_shown,
                cursor_pt,
                &mut off_cursor_since,
                &mut revealed_at,
            );
            std::thread::sleep(forced_poll_duration(
                forced,
                launch.launched_at,
                revealed_at,
            ));
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
            update_forced_settling(
                forced,
                false,
                cursor_pt,
                &mut off_cursor_since,
                &mut revealed_at,
            );
            std::thread::sleep(forced_poll_duration(
                forced,
                launch.launched_at,
                revealed_at,
            ));
            continue;
        }

        // Only after `source` is settled and aligned: size `host`, bind thumbnail, and show `host`.
        //
        // **The host is the whole of `host_rect`, not the box.** The picture a thumbnail composes is
        // the program's visible frame — the client *and* the caption above it — so a host the size of
        // the box alone is a host whose client is smaller than the picture, and the picture is then
        // fitted into it with the background showing either side: bars drawn by the host rather than
        // by the player. Placed at the grown rect the client is the picture exactly, and
        // `the_picture_in` answers with the whole of it.
        if box_of_the_host != Some(host_rect) {
            let _ = unsafe {
                SetWindowPos(
                    host,
                    HWND::default(),
                    host_rect.0,
                    host_rect.1,
                    host_rect.2,
                    host_rect.3,
                    SWP_NOZORDER | SWP_NOACTIVATE,
                )
            };
            box_of_the_host = Some(host_rect);
        }

        if registered == 0 {
            match unsafe { DwmRegisterThumbnail(host, source) } {
                Ok(registration) => {
                    registered = registration;
                    if let Ok(mut held) = THUMBNAIL.lock() {
                        *held = registration;
                    }
                    cp_trace!(
                        "run_it: std_dwm_register id={} reg=0x{:x} host=0x{:x} src=0x{:x}",
                        launch.id,
                        registration as usize,
                        host.0 as usize,
                        source.0 as usize
                    );
                }
                Err(e) => {
                    cp_trace!("run_it: std_dwm_register_err id={} err={:?}", launch.id, e);
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
                if pre_cloaked {
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
                    let _ = unsafe { DwmFlush() };
                    let _ = set_window_cloak(source, false);
                    pre_cloaked = false;
                } else {
                    let _ = unsafe {
                        PostMessageW(host, WM_THE_PICTURE_IS_READY, WPARAM(0), LPARAM(0))
                    };
                }
                cp_trace!(
                    "run_it: std_shown id={} host={:?} picture={:?} placed={:?}",
                    launch.id,
                    host_rect,
                    picture,
                    placed_mode
                );
            }
        }

        update_forced_settling(
            forced,
            SHOWN.load(Ordering::Acquire),
            cursor_pt,
            &mut off_cursor_since,
            &mut revealed_at,
        );
        std::thread::sleep(forced_poll_duration(
            forced,
            launch.launched_at,
            revealed_at,
        ));
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

/// The chrome a program draws around its picture **now**, read off the window itself: `(left, top,
/// extra_w, extra_h)` as [`chrome_of`] defines it, and zeros for a window whose frame is stripped.
///
/// A window that will not say where its client stands — destroyed, or not yet made — is answered with
/// no chrome, which is the box as it stands: the hover is then exactly as it was before this was read.
fn chrome_now(window: HWND, window_mode: CustomWindowMode) -> (i32, i32, i32, i32) {
    if window_mode != CustomWindowMode::Normal {
        return (0, 0, 0, 0);
    }
    let Some(visible) = visible_frame_rect(window) else {
        return (0, 0, 0, 0);
    };
    let mut client = RECT::default();
    if unsafe { GetClientRect(window, &mut client) }.is_err() {
        return (0, 0, 0, 0);
    }
    let mut origin = POINT::default();
    if !unsafe { ClientToScreen(window, &mut origin) }.as_bool() {
        return (0, 0, 0, 0);
    }
    chrome_of(
        visible,
        origin,
        (client.right - client.left, client.bottom - client.top),
    )
}

/// The size the program's window is asked for: the host rect — the box grown by the frame the program
/// draws around its picture — raised to the smallest client the window will take at the box's shape.
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
/// **The ask is the frame the picture is composed from, and `taken_client` is a client.** A caller
/// places the window so that this rect is its *visible frame* (see `behind_host_window_rect`),
/// because the visible frame is what a DWM thumbnail composes and what the host is the size of. So
/// the two halves are in one space each: the box and what the window took are clients, and the frame
/// is added on the outside. Measured in different spaces — a window rect taken against a frame read
/// off the visible one — the ask grew by its own invisible border on every poll, which left the
/// program's client wider than the box and the player letterboxing the film inside it.
fn the_ask_for(
    box_: (i32, i32),
    taken_client: (u32, u32),
    chrome: (i32, i32, i32, i32),
) -> (i32, i32) {
    // **The ask is a frame and the box is a client**, so the frame is added to what the box is asked
    // for (see `chrome_of`). A `Borderless` program has none, and the two halves of this are the
    // identity.
    let extra = (chrome.2.max(0) as u32, chrome.3.max(0) as u32);
    let client = the_client_ask_for(box_, taken_client);
    (client.0 + extra.0 as i32, client.1 + extra.1 as i32)
}

/// [`the_ask_for`] in client space: the box's own shape, raised to the smallest client the window will
/// take at that shape.
fn the_client_ask_for(box_: (i32, i32), taken: (u32, u32)) -> (i32, i32) {
    let box_ = (box_.0.max(1), box_.1.max(1));
    let shape = box_.0 as f32 / box_.1 as f32;

    let width = (taken.0 as i32)
        .max((taken.1 as f32 * shape).ceil() as i32)
        .max(box_.0);
    let height = ((width as f32 / shape).ceil() as i32).max(1);

    (width, height)
}

/// [`the_ask_for`] for a window that is standing, given the box it is to be composed on: the client
/// and the frame are read off the window, which is what every caller inside a placement has.
fn the_ask_for_window(window: HWND, box_: (i32, i32), window_mode: CustomWindowMode) -> (i32, i32) {
    the_ask_for(box_, client_size(window), chrome_now(window, window_mode))
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

/// The size of the client `window` draws its document in: the box a preview puts it on.
fn client_size(window: HWND) -> (u32, u32) {
    let mut rect = RECT::default();
    if unsafe { GetClientRect(window, &mut rect) }.is_err() {
        return (0, 0);
    }
    (
        (rect.right - rect.left).max(0) as u32,
        (rect.bottom - rect.top).max(0) as u32,
    )
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

/// The chrome a program draws around its picture, as `(left, top, extra_w, extra_h)`: where its client
/// sits inside the frame the user sees, and how much of that frame the chrome takes altogether.
///
/// **This is the whole of the black-bar fault, and it is a fact about the window rather than about the
/// file.** A program's *client* is where it draws the document; the caption, the resize border and the
/// invisible Win11 shadow margins are the frame around it. A preview composited from the whole window
/// therefore gives the document only the box *minus* the frame — and a player handed a box 4 % wider
/// than the film's own shape fills the height and draws the remainder as black bars either side.
/// Measured on MPV Player: a 1200x675 box composed as a whole window is a 1184x636 client, a shape of
/// 1.861 against the film's 1.778, and 27 px of black sit at each side of a 16:9 film. The same window
/// with a *client* of 1200x675 draws no bars at all.
///
/// So the host is grown by this and the client is put on the box (see [`the_host_for`]): the caption
/// keeps its strip at the top of the hover and the picture fills everything under it.
fn chrome_of(visible: RECT, client_origin: POINT, client: (i32, i32)) -> (i32, i32, i32, i32) {
    let left = (client_origin.x - visible.left).max(0);
    let top = (client_origin.y - visible.top).max(0);
    let extra_w = ((visible.right - visible.left) - client.0.max(0)).max(0);
    let extra_h = ((visible.bottom - visible.top) - client.1.max(0)).max(0);
    (left, top, extra_w, extra_h)
}

/// The host a box is drawn in: the box's own rectangle, grown by the chrome the program draws around
/// its picture — so the picture lands exactly on the box and the caption lands beside it (see
/// [`chrome_of`] for why the box alone is the box a player letterboxes in).
///
/// `Borderless` is a program with its frame stripped, and `(0, 0, 0, 0)` leaves it exactly the box.
fn the_host_for(box_: (i32, i32, i32, i32), chrome: (i32, i32, i32, i32)) -> (i32, i32, i32, i32) {
    let (left, top, extra_w, extra_h) = chrome;
    (
        box_.0 - left,
        box_.1 - top,
        (box_.2 + extra_w).max(1),
        (box_.3 + extra_h).max(1),
    )
}

/// [`the_host_for`] backwards: the box a host rect was grown from.
fn the_box_of(
    host_rect: (i32, i32, i32, i32),
    chrome: (i32, i32, i32, i32),
) -> (i32, i32, i32, i32) {
    let (left, top, extra_w, extra_h) = chrome;
    (
        host_rect.0 + left,
        host_rect.1 + top,
        (host_rect.2 - extra_w).max(1),
        (host_rect.3 - extra_h).max(1),
    )
}

/// [`the_ask_for`] for a host that is standing: the host rect is the frame the picture is composed
/// from, so the box is read back out of it (see [`the_box_of`]) and the ask is that same frame raised
/// to the smallest client the program will take.
///
/// This is where the two spaces meet, and the black-bar fault lived in the meeting: the host is the
/// box grown by the frame ([`the_host_for`]) and the ask is that same rect again, so a program that
/// takes the ask has its client on the box and its picture on the host — nothing is fitted, and no
/// bar is drawn by anyone. A caller that hands this the *box* rather than the host asks for the box
/// and nothing more, which is what a placement that has not read the frame yet wants.
///
/// `chrome` is handed in rather than read again, so that the host rect and the ask are one reading of
/// the frame: read twice they can answer differently — a window being restyled reports its caption a
/// tick after it has one — and the ask then lands beside the host by the difference.
fn the_ask_for_host(
    window: HWND,
    host_rect: (i32, i32, i32, i32),
    chrome: (i32, i32, i32, i32),
) -> (i32, i32) {
    let (_, _, box_w, box_h) = the_box_of(host_rect, chrome);
    the_ask_for((box_w, box_h), client_size(window), chrome)
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

fn where_cursor(at: Where) -> Option<POINT> {
    match at {
        Where::Pointer { cursor, .. } => Some(POINT {
            x: cursor.0,
            y: cursor.1,
        }),
        Where::Item { .. } => None,
    }
}

fn work_area_rect(pt: POINT) -> RECT {
    let bounds = work_area_at(pt.x, pt.y);
    RECT {
        left: bounds.left,
        top: bounds.top,
        right: bounds.right,
        bottom: bounds.bottom,
    }
}

/// If `actual` overlaps `cursor`, pick the nearest shifted `(x, y)` inside `work` that keeps
/// `actual`'s size clear of `cursor` by `standoff` pixels. Returns `None` if already clear or unshiftable.
fn shift_rect_off_cursor(
    actual: RECT,
    cursor: POINT,
    work: RECT,
    standoff: i32,
) -> Option<(i32, i32)> {
    if !point_in_rect(cursor, actual) {
        return None;
    }
    let aw = (actual.right - actual.left).max(1);
    let ah = (actual.bottom - actual.top).max(1);
    let min_x = work.left;
    let max_x = (work.right - aw).max(work.left);
    let min_y = work.top;
    let max_y = (work.bottom - ah).max(work.top);
    let clamped_x = actual.left.clamp(min_x, max_x);
    let clamped_y = actual.top.clamp(min_y, max_y);

    let candidates = [
        ((cursor.x + standoff).clamp(min_x, max_x), clamped_y),
        ((cursor.x - standoff - aw).clamp(min_x, max_x), clamped_y),
        (clamped_x, (cursor.y + standoff).clamp(min_y, max_y)),
        (clamped_x, (cursor.y - standoff - ah).clamp(min_y, max_y)),
    ];

    candidates
        .into_iter()
        .filter(|&(cx, cy)| {
            !point_in_rect(
                cursor,
                RECT {
                    left: cx,
                    top: cy,
                    right: cx + aw,
                    bottom: cy + ah,
                },
            )
        })
        .min_by_key(|&(cx, cy)| (cx - actual.left).abs() + (cy - actual.top).abs())
}

#[derive(Clone, Copy)]
struct PlacementContext<'a> {
    forced: bool,
    hook: Option<&'a HookSharedView>,
    avoid_cursor: Option<POINT>,
    previous_foreground: isize,
    processes: &'a [u32],
}

fn apply_hook_rect_now(
    hook: Option<&HookSharedView>,
    window: HWND,
    rect: (i32, i32, i32, i32),
    z_order_mode: u8,
) {
    if let Some(view) = hook {
        view.set_z_order_mode(z_order_mode);
        view.set_rect(rect.0, rect.1, rect.2, rect.3);
        let _ = view.enforce_now_sync(window);
    }
}

fn sync_hook_if_shifted(hook: Option<&HookSharedView>, window: HWND, intended_xy: (i32, i32)) {
    if let Some(view) = hook {
        if let Some(actual) = window_rect(window) {
            if actual.left != intended_xy.0 || actual.top != intended_xy.1 {
                view.set_rect(
                    actual.left,
                    actual.top,
                    (actual.right - actual.left).max(1),
                    (actual.bottom - actual.top).max(1),
                );
            }
        }
    }
}

fn force_place_window(
    hwnd: HWND,
    target_rect: (i32, i32, i32, i32),
    z_order: HWND,
    extra_flags: SET_WINDOW_POS_FLAGS,
    ctx: PlacementContext<'_>,
) {
    let tracing = trace::enabled();
    let before_rect = if tracing { window_rect(hwnd) } else { None };
    let before_vis = if tracing {
        visible_frame_rect(hwnd)
    } else {
        None
    };

    let (tx, ty, tw, th) = target_rect;
    let (tw, th) = (tw.max(1), th.max(1));
    let mut placement_res: Option<(bool, u32)> = None;
    if unsafe { IsZoomed(hwnd) }.as_bool() || unsafe { IsIconic(hwnd) }.as_bool() {
        let target = RECT {
            left: tx,
            top: ty,
            right: tx + tw,
            bottom: ty + th,
        };
        let rc_normal = if is_off_screen(&target, virtual_screen()) {
            target
        } else {
            workspace_rect_for_screen(target)
        };
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags: WINDOWPLACEMENT_FLAGS(0),
            showCmd: SW_SHOWNOACTIVATE.0 as u32,
            rcNormalPosition: rc_normal,
            ..Default::default()
        };
        let ok = unsafe { SetWindowPlacement(hwnd, &placement) }.is_ok();
        let err = if ok { 0 } else { win_err() };
        placement_res = Some((ok, err));
    }

    let cur_style = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) } as u32;
    let max_min = WS_MAXIMIZE.0 | WS_MINIMIZE.0;
    if (cur_style & max_min) != 0 {
        let _ = unsafe { SetWindowLongPtrW(hwnd, GWL_STYLE, (cur_style & !max_min) as isize) };
    }

    let base_flags = SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING | extra_flags;
    let pos1_ok = unsafe { SetWindowPos(hwnd, z_order, tx, ty, tw, th, base_flags) }.is_ok();
    let pos1_err = if pos1_ok { 0 } else { win_err() };

    let mut mv_res: Option<(bool, u32, bool, u32)> = None;
    if window_rect(hwnd).is_some_and(|r| r.left != tx || r.top != ty) {
        let size_ignored = (extra_flags.0 & SWP_NOSIZE.0) != 0;
        let cur_size = window_rect(hwnd)
            .map(|r| ((r.right - r.left).max(1), (r.bottom - r.top).max(1)))
            .unwrap_or((tw, th));
        let (mw, mh) = if size_ignored { cur_size } else { (tw, th) };
        let mv_ok = unsafe { MoveWindow(hwnd, tx, ty, mw, mh, true) }.is_ok();
        let mv_err = if mv_ok { 0 } else { win_err() };
        let pos2_ok = unsafe {
            SetWindowPos(
                hwnd,
                z_order,
                tx,
                ty,
                0,
                0,
                SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING | extra_flags,
            )
        }
        .is_ok();
        let pos2_err = if pos2_ok { 0 } else { win_err() };
        mv_res = Some((mv_ok, mv_err, pos2_ok, pos2_err));
    }

    let mut shift_res: Option<(i32, i32, bool, u32)> = None;
    if let Some(cursor) = ctx.avoid_cursor {
        if let Some(actual) = window_rect(hwnd) {
            let work = work_area_rect(cursor);
            if let Some((sx, sy)) =
                shift_rect_off_cursor(actual, cursor, work, FORCED_CURSOR_STANDOFF_PX)
            {
                if let Some(view) = ctx.hook {
                    view.set_rect(
                        sx,
                        sy,
                        (actual.right - actual.left).max(1),
                        (actual.bottom - actual.top).max(1),
                    );
                }
                let shift_ok = unsafe {
                    SetWindowPos(
                        hwnd,
                        z_order,
                        sx,
                        sy,
                        0,
                        0,
                        SWP_NOSIZE
                            | SWP_NOACTIVATE
                            | SWP_NOOWNERZORDER
                            | SWP_NOSENDCHANGING
                            | extra_flags,
                    )
                }
                .is_ok();
                let shift_err = if shift_ok { 0 } else { win_err() };
                shift_res = Some((sx, sy, shift_ok, shift_err));
            }
        }
    }

    if tracing {
        cp_trace!(
            "force_place_window: hwnd=0x{:x} target={:?} z=0x{:x} extra=0x{:x} has_hook={} before={:?} before_vis={:?} after={:?} after_vis={:?} place={:?} pos1=({},{}) mv={:?} shift={:?}",
            hwnd.0 as usize,
            target_rect,
            z_order.0 as isize,
            extra_flags.0,
            ctx.hook.is_some(),
            before_rect,
            before_vis,
            window_rect(hwnd),
            visible_frame_rect(hwnd),
            placement_res,
            pos1_ok,
            pos1_err,
            mv_res,
            shift_res
        );
    }

    restore_foreground(ctx.previous_foreground, ctx.processes);
}

/// Compute `LAST_HOST_RECT` after forced `DirectWindow` placement: keep `host_rect` when
/// unshifted, or adopt the actual placed/shifted window rect when moved off-target/off-cursor.
fn shifted_direct_host_rect(
    intended_outer: (i32, i32, i32, i32),
    actual_window_rect: RECT,
    host_rect: (i32, i32, i32, i32),
) -> (i32, i32, i32, i32) {
    if actual_window_rect.left != intended_outer.0 || actual_window_rect.top != intended_outer.1 {
        (
            actual_window_rect.left,
            actual_window_rect.top,
            (actual_window_rect.right - actual_window_rect.left).max(1),
            (actual_window_rect.bottom - actual_window_rect.top).max(1),
        )
    } else {
        host_rect
    }
}

/// Place `window` directly at `host_rect` as `HWND_TOPMOST` click-through (no DWM host/thumbnail).
fn place_direct_window(
    window: HWND,
    host_rect: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
    ctx: PlacementContext<'_>,
) {
    apply_direct_window_style(window);
    let insets = if window_mode == CustomWindowMode::Normal {
        dwm_frame_insets(window)
    } else {
        (0, 0, 0, 0)
    };
    let (wx, wy, ww, wh) = behind_host_window_rect(host_rect, insets);

    if ctx.forced {
        apply_hook_rect_now(ctx.hook, window, (wx, wy, ww, wh), ZORDER_TOPMOST);
        force_place_window(
            window,
            (wx, wy, ww, wh),
            HWND_TOPMOST,
            SET_WINDOW_POS_FLAGS(0),
            ctx,
        );
        if !unsafe { IsWindowVisible(window) }.as_bool() {
            let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
            apply_hook_rect_now(ctx.hook, window, (wx, wy, ww, wh), ZORDER_TOPMOST);
            force_place_window(
                window,
                (wx, wy, ww, wh),
                HWND_TOPMOST,
                SET_WINDOW_POS_FLAGS(0),
                ctx,
            );
        }
        sync_hook_if_shifted(ctx.hook, window, (wx, wy));
        if let Some(actual) = window_rect(window) {
            let synced = shifted_direct_host_rect((wx, wy, ww, wh), actual, host_rect);
            if let Ok(mut slot) = LAST_HOST_RECT.lock() {
                *slot = Some(synced);
            }
        }
    } else {
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
    chrome: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
    the_ask: &mut Option<(i32, i32)>,
    ctx: PlacementContext<'_>,
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

    let mut final_target = (wx, wy, ww, wh);
    if ctx.forced {
        apply_hook_rect_now(ctx.hook, window, (wx, wy, ww, wh), ZORDER_BOTTOM);
        force_place_window(
            window,
            (wx, wy, ww, wh),
            HWND_BOTTOM,
            SET_WINDOW_POS_FLAGS(0),
            ctx,
        );
    } else {
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
    }

    // If the program's floor refused `(tw, th)`, resize on-monitor to the ask, keeping the box's shape.
    //
    // The ask is taken from `host_rect` and not from `(hw, hh)` as a box: the host rect is already the
    // box grown by the frame, and growing it again by the same frame is what made the ask climb by a
    // caption on every poll (see `the_ask_for_host`).
    let ask = the_ask_for_host(window, host_rect, chrome);
    if ask != (tw, th) {
        *the_ask = (ask != (hw, hh)).then_some(ask);
        let (wx2, wy2, ww2, wh2) = behind_host_window_rect((hx, hy, ask.0, ask.1), insets);
        final_target = (wx2, wy2, ww2, wh2);
        if ctx.forced {
            apply_hook_rect_now(ctx.hook, window, final_target, ZORDER_BOTTOM);
            force_place_window(
                window,
                final_target,
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                ctx,
            );
        } else {
            let _ =
                unsafe { SetWindowPos(window, HWND_BOTTOM, wx2, wy2, ww2, wh2, SWP_NOACTIVATE) };
        }
    }

    if !unsafe { IsWindowVisible(window) }.as_bool() {
        let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
        if ctx.forced {
            let (cw, ch) = the_ask.unwrap_or((hw, hh));
            let (wx3, wy3, ww3, wh3) = behind_host_window_rect((hx, hy, cw, ch), insets);
            final_target = (wx3, wy3, ww3, wh3);
            apply_hook_rect_now(ctx.hook, window, final_target, ZORDER_BOTTOM);
            force_place_window(
                window,
                final_target,
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                ctx,
            );
        }
    }
    if ctx.forced {
        sync_hook_if_shifted(ctx.hook, window, (final_target.0, final_target.1));
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
    ctx: PlacementContext<'_>,
) -> bool {
    let insets = if window_mode == CustomWindowMode::Normal {
        dwm_frame_insets(window)
    } else {
        (0, 0, 0, 0)
    };
    let (wx, wy, ww, wh) = behind_host_window_rect(host_rect, insets);

    if ctx.forced {
        let no_shift_ctx = PlacementContext {
            avoid_cursor: None,
            ..ctx
        };
        apply_hook_rect_now(ctx.hook, window, (wx, wy, ww, wh), ZORDER_BOTTOM);
        force_place_window(
            window,
            (wx, wy, ww, wh),
            HWND_BOTTOM,
            SET_WINDOW_POS_FLAGS(0),
            no_shift_ctx,
        );
        if !unsafe { IsWindowVisible(window) }.as_bool() {
            let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
            apply_hook_rect_now(ctx.hook, window, (wx, wy, ww, wh), ZORDER_BOTTOM);
            force_place_window(
                window,
                (wx, wy, ww, wh),
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                no_shift_ctx,
            );
        }
    } else {
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

/// True when `window` is either DWM-cloaked by the hook or non-topmost (strictly behind topmost `host`).
fn is_window_behind_host(window: HWND, cloaked: bool) -> bool {
    if cloaked {
        return true;
    }
    let ex = unsafe { GetWindowLongPtrW(window, GWL_EXSTYLE) } as u32;
    (ex & WS_EX_TOPMOST.0) == 0
}

/// True when an `OffScreen`-placed window needs another offscreen park attempt.
/// Avoids 8ms `STAGE_COORD` <-> monitor-clamp ping-pong when the app already clamped the previous
/// park attempt on-screen at `expected_rect`, off-cursor (or handled by caller nudge), and
/// cloaked or behind `host`.
fn should_repark_offscreen_state(
    needs_window_restore: bool,
    cur_rect: Option<RECT>,
    expected_rect: Option<RECT>,
    vscreen: RECT,
    forced: bool,
    cursor_pt: Option<POINT>,
    cloaked_or_behind_host: bool,
) -> bool {
    if needs_window_restore {
        return true;
    }
    if cur_rect.is_some_and(|r| is_off_screen(&r, vscreen)) {
        return false;
    }
    let same_rect = cur_rect.is_some() && cur_rect == expected_rect;
    let clamped_onscreen = cur_rect.is_some_and(|r| !is_off_screen(&r, vscreen));
    let off_cursor = cursor_pt.is_none_or(|pt| cur_rect.is_some_and(|r| !point_in_rect(pt, r)));
    !(same_rect && clamped_onscreen && cloaked_or_behind_host && (off_cursor || forced))
}

fn should_repark_offscreen(
    window: HWND,
    expected_rect: Option<RECT>,
    forced: bool,
    cursor_pt: Option<POINT>,
    cloaked: bool,
) -> bool {
    let is_zoomed = unsafe { IsZoomed(window) }.as_bool();
    let is_iconic = unsafe { IsIconic(window) }.as_bool();
    let is_visible = unsafe { IsWindowVisible(window) }.as_bool();
    let cur_rect = window_rect(window);
    let vscreen = virtual_screen();
    should_repark_offscreen_state(
        is_zoomed || is_iconic || !is_visible,
        cur_rect,
        expected_rect,
        vscreen,
        forced,
        cursor_pt,
        is_window_behind_host(window, cloaked),
    )
}

/// If an off-screen parking attempt was clamped onto the monitor and now overlaps `cursor_pt`,
/// shift `window` once off cursor at `HWND_BOTTOM` without bouncing back to `STAGE_COORD`.
fn nudge_clamped_offscreen_off_cursor(
    window: HWND,
    view: Option<&HookSharedView>,
    cursor_pt: Option<POINT>,
    previous_foreground: isize,
    processes: &[u32],
) {
    if is_window_offscreen(window) {
        return;
    }
    let (Some(cursor), Some(actual)) = (cursor_pt, window_rect(window)) else {
        return;
    };
    let work = work_area_rect(cursor);
    let Some((sx, sy)) = shift_rect_off_cursor(actual, cursor, work, FORCED_CURSOR_STANDOFF_PX)
    else {
        return;
    };
    if let Some(view) = view {
        view.set_rect(
            sx,
            sy,
            (actual.right - actual.left).max(1),
            (actual.bottom - actual.top).max(1),
        );
    }
    let _ = unsafe {
        SetWindowPos(
            window,
            HWND_BOTTOM,
            sx,
            sy,
            0,
            0,
            SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING,
        )
    };
    restore_foreground(previous_foreground, processes);
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

/// Fallback off-screen park when `window`'s minimum size exceeds `host`.
fn park_offscreen(window: HWND, width: i32, height: i32, ctx: PlacementContext<'_>) {
    let vscreen = virtual_screen();
    let (px, py) = parking_origin(vscreen);
    let target = RECT {
        left: px,
        top: py,
        right: px + width,
        bottom: py + height,
    };

    if ctx.forced {
        let no_shift_ctx = PlacementContext {
            avoid_cursor: None,
            ..ctx
        };
        apply_hook_rect_now(ctx.hook, window, (px, py, width, height), ZORDER_BOTTOM);
        force_place_window(
            window,
            (px, py, width, height),
            HWND_BOTTOM,
            SET_WINDOW_POS_FLAGS(0),
            no_shift_ctx,
        );
        if should_try_fallback_parking(true, (px, py), window_rect(window), vscreen) {
            apply_hook_rect_now(
                ctx.hook,
                window,
                (FALLBACK_PARKING, FALLBACK_PARKING, width, height),
                ZORDER_BOTTOM,
            );
            force_place_window(
                window,
                (FALLBACK_PARKING, FALLBACK_PARKING, width, height),
                HWND_BOTTOM,
                SWP_NOSIZE,
                no_shift_ctx,
            );
        }
        if !unsafe { IsWindowVisible(window) }.as_bool() {
            let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
            if !is_window_offscreen(window) {
                apply_hook_rect_now(ctx.hook, window, (px, py, width, height), ZORDER_BOTTOM);
                force_place_window(
                    window,
                    (px, py, width, height),
                    HWND_BOTTOM,
                    SET_WINDOW_POS_FLAGS(0),
                    no_shift_ctx,
                );
            }
        }
        nudge_clamped_offscreen_off_cursor(
            window,
            ctx.hook,
            ctx.avoid_cursor,
            ctx.previous_foreground,
            ctx.processes,
        );
    } else {
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
            unsafe { SetWindowPos(window, HWND_BOTTOM, px, py, width, height, SWP_NOACTIVATE) }
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

        if !unsafe { IsWindowVisible(window) }.as_bool() {
            let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
        }
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
    chrome: (i32, i32, i32, i32),
    window_mode: CustomWindowMode,
    ctx: PlacementContext<'_>,
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
            place_invisible_layered(source, target_rect, chrome, window_mode, the_ask, ctx);
            restore_foreground(ctx.previous_foreground, ctx.processes);
            PlacementMode::InvisibleLayered
        }
        SnapAction::PlaceDirectWindow => {
            place_direct_window(source, target_rect, window_mode, ctx);
            restore_foreground(ctx.previous_foreground, ctx.processes);
            PlacementMode::DirectWindow
        }
        SnapAction::KeepOffScreen => PlacementMode::OffScreen,
        SnapAction::ReparkOffScreen => {
            let ask = the_ask.unwrap_or_else(|| the_ask_for_host(source, target_rect, chrome));
            *the_ask = Some(ask);
            park_offscreen(source, ask.0, ask.1, ctx);
            restore_foreground(ctx.previous_foreground, ctx.processes);
            PlacementMode::OffScreen
        }
        SnapAction::TryBehindHost => {
            let mode = if place_behind_host(source, target_rect, window_mode, ctx) {
                *the_ask = None;
                PlacementMode::BehindHost
            } else {
                let ask = the_ask_for_host(source, target_rect, chrome);
                *the_ask = Some(ask);
                park_offscreen(source, ask.0, ask.1, ctx);
                PlacementMode::OffScreen
            };
            restore_foreground(ctx.previous_foreground, ctx.processes);
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
    let cur_pid = pid_of(current);
    if processes.contains(&cur_pid) {
        let ok = unsafe { SetForegroundWindow(prev) }.as_bool();
        cp_trace!(
            "restore_foreground: active prev=0x{:x} prev_class={:?} from=0x{:x} from_pid={} ok={}",
            previous as usize,
            hwnd_class_name(prev),
            current.0 as usize,
            cur_pid,
            ok
        );
    }
}

fn should_restore_foreground_on_teardown(
    previous: isize,
    fg_before: isize,
    fg_pid_before: u32,
    fg_after: isize,
    fg_pid_after: u32,
    dying_pids: &[u32],
) -> bool {
    previous != 0
        && fg_after != previous
        && ((fg_before != previous && dying_pids.contains(&fg_pid_before))
            || dying_pids.contains(&fg_pid_after))
}

/// Restore previous foreground window when closing/terminating a preview that held foreground,
/// preventing Windows from auto-raising another window (such as an already-open terminal).
fn restore_foreground_on_teardown(previous: isize, fg_before: HWND, dying_pids: &[u32]) {
    if previous == 0 || dying_pids.is_empty() {
        return;
    }
    let prev = HWND(previous as *mut _);
    if !unsafe { IsWindow(prev) }.as_bool()
        || !unsafe { IsWindowVisible(prev) }.as_bool()
        || unsafe { IsIconic(prev) }.as_bool()
    {
        return;
    }
    let fg_after = unsafe { GetForegroundWindow() };
    let fg_before_pid = pid_of(fg_before);
    let fg_after_pid = pid_of(fg_after);
    if should_restore_foreground_on_teardown(
        previous,
        fg_before.0 as isize,
        fg_before_pid,
        fg_after.0 as isize,
        fg_after_pid,
        dying_pids,
    ) {
        let ok = unsafe { SetForegroundWindow(prev) }.as_bool();
        cp_trace!(
            "restore_foreground: teardown prev=0x{:x} prev_class={:?} fg_before=0x{:x} fg_before_pid={} fg_after=0x{:x} fg_after_pid={} ok={}",
            previous as usize,
            hwnd_class_name(prev),
            fg_before.0 as usize,
            fg_before_pid,
            fg_after.0 as usize,
            fg_after_pid,
            ok
        );
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DiscoveryMatch {
    TreeOrJob,
    ImageName,
    InstallDir,
}

fn full_path_matches_install_dir(
    pid_is_new: bool,
    install_dir: Option<&Path>,
    full_path: Option<&Path>,
) -> bool {
    if !pid_is_new {
        return false;
    }
    let (Some(dir), Some(exe)) = (install_dir, full_path) else {
        return false;
    };
    exe_is_inside_install_dir(exe, dir)
}

fn discovery_match_for_pid(
    pid: u32,
    launch: &Launch,
    processes: &[u32],
    mode: OwnershipMode,
    images: &mut ProcessImages,
    creation_times: &mut HashMap<u32, Option<u64>>,
) -> Option<DiscoveryMatch> {
    if pid == 0 || pid == std::process::id() {
        return None;
    }
    if processes.contains(&pid) {
        return Some(DiscoveryMatch::TreeOrJob);
    }
    let OwnershipMode::Discover {
        allow_image_fallback: true,
    } = mode
    else {
        return None;
    };
    let full_path = image_path_of(pid, images);
    let image = full_path.as_deref().and_then(image_name_from_path);
    let is_new = !launch.processes_before.contains(&pid);
    let creation = if image.as_deref() == Some(launch.exe.as_str()) && is_new {
        creation_time_of(pid, creation_times)
    } else {
        None
    };
    if belongs_to_preview(
        pid,
        processes,
        mode,
        &launch.exe,
        launch.launch_filetime,
        &launch.processes_before,
        image.as_deref(),
        creation,
    ) {
        return Some(DiscoveryMatch::ImageName);
    }
    if full_path_matches_install_dir(is_new, launch.install_dir.as_deref(), full_path.as_deref()) {
        return Some(DiscoveryMatch::InstallDir);
    }
    None
}

fn discovery_processes(launch: &Launch) -> Vec<u32> {
    let my_pid = std::process::id();
    let mut found = Vec::new();
    for pid in discovery_tree(launch)
        .into_iter()
        .chain(discovery_job_pids(launch))
    {
        if pid != 0 && pid != my_pid && !found.contains(&pid) {
            found.push(pid);
        }
    }
    found
}

/// The launched tree, walked only as far as the preset's `Child Processes` row allows.
fn discovery_tree(launch: &Launch) -> Vec<u32> {
    process_tree_to(launch.pid, launch.child_scope.depth())
}

/// The job's own members — which are processes the preset asked for only where it asked for
/// children at all.
fn discovery_job_pids(launch: &Launch) -> Vec<u32> {
    if launch.child_scope == CustomChildScope::Disabled {
        return Vec::new();
    }
    active_job_pids(launch.id)
}

fn refresh_job_pids(launch: &Launch, processes: &mut Vec<u32>) -> Vec<u32> {
    let my_pid = std::process::id();
    let job_pids = discovery_job_pids(launch);
    let mut changed = false;
    for &pid in &job_pids {
        if pid != 0 && pid != my_pid && !processes.contains(&pid) {
            processes.push(pid);
            changed = true;
        }
    }
    if changed {
        update_owned_pids_if_current(launch, processes);
    }
    job_pids
}

/// Whether a pid the launched tree never yielded may still be adopted: the preset asked for the
/// broker road, and the launcher has either exited or had its grace.
///
/// **Off unless the preset asks for it.** A packaged program (Windows 11's Paint) is begun by an
/// activator that hands the file to a pid this app never started, and a Squirrel/Velopack or NSIS
/// stub exits as soon as it has begun the real program — so the road exists, it is just not one a
/// preview takes without being told to, because the pid it adopts is one the user never named.
fn broker_fallback_allowed(launch: &Launch) -> bool {
    if !launch.broker_fallback {
        return false;
    }
    let grace = if launch.packaged_broker {
        Duration::ZERO
    } else {
        BROKER_FALLBACK_GRACE
    };
    process_has_exited(HANDLE(launch.process.as_raw_handle()))
        || launch.launched_at.elapsed() >= grace
}

fn set_window_cloak(hwnd: HWND, cloak: bool) -> bool {
    if hwnd.0.is_null() || !unsafe { IsWindow(hwnd) }.as_bool() {
        return false;
    }
    let val: u32 = u32::from(cloak);
    unsafe {
        DwmSetWindowAttribute(
            hwnd,
            DWMWA_CLOAK,
            &val as *const u32 as *const _,
            std::mem::size_of::<u32>() as u32,
        )
    }
    .is_ok()
}

fn is_precloak_candidate(hwnd: HWND) -> bool {
    if hwnd.0.is_null() || !unsafe { IsWindow(hwnd) }.as_bool() {
        return false;
    }
    let style = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) } as u32;
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    (style & WS_CHILD.0) == 0 && (ex_style & WS_EX_TOOLWINDOW.0) == 0 && !owner_is_blocking(hwnd)
}

fn is_cloaked_by_pre_hook(pre_hook: Option<&InProcessHookSession>, pid: u32) -> bool {
    pre_hook.is_some_and(|session| {
        session.pid == pid && session.view.is_cloaked() && session.view.hook_hits() > 0
    })
}

type NtGetNextThreadFn = unsafe extern "system" fn(
    process: HANDLE,
    thread: HANDLE,
    desired_access: u32,
    handle_attributes: u32,
    flags: u32,
    new_thread: *mut HANDLE,
) -> i32;

fn nt_get_next_thread_fn() -> Option<NtGetNextThreadFn> {
    static FN: std::sync::OnceLock<Option<NtGetNextThreadFn>> = std::sync::OnceLock::new();
    *FN.get_or_init(|| unsafe {
        let ntdll = GetModuleHandleW(w!("ntdll.dll")).ok()?;
        let proc = GetProcAddress(ntdll, PCSTR(c"NtGetNextThread".as_ptr().cast()))?;
        // SAFETY: `NtGetNextThread` exported by ntdll.dll matches `NtGetNextThreadFn` ABI.
        Some(std::mem::transmute::<
            unsafe extern "system" fn() -> isize,
            NtGetNextThreadFn,
        >(proc))
    })
}

fn threads_of_process_via_nt(pid: u32) -> Option<Vec<u32>> {
    if pid == 0 {
        return Some(Vec::new());
    }
    let nt_get_next_thread = nt_get_next_thread_fn()?;
    let proc_handle = unsafe { OpenProcess(PROCESS_QUERY_INFORMATION, false, pid) }.ok()?;
    let mut tids = Vec::new();
    let mut prev = HANDLE::default();
    for _ in 0..1024 {
        let mut next = HANDLE::default();
        let status = unsafe {
            nt_get_next_thread(
                proc_handle,
                prev,
                THREAD_QUERY_LIMITED_INFORMATION.0,
                0,
                0,
                &mut next,
            )
        };
        if !prev.is_invalid() && !prev.0.is_null() {
            let _ = unsafe { CloseHandle(prev) };
        }
        prev = HANDLE::default();
        if status < 0 || next.is_invalid() || next.0.is_null() {
            if !next.is_invalid() && !next.0.is_null() {
                let _ = unsafe { CloseHandle(next) };
            }
            break;
        }
        let tid = unsafe { GetThreadId(next) };
        if tid != 0 {
            tids.push(tid);
        }
        prev = next;
    }
    if !prev.is_invalid() && !prev.0.is_null() {
        let _ = unsafe { CloseHandle(prev) };
    }
    let _ = unsafe { CloseHandle(proc_handle) };
    Some(tids)
}

fn threads_of_process(pid: u32) -> Vec<u32> {
    if pid == 0 {
        return Vec::new();
    }
    if let Some(tids) = threads_of_process_via_nt(pid) {
        return tids;
    }
    let Ok(snapshot) = (unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0) }) else {
        return Vec::new();
    };
    let mut entry = THREADENTRY32 {
        dwSize: std::mem::size_of::<THREADENTRY32>() as u32,
        ..Default::default()
    };
    let mut tids = Vec::new();
    if unsafe { Thread32First(snapshot, &mut entry) }.is_ok() {
        loop {
            if entry.th32OwnerProcessID == pid && entry.th32ThreadID != 0 {
                tids.push(entry.th32ThreadID);
            }
            if unsafe { Thread32Next(snapshot, &mut entry) }.is_err() {
                break;
            }
        }
    }
    let _ = unsafe { CloseHandle(snapshot) };
    tids
}

struct DiscoveredWindow {
    hwnd: HWND,
    hook_session: Option<InProcessHookSession>,
    pre_cloaked: bool,
}

struct PreCloakTracker {
    enabled: bool,
    cloaked_hwnds: HashSet<isize>,
}

impl PreCloakTracker {
    fn new(
        requested_capture: CustomCaptureTechnique,
        sibling_windows: CustomSiblingWindows,
    ) -> Self {
        // **Two things have to be true before a window is cloaked from the outside**: the preset
        // asked for it (`Sibling Windows` is not `Disabled`) and the technique can survive it. A
        // window that is meant to stand on screen is the one a stray cloak kills outright, which
        // is why `DirectWindow` keeps its carve-out whatever the preset says.
        Self {
            enabled: sibling_windows.cloaks()
                && requested_capture != CustomCaptureTechnique::DirectWindow,
            cloaked_hwnds: HashSet::new(),
        }
    }

    fn try_cloak(&mut self, hwnd: HWND, pre_hook: Option<&InProcessHookSession>, pid: u32) {
        if !self.enabled || pid == 0 || is_cloaked_by_pre_hook(pre_hook, pid) {
            return;
        }
        let raw = hwnd.0 as isize;
        if !self.cloaked_hwnds.contains(&raw) && set_window_cloak(hwnd, true) {
            self.cloaked_hwnds.insert(raw);
        }
    }

    fn finish_with_winner(
        &mut self,
        winner: HWND,
        pre_hook: Option<&InProcessHookSession>,
        winner_pid: u32,
    ) -> bool {
        self.try_cloak(winner, pre_hook, winner_pid);
        let winner_raw = winner.0 as isize;
        let winner_cloaked = self.cloaked_hwnds.contains(&winner_raw);
        for raw in self.cloaked_hwnds.drain() {
            if raw != winner_raw {
                let _ = set_window_cloak(HWND(raw as *mut _), false);
            }
        }
        winner_cloaked
    }
}

impl Drop for PreCloakTracker {
    fn drop(&mut self) {
        for raw in self.cloaked_hwnds.drain() {
            let _ = set_window_cloak(HWND(raw as *mut _), false);
        }
    }
}

/// Pre-attaches `custom_preview_hook.dll` onto 64-bit `PreviewJob` child processes during
/// discovery so stub-spawned grandchildren are hooked & cloaked before `CreateWindowExW` runs.
struct JobChildPreHook {
    enabled: bool,
    dll_missing: bool,
    desired_capture: CustomCaptureTechnique,
    placement: CustomPlacement,
    initial_hook_rect: Option<(i32, i32, i32, i32)>,
    pre_hook: Option<InProcessHookSession>,
    hooked_tids: HashSet<u32>,
    tight_attempted_pids: HashSet<u32>,
    wow64_cache: HashMap<u32, Option<bool>>,
    logged_errs: HashSet<(u32, u32, u32)>,
}

impl JobChildPreHook {
    fn new(launch: &Launch, effective_capture: CustomCaptureTechnique) -> Self {
        // **The gate on writing this app's DLL into another process.** A job child is a pid this
        // app never began, so the pre-attach is a `Child Processes` question and not only a
        // capture-technique one: a preset that asked for the launched pid alone gets no injector.
        let wants_hook = (launch.requested_capture == CustomCaptureTechnique::InProcessHook
            || uses_hook_enforcement(launch.placement))
            && launch.child_scope != CustomChildScope::Disabled;
        let desired_capture = if launch.requested_capture == CustomCaptureTechnique::InProcessHook {
            CustomCaptureTechnique::InProcessHook
        } else {
            effective_capture
        };
        Self {
            enabled: wants_hook,
            dll_missing: false,
            desired_capture,
            placement: launch.placement,
            initial_hook_rect: launch.provisional,
            pre_hook: None,
            hooked_tids: HashSet::new(),
            tight_attempted_pids: HashSet::new(),
            wow64_cache: HashMap::new(),
            logged_errs: HashSet::new(),
        }
    }

    fn poll(&mut self, launch: &Launch, job_pids: &[u32]) {
        if !self.enabled || self.dll_missing || job_pids.is_empty() {
            return;
        }
        if self
            .pre_hook
            .as_ref()
            .is_some_and(|session| session.hooks.is_empty() && !job_pids.contains(&session.pid))
        {
            if let Some(stale) = self.pre_hook.take() {
                stale.view.clear_locks_for_detach();
            }
            self.hooked_tids.clear();
        }
        let my_pid = std::process::id();
        if self.pre_hook.is_none() {
            for &child_pid in job_pids {
                if child_pid == 0 || child_pid == launch.pid || child_pid == my_pid {
                    continue;
                }
                let wow64 = *self
                    .wow64_cache
                    .entry(child_pid)
                    .or_insert_with(|| pid_is_wow64(child_pid));
                if wow64 != Some(false) {
                    continue;
                }
                let Some((hmod, _, _)) = load_hook_dll() else {
                    self.dll_missing = true;
                    return;
                };
                let uncloak_msg = unsafe { RegisterWindowMessageW(HOOK_CONTROL_MSG_NAME) };
                let init =
                    hook_init_for(self.desired_capture, self.placement, self.initial_hook_rect);
                let Some(view) = HookSharedView::create(child_pid, init, uncloak_msg).map(Arc::new)
                else {
                    let _ = unsafe { FreeLibrary(hmod) };
                    continue;
                };
                self.pre_hook = Some(InProcessHookSession {
                    pid: child_pid,
                    view,
                    hooks: Vec::new(),
                    dll_module: hmod,
                });
                break;
            }
        }
        let Some(ref mut session) = self.pre_hook else {
            return;
        };
        let pid = session.pid;
        let Some((cbt_proc, cwp_proc)) = hook_procs_from_module(session.dll_module) else {
            return;
        };

        let attach_start = Instant::now();
        let do_tight_retry = session.hooks.is_empty() && self.tight_attempted_pids.insert(pid);
        if do_tight_retry {
            let tight_deadline = attach_start + Duration::from_millis(60);
            let enum_start = Instant::now();
            let mut tids = threads_of_process(pid);
            let mut enum_ms = enum_start.elapsed().as_millis();
            let mut last_enum = Instant::now();
            let mut attempts: u32 = 0;
            let mut last_err: Option<(&'static str, u32)> = None;
            while !tids.is_empty() && Instant::now() < tight_deadline {
                if !launch_still_wanted(launch) {
                    return;
                }
                if last_enum.elapsed() >= Duration::from_millis(5) {
                    let re_start = Instant::now();
                    tids = threads_of_process(pid);
                    enum_ms = enum_ms.max(re_start.elapsed().as_millis());
                    last_enum = Instant::now();
                    if tids.is_empty() {
                        break;
                    }
                }
                attempts = attempts.saturating_add(1);
                for &tid in &tids {
                    if self.hooked_tids.contains(&tid) {
                        continue;
                    }
                    match try_attach_thread_hooks_result(
                        session.dll_module,
                        cbt_proc,
                        cwp_proc,
                        tid,
                    ) {
                        Ok((cbt, cwp)) => {
                            self.hooked_tids.insert(tid);
                            session.hooks.push((tid, cbt, cwp));
                            cp_trace!(
                                "pre_attach_job_child: ok id={} pid={} tid={} hooks={} attempts={} enum_ms={} elapsed_ms={}",
                                launch.id,
                                pid,
                                tid,
                                session.hooks.len(),
                                attempts,
                                enum_ms,
                                attach_start.elapsed().as_millis()
                            );
                            return;
                        }
                        Err(err) => {
                            last_err = Some(err);
                        }
                    }
                }
                std::thread::yield_now();
            }
            if session.hooks.is_empty() {
                cp_trace!(
                    "pre_attach_job_child: tight_retry_timeout id={} pid={} attempts={} enum_ms={} elapsed_ms={} tids={} last_err={:?}",
                    launch.id,
                    pid,
                    attempts,
                    enum_ms,
                    attach_start.elapsed().as_millis(),
                    tids.len(),
                    last_err
                );
            }
            return;
        }

        let enum_start = Instant::now();
        let tids = threads_of_process(pid);
        let enum_ms = enum_start.elapsed().as_millis();
        for tid in tids {
            if self.hooked_tids.contains(&tid) {
                continue;
            }
            match try_attach_thread_hooks_result(session.dll_module, cbt_proc, cwp_proc, tid) {
                Ok((cbt, cwp)) => {
                    self.hooked_tids.insert(tid);
                    session.hooks.push((tid, cbt, cwp));
                    cp_trace!(
                        "pre_attach_job_child: ok id={} pid={} tid={} hooks={} attempts=1 enum_ms={} elapsed_ms={}",
                        launch.id,
                        pid,
                        tid,
                        session.hooks.len(),
                        enum_ms,
                        attach_start.elapsed().as_millis()
                    );
                }
                Err((stage, err)) => {
                    // err=87 (`ERROR_INVALID_PARAMETER`) means USER32 is not initialized yet; retry next tick.
                    if err != 87 && self.logged_errs.insert((pid, tid, err)) {
                        cp_trace!(
                            "pre_attach_job_child: err id={} pid={} tid={} stage={} err={}",
                            launch.id,
                            pid,
                            tid,
                            stage,
                            err
                        );
                    }
                }
            }
        }
    }

    fn wait_cap_ms(&self, launch: &Launch) -> u128 {
        if self.enabled
            && self.pre_hook.is_none()
            && launch.launched_at.elapsed() < Duration::from_millis(150)
        {
            1
        } else {
            5
        }
    }

    fn session(&self) -> Option<&InProcessHookSession> {
        self.pre_hook.as_ref()
    }

    fn captured_hwnd(&self) -> Option<HWND> {
        self.pre_hook.as_ref().and_then(|s| s.view.captured_hwnd())
    }

    fn take_for_pid(&mut self, pid: u32) -> Option<InProcessHookSession> {
        let matches = self.pre_hook.as_ref().is_some_and(|s| s.pid == pid);
        if matches {
            self.pre_hook.take()
        } else {
            if let Some(stale) = self.pre_hook.take() {
                stale.view.clear_locks_for_detach();
            }
            None
        }
    }
}

impl Drop for JobChildPreHook {
    fn drop(&mut self) {
        if let Some(stale) = self.pre_hook.take() {
            stale.view.clear_locks_for_detach();
        }
    }
}

/// Owned PID set after adoption: launched tree union active job PIDs union adopted brokered tree (only if newly spawned).
fn owned_pids_for(
    launch: &Launch,
    launch_pid: u32,
    source_pid: u32,
    source_is_new: bool,
) -> Vec<u32> {
    let my_pid = std::process::id();
    let mut owned = Vec::new();
    let depth = launch.child_scope.depth();
    for pid in process_tree_to(launch_pid, depth)
        .into_iter()
        .chain(discovery_job_pids(launch))
    {
        if pid != 0 && pid != my_pid && !owned.contains(&pid) {
            owned.push(pid);
        }
    }
    if source_is_new && source_pid != 0 && depth != CustomChildDepth::LaunchedOnly {
        for pid in process_tree_to(source_pid, depth) {
            if pid != 0 && pid != my_pid && !owned.contains(&pid) {
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
    images: &'a mut ProcessImages,
    creation_times: &'a mut HashMap<u32, Option<u64>>,
    /// The window this app has already adopted, which is the picture and not to be hidden.
    keep: isize,
    /// Hide visible non-`keep` windows only after `source` is adopted (never during discovery).
    hide_others: bool,
    mode: OwnershipMode,
    /// Candidate windows ranked by tier (`Main > HiddenOwner > UntitledVisible > ToolWindow`),
    /// then visibility, then area.
    candidates: Vec<(CandidateRank, HWND)>,
}

/// Look at the desktop once, on the terms `Census` states.
fn census<'a>(
    launch: &'a Launch,
    processes: &'a [u32],
    images: &'a mut ProcessImages,
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
        // **The preset's `Sibling Windows` row, and only where the caller asked at all.** A
        // caller that says `false` (every discovery pass) still hides nothing, and a preset that
        // says so hides nothing either — the two are one gate because they are one question.
        hide_others: hide_others && launch.sibling_windows.hides(),
        mode,
        candidates: Vec::new(),
    };
    let _ = unsafe { EnumWindows(Some(visit), LPARAM(&mut census as *mut Census<'_> as isize)) };
    census
}

impl Census<'_> {
    fn best_candidate(&self) -> Option<(CandidateRank, HWND)> {
        select_best_candidate(&self.candidates)
    }

    /// Whether this look found anything a picture could be drawn in at all, settled or not. The
    /// distinction matters where the row asks for the newest: a census that has candidates and has
    /// not settled yet is *about* to answer, and a latch adopted in the meantime is the wrong window.
    fn has_candidates(&self) -> bool {
        !self.candidates.is_empty()
    }

    /// The picture among what this look found, on the terms the preset's `Picture Window` row names,
    /// measured against what this launch has watched appear (see `select_the_picture`).
    fn the_picture(&self) -> Option<(CandidateRank, HWND)> {
        let mode = self.launch.picture_window;
        if !mode.prefers_the_newest() {
            return self.best_candidate();
        }
        SIGHTINGS
            .try_with(|held| select_the_picture(&self.candidates, &held.borrow(), mode))
            .ok()
            .flatten()
    }

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
        let tool_window = ex_style & WS_EX_TOOLWINDOW.0 != 0;
        let owner_hwnd = unsafe { GetWindow(hwnd, GW_OWNER) }
            .ok()
            .filter(|o| !o.0.is_null());
        let strictly_owned = owner_hwnd.is_some();
        let owner_blocking = owner_hwnd.is_some_and(window_is_blocking_owner);
        let titled = unsafe { GetWindowTextLengthW(hwnd) } > 0;
        let visible = unsafe { IsWindowVisible(hwnd) }.as_bool();
        let mut rect = RECT::default();
        let size = if unsafe { GetWindowRect(hwnd, &mut rect) }.is_ok() {
            WindowSize::of(rect)
        } else {
            WindowSize {
                area: 0,
                a_strip: true,
            }
        };

        if let Some(tier) = classify_candidate(
            titled,
            strictly_owned,
            owner_blocking,
            tool_window,
            visible,
            size,
        ) {
            // **When it was first seen, recorded by the look that first saw it.** The ranking is a
            // separate question (see `select_the_picture`); this is the one fact a single look at the
            // desktop cannot answer, kept by the launch's own record rather than in the rank.
            let _ = SIGHTINGS.try_with(|held| held.borrow_mut().note(hwnd));
            self.candidates.push((
                CandidateRank {
                    tier,
                    visible,
                    area: size.area,
                },
                hwnd,
            ));
        }

        if visible && self.hide_others {
            suppress_companion_window(hwnd, self.keep);
        }
    }

    fn belongs(&mut self, pid: u32) -> bool {
        discovery_match_for_pid(
            pid,
            self.launch,
            self.processes,
            self.mode,
            self.images,
            self.creation_times,
        )
        .is_some()
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
fn wait_for_a_window_offscreen_staged(
    launch: &Launch,
    effective_capture: CustomCaptureTechnique,
) -> Option<DiscoveredWindow> {
    forget_sightings();
    let deadline = Instant::now() + WINDOW_WAIT;
    let mut processes = discovery_processes(launch);
    update_owned_pids_if_current(launch, &processes);
    let mut read = Instant::now();
    let mut images: ProcessImages = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();
    let mut first_hidden: Option<(Instant, HWND)> = None;
    let mut pre_cloak = PreCloakTracker::new(launch.requested_capture, launch.sibling_windows);
    let mut pre_hook = JobChildPreHook::new(launch, effective_capture);
    let initial_job_pids = active_job_pids(launch.id);
    pre_hook.poll(launch, &initial_job_pids);
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

    let fallback_allowed = broker_fallback_allowed;

    let check_now = |processes: &[u32],
                     images: &mut ProcessImages,
                     creation_times: &mut HashMap<u32, Option<u64>>,
                     first_hidden: &mut Option<(Instant, HWND)>,
                     pre_cloak: &mut PreCloakTracker,
                     pre_hook: &JobChildPreHook|
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
        if let Some((rank, window)) = look.the_picture() {
            pre_cloak.try_cloak(window, pre_hook.session(), pid_of(window));
            if rank.visible {
                return Some(window);
            }
            let entry = first_hidden.get_or_insert_with(|| (Instant::now(), window));
            entry.1 = window;
            if entry.0.elapsed() >= Duration::from_millis(40) {
                return Some(window);
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
        if let Some((rank, window)) = look.the_picture() {
            pre_cloak.try_cloak(window, pre_hook.session(), pid_of(window));
            if rank.visible {
                return Some(window);
            }
            let entry = first_hidden.get_or_insert_with(|| (Instant::now(), window));
            entry.1 = window;
            if entry.0.elapsed() >= Duration::from_millis(40) {
                return Some(window);
            }
        }
        None
    };

    let finish_found = |window: HWND,
                        pre_cloak: &mut PreCloakTracker,
                        pre_hook: &mut JobChildPreHook|
     -> DiscoveredWindow {
        let winner_pid = pid_of(window);
        let hook_session = pre_hook.take_for_pid(winner_pid);
        let pre_cloaked = pre_cloak.finish_with_winner(window, hook_session.as_ref(), winner_pid);
        DiscoveredWindow {
            hwnd: window,
            hook_session,
            pre_cloaked,
        }
    };

    while Instant::now() < deadline {
        if !launch_still_wanted(launch) {
            cp_trace!("wait_staged: aborted_loop id={}", launch.id);
            return None;
        }
        let job_pids = refresh_job_pids(launch, &mut processes);
        pre_hook.poll(launch, &job_pids);
        if let Some(window) = check_now(
            &processes,
            &mut images,
            &mut creation_times,
            &mut first_hidden,
            &mut pre_cloak,
            &pre_hook,
        ) {
            if !launch_still_wanted(launch) {
                cp_trace!(
                    "wait_staged: aborted_after_check id={} hwnd=0x{:x}",
                    launch.id,
                    window.0 as usize
                );
                cleanup_aborted_source(window, launch);
                return None;
            }
            return Some(finish_found(window, &mut pre_cloak, &mut pre_hook));
        }

        if read.elapsed() >= TREE_POLL {
            processes = discovery_processes(launch);
            update_owned_pids_if_current(launch, &processes);
            read = Instant::now();
        }

        let remaining_ms = deadline
            .saturating_duration_since(Instant::now())
            .as_millis()
            .min(pre_hook.wait_cap_ms(launch)) as u32;
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
                    if !processes.contains(&pid) {
                        let job_pids = refresh_job_pids(launch, &mut processes);
                        pre_hook.poll(launch, &job_pids);
                    }
                    let belongs = discovery_match_for_pid(
                        pid,
                        launch,
                        &processes,
                        mode,
                        &mut images,
                        &mut creation_times,
                    )
                    .is_some();
                    if belongs {
                        if event == EVENT_OBJECT_CREATE || event == EVENT_OBJECT_SHOW {
                            if is_precloak_candidate(hwnd) {
                                pre_cloak.try_cloak(hwnd, pre_hook.session(), pid);
                            }
                            stage_candidate_offscreen(hwnd, stage_w, stage_h);
                            restore_foreground(launch.previous_foreground, &processes);
                        }
                        matched = true;
                    }
                }
            });
            if matched {
                if !launch_still_wanted(launch) {
                    cp_trace!("wait_staged: aborted_msg id={}", launch.id);
                    return None;
                }
                if let Some(window) = check_now(
                    &processes,
                    &mut images,
                    &mut creation_times,
                    &mut first_hidden,
                    &mut pre_cloak,
                    &pre_hook,
                ) {
                    if !launch_still_wanted(launch) {
                        cp_trace!(
                            "wait_staged: aborted_msg_after_check id={} hwnd=0x{:x}",
                            launch.id,
                            window.0 as usize
                        );
                        cleanup_aborted_source(window, launch);
                        return None;
                    }
                    return Some(finish_found(window, &mut pre_cloak, &mut pre_hook));
                }
            }
        }
    }
    if !launch_still_wanted(launch) {
        cp_trace!(
            "wait_staged: aborted_timeout id={} first_hidden={:?}",
            launch.id,
            first_hidden.map(|(_, h)| h.0 as usize)
        );
        if let Some((_, hwnd)) = first_hidden {
            cleanup_aborted_source(hwnd, launch);
        }
        return None;
    }
    cp_trace!(
        "wait_staged: timeout id={} fallback={:?} procs={:?}",
        launch.id,
        first_hidden.map(|(_, h)| h.0 as usize),
        processes
    );
    first_hidden.map(|(_, hwnd)| finish_found(hwnd, &mut pre_cloak, &mut pre_hook))
}

struct OffscreenStagedState<'a> {
    invisible_layered: &'a mut bool,
    pre_cloaked: &'a mut bool,
    placed_mode: &'a mut Option<PlacementMode>,
    placed_host_rect: &'a mut Option<(i32, i32, i32, i32)>,
    expected_rect: &'a mut Option<RECT>,
    the_ask: &'a mut Option<(i32, i32)>,
    box_of_the_host: &'a mut Option<(i32, i32, i32, i32)>,
    registered: &'a mut isize,
}

fn current_want_snapshot() -> Option<(Where, CustomPlacement)> {
    WANT.lock()
        .ok()
        .and_then(|want| want.as_ref().map(|w| (w.at, w.placement)))
}

fn placement_ctx_from_want<'a>(
    launch: &Launch,
    processes: &'a [u32],
    at: Option<Where>,
    hook: Option<&'a HookSharedView>,
) -> PlacementContext<'a> {
    let want = current_want_snapshot();
    let placement = want.map(|(_, p)| p).unwrap_or(launch.placement);
    let forced = enforces_placement(placement);
    let hook = uses_hook_enforcement(placement).then_some(hook).flatten();
    let avoid_cursor = at.or_else(|| want.map(|(w, _)| w)).and_then(where_cursor);
    PlacementContext {
        forced,
        hook,
        avoid_cursor,
        previous_foreground: launch.previous_foreground,
        processes,
    }
}

/// Strictly isolated stage -> settle at `(STAGE_COORD, STAGE_COORD)` -> reveal phase for `OffscreenStaged`.
/// Returns `false` if launch was superseded or `source` was destroyed.
#[allow(clippy::too_many_arguments)]
fn run_offscreen_staged_settle_and_reveal(
    host: HWND,
    source: HWND,
    launch: &Launch,
    processes: &[u32],
    images: &mut ProcessImages,
    creation_times: &mut HashMap<u32, Option<u64>>,
    was_hidden: bool,
    own: (u32, u32),
    window_mode: CustomWindowMode,
    hook_view: Option<&HookSharedView>,
    current_target_rect: &dyn Fn() -> Option<(i32, i32, i32, i32)>,
    state: &mut OffscreenStagedState<'_>,
) -> bool {
    let forced = enforces_placement(launch.placement);
    let active_hook = uses_hook_enforcement(launch.placement)
        .then_some(hook_view)
        .flatten();
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

    let stage_now = || {
        if forced {
            apply_hook_rect_now(
                active_hook,
                source,
                (STAGE_COORD, STAGE_COORD, stage_w, stage_h),
                ZORDER_BOTTOM,
            );
            force_place_window(
                source,
                (STAGE_COORD, STAGE_COORD, stage_w, stage_h),
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                PlacementContext {
                    forced,
                    hook: active_hook,
                    avoid_cursor: None,
                    previous_foreground: launch.previous_foreground,
                    processes,
                },
            );
        } else {
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
    };

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
        stage_now();
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
        stage_now();
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
        if !is_current(launch.id) {
            return false;
        }
        if !unsafe { IsWindow(source) }.as_bool() {
            if is_current(launch.id) {
                end_the_program();
            }
            return false;
        }
        if current_target_rect().is_none() {
            if is_current(launch.id) {
                end_the_program();
            }
            return false;
        }

        let raw_moved = drain_settle_events(source, launch, processes);
        let cur_rect = window_rect(source);
        let drifted = (raw_moved && cur_rect != staged_expected)
            || cur_rect.is_some_and(|r| !is_off_screen(&r, vscreen))
            || (forced && unsafe { IsZoomed(source) }.as_bool());
        if drifted {
            stage_now();
            let _ = drain_settle_events(source, launch, processes);
            staged_expected = window_rect(source);
            if !forced && staged_expected.is_some_and(|r| !is_off_screen(&r, vscreen)) {
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
            if !is_current(launch.id) {
                return false;
            }
            if !unsafe { IsWindow(source) }.as_bool() {
                if is_current(launch.id) {
                    end_the_program();
                }
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

    if !is_current(launch.id) {
        return false;
    }
    if !unsafe { IsWindow(source) }.as_bool() {
        if is_current(launch.id) {
            end_the_program();
        }
        return false;
    }

    let Some(box_) = current_target_rect() else {
        if is_current(launch.id) {
            end_the_program();
        }
        return false;
    };
    // The host is the box grown by the frame the program draws around its picture (see `the_host_for`).
    let chrome = chrome_now(source, window_mode);
    let host_rect = the_host_for(box_, chrome);
    if let Ok(mut slot) = LAST_HOST_RECT.lock() {
        *slot = Some(host_rect);
    }
    let (hx, hy, hw, hh) = host_rect;

    let ctx = placement_ctx_from_want(launch, processes, None, active_hook);

    // Branch A: try `WS_EX_LAYERED` + `LWA_ALPHA = 1` WHILE STILL AT `(STAGE_COORD, STAGE_COORD)`.
    if make_source_invisible(source.0 as isize) {
        *state.invisible_layered = true;
        SOURCE_INVISIBLE_LAYERED.store(true, Ordering::SeqCst);
        place_invisible_layered(source, host_rect, chrome, window_mode, state.the_ask, ctx);
        if window_rect(source).is_some_and(|r| is_off_screen(&r, vscreen)) {
            // The ask is the frame the picture is composed from, so the window is put where that
            // frame — and not the window rect — lands on the host (see `the_ask_for_host`).
            let (aw, ah) = state.the_ask.unwrap_or((hw, hh));
            let insets = if window_mode == CustomWindowMode::Normal {
                dwm_frame_insets(source)
            } else {
                (0, 0, 0, 0)
            };
            let (wx, wy, ww, wh) = behind_host_window_rect((hx, hy, aw, ah), insets);
            if ctx.forced {
                apply_hook_rect_now(active_hook, source, (wx, wy, ww, wh), ZORDER_BOTTOM);
                force_place_window(
                    source,
                    (wx, wy, ww, wh),
                    HWND_BOTTOM,
                    SET_WINDOW_POS_FLAGS(0),
                    ctx,
                );
                sync_hook_if_shifted(active_hook, source, (wx, wy));
            } else {
                let _ =
                    unsafe { SetWindowPos(source, HWND_BOTTOM, wx, wy, ww, wh, SWP_NOACTIVATE) };
            }
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
    apply_hook_rect_now(
        active_hook,
        source,
        (STAGE_COORD, STAGE_COORD, ww, wh),
        ZORDER_BOTTOM,
    );
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
    let mode = if place_behind_host(source, host_rect, window_mode, ctx) {
        PlacementMode::BehindHost
    } else {
        let ask = the_ask_for_host(source, host_rect, chrome);
        *state.the_ask = Some(ask);
        if ctx.forced {
            apply_hook_rect_now(
                active_hook,
                source,
                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                ZORDER_BOTTOM,
            );
            force_place_window(
                source,
                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                PlacementContext {
                    avoid_cursor: None,
                    ..ctx
                },
            );
        } else {
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
        let _ = unsafe { InvalidateRect(source, None, false) };
        let _ = unsafe { UpdateWindow(source) };
        PlacementMode::OffScreen
    };
    restore_foreground(launch.previous_foreground, processes);
    if *state.pre_cloaked {
        let _ = unsafe { DwmFlush() };
        let _ = set_window_cloak(source, false);
        *state.pre_cloaked = false;
    }
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
const CMD_ENFORCE_NOW: usize = 4;
const ZORDER_BOTTOM: u8 = 0;
const ZORDER_TOPMOST: u8 = 1;
#[allow(dead_code)]
const ZORDER_KEEP: u8 = 2;
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
    forced_placement: AtomicU8,
    z_order_mode: AtomicU8,
}

#[derive(Clone, Copy)]
struct HookSharedInit {
    rect: (i32, i32, i32, i32),
    cloak_enabled: bool,
    rect_locked: bool,
    forced_mode: u8,
    z_order_mode: u8,
    captured_hwnd: isize,
}

fn hook_init_for(
    capture_technique: CustomCaptureTechnique,
    placement: CustomPlacement,
    provisional: Option<(i32, i32, i32, i32)>,
) -> HookSharedInit {
    let stage_by_default = matches!(
        capture_technique,
        CustomCaptureTechnique::InProcessHook
            | CustomCaptureTechnique::OffscreenStaged
            | CustomCaptureTechnique::OffScreen
    );
    let rect = match provisional {
        Some((px, py, pw, ph)) if !stage_by_default => (px, py, pw.max(1), ph.max(1)),
        Some((_, _, pw, ph)) => (STAGE_COORD, STAGE_COORD, pw.max(1), ph.max(1)),
        None => (STAGE_COORD, STAGE_COORD, 0, 0),
    };
    HookSharedInit {
        rect,
        cloak_enabled: capture_technique == CustomCaptureTechnique::InProcessHook,
        rect_locked: true,
        forced_mode: sh_forced_mode(placement),
        z_order_mode: z_order_mode_for(capture_technique),
        captured_hwnd: 0,
    }
}

/// Shared memory mapping for `InProcessHook` & `CustomPlacement::Hook`; safe to clone across `hide()` and worker thread.
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
    fn create(pid: u32, init: HookSharedInit, uncloak_msg: u32) -> Option<Self> {
        let name = wide_null(OsStr::new(&format!("Local\\RhpCustomPreviewHook_{pid}")));
        let size = std::mem::size_of::<HookSharedState>() as u32;
        let mapping = match unsafe {
            CreateFileMappingW(
                INVALID_HANDLE_VALUE,
                None,
                PAGE_READWRITE,
                0,
                size,
                PCWSTR(name.as_ptr()),
            )
        } {
            Ok(h) => h,
            Err(e) => {
                cp_trace!(
                    "shm: CreateFileMappingW failed pid={} err={} code={:?}",
                    pid,
                    win_err(),
                    e
                );
                return None;
            }
        };
        let already_exists = unsafe { GetLastError() } == ERROR_ALREADY_EXISTS;
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
            cp_trace!("shm: MapViewOfFile failed pid={} err={}", pid, win_err());
            return None;
        }
        let ptr = mapped.Value as *mut HookSharedState;
        let (x, y, w, h) = init.rect;
        let existing_valid = unsafe {
            let existing_valid = already_exists
                && (*ptr).magic == HOOK_MAGIC
                && (*ptr).version == HOOK_VERSION
                && (*ptr).target_pid == pid;
            if existing_valid {
                let s = &*ptr;
                s.cloak_enabled
                    .store(u32::from(init.cloak_enabled), Ordering::Release);
                s.rect_locked
                    .store(u32::from(init.rect_locked), Ordering::Release);
                s.x.store(x, Ordering::Release);
                s.y.store(y, Ordering::Release);
                s.w.store(w, Ordering::Release);
                s.h.store(h, Ordering::Release);
                s.captured_hwnd.store(init.captured_hwnd, Ordering::Release);
                s.forced_placement
                    .store(init.forced_mode, Ordering::Release);
                s.z_order_mode.store(init.z_order_mode, Ordering::Release);
            } else {
                ptr.write(HookSharedState {
                    magic: HOOK_MAGIC,
                    version: HOOK_VERSION,
                    target_pid: pid,
                    uncloak_msg,
                    cloak_enabled: AtomicU32::new(u32::from(init.cloak_enabled)),
                    rect_locked: AtomicU32::new(u32::from(init.rect_locked)),
                    x: AtomicI32::new(x),
                    y: AtomicI32::new(y),
                    w: AtomicI32::new(w),
                    h: AtomicI32::new(h),
                    captured_hwnd: AtomicIsize::new(init.captured_hwnd),
                    hook_hits: AtomicU32::new(0),
                    forced_placement: AtomicU8::new(init.forced_mode),
                    z_order_mode: AtomicU8::new(init.z_order_mode),
                });
            }
            existing_valid
        };
        cp_trace!(
            "shm: create pid={} already_exists={} existing_valid={} rect={:?} cloak={} locked={} forced={} z_mode={} msg=0x{:x}",
            pid,
            already_exists,
            existing_valid,
            init.rect,
            init.cloak_enabled,
            init.rect_locked,
            init.forced_mode,
            init.z_order_mode,
            uncloak_msg
        );
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

    fn set_z_order_mode(&self, mode: u8) {
        self.state().z_order_mode.store(mode, Ordering::Release);
    }

    fn set_forced_mode(&self, mode: u8) {
        self.state().forced_placement.store(mode, Ordering::Release);
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
        let err = if sent { 0 } else { win_err() };
        let hits_after = self.hook_hits();
        cp_trace!(
            "shm: cloak_sync hwnd=0x{:x} sent={} err={} hits={}->{}",
            hwnd.0 as usize,
            sent,
            err,
            hits_before,
            hits_after
        );
        sent && hits_after > hits_before
    }

    fn uncloak_sync(&self, hwnd: HWND) {
        self.state().cloak_enabled.store(0, Ordering::Release);
        if !hwnd.0.is_null() && unsafe { IsWindow(hwnd) }.as_bool() {
            let hits_before = self.hook_hits();
            let sent = unsafe {
                SendMessageTimeoutW(
                    hwnd,
                    self.uncloak_msg,
                    WPARAM(CMD_UNCLOAK),
                    LPARAM(0),
                    SMTO_ABORTIFHUNG | SMTO_NORMAL,
                    150,
                    None,
                )
            }
            .0 != 0;
            let err = if sent { 0 } else { win_err() };
            let hits_after = self.hook_hits();
            cp_trace!(
                "shm: uncloak_sync hwnd=0x{:x} sent={} err={} hits={}->{}",
                hwnd.0 as usize,
                sent,
                err,
                hits_before,
                hits_after
            );
        } else {
            cp_trace!("shm: uncloak_sync skipped hwnd=0x{:x}", hwnd.0 as usize);
        }
    }

    /// Synchronously ensure subclass installation and snap `hwnd` to locked SHM rect from in-process.
    fn enforce_now_sync(&self, hwnd: HWND) -> bool {
        if DETACHING.load(Ordering::Acquire) {
            cp_trace!(
                "shm: enforce_now_sync skipped hwnd=0x{:x} reason=detaching",
                hwnd.0 as usize
            );
            return false;
        }
        self.state().rect_locked.store(1, Ordering::Release);
        if DETACHING.load(Ordering::Acquire) {
            self.clear_locks_for_detach();
            cp_trace!(
                "shm: enforce_now_sync cleared hwnd=0x{:x} reason=detaching_race",
                hwnd.0 as usize
            );
            return false;
        }
        if hwnd.0.is_null() || !unsafe { IsWindow(hwnd) }.as_bool() {
            cp_trace!(
                "shm: enforce_now_sync skipped hwnd=0x{:x} reason=not_window",
                hwnd.0 as usize
            );
            return false;
        }
        let hits_before = self.hook_hits();
        let sent = unsafe {
            SendMessageTimeoutW(
                hwnd,
                self.uncloak_msg,
                WPARAM(CMD_ENFORCE_NOW),
                LPARAM(0),
                SMTO_ABORTIFHUNG | SMTO_NORMAL,
                100,
                None,
            )
        }
        .0 != 0;
        let err = if sent { 0 } else { win_err() };
        let hits_after = self.hook_hits();
        let s = self.state();
        cp_trace!(
            "shm: enforce_now_sync hwnd=0x{:x} sent={} err={} hits={}->{} rect=({},{},{},{}) forced={} cloak={} z_mode={}",
            hwnd.0 as usize,
            sent,
            err,
            hits_before,
            hits_after,
            s.x.load(Ordering::Relaxed),
            s.y.load(Ordering::Relaxed),
            s.w.load(Ordering::Relaxed),
            s.h.load(Ordering::Relaxed),
            s.forced_placement.load(Ordering::Relaxed),
            s.cloak_enabled.load(Ordering::Relaxed),
            s.z_order_mode.load(Ordering::Relaxed)
        );
        sent && hits_after > hits_before
    }

    /// Clear cloak + rect and forced-placement locks before promotion moves or restores `hwnd`.
    fn clear_locks_for_detach(&self) {
        let s = self.state();
        s.cloak_enabled.store(0, Ordering::Release);
        s.forced_placement.store(0, Ordering::Release);
        s.rect_locked.store(0, Ordering::Release);
    }

    /// Clear cloak + locks and synchronously remove the in-process subclass hook (`CMD_DETACH`).
    fn detach_sync(&self, hwnd: HWND) -> bool {
        self.clear_locks_for_detach();
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

/// Active hook handles + DLL module + shared memory for `InProcessHook` & `CustomPlacement::Hook`.
struct InProcessHookSession {
    pid: u32,
    view: Arc<HookSharedView>,
    hooks: Vec<(u32, HHOOK, HHOOK)>,
    dll_module: HMODULE,
}

// SAFETY: `HHOOK` and `HMODULE` are process-wide Win32 handles; `UnhookWindowsHookEx` / `FreeLibrary`
// are safe to call from the worker thread on teardown.
unsafe impl Send for InProcessHookSession {}

impl InProcessHookSession {
    fn attach_thread(&mut self, tid: u32) -> bool {
        if tid == 0 || self.hooks.iter().any(|&(t, _, _)| t == tid) {
            return true;
        }
        let Some((cbt_proc, cwp_proc)) = hook_procs_from_module(self.dll_module) else {
            cp_trace!(
                "hook_attach: thread_missing_procs pid={} tid={}",
                self.pid,
                tid
            );
            return false;
        };
        match try_attach_thread_hooks_result(self.dll_module, cbt_proc, cwp_proc, tid) {
            Ok((cbt, cwp)) => {
                cp_trace!(
                    "hook_attach: thread ok pid={} tid={} cbt=0x{:x} cwp=0x{:x}",
                    self.pid,
                    tid,
                    cbt.0 as usize,
                    cwp.0 as usize
                );
                self.hooks.push((tid, cbt, cwp));
                true
            }
            Err((stage, err)) => {
                cp_trace!(
                    "hook_attach: thread failed pid={} tid={} stage={} err={}",
                    self.pid,
                    tid,
                    stage,
                    err
                );
                false
            }
        }
    }
}

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
        for &(_, cbt_hook, callwnd_hook) in &self.hooks {
            if !cbt_hook.0.is_null() {
                let _ = unsafe { UnhookWindowsHookEx(cbt_hook) };
            }
            if !callwnd_hook.0.is_null() {
                let _ = unsafe { UnhookWindowsHookEx(callwnd_hook) };
            }
        }
        if !self.dll_module.0.is_null() {
            let _ = unsafe { FreeLibrary(self.dll_module) };
        }
    }
}

/// The hook DLL this build would load, if the machine has one (see [`crate::paths::hook_dll`]).
///
/// It is the path module's search rather than this one's, because the configuration asks the same
/// question about the same file: a preset's `Placement` row reads whether the hook can be used, and
/// a fresh preset's default is the hook where it can be. One search, so "the DLL is there" cannot
/// mean two things in two places.
fn find_hook_dll() -> Option<PathBuf> {
    crate::paths::hook_dll()
}

/// Whether a launch would find the hook DLL, read once for the preset editor's Capture Technique
/// row. The DLL does not move while this app runs, and a hover searches for it again anyway.
pub(crate) fn hook_dll_present() -> bool {
    *HOOK_DLL_PRESENT.get_or_init(|| find_hook_dll().is_some())
}

static HOOK_DLL_PRESENT: std::sync::OnceLock<bool> = std::sync::OnceLock::new();

fn hook_procs_from_module(hmod: HMODULE) -> Option<(HOOKPROC, HOOKPROC)> {
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
            Some((cbt_proc, cwp_proc))
        }
        _ => None,
    }
}

fn load_hook_dll() -> Option<(HMODULE, HOOKPROC, HOOKPROC)> {
    let Some(path) = find_hook_dll() else {
        cp_trace!("hook_dll: find_hook_dll -> None");
        return None;
    };
    let wide = wide_null(path.as_os_str());
    let hmod = match unsafe { LoadLibraryW(PCWSTR(wide.as_ptr())) } {
        Ok(h) => h,
        Err(e) => {
            cp_trace!(
                "hook_dll: LoadLibraryW failed path={} err={} code={:?}",
                path.display(),
                win_err(),
                e
            );
            return None;
        }
    };
    match hook_procs_from_module(hmod) {
        Some((cbt_proc, cwp_proc)) => {
            cp_trace!(
                "hook_dll: loaded path={} hmod=0x{:x}",
                path.display(),
                hmod.0 as usize
            );
            Some((hmod, cbt_proc, cwp_proc))
        }
        None => {
            cp_trace!(
                "hook_dll: GetProcAddress missing exports path={} err={}",
                path.display(),
                win_err()
            );
            let _ = unsafe { FreeLibrary(hmod) };
            None
        }
    }
}

#[cfg_attr(not(test), allow(dead_code))]
fn try_attach_thread_hooks(
    hmod: HMODULE,
    cbt_proc: HOOKPROC,
    cwp_proc: HOOKPROC,
    tid: u32,
) -> Option<(HHOOK, HHOOK)> {
    try_attach_thread_hooks_result(hmod, cbt_proc, cwp_proc, tid).ok()
}

fn try_attach_thread_hooks_result(
    hmod: HMODULE,
    cbt_proc: HOOKPROC,
    cwp_proc: HOOKPROC,
    tid: u32,
) -> Result<(HHOOK, HHOOK), (&'static str, u32)> {
    let hinst = HINSTANCE(hmod.0);
    let cbt = unsafe { SetWindowsHookExW(WH_CBT, cbt_proc, hinst, tid) }
        .map_err(|_| ("WH_CBT", win_err()))?;
    match unsafe { SetWindowsHookExW(WH_CALLWNDPROC, cwp_proc, hinst, tid) } {
        Ok(cwp) => Ok((cbt, cwp)),
        Err(_) => {
            let err = win_err();
            let _ = unsafe { UnhookWindowsHookEx(cbt) };
            Err(("WH_CALLWNDPROC", err))
        }
    }
}

/// Attach a hook session directly onto `(pid, tid)` for child/brokered processes or late threads.
fn attach_dynamic_hook_session(
    pid: u32,
    tid: u32,
    init: HookSharedInit,
) -> Option<InProcessHookSession> {
    if pid == 0 || tid == 0 {
        cp_trace!("hook_attach: dynamic skipped pid={} tid={}", pid, tid);
        return None;
    }
    if pid_is_wow64(pid) == Some(true) {
        cp_trace!(
            "dynamic_hook: wow64_skip_64bit_hook pid={} tid={}",
            pid,
            tid
        );
        return None;
    }
    let (hmod, cbt_proc, cwp_proc) = load_hook_dll()?;
    let uncloak_msg = unsafe { RegisterWindowMessageW(HOOK_CONTROL_MSG_NAME) };
    let view = HookSharedView::create(pid, init, uncloak_msg).map(Arc::new);
    let hooks = if view.is_some() {
        match try_attach_thread_hooks_result(hmod, cbt_proc, cwp_proc, tid) {
            Ok((cbt, cwp)) => {
                cp_trace!(
                    "hook_attach: dynamic ok pid={} tid={} wow64={:?} cbt=0x{:x} cwp=0x{:x}",
                    pid,
                    tid,
                    pid_is_wow64(pid),
                    cbt.0 as usize,
                    cwp.0 as usize
                );
                Some((cbt, cwp))
            }
            Err((stage, err)) => {
                cp_trace!(
                    "hook_attach: dynamic failed pid={} tid={} wow64={:?} stage={} err={}",
                    pid,
                    tid,
                    pid_is_wow64(pid),
                    stage,
                    err
                );
                None
            }
        }
    } else {
        cp_trace!("hook_attach: dynamic shm_failed pid={} tid={}", pid, tid);
        None
    };
    match (view, hooks) {
        (Some(view), Some((cbt_hook, callwnd_hook))) => Some(InProcessHookSession {
            pid,
            view,
            hooks: vec![(tid, cbt_hook, callwnd_hook)],
            dll_module: hmod,
        }),
        _ => {
            let _ = unsafe { FreeLibrary(hmod) };
            None
        }
    }
}

const IN_PROCESS_HOOK_SHOW_CMD: u16 = SW_HIDE.0 as u16;

/// Spawn target with `CREATE_SUSPENDED | CREATE_NO_WINDOW`, initialize per-PID shared memory,
/// and attach `WH_CBT` + `WH_CALLWNDPROC` before/immediately after `ResumeThread`. Falls back to `None` session
/// if DLL is missing or hook cannot attach.
fn launch_with_hook(
    exe: &str,
    arguments: &str,
    path: &Path,
    method: CustomPathMethod,
    provisional: Option<(i32, i32, i32, i32)>,
    capture_technique: CustomCaptureTechnique,
    init: HookSharedInit,
) -> Option<(OwnedHandle, u32, u32, Option<InProcessHookSession>)> {
    let Some((hmod, cbt_proc, cwp_proc)) = load_hook_dll() else {
        let fallback_tech = if capture_technique == CustomCaptureTechnique::InProcessHook {
            CustomCaptureTechnique::InvisibleLayered
        } else {
            capture_technique
        };
        cp_trace!(
            "launch_with_hook: no_dll fallback_tech={:?} exe={}",
            fallback_tech,
            path_basename(Path::new(exe))
        );
        let (proc_handle, pid) = launch(exe, arguments, path, method, provisional, fallback_tech)?;
        return Some((proc_handle, pid, 0, None));
    };

    let application = wide_null(OsStr::new(exe));
    let mut cmdline = build_command_line(Path::new(exe), arguments, path, method);
    let cwd_wide = Path::new(exe)
        .parent()
        .filter(|dir| !dir.as_os_str().is_empty())
        .map(|dir| wide_null(dir.as_os_str()));
    let cwd_ptr = cwd_wide
        .as_ref()
        .map_or(PCWSTR::null(), |buf| PCWSTR(buf.as_ptr()));

    let mut si = launch_startup_info_for(provisional, capture_technique, virtual_screen());
    if capture_technique == CustomCaptureTechnique::InProcessHook {
        si.wShowWindow = IN_PROCESS_HOOK_SHOW_CMD;
    }
    let nul = open_inheritable_nul();
    let inherit_handles = if let Some(ref file) = nul {
        attach_nul_stdio(&mut si, HANDLE(file.as_raw_handle()));
        BOOL(1)
    } else {
        BOOL(0)
    };

    let mut pi = PROCESS_INFORMATION::default();
    let created = unsafe {
        CreateProcessW(
            PCWSTR(application.as_ptr()),
            PWSTR(cmdline.as_mut_ptr()),
            None,
            None,
            inherit_handles,
            preview_creation_flags(true),
            None,
            cwd_ptr,
            &si,
            &mut pi,
        )
    };
    drop(nul);
    if created.is_err() {
        cp_trace!(
            "launch_with_hook: CreateProcessW failed exe={} err={}",
            path_basename(Path::new(exe)),
            win_err()
        );
        let _ = unsafe { FreeLibrary(hmod) };
        return None;
    }

    // SAFETY: `pi.hProcess` is uniquely owned and closed on `OwnedHandle` drop.
    let process = unsafe { OwnedHandle::from_raw_handle(pi.hProcess.0) };
    let pid = pi.dwProcessId;
    let tid = pi.dwThreadId;
    let wow64 = process_is_wow64(HANDLE(process.as_raw_handle()));
    cp_trace!(
        "launch_with_hook: created pid={} tid={} wow64={:?} cap={:?} si_flags=0x{:x} si_show={} si_xy=({},{}) si_wh=({},{})",
        pid,
        tid,
        wow64,
        capture_technique,
        si.dwFlags.0,
        si.wShowWindow,
        si.dwX,
        si.dwY,
        si.dwXSize,
        si.dwYSize
    );
    if wow64 == Some(true) {
        cp_trace!(
            "launch_with_hook: wow64_skip_64bit_hook pid={} tid={}",
            pid,
            tid
        );
        let _ = unsafe { ResumeThread(pi.hThread) };
        let _ = unsafe { CloseHandle(pi.hThread) };
        let _ = unsafe { FreeLibrary(hmod) };
        return Some((process, pid, tid, None));
    }
    let uncloak_msg = unsafe { RegisterWindowMessageW(HOOK_CONTROL_MSG_NAME) };
    let view = HookSharedView::create(pid, init, uncloak_msg).map(Arc::new);

    // Attempt hook install before `ResumeThread`; if Win11 rejects a suspended thread before
    // `W32THREAD` init (`ERROR_INVALID_THREAD_ID`), spin-yield 30ms (sub-0.1ms attach; avoids
    // 15.6ms timer-tick sleep) then sleep(1ms) up to 200ms.
    let mut hooks = if view.is_some() {
        match try_attach_thread_hooks_result(hmod, cbt_proc, cwp_proc, tid) {
            Ok((cbt, cwp)) => {
                cp_trace!(
                    "launch_with_hook: pre_resume_hook ok pid={} tid={} cbt=0x{:x} cwp=0x{:x}",
                    pid,
                    tid,
                    cbt.0 as usize,
                    cwp.0 as usize
                );
                Some((cbt, cwp))
            }
            Err((stage, err)) => {
                cp_trace!(
                    "launch_with_hook: pre_resume_hook failed pid={} tid={} stage={} err={}",
                    pid,
                    tid,
                    stage,
                    err
                );
                None
            }
        }
    } else {
        None
    };

    let _ = unsafe { ResumeThread(pi.hThread) };
    let _ = unsafe { CloseHandle(pi.hThread) };

    if view.is_some() && hooks.is_none() {
        let retry_start = Instant::now();
        let spin_deadline = retry_start + Duration::from_millis(30);
        let retry_deadline = retry_start + Duration::from_millis(200);
        let mut attempts: u32 = 0;
        let mut last_err: Option<(&'static str, u32)> = None;
        let mut exited = false;
        while Instant::now() < retry_deadline {
            if process_has_exited(HANDLE(process.as_raw_handle())) {
                exited = true;
                break;
            }
            attempts = attempts.saturating_add(1);
            match try_attach_thread_hooks_result(hmod, cbt_proc, cwp_proc, tid) {
                Ok((cbt, cwp)) => {
                    cp_trace!(
                        "launch_with_hook: post_resume_hook ok pid={} tid={} attempts={} elapsed_ms={} cbt=0x{:x} cwp=0x{:x}",
                        pid,
                        tid,
                        attempts,
                        retry_start.elapsed().as_millis(),
                        cbt.0 as usize,
                        cwp.0 as usize
                    );
                    hooks = Some((cbt, cwp));
                    break;
                }
                Err(e) => {
                    last_err = Some(e);
                }
            }
            if Instant::now() < spin_deadline {
                std::thread::yield_now();
            } else {
                std::thread::sleep(Duration::from_millis(1));
            }
        }
        if hooks.is_none() {
            cp_trace!(
                "launch_with_hook: post_resume_hook failed pid={} tid={} attempts={} elapsed_ms={} exited={} last_err={:?}",
                pid,
                tid,
                attempts,
                retry_start.elapsed().as_millis(),
                exited,
                last_err
            );
        }
    }

    let session = match (view, hooks) {
        (Some(view), Some((cbt_hook, callwnd_hook))) => Some(InProcessHookSession {
            pid,
            view,
            hooks: vec![(tid, cbt_hook, callwnd_hook)],
            dll_module: hmod,
        }),
        _ => {
            let _ = unsafe { FreeLibrary(hmod) };
            None
        }
    };
    Some((process, pid, tid, session))
}

/// Validate a `captured_hwnd` candidate: live, non-blocked owner, non-toolwindow, and titled or >= 64x64.
fn validate_captured_hwnd(hwnd: HWND) -> Option<(i64, bool, HWND)> {
    if !unsafe { IsWindow(hwnd) }.as_bool() {
        return None;
    }
    if owner_is_blocking(hwnd) {
        return None;
    }
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    if (ex_style & WS_EX_TOOLWINDOW.0) != 0 {
        return None;
    }
    let mut rect = RECT::default();
    let size = if unsafe { GetWindowRect(hwnd, &mut rect) }.is_ok() {
        WindowSize::of(rect)
    } else {
        WindowSize {
            area: 0,
            a_strip: true,
        }
    };
    let titled = unsafe { GetWindowTextLengthW(hwnd) } > 0;
    let visible = unsafe { IsWindowVisible(hwnd) }.as_bool();
    // The rule a census candidate is held to, and the letter of it: a titled window that has never
    // been seen and is a strip is a tooltip rather than a picture (see `THE_SMALLEST_SIDE`).
    if !visible && size.a_strip {
        return None;
    }
    if titled || size.area >= THE_SMALLEST_PICTURE {
        Some((size.area, titled, hwnd))
    } else {
        None
    }
}

fn select_hook_or_census_candidate(
    best: Option<(CandidateRank, HWND)>,
    valid_captured: Option<(i64, bool, HWND)>,
    first_candidate: &mut Option<(Instant, HWND)>,
    hidden_settle: Duration,
    the_latch_has_no_say: bool,
) -> Option<HWND> {
    // **The hook's latched handle is a guess at the picture, not the picture.** It is whatever the
    // DLL saw first — the first window whose `WM_WINDOWPOSCHANGING` carried `SWP_SHOWWINDOW`, or one
    // 64 pixels across with a caption — which for a program that makes a helper before its document
    // is the helper. Where the preset asks for the window made last and the census found candidates
    // of its own, the census is what answers; the latch is left for a look that found nothing it
    // could classify.
    let valid_captured = if the_latch_has_no_say {
        None
    } else {
        valid_captured
    };
    if let Some((rank, best_hwnd)) = best {
        if rank.visible {
            if let Some((cap_area, true, cap_hwnd)) = valid_captured {
                if cap_area > rank.area {
                    Some(cap_hwnd)
                } else {
                    Some(best_hwnd)
                }
            } else {
                Some(best_hwnd)
            }
        } else if let Some((cap_area, true, cap_hwnd)) = valid_captured {
            if rank.area > cap_area {
                Some(best_hwnd)
            } else if cap_area >= THE_SMALLEST_PICTURE {
                Some(cap_hwnd)
            } else {
                None
            }
        } else {
            let entry = first_candidate.get_or_insert_with(|| (Instant::now(), best_hwnd));
            entry.1 = best_hwnd;
            (entry.0.elapsed() >= hidden_settle).then_some(best_hwnd)
        }
    } else if let Some((cap_area, true, cap_hwnd)) = valid_captured {
        (cap_area >= THE_SMALLEST_PICTURE).then_some(cap_hwnd)
    } else {
        if let Some((_, _, cap_hwnd)) = valid_captured {
            first_candidate.get_or_insert_with(|| (Instant::now(), cap_hwnd));
        }
        None
    }
}

/// Fast discovery for `InProcessHook`: validates `session.captured_hwnd` alongside `census` and
/// prefers a larger titled main window in `processes` when one exists.
fn wait_for_a_window_in_process_hook(
    launch: &Launch,
    session: &InProcessHookSession,
    effective_capture: CustomCaptureTechnique,
) -> Option<DiscoveredWindow> {
    forget_sightings();
    let started = Instant::now();
    let deadline = started + WINDOW_WAIT;
    let mut processes = discovery_processes(launch);
    update_owned_pids_if_current(launch, &processes);
    let mut read = started;
    let mut images: ProcessImages = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();
    let mut first_candidate: Option<(Instant, HWND)> = None;
    let mut dumped_300ms = false;

    while Instant::now() < deadline {
        if !launch_still_wanted(launch) {
            cp_trace!(
                "wait_hook: aborted_loop id={} hits={}",
                launch.id,
                session.view.hook_hits()
            );
            return None;
        }
        let _ = refresh_job_pids(launch, &mut processes);
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
        let mut best = look.the_picture();
        let mut looked = look.has_candidates();
        if best.is_none() && broker_fallback_allowed(launch) {
            let look_fb = census(
                launch,
                &processes,
                &mut images,
                &mut creation_times,
                0,
                false,
                OwnershipMode::Discover {
                    allow_image_fallback: true,
                },
            );
            best = look_fb.the_picture();
            looked |= look_fb.has_candidates();
        }

        let found = select_hook_or_census_candidate(
            best,
            valid_captured,
            &mut first_candidate,
            Duration::from_millis(40),
            launch.picture_window.prefers_the_newest() && looked,
        );
        if let Some(window) = found {
            if !launch_still_wanted(launch) {
                cp_trace!(
                    "wait_hook: aborted_found id={} hwnd=0x{:x} hits={}",
                    launch.id,
                    window.0 as usize,
                    session.view.hook_hits()
                );
                cleanup_aborted_source(window, launch);
                return None;
            }
            cp_trace!(
                "wait_hook: found id={} hwnd=0x{:x} hits={} cap={:?} elapsed_ms={}",
                launch.id,
                window.0 as usize,
                session.view.hook_hits(),
                session.view.captured_hwnd().map(|h| h.0 as usize),
                launch.launched_at.elapsed().as_millis()
            );
            return Some(DiscoveredWindow {
                hwnd: window,
                hook_session: None,
                pre_cloaked: false,
            });
        }

        maybe_trace_300ms_no_match(
            "wait_hook",
            launch,
            started,
            &mut dumped_300ms,
            &processes,
            &mut images,
        );

        if read.elapsed() >= TREE_POLL {
            processes = discovery_processes(launch);
            update_owned_pids_if_current(launch, &processes);
            read = Instant::now();
        }
        std::thread::sleep(Duration::from_millis(5));
    }
    if !launch_still_wanted(launch) {
        cp_trace!(
            "wait_hook: aborted_timeout id={} first_candidate={:?} hits={}",
            launch.id,
            first_candidate.map(|(_, h)| h.0 as usize),
            session.view.hook_hits()
        );
        if let Some((_, hwnd)) = first_candidate {
            cleanup_aborted_source(hwnd, launch);
        }
        return None;
    }
    cp_trace!(
        "wait_hook: timeout id={} first_candidate={:?} hits={} procs={:?}",
        launch.id,
        first_candidate.map(|(_, h)| h.0 as usize),
        session.view.hook_hits(),
        processes
    );
    first_candidate
        .map(|(_, hwnd)| DiscoveredWindow {
            hwnd,
            hook_session: None,
            pre_cloaked: false,
        })
        .or_else(|| wait_for_a_window(launch, effective_capture))
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
    // One reading of the frame answers every ask below, so a window that is still being restyled
    // cannot have the ask land beside the host it was grown from.
    let chrome = chrome_now(source, window_mode);
    let ctx = placement_ctx_from_want(launch, processes, None, Some(&session.view));
    cp_trace!(
        "reveal_hook: start id={} src=0x{:x} host=0x{:x} host_rect={:?} hook_hits={} inv_layered={} forced={} cloaked={}",
        launch.id,
        source.0 as usize,
        host.0 as usize,
        host_rect,
        session.view.hook_hits(),
        *state.invisible_layered,
        ctx.forced,
        session.view.is_cloaked()
    );
    // Late-hooked fallback (`hook_hits == 0`): source is not cloaked; use InvisibleLayered if active.
    if session.view.hook_hits() == 0 && *state.invisible_layered {
        session
            .view
            .state()
            .cloak_enabled
            .store(0, Ordering::Release);
        let fallback_ctx = PlacementContext { hook: None, ..ctx };
        place_invisible_layered(
            source,
            host_rect,
            chrome,
            window_mode,
            state.the_ask,
            fallback_ctx,
        );
        if *state.pre_cloaked {
            let _ = set_window_cloak(source, false);
            *state.pre_cloaked = false;
        }
        restore_foreground(launch.previous_foreground, processes);
        *state.placed_mode = Some(PlacementMode::InvisibleLayered);
        *state.placed_host_rect = Some(host_rect);
        *state.expected_rect = window_rect(source);
        cp_trace!(
            "reveal_hook: fallback_invisible_layered id={} expected={:?}",
            launch.id,
            *state.expected_rect
        );
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
    let no_shift_ctx = PlacementContext {
        avoid_cursor: None,
        ..ctx
    };
    if ctx.forced {
        apply_hook_rect_now(
            ctx.hook,
            source,
            (STAGE_COORD, STAGE_COORD, ww, wh),
            ZORDER_BOTTOM,
        );
        force_place_window(
            source,
            (STAGE_COORD, STAGE_COORD, ww, wh),
            HWND_BOTTOM,
            SET_WINDOW_POS_FLAGS(0),
            no_shift_ctx,
        );
    } else {
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
    }
    if !unsafe { IsWindowVisible(source) }.as_bool() {
        let _ = unsafe { ShowWindow(source, SW_SHOWNOACTIVATE) };
    }

    // Check fit using off-screen dimensions so `source` never touches the monitor before `host`.
    let stage_vis = visible_frame_rect(source);
    let stage_win = window_rect(source);
    let stage_shown = unsafe { IsWindowVisible(source) }.as_bool();
    let stage_zoomed = unsafe { IsZoomed(source) }.as_bool();
    let stage_iconic = unsafe { IsIconic(source) }.as_bool();
    let fits_host = stage_vis.is_some_and(|r| {
        let aw = r.right - r.left;
        let ah = r.bottom - r.top;
        aw > 0 && ah > 0 && aw <= hw + ALIGN_TOLERANCE * 2 && ah <= hh + ALIGN_TOLERANCE * 2
    }) && stage_shown
        && !stage_zoomed
        && !stage_iconic;
    cp_trace!(
        "reveal_hook: stage_fit id={} target_behind=({},{},{},{}) fits={} stage_vis={:?} stage_win={:?} vis={} zoom={} min={}",
        launch.id,
        wx,
        wy,
        ww,
        wh,
        fits_host,
        stage_vis,
        stage_win,
        stage_shown,
        stage_zoomed,
        stage_iconic
    );

    if !fits_host {
        let ask = the_ask_for_host(source, host_rect, chrome);
        *state.the_ask = Some(ask);
        cp_trace!(
            "reveal_hook: stage_ask id={} ask={:?} hw_hh=({},{})",
            launch.id,
            ask,
            hw,
            hh
        );
        session
            .view
            .set_rect(STAGE_COORD, STAGE_COORD, ask.0, ask.1);
        if ctx.forced {
            apply_hook_rect_now(
                ctx.hook,
                source,
                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                ZORDER_BOTTOM,
            );
            force_place_window(
                source,
                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                no_shift_ctx,
            );
        } else {
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
    }

    // Size `host`, register thumbnail, and show opaque topmost `host` FIRST.
    if *state.box_of_the_host != Some(host_rect) {
        let host_ok = unsafe {
            SetWindowPos(
                host,
                HWND::default(),
                hx,
                hy,
                hw,
                hh,
                SWP_NOZORDER | SWP_NOACTIVATE,
            )
        }
        .is_ok();
        let host_err = if host_ok { 0 } else { win_err() };
        cp_trace!(
            "reveal_hook: host_pos id={} host_rect={:?} ok={} err={}",
            launch.id,
            host_rect,
            host_ok,
            host_err
        );
        *state.box_of_the_host = Some(host_rect);
    }
    if *state.registered == 0 {
        match unsafe { DwmRegisterThumbnail(host, source) } {
            Ok(registration) => {
                *state.registered = registration;
                if let Ok(mut held) = THUMBNAIL.lock() {
                    *held = registration;
                }
                cp_trace!(
                    "reveal_hook: dwm_register id={} reg=0x{:x} host=0x{:x} src=0x{:x}",
                    launch.id,
                    registration as usize,
                    host.0 as usize,
                    source.0 as usize
                );
            }
            Err(e) => {
                cp_trace!(
                    "reveal_hook: dwm_register_err id={} host=0x{:x} src=0x{:x} err={:?}",
                    launch.id,
                    host.0 as usize,
                    source.0 as usize,
                    e
                );
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
        match unsafe { DwmUpdateThumbnailProperties(*state.registered, &properties) } {
            Ok(()) => {
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
                cp_trace!(
                    "reveal_hook: dwm_update ok id={} client=({},{}) picture={:?}",
                    launch.id,
                    client.right,
                    client.bottom,
                    picture
                );
            }
            Err(e) => {
                cp_trace!(
                    "reveal_hook: dwm_update_err id={} client=({},{}) picture={:?} err={:?}",
                    launch.id,
                    client.right,
                    client.bottom,
                    picture,
                    e
                );
            }
        }
    }

    // Wait for DWM to physically composite opaque `host` on screen before moving/uncloaking `source`.
    let _ = unsafe { DwmFlush() };

    let mode = if fits_host {
        session.view.set_rect(wx, wy, ww, wh);
        if ctx.forced {
            apply_hook_rect_now(ctx.hook, source, (wx, wy, ww, wh), ZORDER_BOTTOM);
            force_place_window(
                source,
                (wx, wy, ww, wh),
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                no_shift_ctx,
            );
            session.view.uncloak_sync(source);
            if *state.pre_cloaked {
                let _ = set_window_cloak(source, false);
                *state.pre_cloaked = false;
            }
            if !aligned_behind_host(source, host_rect) {
                cp_trace!(
                    "reveal_hook: behind_host_retry id={} cur_win={:?} cur_vis={:?} host={:?}",
                    launch.id,
                    window_rect(source),
                    visible_frame_rect(source),
                    host_rect
                );
                apply_hook_rect_now(ctx.hook, source, (wx, wy, ww, wh), ZORDER_BOTTOM);
                force_place_window(
                    source,
                    (wx, wy, ww, wh),
                    HWND_BOTTOM,
                    SET_WINDOW_POS_FLAGS(0),
                    no_shift_ctx,
                );
            }
            if aligned_behind_host(source, host_rect) {
                let _ = unsafe { InvalidateRect(source, None, false) };
                let _ = unsafe { UpdateWindow(source) };
                PlacementMode::BehindHost
            } else {
                let ask = the_ask_for_host(source, host_rect, chrome);
                cp_trace!(
                    "reveal_hook: behind_host_degrade_offscreen id={} ask={:?} cur_win={:?} cur_vis={:?}",
                    launch.id,
                    ask,
                    window_rect(source),
                    visible_frame_rect(source)
                );
                *state.the_ask = Some(ask);
                session
                    .view
                    .set_rect(STAGE_COORD, STAGE_COORD, ask.0, ask.1);
                apply_hook_rect_now(
                    ctx.hook,
                    source,
                    (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                    ZORDER_BOTTOM,
                );
                force_place_window(
                    source,
                    (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                    HWND_BOTTOM,
                    SET_WINDOW_POS_FLAGS(0),
                    no_shift_ctx,
                );
                nudge_clamped_offscreen_off_cursor(
                    source,
                    Some(&session.view),
                    ctx.avoid_cursor,
                    launch.previous_foreground,
                    processes,
                );
                let _ = unsafe { InvalidateRect(source, None, false) };
                let _ = unsafe { UpdateWindow(source) };
                PlacementMode::OffScreen
            }
        } else {
            let _ = unsafe { SetWindowPos(source, HWND_BOTTOM, wx, wy, ww, wh, SWP_NOACTIVATE) };
            session.view.uncloak_sync(source);
            if *state.pre_cloaked {
                let _ = set_window_cloak(source, false);
                *state.pre_cloaked = false;
            }
            let _ = unsafe { InvalidateRect(source, None, false) };
            let _ = unsafe { UpdateWindow(source) };
            PlacementMode::BehindHost
        }
    } else {
        session.view.uncloak_sync(source);
        if *state.pre_cloaked {
            let _ = set_window_cloak(source, false);
            *state.pre_cloaked = false;
        }
        if ctx.forced && !is_window_offscreen(source) {
            let ask = state
                .the_ask
                .unwrap_or_else(|| the_ask_for_host(source, host_rect, chrome));
            *state.the_ask = Some(ask);
            session
                .view
                .set_rect(STAGE_COORD, STAGE_COORD, ask.0, ask.1);
            apply_hook_rect_now(
                ctx.hook,
                source,
                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                ZORDER_BOTTOM,
            );
            force_place_window(
                source,
                (STAGE_COORD, STAGE_COORD, ask.0, ask.1),
                HWND_BOTTOM,
                SET_WINDOW_POS_FLAGS(0),
                no_shift_ctx,
            );
            nudge_clamped_offscreen_off_cursor(
                source,
                Some(&session.view),
                ctx.avoid_cursor,
                launch.previous_foreground,
                processes,
            );
        }
        let _ = unsafe { InvalidateRect(source, None, false) };
        let _ = unsafe { UpdateWindow(source) };
        PlacementMode::OffScreen
    };

    restore_foreground(launch.previous_foreground, processes);
    if *state.pre_cloaked {
        let _ = unsafe { DwmFlush() };
        let _ = set_window_cloak(source, false);
        *state.pre_cloaked = false;
    }
    *state.placed_mode = Some(mode);
    *state.placed_host_rect = Some(host_rect);
    *state.expected_rect = window_rect(source);
    cp_trace!(
        "reveal_hook: done id={} mode={:?} fits_host={} host_rect={:?} expected={:?} vis={:?} offscreen={} hits={}",
        launch.id,
        mode,
        fits_host,
        host_rect,
        *state.expected_rect,
        visible_frame_rect(source),
        is_window_offscreen(source),
        session.view.hook_hits()
    );
}

#[allow(clippy::too_many_arguments)]
fn run_in_process_hook_settle_and_reveal(
    host: HWND,
    source: HWND,
    launch: &Launch,
    session: &InProcessHookSession,
    processes: &[u32],
    images: &mut ProcessImages,
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

    let forced = enforces_placement(launch.placement);
    let hook_mode = uses_hook_enforcement(launch.placement);
    cp_trace!(
        "settle_hook: start id={} src=0x{:x} was_hidden={} forced={} hook_mode={} hits={} cloaked={} win={:?}",
        launch.id,
        source.0 as usize,
        was_hidden,
        forced,
        hook_mode,
        session.view.hook_hits(),
        session.view.is_cloaked(),
        window_rect(source)
    );
    if was_hidden {
        let Some(box_) = current_target_rect() else {
            cp_trace!(
                "settle_hook: abort id={} reason=no_target_rect_hidden",
                launch.id
            );
            if is_current(launch.id) {
                end_the_program();
            }
            return false;
        };
        let target_rect = the_host_for(box_, chrome_now(source, window_mode));
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
        if hook_mode {
            let _ = session.view.enforce_now_sync(source);
        }
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
        cp_trace!(
            "settle_hook: degrade_uncloaked id={} inv_layered={} offscreen={}",
            launch.id,
            ok,
            is_window_offscreen(source)
        );
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
    let mut settle_reason = "deadline";

    while Instant::now() < settle_deadline {
        if !is_current(launch.id) {
            cp_trace!(
                "settle_hook: exit id={} reason=superseded elapsed_ms={}",
                launch.id,
                settle_start.elapsed().as_millis()
            );
            return false;
        }
        if !unsafe { IsWindow(source) }.as_bool() {
            cp_trace!(
                "settle_hook: exit id={} reason=window_destroyed elapsed_ms={}",
                launch.id,
                settle_start.elapsed().as_millis()
            );
            if is_current(launch.id) {
                end_the_program();
            }
            return false;
        }
        let Some(box_) = current_target_rect() else {
            cp_trace!(
                "settle_hook: exit id={} reason=want_cleared elapsed_ms={}",
                launch.id,
                settle_start.elapsed().as_millis()
            );
            if !is_current(launch.id) {
                return false;
            }
            cloak_or_park(*state.invisible_layered);
            if is_current(launch.id) {
                end_the_program();
            }
            return false;
        };
        let target_rect = the_host_for(box_, chrome_now(source, window_mode));
        let insets = if window_mode == CustomWindowMode::Normal {
            dwm_frame_insets(source)
        } else {
            (0, 0, 0, 0)
        };
        let (_, _, ww, wh) = behind_host_window_rect(target_rect, insets);
        session.view.set_rect(STAGE_COORD, STAGE_COORD, ww, wh);

        let cur = window_rect(source);
        let needs_hook_snap = hook_mode
            && !cur.is_some_and(|r| {
                r.left == STAGE_COORD
                    && r.top == STAGE_COORD
                    && (r.right - r.left) == ww
                    && (r.bottom - r.top) == wh
            });
        if cur != last_rect || (forced && unsafe { IsZoomed(source) }.as_bool()) || needs_hook_snap
        {
            if hook_mode {
                let _ = session.view.enforce_now_sync(source);
            }
            if forced {
                let _ = unsafe {
                    SetWindowPos(
                        source,
                        HWND_BOTTOM,
                        STAGE_COORD,
                        STAGE_COORD,
                        ww,
                        wh,
                        SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING,
                    )
                };
            }
            last_rect = window_rect(source);
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
            if !is_current(launch.id) {
                cp_trace!(
                    "settle_hook: exit id={} reason=superseded_wm_null elapsed_ms={}",
                    launch.id,
                    settle_start.elapsed().as_millis()
                );
                return false;
            }
            if !unsafe { IsWindow(source) }.as_bool() {
                cp_trace!(
                    "settle_hook: exit id={} reason=destroyed_wm_null elapsed_ms={}",
                    launch.id,
                    settle_start.elapsed().as_millis()
                );
                if is_current(launch.id) {
                    end_the_program();
                }
                return false;
            }
            if answered && window_rect(source) == last_rect {
                settle_reason = "stable";
                break;
            }
        }
        std::thread::sleep(Duration::from_millis(10));
    }

    cp_trace!(
        "settle_hook: settled id={} reason={} elapsed_ms={} last_rect={:?} hits={}",
        launch.id,
        settle_reason,
        settle_start.elapsed().as_millis(),
        last_rect,
        session.view.hook_hits()
    );

    if !is_current(launch.id) {
        return false;
    }
    if !unsafe { IsWindow(source) }.as_bool() {
        cp_trace!(
            "settle_hook: post_exit id={} reason=window_destroyed",
            launch.id
        );
        if is_current(launch.id) {
            end_the_program();
        }
        return false;
    }
    let Some(box_) = current_target_rect() else {
        cp_trace!(
            "settle_hook: post_exit id={} reason=no_target_rect",
            launch.id
        );
        cloak_or_park(*state.invisible_layered);
        if is_current(launch.id) {
            end_the_program();
        }
        return false;
    };
    // The host is the box grown by the frame, and the reveal places the program's window so that
    // frame lands on it (see `the_host_for` and `reveal_in_process_hook`).
    let host_rect = the_host_for(box_, chrome_now(source, window_mode));
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
fn wait_for_a_window(
    launch: &Launch,
    effective_capture: CustomCaptureTechnique,
) -> Option<DiscoveredWindow> {
    forget_sightings();
    let started = Instant::now();
    let deadline = started + WINDOW_WAIT;
    let mut processes = discovery_processes(launch);
    update_owned_pids_if_current(launch, &processes);
    let mut read = started;
    let mut images: ProcessImages = HashMap::new();
    let mut creation_times: HashMap<u32, Option<u64>> = HashMap::new();
    let mut first_hook_candidate: Option<(Instant, HWND)> = None;
    let mut dumped_300ms = false;
    let mut pre_cloak = PreCloakTracker::new(launch.requested_capture, launch.sibling_windows);
    let mut pre_hook = JobChildPreHook::new(launch, effective_capture);
    let initial_job_pids = active_job_pids(launch.id);
    pre_hook.poll(launch, &initial_job_pids);

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

    let fallback_allowed = broker_fallback_allowed;

    let check_now = |processes: &[u32],
                     images: &mut ProcessImages,
                     creation_times: &mut HashMap<u32, Option<u64>>,
                     first_hook_candidate: &mut Option<(Instant, HWND)>,
                     pre_cloak: &mut PreCloakTracker,
                     pre_hook: &JobChildPreHook|
     -> Option<HWND> {
        let valid_captured = pre_hook.captured_hwnd().and_then(validate_captured_hwnd);

        // Pass 1: ALWAYS check process_tree(launch.pid) + job PIDs first; never adopt sibling PIDs when tree/job has a window.
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
        let best = look.the_picture();
        if let Some((_, window)) = best {
            pre_cloak.try_cloak(window, pre_hook.session(), pid_of(window));
        }
        if pre_hook.session().is_some() {
            if let Some(window) = select_hook_or_census_candidate(
                best,
                valid_captured,
                first_hook_candidate,
                Duration::from_millis(40),
                launch.picture_window.prefers_the_newest() && look.has_candidates(),
            ) {
                pre_cloak.try_cloak(window, pre_hook.session(), pid_of(window));
                return Some(window);
            }
        } else if let Some((rank, window)) = best {
            if rank.visible {
                return Some(window);
            }
            wait_for_input_idle_handle(HANDLE(launch.process.as_raw_handle()), 250);
            if !launch_still_wanted(launch) {
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
            let chosen = look2.the_picture().map(|(_, w)| w).or(Some(window));
            if let Some(w) = chosen {
                pre_cloak.try_cloak(w, pre_hook.session(), pid_of(w));
            }
            return chosen;
        }

        // Pass 2: Brokered/packaged app + same-install-dir stub fallback ONLY when tree/job has
        // no window AND launcher exited or grace elapsed.
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
        let best_fb = look.the_picture();
        if let Some((_, window)) = best_fb {
            pre_cloak.try_cloak(window, pre_hook.session(), pid_of(window));
        }
        if pre_hook.session().is_some() {
            if let Some(window) = select_hook_or_census_candidate(
                best_fb,
                valid_captured,
                first_hook_candidate,
                Duration::from_millis(40),
                launch.picture_window.prefers_the_newest() && look.has_candidates(),
            ) {
                pre_cloak.try_cloak(window, pre_hook.session(), pid_of(window));
                return Some(window);
            }
            return None;
        }
        if let Some((rank, window)) = best_fb {
            if rank.visible {
                return Some(window);
            }
            wait_for_input_idle_handle(HANDLE(launch.process.as_raw_handle()), 250);
            if !launch_still_wanted(launch) {
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
            let chosen = look2.the_picture().map(|(_, w)| w).or(Some(window));
            if let Some(w) = chosen {
                pre_cloak.try_cloak(w, pre_hook.session(), pid_of(w));
            }
            return chosen;
        }
        None
    };

    let finish_found = |window: HWND,
                        pre_cloak: &mut PreCloakTracker,
                        pre_hook: &mut JobChildPreHook|
     -> DiscoveredWindow {
        let winner_pid = pid_of(window);
        let hook_session = pre_hook.take_for_pid(winner_pid);
        let pre_cloaked = pre_cloak.finish_with_winner(window, hook_session.as_ref(), winner_pid);
        DiscoveredWindow {
            hwnd: window,
            hook_session,
            pre_cloaked,
        }
    };

    while Instant::now() < deadline {
        if !launch_still_wanted(launch) {
            cp_trace!("wait_std: aborted_loop id={}", launch.id);
            return None;
        }

        let job_pids = refresh_job_pids(launch, &mut processes);
        pre_hook.poll(launch, &job_pids);

        if let Some(window) = check_now(
            &processes,
            &mut images,
            &mut creation_times,
            &mut first_hook_candidate,
            &mut pre_cloak,
            &pre_hook,
        ) {
            if !launch_still_wanted(launch) {
                cp_trace!(
                    "wait_std: aborted_after_check id={} hwnd=0x{:x}",
                    launch.id,
                    window.0 as usize
                );
                drop(pre_hook);
                drop(pre_cloak);
                cleanup_aborted_source(window, launch);
                return None;
            }
            return Some(finish_found(window, &mut pre_cloak, &mut pre_hook));
        }
        maybe_trace_300ms_no_match(
            "wait_std",
            launch,
            started,
            &mut dumped_300ms,
            &processes,
            &mut images,
        );

        if read.elapsed() >= TREE_POLL {
            processes = discovery_processes(launch);
            update_owned_pids_if_current(launch, &processes);
            read = Instant::now();
        }

        let remaining_ms = deadline
            .saturating_duration_since(Instant::now())
            .as_millis()
            .min(pre_hook.wait_cap_ms(launch)) as u32;
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
                    if !processes.contains(&pid) {
                        let job_pids = refresh_job_pids(launch, &mut processes);
                        pre_hook.poll(launch, &job_pids);
                    }
                    let belongs = discovery_match_for_pid(
                        pid,
                        launch,
                        &processes,
                        mode,
                        &mut images,
                        &mut creation_times,
                    )
                    .is_some();
                    if belongs {
                        if event == EVENT_OBJECT_CREATE || event == EVENT_OBJECT_SHOW {
                            if is_precloak_candidate(hwnd) {
                                pre_cloak.try_cloak(hwnd, pre_hook.session(), pid);
                            }
                            if event == EVENT_OBJECT_CREATE {
                                if let Some(provisional) = launch.provisional {
                                    preposition_hidden_candidate(hwnd, provisional);
                                }
                            }
                        }
                        matched = true;
                    }
                }
            });
            if matched {
                if !launch_still_wanted(launch) {
                    cp_trace!("wait_std: aborted_msg id={}", launch.id);
                    return None;
                }
                if let Some(window) = check_now(
                    &processes,
                    &mut images,
                    &mut creation_times,
                    &mut first_hook_candidate,
                    &mut pre_cloak,
                    &pre_hook,
                ) {
                    if !launch_still_wanted(launch) {
                        cp_trace!(
                            "wait_std: aborted_msg_after_check id={} hwnd=0x{:x}",
                            launch.id,
                            window.0 as usize
                        );
                        drop(pre_hook);
                        drop(pre_cloak);
                        cleanup_aborted_source(window, launch);
                        return None;
                    }
                    return Some(finish_found(window, &mut pre_cloak, &mut pre_hook));
                }
            }
        }
    }
    cp_trace!("wait_std: timeout id={} procs={:?}", launch.id, processes);
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
    process_tree_to(pid, CustomChildDepth::WholeTree)
}

/// The same walk, to the depth a preset's `Child Processes` row asks for.
///
/// **`LaunchedOnly` is the whole of the gate on discovery**: no pid past the launched one is
/// returned, so `discovery_processes`, `owned_pids_for` and the teardown sweeps — every reader of
/// this walk — are bounded by it without knowing it exists.
fn process_tree_to(pid: u32, depth: CustomChildDepth) -> Vec<u32> {
    if pid == 0 {
        return Vec::new();
    }
    if depth == CustomChildDepth::LaunchedOnly {
        return vec![pid];
    }
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
        // One level is the launched pid and what it started directly, which is the first pass of
        // this loop and nothing after it.
        if depth == CustomChildDepth::OneLevel && at >= 1 {
            break;
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

/// Full image path of `pid` (cached once per discovery/launch).
fn image_path_of(pid: u32, images: &mut ProcessImages) -> Option<PathBuf> {
    if let Some(known) = images.get(&pid) {
        return known.clone();
    }
    let path = unsafe { process_image_path(pid) };
    images.insert(pid, path.clone());
    path
}

fn image_name_from_path(path: &Path) -> Option<String> {
    path.file_name()
        .map(|name| name.to_string_lossy().to_lowercase())
}

/// The image a process is running, as its file name and lowercased, or nothing where the machine
/// will not say.
///
/// Asked once per process per launch: the answer does not change while the process lives, and
/// every window on the desktop is asked about on every look.
fn image_name_of(pid: u32, images: &mut ProcessImages) -> Option<String> {
    image_path_of(pid, images)
        .as_deref()
        .and_then(image_name_from_path)
}

/// Safety: opens the process for the query alone and closes it again; nothing is dereferenced.
unsafe fn process_image_path(pid: u32) -> Option<PathBuf> {
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
    (!path.is_empty()).then(|| PathBuf::from(path))
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
        // with. No chrome: this is the `Borderless` window, and the two halves of `the_ask_for` are
        // then the identity.
        let asked = the_ask_for((80, 60), (704, 500), (0, 0, 0, 0));
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
        let asked = the_ask_for((490, 1380), (704, 1380), (0, 0, 0, 0));
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

        // **And the frame a program draws around its picture is counted.** MPV Player's window takes
        // a 16x32 frame over its client, so a 1200x675 box is a window asked for 1216x707 — which is
        // the whole of the black-bar fix: the *client* is the box's shape, so the player draws no
        // bars in it (see `chrome_of`).
        let chrome = (8, 31, 16, 32);
        let asked = the_ask_for((1200, 675), (1200, 675), chrome);
        assert_eq!(
            asked,
            (1216, 707),
            "a framed window is asked for its client plus its frame"
        );
        assert_eq!(
            (asked.0 - chrome.2, asked.1 - chrome.3),
            (1200, 675),
            "and the client inside that ask is the box exactly"
        );

        // A window that took the box is asked for nothing more, and one that will not say what it
        // took leaves the box it was asked for.
        assert_eq!(
            the_ask_for((600, 400), (600, 400), (0, 0, 0, 0)),
            (600, 400)
        );
        assert_eq!(the_ask_for((600, 400), (0, 0), (0, 0, 0, 0)), (600, 400));
    }

    #[test]
    fn the_ask_is_the_host_rect_so_the_client_lands_on_the_box_and_the_picture_on_the_host() {
        // MPV Player's frame as measured on the running app: the client stands 1 px in and 31 px
        // down from the visible frame, which is 2 x 32 larger than it — and Win11's invisible shadow
        // border puts another 7 px round the window rect, so all three rects are different.
        let chrome = (1, 31, 2, 32);
        let insets = (7, 0, 7, 7);
        let box_ = (300, 200, 1200, 675);

        let host = the_host_for(box_, chrome);
        let asked = the_ask_for((box_.2, box_.3), (1200, 675), chrome);
        assert_eq!(
            (asked.0, asked.1),
            (host.2, host.3),
            "the ask is the host rect: the frame the picture is composed from"
        );

        // Put the window where the ask says, and read back the three rects a placement sees.
        let (wx, wy, ww, wh) = behind_host_window_rect((host.0, host.1, asked.0, asked.1), insets);
        let visible = (
            wx + insets.0,
            wy + insets.1,
            ww - insets.0 - insets.2,
            wh - insets.1 - insets.3,
        );
        assert_eq!(
            visible, host,
            "the window's visible frame lands on the host, which is what a thumbnail composes"
        );
        let client = (
            visible.0 + chrome.0,
            visible.1 + chrome.1,
            visible.2 - chrome.2,
            visible.3 - chrome.3,
        );
        assert_eq!(
            client, box_,
            "and its client is the box exactly, which is the shape the player needs"
        );

        // The picture is that visible frame and the host's client is the whole of it: drawn 1:1,
        // with no letterbox for either the player or the host to draw.
        let picture = (visible.2, visible.3);
        let destination = the_picture_in(host.2, host.3, picture);
        assert_eq!(
            (
                destination.left,
                destination.top,
                destination.right,
                destination.bottom
            ),
            (0, 0, host.2, host.3),
            "the whole of the host's client is the picture, and none of the background shows"
        );

        // The box a host rect was grown from is read back out of it, which is how a placement handed
        // the host rect asks for the box it stands on.
        assert_eq!(the_box_of(host, chrome), box_);
        assert_eq!(
            the_box_of(host, (0, 0, 0, 0)),
            host,
            "and no frame leaves it alone"
        );
    }

    #[test]
    fn the_ask_is_a_fixed_point_once_the_window_has_taken_it() {
        // The reported hover: a 2406x1353 box on a 1920x1080 film. The ask is the visible frame, so
        // the client inside it is the box — and asking again with that client answers the same
        // thing. Taking the *window* rect for the client instead is what made the ask grow by its own
        // frame on every poll, leaving the player's client wider than the box.
        let chrome = (1, 31, 2, 32);
        let box_ = (2406, 1353);
        let first = the_ask_for(box_, (1936, 1119), chrome);
        assert_eq!(
            first,
            (2408, 1385),
            "the box plus the frame, and nothing else"
        );

        let client = (first.0 - chrome.2, first.1 - chrome.3);
        assert_eq!(client, box_, "the client inside the ask is the box");
        let again = the_ask_for(box_, (client.0 as u32, client.1 as u32), chrome);
        assert_eq!(
            again, first,
            "so the next poll asks for exactly the same window"
        );
    }

    #[test]
    fn the_host_is_the_box_grown_by_the_chrome_and_the_client_lands_on_the_box() {
        // A player's frame: the client stands 8 px in and 31 px down from the visible frame, and the
        // frame is 16 x 32 larger than the client — the numbers measured on MPV Player.
        let chrome = (8, 31, 16, 32);
        let box_ = (300, 200, 1200, 675);
        let host = the_host_for(box_, chrome);
        assert_eq!(
            host,
            (292, 169, 1216, 707),
            "the host is grown up and left by the chrome, and by it in size"
        );
        assert_eq!(
            (host.0 + chrome.0, host.1 + chrome.1),
            (box_.0, box_.1),
            "so the client's own corner lands on the box's corner"
        );
        assert_eq!(
            (host.2 - chrome.2, host.3 - chrome.3),
            (box_.2, box_.3),
            "and the client is the box exactly, which is the shape the player needs"
        );

        // A `Borderless` program has no frame, so the host is the box and nothing moves.
        assert_eq!(the_host_for(box_, (0, 0, 0, 0)), box_);

        // `chrome_of` is the other half: where the client stands inside the frame, and how much of it
        // the frame takes.
        let visible = RECT {
            left: 300,
            top: 200,
            right: 1500,
            bottom: 875,
        };
        let origin = POINT { x: 308, y: 231 };
        assert_eq!(
            chrome_of(visible, origin, (1184, 636)),
            (8, 31, 16, 39),
            "the client's inset, and the difference between the frame and the client"
        );
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
        let decode = |exe: &str, args: &str, path: &str| {
            let wide = build_command_line(
                Path::new(exe),
                args,
                Path::new(path),
                CustomPathMethod::Plain,
            );
            assert_eq!(wide.last(), Some(&0), "must be NUL-terminated");
            String::from_utf16_lossy(&wide[..wide.len() - 1])
        };

        assert_eq!(
            decode("viewer.exe", "", "photo.png"),
            r#""viewer.exe" "photo.png""#
        );
        assert_eq!(
            decode(
                r"Program Files\My Viewer\app.exe",
                "",
                r"my folder\pic 1.png"
            ),
            r#""Program Files\My Viewer\app.exe" "my folder\pic 1.png""#
        );
        assert_eq!(
            decode("app.exe", "", r#"say "hi".txt"#),
            r#""app.exe" "say \"hi\".txt""#
        );
        assert_eq!(
            decode("app.exe", "", r#"dir\"quoted".png"#),
            r#""app.exe" "dir\\\"quoted\".png""#
        );
        assert_eq!(
            decode("app.exe", "", r"trailing\folder\\"),
            r#""app.exe" "trailing\folder\\\\""#
        );
        assert_eq!(
            decode("viewer.exe", "--fullscreen --silent", "photo.png"),
            r#""viewer.exe" "--fullscreen" "--silent" "photo.png""#
        );
    }

    #[test]
    fn command_line_round_trips_through_win32_command_line_to_argv_w() {
        use windows::Win32::Foundation::{LocalFree, HLOCAL};
        use windows::Win32::UI::Shell::CommandLineToArgvW;

        fn parse_with_win32(wide_nul: &[u16]) -> Vec<String> {
            assert_eq!(wide_nul.last(), Some(&0), "must be NUL-terminated");
            let mut argc = 0i32;
            let argv = unsafe { CommandLineToArgvW(PCWSTR(wide_nul.as_ptr()), &mut argc) };
            assert!(!argv.is_null(), "CommandLineToArgvW failed");
            let mut out = Vec::with_capacity(argc.max(0) as usize);
            unsafe {
                for i in 0..(argc.max(0) as usize) {
                    let pwstr = *argv.add(i);
                    let mut len = 0usize;
                    while *pwstr.0.add(len) != 0 {
                        len += 1;
                    }
                    let slice = std::slice::from_raw_parts(pwstr.0, len);
                    out.push(String::from_utf16_lossy(slice));
                }
                let _ = LocalFree(HLOCAL(argv.cast()));
            }
            out
        }

        let cases: [(&str, &str, &str, &[&str]); 6] = [
            (
                "viewer.exe",
                "",
                r"samples\pic.png",
                &["viewer.exe", r"samples\pic.png"],
            ),
            (
                r"Program Files\Viewer\app.exe",
                "--fullscreen --silent",
                r"media\my video.mp4",
                &[
                    r"Program Files\Viewer\app.exe",
                    "--fullscreen",
                    "--silent",
                    r"media\my video.mp4",
                ],
            ),
            (
                "app.exe",
                r#"--title="My Preview Title" --zoom=100%"#,
                "a b.png",
                &[
                    "app.exe",
                    "--title=My Preview Title",
                    "--zoom=100%",
                    "a b.png",
                ],
            ),
            (
                "app.exe",
                r#"--dir=sub\dir\file "trailing\dir\\" "\"escaped\"""#,
                "pic.jpg",
                &[
                    "app.exe",
                    r"--dir=sub\dir\file",
                    r"trailing\dir\",
                    "\"escaped\"",
                    "pic.jpg",
                ],
            ),
            (
                "app.exe",
                r#"--input={file} --mode=fast"#,
                r"folder with spaces\clip 1.mkv",
                &[
                    "app.exe",
                    r"--input=folder with spaces\clip 1.mkv",
                    "--mode=fast",
                ],
            ),
            (
                "app.exe",
                r#""" --empty='""' --pair={file}:{file}"#,
                "shot.png",
                &["app.exe", "", r#"--empty="""#, "--pair=shot.png:shot.png"],
            ),
        ];

        for (exe, arguments, file, expected) in cases {
            let wide = build_command_line(
                Path::new(exe),
                arguments,
                Path::new(file),
                CustomPathMethod::Plain,
            );
            let actual = parse_with_win32(&wide);
            assert_eq!(
                actual.iter().map(String::as_str).collect::<Vec<_>>(),
                expected,
                "CommandLineToArgvW mismatch for args=`{arguments}` file=`{file}`"
            );
        }
    }

    #[test]
    fn preview_creation_flags_always_include_create_no_window() {
        let normal = preview_creation_flags(false);
        assert_ne!(normal.0 & CREATE_NO_WINDOW.0, 0);
        assert_eq!(normal.0 & CREATE_SUSPENDED.0, 0);

        let suspended = preview_creation_flags(true);
        assert_ne!(suspended.0 & CREATE_NO_WINDOW.0, 0);
        assert_ne!(suspended.0 & CREATE_SUSPENDED.0, 0);
    }

    #[test]
    fn nul_stdio_redirection_sets_usestdhandles_and_handles() {
        let nul = open_inheritable_nul().expect("opening NUL must succeed on Windows");
        let nul_handle = HANDLE(nul.as_raw_handle());
        let mut si = launch_startup_info(Some((10, 20, 300, 200)));
        attach_nul_stdio(&mut si, nul_handle);
        assert!(si.dwFlags.contains(STARTF_USESTDHANDLES));
        assert!(si.dwFlags.contains(STARTF_USESHOWWINDOW));
        assert_eq!(si.hStdInput, nul_handle);
        assert_eq!(si.hStdOutput, nul_handle);
        assert_eq!(si.hStdError, nul_handle);
    }

    #[test]
    fn should_restore_foreground_on_teardown_respects_dying_pids() {
        let prev = 0x1234isize;
        let other = 0x5678isize;
        let dying = [4001u32, 4002];

        // Dying process held foreground before teardown -> restore even if Windows moved fg to another window.
        assert!(should_restore_foreground_on_teardown(
            prev, 0x9999, 4001, other, 9999, &dying
        ));
        // Dying process still holds foreground after teardown -> restore.
        assert!(should_restore_foreground_on_teardown(
            prev, prev, 100, 0x9999, 4002, &dying
        ));
        // Already back on previous foreground -> no-op.
        assert!(!should_restore_foreground_on_teardown(
            prev, 0x9999, 4001, prev, 100, &dying
        ));
        // User had switched to unrelated app before teardown -> do not steal focus.
        assert!(!should_restore_foreground_on_teardown(
            prev, other, 9999, other, 9999, &dying
        ));
        // Zero previous handle -> no-op.
        assert!(!should_restore_foreground_on_teardown(
            0, 0x9999, 4001, other, 9999, &dying
        ));
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
        let candidates = crate::paths::hook_dll_candidates(dir, Some(user));
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
        assert_eq!(crate::paths::hook_dll_candidates(dir, None).len(), 2);
    }

    #[test]
    fn in_process_hook_dll_exports_shm_layout_and_thread_hook_attachment() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        // 1. Companion DLL must exist beside the test binary (or parent target dir) and export both hooks.
        let (hmod, cbt_proc, cwp_proc) = load_hook_dll().expect(
            "custom_preview_hook.dll must be built and export rhp_cbt_proc + rhp_callwnd_proc",
        );

        // 2. Verify shared memory creation and exact 56-byte C field layout.
        let pid = std::process::id();
        let msg = unsafe { RegisterWindowMessageW(HOOK_CONTROL_MSG_NAME) };
        assert_ne!(msg, 0);
        let view = HookSharedView::create(
            pid,
            HookSharedInit {
                rect: (STAGE_COORD, STAGE_COORD, 640, 480),
                cloak_enabled: true,
                rect_locked: true,
                forced_mode: 1,
                z_order_mode: ZORDER_BOTTOM,
                captured_hwnd: 0,
            },
            msg,
        )
        .expect("HookSharedView::create must succeed");
        assert_eq!(std::mem::size_of::<HookSharedState>(), 56);
        assert_eq!(std::mem::offset_of!(HookSharedState, forced_placement), 52);
        assert_eq!(std::mem::offset_of!(HookSharedState, z_order_mode), 53);
        assert_eq!(view.state().magic, HOOK_MAGIC);
        assert_eq!(view.state().version, HOOK_VERSION);
        assert_eq!(view.state().target_pid, pid);
        assert_eq!(view.state().forced_placement.load(Ordering::Acquire), 1);
        assert_eq!(
            view.state().z_order_mode.load(Ordering::Acquire),
            ZORDER_BOTTOM
        );
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
                preview_creation_flags(true),
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

        // CMD_DETACH clears cloak_enabled + forced_placement + rect_locked and removes the subclass hook.
        assert!(
            view.detach_sync(main_hwnd),
            "detach_sync must deliver CMD_DETACH synchronously and increment hook_hits"
        );
        assert_eq!(view.state().cloak_enabled.load(Ordering::Acquire), 0);
        assert_eq!(view.state().forced_placement.load(Ordering::Acquire), 0);
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

        // Test CustomPlacement::Hook (forced_placement = 2, ZORDER_TOPMOST, invisible owner,
        // WM_GETMINMAXINFO override, SWP_NOSENDCHANGING re-snap via WM_WINDOWPOSCHANGED, and CMD_DETACH).
        let hook_view = HookSharedView::create(
            pid,
            HookSharedInit {
                rect: (150, 110, 260, 180),
                cloak_enabled: false,
                rect_locked: true,
                forced_mode: 2,
                z_order_mode: ZORDER_TOPMOST,
                captured_hwnd: 0,
            },
            msg,
        )
        .expect("re-init HookSharedView in Hook mode");
        let hook_hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!("Hook Mode Test"),
                WS_CAPTION
                    | WS_MAXIMIZE
                    | WS_THICKFRAME
                    | windows::Win32::UI::WindowsAndMessaging::WS_VISIBLE,
                10,
                10,
                900,
                700,
                helper_hwnd,
                None,
                None,
                None,
            )
            .expect("create invisible-owner Hook Mode Test window")
        };
        assert!(
            !unsafe { IsZoomed(hook_hwnd) }.as_bool(),
            "Hook HCBT_CREATEWND must strip WS_MAXIMIZE"
        );
        let created_rect = window_rect(hook_hwnd).expect("created hook_hwnd rect");
        assert_eq!(
            (
                created_rect.left,
                created_rect.top,
                created_rect.right - created_rect.left,
                created_rect.bottom - created_rect.top,
            ),
            (150, 110, 260, 180),
            "Hook mode must lock creation origin and size even with invisible owner"
        );

        hook_view.set_rect(170, 130, 110, 90);
        assert!(hook_view.enforce_now_sync(hook_hwnd));
        let enforced_rect = window_rect(hook_hwnd).expect("enforced hook_hwnd rect");
        assert_eq!(
            (
                enforced_rect.left,
                enforced_rect.top,
                enforced_rect.right - enforced_rect.left,
                enforced_rect.bottom - enforced_rect.top,
            ),
            (170, 130, 110, 90),
            "CMD_ENFORCE_NOW + WM_GETMINMAXINFO override must shrink below default track floor"
        );
        let ex_after_enforce = unsafe { GetWindowLongPtrW(hook_hwnd, GWL_EXSTYLE) } as u32;
        assert_ne!(
            ex_after_enforce & WS_EX_TOPMOST.0,
            0,
            "ZORDER_TOPMOST must elevate window for DirectWindow"
        );

        let _ = unsafe {
            SetWindowPos(
                hook_hwnd,
                HWND_BOTTOM,
                40,
                40,
                500,
                400,
                SWP_NOSENDCHANGING | SWP_NOACTIVATE,
            )
        };
        let snapped_back = window_rect(hook_hwnd).expect("snapped_back hook_hwnd rect");
        assert_eq!(
            (
                snapped_back.left,
                snapped_back.top,
                snapped_back.right - snapped_back.left,
                snapped_back.bottom - snapped_back.top,
            ),
            (170, 130, 110, 90),
            "WM_WINDOWPOSCHANGED must snap back even when SWP_NOSENDCHANGING bypasses CHANGING"
        );

        // Cursor-shift in force_place_window must update SHM BEFORE SetWindowPos(SWP_NOSENDCHANGING)
        // so WM_WINDOWPOSCHANGED never pulls the shifted window back onto the cursor.
        let cursor_on_hook = POINT { x: 200, y: 160 };
        force_place_window(
            hook_hwnd,
            (170, 130, 110, 90),
            HWND_TOPMOST,
            SET_WINDOW_POS_FLAGS(0),
            PlacementContext {
                forced: true,
                hook: Some(&hook_view),
                avoid_cursor: Some(cursor_on_hook),
                previous_foreground: 0,
                processes: &[],
            },
        );
        let shifted_hook = window_rect(hook_hwnd).expect("shifted hook_hwnd rect");
        assert!(
            !point_in_rect(cursor_on_hook, shifted_hook),
            "Hook window must stay off-cursor without WM_WINDOWPOSCHANGED snap-back"
        );
        assert_eq!(
            hook_view.state().x.load(Ordering::Acquire),
            shifted_hook.left
        );
        assert_eq!(
            hook_view.state().y.load(Ordering::Acquire),
            shifted_hook.top
        );

        // enforce_now_sync must never re-lock SHM while DETACHING is active.
        hook_view.clear_locks_for_detach();
        DETACHING.store(true, Ordering::SeqCst);
        let detach_guard = DetachGuard;
        assert!(!hook_view.enforce_now_sync(hook_hwnd));
        assert_eq!(hook_view.state().rect_locked.load(Ordering::Acquire), 0);
        drop(detach_guard);

        assert!(hook_view.detach_sync(hook_hwnd));
        let _ = unsafe {
            SetWindowPos(
                hook_hwnd,
                HWND_NOTOPMOST,
                220,
                150,
                320,
                220,
                SWP_NOACTIVATE,
            )
        };
        let freed_rect = window_rect(hook_hwnd).expect("freed hook_hwnd rect");
        assert_eq!(
            (
                freed_rect.left,
                freed_rect.top,
                freed_rect.right - freed_rect.left,
                freed_rect.bottom - freed_rect.top,
            ),
            (220, 150, 320, 220),
            "CMD_DETACH must restore full move and resize freedom"
        );
        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(hook_hwnd) };
        drop(hook_view);
        drop(view);

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
            arguments: String::new(),
            previous_foreground: 0,
            process: Arc::new(unsafe { OwnedHandle::from_raw_handle(handle.0) }),
            pid: 0,
            job: None,
            processes_before: Vec::new(),
            child_scope: CustomChildScope::ChildrenAndGrandchildren,
        }
    }

    #[test]
    fn active_path_tracks_want_and_program_and_clears_on_hide() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        let path = PathBuf::from("sample.custom");

        end_the_program();
        assert_eq!(active_path(), None);
        assert!(!is_active());
        assert!(!has_live_preview());

        if let Ok(mut p) = PROGRAM.lock() {
            *p = Some(dummy_program(path.clone()));
        }
        // Program without WANT must not report active.
        assert_eq!(active_path(), None);
        assert!(!is_active());

        if let Ok(mut w) = WANT.lock() {
            *w = Some(Want {
                at: pointer((100, 100)),
                placement: CustomPlacement::FollowTray,
            });
        }
        LIVE.store(true, Ordering::Release);
        assert_eq!(active_path().as_deref(), Some(path.as_path()));
        assert!(is_active());
        assert!(has_live_preview());

        // Pin-press routing must emit PinCold for the active custom preview path.
        assert!(matches!(
            pin_request_if_active(),
            Some(super::super::PreviewMessage::PinCold { path: routed }) if routed == path
        ));

        hide();
        assert_eq!(active_path(), None);
        assert!(!is_active());
        assert!(!has_live_preview());
        assert!(WANT.lock().unwrap().is_none());
        assert!(PROGRAM.lock().unwrap().is_none());
        assert!(OWNED_PIDS.lock().unwrap().is_empty());
        assert!(pin_request_if_active().is_none());
    }

    #[test]
    fn hide_while_detaching_preserves_promoted_state_and_owned_pid_cleanup_excludes_preexisting() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        end_the_program();

        let path = PathBuf::from("detach-guard.custom");
        if let Ok(mut p) = PROGRAM.lock() {
            *p = Some(dummy_program(path.clone()));
        }
        if let Ok(mut w) = WANT.lock() {
            *w = Some(Want {
                at: pointer((120, 120)),
                placement: CustomPlacement::FollowTray,
            });
        }
        LIVE.store(true, Ordering::Release);
        if let Ok(mut owned) = OWNED_PIDS.lock() {
            *owned = vec![4001, 4002];
        }

        // While DETACHING is true, hide() must return immediately without touching state.
        DETACHING.store(true, Ordering::SeqCst);
        assert!(!has_live_preview());
        hide();
        assert_eq!(active_path().as_deref(), Some(path.as_path()));
        assert!(PROGRAM.lock().unwrap().is_some());
        assert_eq!(*OWNED_PIDS.lock().unwrap(), vec![4001, 4002]);
        DETACHING.store(false, Ordering::SeqCst);
        assert!(has_live_preview());

        // Once DETACHING clears, hide() tears down immediately.
        hide();
        assert!(!has_live_preview());
        assert!(PROGRAM.lock().unwrap().is_none());
        assert!(OWNED_PIDS.lock().unwrap().is_empty());

        // Owned PID merge must include launched/job/new source trees + tracked PIDs, and never pre-existing or own PID.
        let merged_new = merge_owned_pids_to_end(OwnedPidInputs {
            launched_pid: 100,
            launched_tree: &[100, 101, 500, 999],
            job_pids: &[150, 151, 500, 999],
            source_pid: 200,
            source_is_new: true,
            source_tree: &[200, 201, 500],
            tracked_owned: &[300, 500, 999],
            processes_before: &[500, 100],
            my_pid: 999,
        });
        assert_eq!(merged_new, vec![100, 101, 150, 151, 200, 201, 300]);

        // Reused broker (!source_is_new) must NEVER terminate source_pid or its tree.
        let merged_reused = merge_owned_pids_to_end(OwnedPidInputs {
            launched_pid: 100,
            launched_tree: &[100, 101],
            job_pids: &[150],
            source_pid: 500,
            source_is_new: false,
            source_tree: &[500, 501],
            tracked_owned: &[300],
            processes_before: &[500],
            my_pid: 999,
        });
        assert_eq!(merged_reused, vec![100, 101, 150, 300]);
    }

    #[test]
    fn custom_preview_rect_and_visibility_helpers_filter_staging_and_offscreen() {
        let _guard = TEST_STATE_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        end_the_program();

        let vscreen = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        assert!(point_in_xywh(
            POINT { x: 100, y: 100 },
            (100, 100, 200, 150)
        ));
        assert!(!point_in_xywh(
            POINT { x: 300, y: 100 },
            (100, 100, 200, 150)
        ));
        assert!(!point_in_xywh(
            POINT { x: 150, y: 250 },
            (100, 100, 200, 150)
        ));
        assert!(!point_in_xywh(POINT { x: 100, y: 100 }, (100, 100, 0, 150)));

        let onscreen = RECT {
            left: 200,
            top: 150,
            right: 500,
            bottom: 450,
        };
        assert!(point_in_rect(POINT { x: 200, y: 150 }, onscreen));
        assert!(!point_in_rect(POINT { x: 500, y: 450 }, onscreen));
        assert!(is_rect_visible_on_screen(onscreen, vscreen));

        let staged = stage_rect(400, 300);
        assert!(!is_rect_visible_on_screen(staged, vscreen));
        assert!(!is_rect_visible_on_screen(
            RECT {
                left: 2000,
                top: 100,
                right: 2400,
                bottom: 400,
            },
            vscreen
        ));
        assert!(!is_rect_visible_on_screen(
            RECT {
                left: 100,
                top: 100,
                right: 100,
                bottom: 200,
            },
            vscreen
        ));

        // cursor_is_over_preview matches visible HOST rect (SHOWN + LAST_HOST_RECT) and ignores invisible/staged SOURCE.
        let bounds = work_area_at(0, 0);
        let hx = bounds.left + 300;
        let hy = bounds.top + 200;
        let hidden_hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(INVISIBLE_EX_STYLE),
                w!("STATIC"),
                w!("Hidden Staging"),
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
            .expect("create hidden staging window")
        };
        SOURCE.store(hidden_hwnd.0 as isize, Ordering::SeqCst);
        SOURCE_INVISIBLE_LAYERED.store(true, Ordering::SeqCst);
        assert!(!window_is_visible_on_screen(hidden_hwnd));

        LIVE.store(true, Ordering::Release);
        SHOWN.store(true, Ordering::SeqCst);
        if let Ok(mut r) = LAST_HOST_RECT.lock() {
            *r = Some((hx, hy, 400, 300));
        }
        assert!(cursor_is_over_preview(POINT {
            x: hx + 10,
            y: hy + 10
        }));
        assert!(!cursor_is_over_preview(POINT {
            x: hx - 5,
            y: hy + 10
        }));
        // Staging coordinates must not trigger cursor_is_over_preview.
        assert!(!cursor_is_over_preview(POINT {
            x: STAGE_COORD + 10,
            y: STAGE_COORD + 10
        }));

        // During FORCED_SETTLING, Check 1 (visible HOST rect) STILL triggers immediately, while
        // SOURCE / OWNED_PIDS (Checks 2 & 3) are deferred until settling finishes.
        // Place on_screen_hwnd disjoint from LAST_HOST_RECT ([hx..hx+400, hy..hy+300]) inside bounds.
        let sx = bounds.left + 40;
        let sy = hy;
        let on_screen_pt = POINT {
            x: sx + 80,
            y: sy + 50,
        };
        let on_screen_hwnd = unsafe {
            windows::Win32::UI::WindowsAndMessaging::CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!("OnScreen"),
                WS_CAPTION | windows::Win32::UI::WindowsAndMessaging::WS_VISIBLE,
                sx,
                sy,
                220,
                160,
                None,
                None,
                None,
                None,
            )
            .expect("create visible on-screen window")
        };
        SOURCE.store(on_screen_hwnd.0 as isize, Ordering::SeqCst);
        SOURCE_INVISIBLE_LAYERED.store(false, Ordering::SeqCst);
        FORCED_SETTLING.store(true, Ordering::Release);
        assert!(is_forced_settling());
        assert!(cursor_is_over_preview(POINT {
            x: hx + 10,
            y: hy + 10
        }));
        assert!(!cursor_is_over_preview(on_screen_pt));
        FORCED_SETTLING.store(false, Ordering::Release);
        assert!(!is_forced_settling());
        assert!(cursor_is_over_preview(on_screen_pt));

        end_the_program();
        assert!(!FORCED_SETTLING.load(Ordering::Acquire));
        assert!(!cursor_is_over_preview(POINT {
            x: hx + 10,
            y: hy + 10
        }));
        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(on_screen_hwnd) };
        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(hidden_hwnd) };
    }

    #[test]
    fn forced_placement_min_size_cursor_shift_keeps_window_off_cursor() {
        let work = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1040,
        };
        let cursor = POINT { x: 500, y: 400 };

        // Already clear of cursor -> returns None.
        let clear = RECT {
            left: 520,
            top: 200,
            right: 920,
            bottom: 500,
        };
        assert_eq!(shift_rect_off_cursor(clear, cursor, work, 12), None);

        // App enforced wide minimum width overlapping cursor from the left -> shifts cleanly off cursor.
        let overlapping = RECT {
            left: 200,
            top: 300,
            right: 750,
            bottom: 700,
        };
        let (sx, sy) = shift_rect_off_cursor(overlapping, cursor, work, 12)
            .expect("overlapping min-size rect must shift away from cursor");
        let shifted = RECT {
            left: sx,
            top: sy,
            right: sx + (overlapping.right - overlapping.left),
            bottom: sy + (overlapping.bottom - overlapping.top),
        };
        assert!(!point_in_rect(cursor, shifted));
        assert!(shifted.left >= work.left && shifted.right <= work.right);
        assert!(shifted.top >= work.top && shifted.bottom <= work.bottom);

        // Work-area-filling minimum size -> cannot move off cursor, returns None gracefully.
        let fullscreen = work;
        assert_eq!(shift_rect_off_cursor(fullscreen, cursor, work, 12), None);
    }

    #[test]
    fn forced_direct_window_syncs_last_host_rect_and_offscreen_avoids_stage_oscillation() {
        let host_rect = (300, 200, 360, 240);
        // Unshifted rect preserves exact host_rect; shifted rect adopts actual placed coords.
        assert_eq!(
            shifted_direct_host_rect(
                host_rect,
                RECT {
                    left: 300,
                    top: 200,
                    right: 660,
                    bottom: 440,
                },
                host_rect,
            ),
            host_rect
        );
        assert_eq!(
            shifted_direct_host_rect(
                host_rect,
                RECT {
                    left: 720,
                    top: 200,
                    right: 1120,
                    bottom: 500,
                },
                host_rect,
            ),
            (720, 200, 400, 300)
        );

        // End-to-end forced DirectWindow overlapping cursor shifts off-cursor and updates LAST_HOST_RECT.
        let bounds = work_area_at(0, 0);
        let hx = bounds.left + 300;
        let hy = bounds.top + 200;
        let test_host = (hx, hy, 320, 220);
        let cursor = POINT {
            x: hx + 80,
            y: hy + 80,
        };
        let hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!("DirectShift"),
                WS_POPUP,
                hx,
                hy,
                320,
                220,
                None,
                None,
                None,
                None,
            )
            .expect("create DirectWindow test hwnd")
        };
        if let Ok(mut slot) = LAST_HOST_RECT.lock() {
            *slot = None;
        }
        place_direct_window(
            hwnd,
            test_host,
            CustomWindowMode::Borderless,
            PlacementContext {
                forced: true,
                hook: None,
                avoid_cursor: Some(cursor),
                previous_foreground: 0,
                processes: &[],
            },
        );
        let actual = window_rect(hwnd).expect("actual DirectWindow rect");
        assert!(!point_in_rect(cursor, actual));
        assert_eq!(
            LAST_HOST_RECT.lock().unwrap().take(),
            Some((
                actual.left,
                actual.top,
                (actual.right - actual.left).max(1),
                (actual.bottom - actual.top).max(1),
            ))
        );
        let _ = unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(hwnd) };

        // HookSharedView::clear_locks_for_detach immediately zeroes cloak, forced_placement, and rect_locked.
        let msg = unsafe { RegisterWindowMessageW(w!("RhpDetachLocksUnitTest")) };
        let view = HookSharedView::create(
            0xDEAD_BEEF,
            HookSharedInit {
                rect: (STAGE_COORD, STAGE_COORD, 320, 240),
                cloak_enabled: true,
                rect_locked: true,
                forced_mode: 2,
                z_order_mode: ZORDER_TOPMOST,
                captured_hwnd: 0,
            },
            msg,
        )
        .expect("create HookSharedView");
        assert_eq!(view.state().cloak_enabled.load(Ordering::Acquire), 1);
        assert_eq!(view.state().forced_placement.load(Ordering::Acquire), 2);
        assert_eq!(
            view.state().z_order_mode.load(Ordering::Acquire),
            ZORDER_TOPMOST
        );
        assert_eq!(view.state().rect_locked.load(Ordering::Acquire), 1);
        view.clear_locks_for_detach();
        assert_eq!(view.state().cloak_enabled.load(Ordering::Acquire), 0);
        assert_eq!(view.state().forced_placement.load(Ordering::Acquire), 0);
        assert_eq!(view.state().rect_locked.load(Ordering::Acquire), 0);

        // Verify Hook placement helpers across capture techniques and non-media provisional fallback.
        assert_eq!(
            z_order_mode_for(CustomCaptureTechnique::DirectWindow),
            ZORDER_TOPMOST
        );
        for tech in [
            CustomCaptureTechnique::BehindHost,
            CustomCaptureTechnique::InvisibleLayered,
            CustomCaptureTechnique::OffScreen,
            CustomCaptureTechnique::OffscreenStaged,
            CustomCaptureTechnique::InProcessHook,
        ] {
            assert_eq!(z_order_mode_for(tech), ZORDER_BOTTOM);
        }

        let target = (240, 160, 420, 300);
        let direct_init = hook_init_for(
            CustomCaptureTechnique::DirectWindow,
            CustomPlacement::Hook,
            Some(target),
        );
        assert_eq!(direct_init.rect, target);
        assert!(!direct_init.cloak_enabled);
        assert_eq!(direct_init.forced_mode, 2);
        assert_eq!(direct_init.z_order_mode, ZORDER_TOPMOST);

        let staged_init = hook_init_for(
            CustomCaptureTechnique::OffscreenStaged,
            CustomPlacement::Hook,
            Some(target),
        );
        assert_eq!(staged_init.rect, (STAGE_COORD, STAGE_COORD, 420, 300));
        assert!(!staged_init.cloak_enabled);
        assert_eq!(staged_init.forced_mode, 2);
        assert_eq!(staged_init.z_order_mode, ZORDER_BOTTOM);

        let in_proc_init = hook_init_for(
            CustomCaptureTechnique::InProcessHook,
            CustomPlacement::Hook,
            Some(target),
        );
        assert_eq!(in_proc_init.rect, (STAGE_COORD, STAGE_COORD, 420, 300));
        assert!(in_proc_init.cloak_enabled);
        assert_eq!(in_proc_init.forced_mode, 2);
        assert_eq!(in_proc_init.z_order_mode, ZORDER_BOTTOM);

        let fallback_rect = hook_provisional_rect(
            pointer((bounds.left + 100, bounds.top + 100)),
            Path::new("document.unknown_custom_ext"),
            CustomScaling::FollowTray,
            None,
        )
        .expect("Hook placement must compute concrete non-zero fallback rect");
        assert!(fallback_rect.2 > 0 && fallback_rect.3 > 0);

        // Offscreen clamp guard avoids 8ms STAGE_COORD oscillation when already settled off-cursor behind host.
        let vscreen = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let offscreen_parked = RECT {
            left: STAGE_COORD,
            top: STAGE_COORD,
            right: STAGE_COORD + 400,
            bottom: STAGE_COORD + 300,
        };
        let clamped = RECT {
            left: 100,
            top: 100,
            right: 900,
            bottom: 700,
        };
        let off_cursor = Some(POINT { x: 50, y: 50 });
        assert!(!should_repark_offscreen_state(
            false,
            Some(offscreen_parked),
            Some(offscreen_parked),
            vscreen,
            true,
            off_cursor,
            true,
        ));
        assert!(!should_repark_offscreen_state(
            false,
            Some(clamped),
            Some(clamped),
            vscreen,
            true,
            off_cursor,
            true,
        ));
        // Drifted rect, zoomed/hidden, or uncloaked+topmost must still repark.
        assert!(should_repark_offscreen_state(
            false,
            Some(RECT {
                left: 140,
                ..clamped
            }),
            Some(clamped),
            vscreen,
            true,
            off_cursor,
            true,
        ));
        assert!(should_repark_offscreen_state(
            true,
            Some(clamped),
            Some(clamped),
            vscreen,
            true,
            off_cursor,
            true,
        ));
        assert!(should_repark_offscreen_state(
            false,
            Some(clamped),
            Some(clamped),
            vscreen,
            true,
            off_cursor,
            false,
        ));
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
                placement: CustomPlacement::FollowTray,
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
            "--silent",
            CustomPathMethod::Plain,
            |p| {
                assert_eq!(p, path.as_path());
                detached = true;
                true
            },
            |_, _, _, _| spawned = true,
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
            "--fullscreen --silent",
            CustomPathMethod::FileUrl,
            |_| {
                detached = true;
                false
            },
            |exe, args, p, method| {
                assert_eq!(exe, "viewer.exe");
                assert_eq!(args, "--fullscreen --silent");
                assert_eq!(p, path.as_path());
                assert_eq!(method, CustomPathMethod::FileUrl);
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
                placement: CustomPlacement::FollowTray,
            });
        }
        SHOWN.store(true, Ordering::SeqCst);
        let mut detached = false;
        let mut spawned = false;
        assert!(dispatch_open_fully(
            CustomPromotionTechnique::OpenAsNewProcess,
            &path,
            "viewer.exe",
            "--windowed",
            CustomPathMethod::Plain,
            |_| {
                detached = true;
                true
            },
            |exe, args, p, _| {
                assert_eq!(exe, "viewer.exe");
                assert_eq!(args, "--windowed");
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

    #[test]
    fn trace_env_flag_parsing_and_write_truncates_once_then_appends() {
        assert!(!trace::env_flag_enabled(None));
        assert!(!trace::env_flag_enabled(Some(OsStr::new(""))));
        assert!(!trace::env_flag_enabled(Some(OsStr::new("0"))));
        assert!(!trace::env_flag_enabled(Some(OsStr::new("false"))));
        assert!(!trace::env_flag_enabled(Some(OsStr::new("FALSE"))));
        assert!(trace::env_flag_enabled(Some(OsStr::new("1"))));
        assert!(trace::env_flag_enabled(Some(OsStr::new("true"))));

        let unique = format!(
            "rhp_cp_trace_test_{}_{}.log",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos()
        );
        let log = std::env::temp_dir().join(unique);
        std::fs::write(&log, "stale-previous-session\n").unwrap();

        let started = Once::new();
        trace::write_with_once(&started, &log, "first line");
        trace::write_with_once(&started, &log, "second line");

        let content = std::fs::read_to_string(&log).unwrap();
        let _ = std::fs::remove_file(&log);

        assert!(!content.contains("stale-previous-session"));
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 2);
        assert!(lines[0].ends_with(" first line"));
        assert!(lines[1].ends_with(" second line"));
    }

    /// The walk the `Child Processes` row stands for: the launched pid, one level, or the tree.
    #[test]
    fn child_scope_depth_bounds_the_process_walk() {
        // Nothing is walked for a pid of zero, at any depth.
        for depth in [
            CustomChildDepth::LaunchedOnly,
            CustomChildDepth::OneLevel,
            CustomChildDepth::WholeTree,
        ] {
            assert!(process_tree_to(0, depth).is_empty());
        }

        let me = std::process::id();
        // `LaunchedOnly` is the whole of the gate: the answer is the pid asked about and nothing
        // else, whatever the machine says is below it.
        assert_eq!(
            process_tree_to(me, CustomChildDepth::LaunchedOnly),
            vec![me]
        );
        // One level is the launched pid first, exactly as the whole tree begins, so a caller that
        // reads the head of either gets the pid it started from.
        let one = process_tree_to(me, CustomChildDepth::OneLevel);
        let whole = process_tree_to(me, CustomChildDepth::WholeTree);
        assert_eq!(one.first(), Some(&me));
        assert_eq!(whole.first(), Some(&me));
        assert!(
            one.len() <= whole.len(),
            "{one:?} must not exceed {whole:?}"
        );
        assert_eq!(process_tree(me), whole);

        // `process_tree_to` is monotone in depth, which is what lets the row be a widening.
        let scope_depths = [
            CustomChildScope::Disabled.depth(),
            CustomChildScope::Children.depth(),
            CustomChildScope::ChildrenAndGrandchildren.depth(),
        ];
        assert_eq!(
            scope_depths,
            [
                CustomChildDepth::LaunchedOnly,
                CustomChildDepth::OneLevel,
                CustomChildDepth::WholeTree,
            ]
        );
    }

    #[test]
    fn tiered_candidate_classification_and_ranking() {
        // Owner blocking logic: only visible owners larger than 1x1 block candidates.
        assert!(!owner_rect_is_blocking(false, false, None));
        assert!(!owner_rect_is_blocking(true, false, Some((800, 600))));
        assert!(!owner_rect_is_blocking(true, true, Some((0, 0))));
        assert!(!owner_rect_is_blocking(true, true, Some((1, 1))));
        assert!(!owner_rect_is_blocking(true, true, Some((1, 200))));
        assert!(!owner_rect_is_blocking(true, true, Some((200, 1))));
        assert!(!owner_rect_is_blocking(true, true, None));
        assert!(owner_rect_is_blocking(true, true, Some((2, 2))));
        assert!(owner_rect_is_blocking(true, true, Some((800, 600))));
        assert!(!owner_is_blocking(HWND::default()));

        // Tier classification rules, asked in the shape a window was found in: VLC's own two
        // windows are the cases this is written against.
        let size = |w: i32, h: i32| WindowSize {
            area: i64::from(w) * i64::from(h),
            a_strip: w.min(h) < THE_SMALLEST_SIDE,
        };
        let picture = size(640, 480);
        let tool_area = size(160, 120);

        assert_eq!(
            classify_candidate(true, false, false, false, false, picture),
            Some(CandidateTier::Main)
        );
        assert_eq!(
            classify_candidate(true, true, false, false, false, picture),
            Some(CandidateTier::HiddenOwner)
        );

        // **A window never seen visible is held to a picture's shape.** VLC puts up a titled,
        // unowned, non-tool `Static` at 136x39 (`VLC ghk 3.0.24`) before its media window exists, and
        // adopting it was the white box: at 5304 px² it cleared 64x64 and lost to nothing.
        assert_eq!(
            classify_candidate(true, false, false, false, false, size(136, 39)),
            None,
            "VLC's own `VLC ghk 3.0.24` static is a strip, not a picture"
        );
        // The tooltip clears the *old* floor (it is 5304 px2, and 64x64 is 4096), which is the whole
        // of why it used to win; what stops it now is its 39-pixel height.
        let tooltip = size(136, 39);
        assert!(tooltip.area >= THE_SMALLEST_PICTURE);
        assert!(tooltip.a_strip);
        // The windows that already worked are untouched: mpv's first hidden window is 136x100.
        assert_eq!(
            classify_candidate(true, false, false, false, false, size(136, 100)),
            Some(CandidateTier::Main)
        );
        // Seen visible, a small window is a small picture rather than no picture at all.
        assert_eq!(
            classify_candidate(true, false, false, false, true, size(136, 39)),
            Some(CandidateTier::Main)
        );
        assert_eq!(
            classify_candidate(false, false, false, false, true, size(64, 64)),
            Some(CandidateTier::UntitledVisible)
        );
        assert_eq!(
            classify_candidate(false, true, false, false, true, size(64, 64)),
            Some(CandidateTier::UntitledVisible)
        );
        assert_eq!(
            classify_candidate(false, false, false, false, false, size(640, 480)),
            None,
            "hidden untitled window must be rejected"
        );
        assert_eq!(
            classify_candidate(false, false, false, true, true, size(159, 119)),
            None,
            "small toolwindow below 160x120 must be rejected"
        );
        assert_eq!(
            classify_candidate(true, true, false, true, true, tool_area),
            Some(CandidateTier::ToolWindow)
        );
        assert_eq!(
            classify_candidate(true, false, false, true, false, tool_area),
            None,
            "hidden toolwindow must be rejected"
        );
        assert_eq!(
            classify_candidate(true, true, true, false, true, size(1920, 1080)),
            None,
            "dialog/popup owned by visible non-tiny window must be rejected across all tiers"
        );
        assert_eq!(
            classify_candidate(false, true, true, true, true, size(1920, 1080)),
            None
        );

        // Priority ranking: Tier 1 > Tier 2 > Tier 3 > Tier 4; within tier: visible > hidden; then area.
        let ranked = vec![
            (
                CandidateRank {
                    tier: CandidateTier::ToolWindow,
                    visible: true,
                    area: 1_000_000,
                },
                "tier4",
            ),
            (
                CandidateRank {
                    tier: CandidateTier::UntitledVisible,
                    visible: true,
                    area: 800_000,
                },
                "tier3",
            ),
            (
                CandidateRank {
                    tier: CandidateTier::HiddenOwner,
                    visible: false,
                    area: 500_000,
                },
                "tier2_hidden_large",
            ),
            (
                CandidateRank {
                    tier: CandidateTier::HiddenOwner,
                    visible: true,
                    area: 100_000,
                },
                "tier2_visible_small",
            ),
        ];
        assert_eq!(
            select_best_candidate(&ranked).map(|(_, id)| id),
            Some("tier2_visible_small")
        );
        assert_eq!(
            select_best_candidate(&ranked[..3]).map(|(_, id)| id),
            Some("tier2_hidden_large")
        );
        assert_eq!(
            select_best_candidate(&ranked[..2]).map(|(_, id)| id),
            Some("tier3")
        );
        assert_eq!(
            select_best_candidate(&ranked[..1]).map(|(_, id)| id),
            Some("tier4")
        );

        let with_main_hidden = vec![
            ranked[3],
            (
                CandidateRank {
                    tier: CandidateTier::Main,
                    visible: false,
                    area: 50_000,
                },
                "tier1_hidden",
            ),
        ];
        assert_eq!(
            select_best_candidate(&with_main_hidden).map(|(_, id)| id),
            Some("tier1_hidden")
        );
    }

    /// The row that answers MPV Player: a helper window born first and 1.4 % larger must not win when
    /// the preset asks for the window the program made last.
    ///
    /// The two handles stand for `mpv smtc` (1976x1112, created first) and the film's own window
    /// (1936x1119, created second): the same tier, both unseen, and the helper the larger of the two —
    /// which is exactly the case measured on `C:\Program Files\MPV Player\mpv.exe`.
    #[test]
    fn the_newest_comparable_picture_is_the_film_and_not_the_helper_before_it() {
        let helper = HWND(0x10 as *mut _);
        let film = HWND(0x20 as *mut _);
        let toast = HWND(0x30 as *mut _);

        let rank = |area: i64| CandidateRank {
            tier: CandidateTier::Main,
            visible: false,
            area,
        };
        let candidates = vec![
            (rank(1976 * 1112), helper),
            (rank(1936 * 1119), film),
            // A splash of a different size class: comparable on tier, not on size, so it never
            // competes on creation order however late it arrives.
            (rank(320 * 200), toast),
        ];

        // A record where the helper was seen first and the film second, long enough ago that the
        // program has stopped making windows.
        let now = Instant::now();
        let settled = Sightings {
            first_seen: HashMap::from([
                (helper.0 as isize, now - Duration::from_millis(700)),
                (film.0 as isize, now - Duration::from_millis(400)),
                (toast.0 as isize, now - Duration::from_millis(100)),
            ]),
            last_new: Some(now - Duration::from_millis(400)),
        };

        assert_eq!(
            select_the_picture(&candidates, &settled, CustomPictureWindow::Largest)
                .map(|(_, hwnd)| hwnd),
            Some(helper),
            "`Largest` is the ranking every preset had, and it is what MPV Player fell over"
        );
        assert_eq!(
            CustomPictureWindow::default(),
            CustomPictureWindow::NewestComparable,
            "the window made last is what a preset reads with no key"
        );
        assert_eq!(
            select_the_picture(&candidates, &settled, CustomPictureWindow::NewestComparable)
                .map(|(_, hwnd)| hwnd),
            Some(film),
            "the film's window is the one made last of the same size class"
        );

        // **A program still making windows has not said which is its last.** With a window born
        // 20 ms ago, the hidden picture waits instead of adopting the helper.
        let unsettled = Sightings {
            first_seen: settled.first_seen.clone(),
            last_new: Some(Instant::now()),
        };
        assert_eq!(
            select_the_picture(
                &candidates,
                &unsettled,
                CustomPictureWindow::NewestComparable
            ),
            None,
            "a hidden winner waits for the program to stop making windows"
        );

        // **A lone hidden candidate is still a guess**, and the case this row exists for has exactly
        // one candidate for the first 200 ms: MPV Player's helper stands alone until the film's
        // window arrives. So one is enough to make the wait necessary.
        let alone = vec![(rank(1976 * 1112), helper)];
        assert_eq!(
            select_the_picture(&alone, &unsettled, CustomPictureWindow::NewestComparable),
            None,
            "one hidden candidate is not yet the program's last window"
        );
        assert_eq!(
            select_the_picture(&alone, &settled, CustomPictureWindow::NewestComparable)
                .map(|(_, hwnd)| hwnd),
            Some(helper)
        );

        // A window the program *shows* is its own answer and is taken at once, settle or not.
        let visible = vec![
            (
                CandidateRank {
                    visible: true,
                    ..rank(1976 * 1112)
                },
                helper,
            ),
            (
                CandidateRank {
                    visible: true,
                    ..rank(1936 * 1119)
                },
                film,
            ),
        ];
        assert_eq!(
            select_the_picture(&visible, &unsettled, CustomPictureWindow::NewestComparable)
                .map(|(_, hwnd)| hwnd),
            Some(film)
        );

        // The size class is the largest candidate's, so a splash well under it is not a candidate
        // for "the last window the program made" even when it arrived last of all.
        let only_big = vec![(rank(1976 * 1112), helper), (rank(1936 * 1119), film)];
        let late_toast = Sightings {
            first_seen: HashMap::from([
                (helper.0 as isize, now - Duration::from_millis(700)),
                (film.0 as isize, now - Duration::from_millis(400)),
                (toast.0 as isize, now - Duration::from_millis(10)),
            ]),
            last_new: Some(now - Duration::from_millis(400)),
        };
        assert_eq!(
            select_the_picture(
                &only_big,
                &late_toast,
                CustomPictureWindow::NewestComparable
            )
            .map(|(_, hwnd)| hwnd),
            Some(film)
        );

        // Nothing at all is still nothing.
        assert_eq!(
            select_the_picture(&[], &settled, CustomPictureWindow::NewestComparable),
            None
        );
    }

    #[test]
    fn safe_app_install_dir_rejects_shallow_and_generic_containers() {
        // Shallow paths (< 3 normal components beyond root/prefix) are rejected.
        assert!(!is_safe_app_install_dir(Path::new(r"\")));
        assert!(!is_safe_app_install_dir(Path::new(r"\myapp")));
        assert!(!is_safe_app_install_dir(Path::new(r"\apps\myapp")));
        assert!(!is_safe_app_install_dir(Path::new(r"apps\myapp")));
        assert!(!is_safe_app_install_dir(Path::new(r".\apps\vendor\myapp")));
        assert!(!is_safe_app_install_dir(Path::new(
            r"\apps\vendor\..\myapp"
        )));

        // Generic/system/shared containers rejected case-insensitively regardless of depth.
        for generic in [
            "windows",
            "System32",
            "syswow64",
            "Program Files",
            "program files (x86)",
            "ProgramData",
            "Local",
            "LocalLow",
            "Roaming",
            "Programs",
            "Desktop",
            "Downloads",
            "Documents",
            "Temp",
            "tmp",
            "Users",
            "bin",
            "WinApps",
            "WindowsApps",
        ] {
            let candidate = format!(r"\root\container\{generic}");
            assert!(
                !is_safe_app_install_dir(Path::new(&candidate)),
                "generic container {generic} must be rejected"
            );
        }

        // Dedicated application folders (>= 3 normal components, non-generic leaf) accepted.
        assert!(is_safe_app_install_dir(Path::new(r"\apps\vendor\MyApp")));
        assert!(is_safe_app_install_dir(Path::new(
            r"\Program Files\Vendor\MyApp"
        )));
        assert!(is_safe_app_install_dir(Path::new(
            r"\AppData\Local\Programs\MyApp"
        )));
        assert!(!is_safe_app_install_dir(Path::new(
            r"\\server\share\apps\MyApp"
        )));
        assert!(is_safe_app_install_dir(Path::new(
            r"\\server\share\apps\vendor\MyApp"
        )));
    }

    #[test]
    fn exe_is_inside_install_dir_checks_components_case_insensitively_and_rejects_prefix_traps() {
        let dir = Path::new(r"\apps\vendor\SomeApp");

        // Direct and nested Squirrel/Velopack subdirectories accepted case-insensitively.
        assert!(exe_is_inside_install_dir(
            Path::new(r"\apps\vendor\SomeApp\worker.exe"),
            dir
        ));
        assert!(exe_is_inside_install_dir(
            Path::new(r"\APPS\Vendor\someapp\current\SubApp.EXE"),
            dir
        ));
        assert!(exe_is_inside_install_dir(
            Path::new("/apps/vendor/someapp/app-1.2.3/RealApp.exe"),
            dir
        ));

        // Sibling prefix trap MUST be rejected (`SomeApp2` != `SomeApp`).
        assert!(!exe_is_inside_install_dir(
            Path::new(r"\apps\vendor\SomeApp2\other.exe"),
            dir
        ));
        assert!(!exe_is_inside_install_dir(
            Path::new(r"\apps\vendor\SomeApp-updater\updater.exe"),
            dir
        ));

        // Directory itself, parent traversal, and unrelated directories rejected.
        assert!(!exe_is_inside_install_dir(dir, dir));
        assert!(!exe_is_inside_install_dir(
            Path::new(r"\apps\vendor\SomeApp\..\OtherApp\app.exe"),
            dir
        ));
        assert!(!exe_is_inside_install_dir(
            Path::new(r"\apps\other\app.exe"),
            dir
        ));

        // Verbatim UNC vs standard UNC normalizes identically.
        assert!(exe_is_inside_install_dir(
            Path::new(r"\\?\UNC\server\share\apps\SomeApp\current\app.exe"),
            Path::new(r"\\server\share\apps\someapp"),
        ));

        // `full_path_matches_install_dir` strictly requires newly spawned PID and `Some(install_dir)`.
        assert!(full_path_matches_install_dir(
            true,
            Some(dir),
            Some(Path::new(r"\apps\vendor\someapp\current\app.exe")),
        ));
        assert!(!full_path_matches_install_dir(
            false,
            Some(dir),
            Some(Path::new(r"\apps\vendor\someapp\current\app.exe")),
        ));
        assert!(!full_path_matches_install_dir(
            true,
            None,
            Some(Path::new(r"\apps\vendor\someapp\current\app.exe")),
        ));
        assert!(!full_path_matches_install_dir(
            true,
            Some(dir),
            Some(Path::new(r"\apps\vendor\SomeApp2\app.exe")),
        ));
    }

    #[test]
    fn preview_job_process_ids_queries_job_and_filters_pid_list_buffer() {
        // Default/invalid job returns empty cleanly.
        let invalid = PreviewJob {
            handle: HANDLE::default(),
        };
        assert!(invalid.process_ids().is_empty());

        // Fresh unassigned job returns empty cleanly.
        let empty_handle =
            unsafe { CreateJobObjectW(None, PCWSTR::null()) }.expect("create empty job object");
        let empty_job = PreviewJob {
            handle: empty_handle,
        };
        assert!(empty_job.process_ids().is_empty());

        // Assigned child process appears in process_ids(); 0 and self PID are excluded.
        let mut child = engine_processes::hidden_command("cmd.exe")
            .args(["/C", "ping -n 6 127.0.0.1 >nul"])
            .spawn()
            .expect("spawn child for PreviewJob::process_ids test");
        let child_pid = child.id();
        let job = PreviewJob::create_and_assign(child_pid).expect("assign child to job");
        let pids = job.process_ids();
        assert!(
            pids.contains(&child_pid),
            "job pids {pids:?} must include {child_pid}"
        );
        assert!(!pids.contains(&0));
        assert!(!pids.contains(&std::process::id()));
        drop(job);
        let _ = child.kill();
        let _ = child.wait();

        // Buffer parser filters 0, self PID, duplicates, out-of-u32 entries, and clamps to buffer len.
        let header_size = std::mem::size_of::<JOBOBJECT_BASIC_PROCESS_ID_LIST>();
        let ptr_size = std::mem::size_of::<usize>();
        let header_usizes = header_size.div_ceil(ptr_size);
        let mut buf = vec![0usize; header_usizes + 6];
        let list_ptr = std::ptr::addr_of!(
            unsafe { &*(buf.as_ptr() as *const JOBOBJECT_BASIC_PROCESS_ID_LIST) }.ProcessIdList
        ) as usize;
        let offset_usizes = (list_ptr - buf.as_ptr() as usize) / ptr_size;
        {
            let list = unsafe { &mut *(buf.as_mut_ptr() as *mut JOBOBJECT_BASIC_PROCESS_ID_LIST) };
            list.NumberOfAssignedProcesses = 10;
            list.NumberOfProcessIdsInList = 999; // deliberately larger than buffer slots
        }
        let slots = &mut buf[offset_usizes..];
        slots[0] = 0;
        slots[1] = 777; // my_pid
        slots[2] = 4201;
        slots[3] = 4201; // duplicate
        slots[4] = 4202;
        slots[5] = usize::MAX; // non-u32 on 64-bit; or 0xFFFF_FFFF on 32-bit
        let parsed = parse_job_process_id_list(&buf, 777);
        let expected = if usize::BITS > 32 {
            vec![4201, 4202]
        } else {
            vec![4201, 4202, u32::MAX]
        };
        assert_eq!(parsed, expected);
    }

    #[test]
    fn threads_of_process_precloak_and_hook_candidate_selection() {
        assert!(threads_of_process(0).is_empty());
        assert_eq!(threads_of_process_via_nt(0), Some(Vec::new()));
        let my_tid = unsafe { windows::Win32::System::Threading::GetCurrentThreadId() };
        let nt_tids = threads_of_process_via_nt(std::process::id())
            .expect("NtGetNextThread must enumerate threads of current process");
        assert!(
            nt_tids.contains(&my_tid),
            "threads_of_process_via_nt(self) {nt_tids:?} must include current thread {my_tid}"
        );
        let my_tids = threads_of_process(std::process::id());
        assert!(
            my_tids.contains(&my_tid),
            "threads_of_process(self) {my_tids:?} must include current thread {my_tid}"
        );

        // Today's behaviour is `CloakAndHide`, which cloaks everything but `DirectWindow`.
        let direct_tracker = PreCloakTracker::new(
            CustomCaptureTechnique::DirectWindow,
            CustomSiblingWindows::CloakAndHide,
        );
        assert!(!direct_tracker.enabled);
        let hook_tracker = PreCloakTracker::new(
            CustomCaptureTechnique::InProcessHook,
            CustomSiblingWindows::CloakAndHide,
        );
        assert!(hook_tracker.enabled);
        let layered_tracker = PreCloakTracker::new(
            CustomCaptureTechnique::InvisibleLayered,
            CustomSiblingWindows::CloakAndHide,
        );
        assert!(layered_tracker.enabled);

        // The preset's own row: nothing is cloaked at all where it says so, and `Cloak` alone is
        // still enough to cloak.
        assert!(
            !PreCloakTracker::new(
                CustomCaptureTechnique::InProcessHook,
                CustomSiblingWindows::Disabled
            )
            .enabled
        );
        assert!(
            PreCloakTracker::new(
                CustomCaptureTechnique::InProcessHook,
                CustomSiblingWindows::Cloak
            )
            .enabled
        );

        let dummy_launch_with =
            |requested_capture: CustomCaptureTechnique, child_scope: CustomChildScope| -> Launch {
                let handle = unsafe {
                    OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, std::process::id())
                        .expect("open current process query handle")
                };
                Launch {
                    path: PathBuf::from("sample.custom"),
                    initial_own: None,
                    provisional: None,
                    packaged_broker: false,
                    process: Arc::new(unsafe { OwnedHandle::from_raw_handle(handle.0) }),
                    pid: 1,
                    tid: 1,
                    exe: "stub.exe".to_owned(),
                    previous_foreground: 0,
                    windows_before: Vec::new(),
                    processes_before: Vec::new(),
                    launch_filetime: 0,
                    launched_at: Instant::now(),
                    placement: CustomPlacement::FollowTray,
                    requested_capture,
                    child_scope,
                    broker_fallback: false,
                    sibling_windows: CustomSiblingWindows::CloakAndHide,
                    picture_window: CustomPictureWindow::Largest,
                    install_dir: None,
                    id: 1,
                }
            };
        // A launch as the tests below have always read it: the walk on, which is where the two
        // rows were absent and the behaviour was today's.
        let dummy_launch = |requested_capture: CustomCaptureTechnique| -> Launch {
            dummy_launch_with(
                requested_capture,
                CustomChildScope::ChildrenAndGrandchildren,
            )
        };
        let dummy_launch_with_scope =
            |requested_capture: CustomCaptureTechnique, scope: CustomChildScope| -> Launch {
                dummy_launch_with(requested_capture, scope)
            };

        // Wait cap shrinks from 5ms to 1ms for the first 150ms when waiting for a 64-bit job child.
        let offscreen_launch = dummy_launch(CustomCaptureTechnique::OffScreen);
        let offscreen_pre =
            JobChildPreHook::new(&offscreen_launch, CustomCaptureTechnique::OffScreen);
        assert!(!offscreen_pre.enabled);
        assert_eq!(offscreen_pre.wait_cap_ms(&offscreen_launch), 5);
        let hook_launch = dummy_launch(CustomCaptureTechnique::InProcessHook);
        let pre = JobChildPreHook::new(&hook_launch, CustomCaptureTechnique::InvisibleLayered);
        assert!(pre.enabled);
        assert_eq!(pre.wait_cap_ms(&hook_launch), 1);

        // `Child Processes = Disabled` is the gate on writing this app's DLL into a pid it never
        // began, so the pre-attach does not happen however the capture method is set.
        let no_children_launch = dummy_launch_with_scope(
            CustomCaptureTechnique::InProcessHook,
            CustomChildScope::Disabled,
        );
        assert!(
            !JobChildPreHook::new(&no_children_launch, CustomCaptureTechnique::InProcessHook)
                .enabled
        );
        // ... and a hook-enforced placement does not route around it either.
        let mut hook_placement_launch = dummy_launch_with_scope(
            CustomCaptureTechnique::OffScreen,
            CustomChildScope::Disabled,
        );
        hook_placement_launch.placement = CustomPlacement::Hook;
        assert!(
            !JobChildPreHook::new(&hook_placement_launch, CustomCaptureTechnique::OffScreen)
                .enabled
        );
        // A preset that asked for children at all keeps the pre-attach.
        let children_launch = dummy_launch_with_scope(
            CustomCaptureTechnique::InProcessHook,
            CustomChildScope::Children,
        );
        assert!(
            JobChildPreHook::new(&children_launch, CustomCaptureTechnique::InProcessHook).enabled
        );

        // `select_hook_or_census_candidate` immediately adopts a titled captured HWND >= 64x64
        // when no larger visible Census window is present.
        let cap_hwnd = HWND(0x1000 as *mut _);
        let census_hwnd = HWND(0x2000 as *mut _);
        let mut first_cand = None;
        assert_eq!(
            select_hook_or_census_candidate(
                None,
                Some((THE_SMALLEST_PICTURE, true, cap_hwnd)),
                &mut first_cand,
                Duration::from_millis(40),
                false,
            ),
            Some(cap_hwnd)
        );
        assert_eq!(
            select_hook_or_census_candidate(
                Some((
                    CandidateRank {
                        tier: CandidateTier::Main,
                        visible: true,
                        area: THE_SMALLEST_PICTURE + 1000,
                    },
                    census_hwnd,
                )),
                Some((THE_SMALLEST_PICTURE, true, cap_hwnd)),
                &mut first_cand,
                Duration::from_millis(40),
                false,
            ),
            Some(census_hwnd)
        );

        // **Where the preset asks for the window made last, the latch does not get a say.** The
        // hook's handle is whatever the DLL saw first, so a census that has candidates wins even
        // where the latch is the larger of the two — the case MPV Player's `mpv smtc` helper is.
        let hidden_census = Some((
            CandidateRank {
                tier: CandidateTier::Main,
                visible: false,
                area: THE_SMALLEST_PICTURE + 1000,
            },
            census_hwnd,
        ));
        let bigger_latch = Some((THE_SMALLEST_PICTURE + 5000, true, cap_hwnd));
        // A hidden winner still settles for 40 ms, which is the wait this branch has always had.
        let mut first_cand = None;
        assert_eq!(
            select_hook_or_census_candidate(
                hidden_census,
                bigger_latch,
                &mut first_cand,
                Duration::from_millis(40),
                true,
            ),
            None
        );
        // And once it has, the census's window is the one taken — not the larger latch.
        let settled = Some((Instant::now() - Duration::from_millis(45), census_hwnd));
        let mut first_cand = settled;
        assert_eq!(
            select_hook_or_census_candidate(
                hidden_census,
                bigger_latch,
                &mut first_cand,
                Duration::from_millis(40),
                true,
            ),
            Some(census_hwnd)
        );
        // With the row at its default, the same two windows give the latch, which is what every
        // preset did before the row existed.
        let mut first_cand = settled;
        assert_eq!(
            select_hook_or_census_candidate(
                hidden_census,
                bigger_latch,
                &mut first_cand,
                Duration::from_millis(40),
                false,
            ),
            Some(cap_hwnd)
        );
        // And where the look found nothing to classify at all, the caller passes `false` and the
        // latch is still what answers — a program whose only window is one the census refuses is not
        // left with no preview.
        let mut first_cand = None;
        assert_eq!(
            select_hook_or_census_candidate(
                None,
                bigger_latch,
                &mut first_cand,
                Duration::from_millis(40),
                false,
            ),
            Some(cap_hwnd)
        );
    }
}

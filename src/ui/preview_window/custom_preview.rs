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
//! **Why the program outlives the hover.** A program makes its own window and shows it at its own
//! size, and it does so *after* this app has started it — so for the first moment of its life a
//! window of somebody else's stands on the desktop, and often under the hand that asked for the
//! preview. That is a window the pointer has left the file for: the hook takes the hover down, and
//! a hover that killed its program on the way down would start it again on the way back up, and
//! again, and again — a preview that blinks instead of one that is read, and a program killed
//! before it ever opens the file it was given. So the two lifetimes are not the same one: the
//! *host* goes with the hover at once, and the *program* is kept for [`LINGER`] after it, so that
//! a hover the pointer gives back is the same program shown again rather than a second launch. It
//! is killed when the moment passes, when another file is hovered, or when the pin opens it.
//!
//! **And why it is hidden before it is ever seen, and parked before it is ever drawn.** The launch
//! carries `SW_HIDE` in its `STARTUPINFO` (the road a model's viewer already takes, see
//! `never_shown`), and a program that shows its window anyway — one that calls `ShowWindow` itself
//! with a state of its own — is hidden again on every poll of the wait, every 10 ms, from the
//! moment its first window exists. A window that opens maximized is *restored* rather than
//! resized: a `SetWindowPos` on a maximized window changes the size it would have been restored to
//! and leaves it covering the display it is maximized on. And it is shown only where it
//! **verifiably stands at the parked box** (see `park`): a window this app could not move is one it
//! must not put in front of the user, because a frame of somebody's program over the listing is
//! worse than no picture at all.
//!
//! **Where the box is, and how big.** The placement is the one every other preview takes: the
//! pointer's own (`compute_mouse_layout`) or the focused item's (`compute_keyboard_layout`),
//! which is what reads `Tray → Placement` — the position mode, the pointer standoff and the
//! `Avoid` region — and the room of the display the hover is on. `scaling` is then read the way
//! every other kind's is: a percentage is a share of the size the program's own window opens at,
//! so `100%` is that window at the size it is rather than the whole display, and `Fit to Screen`
//! is the room the placement left. That size is read **once**, before the window is parked,
//! because parking writes the parked box into the window's own placement and a second reading
//! would be a share of the last preview rather than of the program. `background` is what the host
//! paints where the program's window is smaller than the box. `volume` is kept in the preset and
//! written back; a program of the user's own is not told what level to play at — there is no
//! portable way to set another process's audio — so it is a setting a future engine-backed preset
//! can use and nothing on this road reads it.
//!
//! **Why the host is a window of the hover's own.** The preview window paints a card and must
//! never take the focus; the host is the same shape as a hovered film's (`hover_film_host`):
//! a borderless, non-activating, topmost window that is shown only once the Shell has the
//! picture composed in it, and taken down with the hover. It is disabled, so a press on it is the
//! hover's press and lands on the listing behind it. It is made once and kept hidden between
//! hovers: making it again for every hover is a second window where one will do, and one more
//! thing that can flicker.
//!
//! **And the pin is not a pin.** Pressing the pin key on one of these does not freeze a picture
//! of the program: it starts the program properly, with its own frame, at its own place on the
//! desktop, and this app's preview goes away — the user asked for the file, not for a picture of
//! it (see [`open_fully`]).

use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU64, Ordering};
use std::sync::{Mutex, Once};
use std::time::{Duration, Instant};

use windows::core::w;
use windows::Win32::Foundation::{BOOL, COLORREF, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{
    CreateSolidBrush, DeleteObject, FillRect, GetStockObject, HBRUSH, HGDIOBJ, NULL_BRUSH,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, EnumWindows, GetClientRect, GetWindow, GetWindowLongPtrW,
    GetWindowPlacement, GetWindowRect, GetWindowTextLengthW, GetWindowThreadProcessId, IsIconic,
    IsWindow, IsWindowVisible, IsZoomed, PostMessageW, RegisterClassExW, SetWindowLongPtrW,
    SetWindowPlacement, SetWindowPos, ShowWindow, GWL_EXSTYLE, GWL_STYLE, GW_OWNER,
    SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE, SW_SHOWNORMAL,
    WINDOWPLACEMENT, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_ERASEBKGND, WM_NCDESTROY,
    WNDCLASSEXW, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW,
    WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU,
    WS_THICKFRAME,
};

use super::displays::{dpi_at, work_area_at};
use super::layout::{
    compute_keyboard_layout, compute_mouse_layout, KeyboardPlacement, PreviewLayout,
};
use super::load_dispatch::HoverPlacement;
use super::never_shown::{self, NeverShown};
use super::pointer::{AvoidRegion, ScreenRegion};
use crate::app::engine_processes;
use crate::config::config::{custom_preset_for, CustomPreset, PreviewScale, TransparentBackground};
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
const WINDOW_POLL: Duration = Duration::from_millis(10);

/// How often the launched process's own tree is read again while its window is looked for: a
/// program that hands its file to a process of its own is a window under that process's id.
const TREE_POLL: Duration = Duration::from_millis(200);

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

/// Where a parked window is put: past the right edge of every display, at a coordinate Windows
/// still accepts. The window is real, visible and drawing there — which is what the Shell needs
/// to compose it — and no user can see it.
const PARKING: i32 = -30000;

/// What a hover of a custom preview is asked to be drawn at: the pointer's own placement, or the
/// focused item's.
///
/// It is the whole of what the hover knows about *where* the preview goes, and it is read on the
/// worker thread rather than here because the box cannot be worked out before the program's
/// window exists: a percentage is a share of the size that window opens at (see the module note).
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

/// The program this app started, while it is up: the file it was begun for, the program itself,
/// and the process it is.
///
/// It is not the hover's: it outlives one by [`LINGER`], and it is what a hover that comes back
/// finds rather than a second launch (see the module note).
struct Program {
    path: PathBuf,
    exe: String,
    child: NeverShown,
    pid: u32,
}

/// The host window, made once and kept off the screen between hovers.
static HOST: AtomicIsize = AtomicIsize::new(0);
/// Whether the host is on the screen.
static SHOWN: AtomicBool = AtomicBool::new(false);
/// The program's window registered into it, while one is.
static SOURCE: AtomicIsize = AtomicIsize::new(0);
/// The thumbnail registration the Shell is drawing through, while one stands.
static THUMBNAIL: Mutex<isize> = Mutex::new(0);
/// The hover on screen, while there is one.
static WANT: Mutex<Option<Want>> = Mutex::new(None);
/// The program this app started for it, which outlives the hover by [`LINGER`].
static PROGRAM: Mutex<Option<Program>> = Mutex::new(None);
/// The colour the host paints where the program's window does not reach.
static BACKGROUND: AtomicIsize = AtomicIsize::new(0);
/// How many programs have been begun: a worker that was begun for an earlier one is one whose
/// program has been ended, and it neither places nor shows anything.
static LAUNCH: AtomicU64 = AtomicU64::new(0);
static CLASS_REGISTERED: Once = Once::new();

/// Whether a custom preview is on screen.
pub(super) fn is_showing() -> bool {
    WANT.lock().map(|want| want.is_some()).unwrap_or(false)
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

    let Some(child) = launch(&exe, path) else {
        return false;
    };
    let Some(host) = (unsafe { ensure_host() }) else {
        let _ = child.kill();
        return false;
    };

    BACKGROUND.store(color_of(preset.background).0 as isize, Ordering::SeqCst);
    let pid = child.pid;
    let scale = preset.scaling;
    if let Ok(mut held) = PROGRAM.lock() {
        *held = Some(Program {
            path: path.to_path_buf(),
            exe: exe.clone(),
            child,
            pid,
        });
    }
    if let Ok(mut want) = WANT.lock() {
        *want = Some(Want { at });
    }
    // **The program is this app's while it draws the hover.** It is put in the run's job so a
    // crash takes it with the app, and written down so a run that never got to end it is
    // answered for by the next one — the same charge every engine of this app's is under
    // (see `engine_processes`). A program opened by the pin is not: that one is the user's,
    // opened the way the listing opens it.
    let _ = engine_processes::adopt(pid);
    engine_processes::record_player(&exe, pid);

    // Read after the launch is written down, so a worker begun for it is current exactly when it
    // is the program on screen.
    let launch = LAUNCH.fetch_add(1, Ordering::SeqCst) + 1;
    // The window is carried as a number: `HWND` is a pointer and a pointer is not `Send`, and
    // what the worker does with it is a placement the Shell answers for either way.
    let host = host.0 as isize;
    std::thread::spawn(move || run_it(HWND(host as *mut _), pid, scale, launch));
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
pub(super) fn open_fully(path: &Path) -> bool {
    let Some(preset) = preset_for(path) else {
        return false;
    };
    let Some(exe) = preset.exe_path().map(str::to_string) else {
        return false;
    };

    end_the_program();
    let _ = Command::new(exe).arg(path).spawn();
    true
}

/// End the program, its window's registration, and the hover it was drawing.
fn end_the_program() {
    // First, so that a worker still waiting on the program's window stops placing and showing.
    LAUNCH.fetch_add(1, Ordering::SeqCst);

    if let Ok(mut want) = WANT.lock() {
        *want = None;
    }

    let program = PROGRAM.lock().ok().and_then(|mut held| held.take());
    if let Some(program) = program {
        let _ = program.child.kill();
        engine_processes::forget(program.pid);
    }

    let thumbnail = THUMBNAIL
        .lock()
        .map(|mut held| std::mem::replace(&mut *held, 0))
        .unwrap_or(0);
    if thumbnail != 0 {
        let _ = unsafe { DwmUnregisterThumbnail(thumbnail) };
    }
    SOURCE.store(0, Ordering::SeqCst);
    hide_the_host();
}

/// Take the host off the screen, on the thread that made it.
fn hide_the_host() {
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

/// The program as the desktop's own loader wants it: **a full path**.
///
/// `CreateProcessW` is handed the program as `lpApplicationName`, and a name there is a place and
/// never a search — so a preset that says `mspaint.exe` is a launch the desktop refuses, which is
/// a preview with nothing on screen. The search `Command` would have done is done here instead,
/// so a preset that names a program is a preset that runs.
fn resolved(exe: &str) -> Option<String> {
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

/// Begin the program with its window never shown, and answer the process it is.
///
/// It is `CreateProcessW` and not `Command`'s spawn for the one field `Command` cannot carry:
/// `STARTUPINFO.wShowWindow` of `SW_HIDE`, which is what keeps a window off every screen from its
/// first instruction (see `never_shown`). A program that ignores it is one whose window is hidden
/// again as soon as it is found, which is the same road a frame later.
fn launch(exe: &str, path: &Path) -> Option<NeverShown> {
    let mut command = Command::new(exe);
    command.arg(path);
    command.creation_flags(engine_processes::CREATE_NO_WINDOW);
    let mut child = never_shown::spawn(&mut command)?;

    // **A pipe nobody reads is a program blocked on its own writing.** What a program of the
    // user's prints is emptied and dropped, so a chatty one is a preview rather than a hang.
    if let Some(report) = child.report.take() {
        std::thread::spawn(move || {
            let mut report = report;
            let _ = std::io::copy(&mut report, &mut std::io::sink());
        });
    }

    Some(child)
}

/// The colour a backdrop names, as a host window fills it.
fn color_of(background: TransparentBackground) -> COLORREF {
    match background {
        TransparentBackground::Black => COLORREF(0),
        TransparentBackground::White => COLORREF(0x00FF_FFFF),
        // A transparency is shown as the dark grey a checkered page's own squares are drawn
        // over, because a host window cannot be transparent: the Shell draws the picture *over*
        // it, and what is behind the picture is a window rather than the desktop.
        TransparentBackground::Transparent | TransparentBackground::Checkerboard => {
            COLORREF(0x0020_2020)
        }
    }
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
            let brush =
                unsafe { CreateSolidBrush(COLORREF(BACKGROUND.load(Ordering::SeqCst) as u32)) };
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

    let _ = unsafe { EnableWindow(host, false) };
    HOST.store(host.0 as isize, Ordering::SeqCst);
    Some(host)
}

/// Whether the program this worker was begun for is still the one on screen.
fn is_current(launch: u64) -> bool {
    LAUNCH.load(Ordering::SeqCst) == launch
}

/// Find the program's window, keep the box and the picture in step with the hover, and keep the
/// window parked until the hover has been gone for [`LINGER`].
///
/// **One worker per program**, not per hover: the hover can come and go while the program is
/// starting, and what that must cost is a hidden host and a kept window rather than a launch.
fn run_it(host: HWND, pid: u32, scale: PreviewScale, launch: u64) {
    let Some(source) = wait_for_a_window(pid, launch) else {
        // A program that made no window is a preview with nothing to show: it goes, and the host
        // is already off the screen.
        if is_current(launch) {
            end_the_program();
        }
        return;
    };
    if !is_current(launch) {
        return;
    }

    // **Hidden before anything else is done with it.** A program that showed its own window on
    // the desktop is a window the hand is over, and a hover whose pointer left the file is a
    // hover this app is asked to take down (see the module note).
    let _ = unsafe { ShowWindow(source, SW_HIDE) };
    SOURCE.store(source.0 as isize, Ordering::SeqCst);
    strip_the_frame(source);

    // **How large the program opens at, read once and never again**: parking writes the parked
    // box into the window's own placement, so a second reading would be a share of the last
    // preview rather than of the program (see the module note).
    let own = own_size(source);

    let mut box_of_the_host: Option<(i32, i32, i32, i32)> = None;
    let mut registered: isize = 0;
    let mut lingering: Option<Instant> = None;

    loop {
        if !is_current(launch) {
            return;
        }
        if !unsafe { IsWindow(source) }.as_bool() {
            // A program that closed its own window has nothing left to draw, and nothing left to
            // keep: it is ended here rather than left running behind a hover that shows nothing.
            if is_current(launch) {
                end_the_program();
            }
            return;
        }

        let at = WANT
            .lock()
            .ok()
            .and_then(|want| want.as_ref().map(|want| want.at));
        let Some(at) = at else {
            // **The hover is gone; the program is not.** The host leaves the screen at once, and
            // the window is kept where it is for a moment, so that a hover the pointer gives back
            // is this program shown again rather than a second launch.
            let since = *lingering.get_or_insert_with(Instant::now);
            hide_the_host();
            if since.elapsed() >= LINGER {
                end_the_program();
                return;
            }
            std::thread::sleep(KEEP_POLL);
            continue;
        };
        lingering = None;

        let Some(layout) = layout_for(at, own, scale, follow_cursor()) else {
            hide_the_host();
            std::thread::sleep(KEEP_POLL);
            continue;
        };
        let (x, y) = (layout.pos_x, layout.pos_y);
        let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);

        if box_of_the_host != Some((x, y, width, height)) {
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
            box_of_the_host = Some((x, y, width, height));
        }

        // **The window is put back at the box before it is drawn, and never drawn anywhere
        // else**: a program that restored itself, showed itself again or resized itself is a
        // window on the desktop rather than a picture, and it is hidden before it is moved.
        if !parked(source, width, height) {
            park(source, width, height);
        }

        // The registration is the Shell's, and a program that has not presented yet is a window
        // with nothing to compose: it is asked for again on the next look rather than given up on.
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
            let properties = DWM_THUMBNAIL_PROPERTIES {
                dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE,
                rcDestination: client,
                rcSource: RECT::default(),
                opacity: 255,
                fVisible: true.into(),
                // The program's client area rather than its window: the frame has been stripped
                // above, and a program that draws its own border keeps it — that is part of its
                // picture.
                fSourceClientAreaOnly: true.into(),
            };
            let _ = unsafe { DwmUpdateThumbnailProperties(registered, &properties) };

            // **The host is shown by its own thread, after the picture is in it**: a host shown
            // before the Shell has composed anything is its own black rectangle over the listing.
            if !SHOWN.load(Ordering::SeqCst) {
                SHOWN.store(true, Ordering::SeqCst);
                let _ =
                    unsafe { PostMessageW(host, WM_THE_PICTURE_IS_READY, WPARAM(0), LPARAM(0)) };
            }
        }

        std::thread::sleep(KEEP_POLL);
    }
}

/// The box the preview is drawn in: the placement the tray's settings and the display ask for,
/// at the size the preset's scale makes of the program's own window.
fn layout_for(
    at: Where,
    own: (u32, u32),
    scale: PreviewScale,
    follow_cursor: bool,
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
                    preview_scale: scale,
                },
                bounds,
                dpi,
            )
        }
    }
}

/// The size the program's own window opens at: what `100%` is a hundred percent of.
///
/// It is the window's *restored* box and not the one it is showing at, because a program that
/// remembers being maximized is showing at the whole display — and the display is what the fit
/// is for, not what a percentage means. A window that never had a restored box is measured where
/// it stands.
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

/// The box a window parked at this size stands in.
fn the_parked_box(width: i32, height: i32) -> RECT {
    RECT {
        left: PARKING - width,
        top: PARKING,
        right: PARKING,
        bottom: PARKING + height,
    }
}

/// Whether a window stands, visibly, at the box a window of this size is parked in.
fn parked(window: HWND, width: i32, height: i32) -> bool {
    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_err() {
        return false;
    }
    let parked = the_parked_box(width, height);

    rect.left == parked.left
        && rect.top == parked.top
        && rect.right == parked.right
        && rect.bottom == parked.bottom
        && unsafe { IsWindowVisible(window) }.as_bool()
        && !unsafe { IsZoomed(window) }.as_bool()
        && !unsafe { IsIconic(window) }.as_bool()
}

/// Restore the window out of any maximized or minimized state, put it at the parked box at the
/// preview's size, and show it where no display reaches — **only where it verifiably stands
/// there**.
///
/// **`SetWindowPlacement` is what restores it**, and it is given the parked box as the box to
/// restore to: a `SetWindowPos` on a maximized window only changes the size it *would* have been
/// restored to and leaves the window covering the display it is maximized on, and a restore to the
/// window's own rect would be a full screen for as long as it takes to hide it again. The window
/// is hidden before the move and hidden again after it, so nothing of somebody's program is on a
/// screen at any point of it, and a window that could not be moved is left hidden rather than
/// shown where it stands — a program's frame over the listing is worse than no picture at all.
fn park(window: HWND, width: i32, height: i32) {
    let parked = the_parked_box(width, height);

    let _ = unsafe { ShowWindow(window, SW_HIDE) };

    if unsafe { IsZoomed(window) }.as_bool() || unsafe { IsIconic(window) }.as_bool() {
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            showCmd: SW_SHOWNORMAL.0 as u32,
            rcNormalPosition: parked,
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(window, &placement) };
        let _ = unsafe { ShowWindow(window, SW_HIDE) };
    }

    let _ = unsafe {
        SetWindowPos(
            window,
            HWND::default(),
            parked.left,
            parked.top,
            width,
            height,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };

    if !parked_at_the_box(window, &parked) {
        // **A window this app could not move is a window it must not show.** The next look hides
        // it again, and the hover is the host's background until the program lets itself be
        // parked.
        let _ = unsafe { ShowWindow(window, SW_HIDE) };
        return;
    }
    let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
}

/// Whether a window's own box is the one it was asked to stand in.
fn parked_at_the_box(window: HWND, box_of: &RECT) -> bool {
    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_err() {
        return false;
    }

    rect.left == box_of.left
        && rect.top == box_of.top
        && rect.right == box_of.right
        && rect.bottom == box_of.bottom
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
}

/// The program's first top-level window with a title, waited for within [`WINDOW_WAIT`].
///
/// It is the window the user would have seen, which is the only window a program can be assumed
/// to draw its document in: a splash screen has no title, and a tool window is not what a program
/// opens a file with. **It need not be visible**: the launch asks for a hidden window, so the one
/// this is looking for is a window that has not been shown at all.
///
/// **Every window the program shows is hidden on the way**, on every poll: a program that shows
/// its own window whatever its launch said is one whose window must be off the desktop within a
/// frame of it being on it — the hand is over that window, and a hand over somebody's window is a
/// hover this app is asked to take down (see the module note). The window the wait is *for* is
/// hidden as soon as it is found; the others are hidden so that they are never a flash.
fn wait_for_a_window(pid: u32, launch: u64) -> Option<HWND> {
    let deadline = Instant::now() + WINDOW_WAIT;
    let mut processes = process_tree(pid);
    let mut read = Instant::now();

    while Instant::now() < deadline {
        if !is_current(launch) {
            return None;
        }
        hide_every_window_of(&processes);
        if let Some(window) = first_window_of(&processes) {
            return Some(window);
        }
        if read.elapsed() >= TREE_POLL {
            processes = process_tree(pid);
            read = Instant::now();
        }
        std::thread::sleep(WINDOW_POLL);
    }
    None
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

/// Hide every visible top-level window of the processes, so that a program that shows its own
/// window anyway is a program whose window is on the desktop for no longer than one poll.
fn hide_every_window_of(processes: &[u32]) {
    struct Search<'a> {
        processes: &'a [u32],
    }

    unsafe extern "system" fn visit(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let search = unsafe { &*(lparam.0 as *const Search) };
        let mut owner = 0u32;
        unsafe { GetWindowThreadProcessId(hwnd, Some(&mut owner)) };
        if search.processes.contains(&owner) && unsafe { IsWindowVisible(hwnd) }.as_bool() {
            let _ = unsafe { ShowWindow(hwnd, SW_HIDE) };
        }
        BOOL(1)
    }

    let search = Search { processes };
    let _ = unsafe { EnumWindows(Some(visit), LPARAM(&search as *const Search as isize)) };
}

/// The first titled, top-level, non-tool window of one of `processes`.
fn first_window_of(processes: &[u32]) -> Option<HWND> {
    struct Search<'a> {
        processes: &'a [u32],
        found: Option<HWND>,
    }

    unsafe extern "system" fn visit(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let search = unsafe { &mut *(lparam.0 as *mut Search) };
        let mut owner = 0u32;
        unsafe { GetWindowThreadProcessId(hwnd, Some(&mut owner)) };
        if !search.processes.contains(&owner) {
            return BOOL(1);
        }
        // A tool window is a palette or a tip, and an owned one is a dialog over something else:
        // neither is the window a document is drawn in.
        let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
        let owned = unsafe { GetWindow(hwnd, GW_OWNER) }.is_ok_and(|owner| !owner.0.is_null());
        if ex_style & WS_EX_TOOLWINDOW.0 != 0 || owned || unsafe { GetWindowTextLengthW(hwnd) } == 0
        {
            return BOOL(1);
        }
        search.found = Some(hwnd);
        BOOL(0)
    }

    let mut search = Search {
        processes,
        found: None,
    };
    let _ = unsafe { EnumWindows(Some(visit), LPARAM(&mut search as *mut Search as isize)) };
    search.found
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
    fn a_hundred_percent_is_the_size_the_program_opens_at_not_the_display() {
        let layout = layout_for(
            pointer((0, 0)),
            (600, 400),
            PreviewScale::Percent(100),
            true,
        )
        .expect("a 600x400 window is a preview a display can hold");
        assert_eq!((layout.preview_w, layout.preview_h), (600, 400));
    }

    #[test]
    fn a_percentage_is_that_share_of_the_programs_own_window() {
        let layout = layout_for(pointer((0, 0)), (600, 400), PreviewScale::Percent(50), true)
            .expect("half of a 600x400 window is a preview a display can hold");
        assert_eq!((layout.preview_w, layout.preview_h), (300, 200));
    }

    #[test]
    fn the_fit_takes_the_room_and_the_box_stays_on_the_display() {
        let room = work_area_at(0, 0);
        let layout = layout_for(
            pointer((0, 0)),
            (600, 400),
            PreviewScale::FitToScreen,
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
    }
}

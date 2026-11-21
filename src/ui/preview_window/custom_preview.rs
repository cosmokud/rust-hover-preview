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
//! `Avoid` region — and the room of the display the hover is on. The **size and the shape** of the
//! box are the file's own: a picture says how wide and how tall it is in its own header, so a
//! picture that is taller than it is wide is a tall box, and the box is that picture's size before
//! the preset's `scaling` is read (see `base_size`). `scaling` is then read the way every other
//! picture of this app's is — the picture's rule, `bitmap_at_display_scale`: `100%` is the picture
//! at its own size and `Fit to Screen` is that same size, never a photograph stretched over the
//! display, which is what a preview of one filling the screen would be (see `share_for`).
//! **And the box is the size the program's window took.** A window is not obliged to draw itself in
//! the box a preview asks for: Windows 11's Paint answers `WM_GETMINMAXINFO` with a floor of 704x500
//! and takes anything smaller as 704x500 — measured with its frame stripped, as it is here, and by a
//! `SetWindowPlacement` as well as a `SetWindowPos` — so a box cut to a tall picture's own shape can
//! be one the window refuses the width of, and a window that takes a size of its own inside a narrow
//! box is the host's background in a band above and below the picture. The size is read off the
//! window itself rather than from a list of programs (see `the_size_it_took`), and where it is larger
//! than the box the box is that size at `100%`: the hover shows the program's picture filling it (see
//! `the_box_around`). A program that draws smaller than the box it was asked for keeps that box, and
//! `background` is what the host paints around its picture (see `the_picture_in`). It is one rule for
//! every program rather than a list of programs: what the window took is the box it gets.
//! **A program's window is not that size**: it opens at whatever size it was last left at — one
//! left maximized is the whole display — so a box read off it is the size of the last file
//! somebody opened with the program, and a tall picture is shown in a wide window that fills the
//! screen. A file that declares no size of its own — a document, an archive, a sound — is the one
//! case where the program's window is the only size there is, and that size is read **once**,
//! before the window is parked, because parking writes the parked box into the window's own
//! placement and a second reading would be the shape of the last preview rather than of the
//! program. Such a file keeps the document's rule, `fit_reduced`: a percentage is a share of the
//! room the placement leaves, and `Fit to Screen` is the whole of it, which is what the tray's own
//! scaling menu calls `Screen 100%`.
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

use std::collections::HashMap;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::{Mutex, Once};
use std::time::{Duration, Instant};

use windows::core::w;
use windows::Win32::Foundation::{
    CloseHandle, BOOL, COLORREF, HWND, LPARAM, LRESULT, RECT, WPARAM,
};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{
    CreateSolidBrush, DeleteObject, FillRect, GetStockObject, HBRUSH, HGDIOBJ, NULL_BRUSH,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::{
    OpenProcess, QueryFullProcessImageNameW, TerminateProcess, PROCESS_NAME_WIN32,
    PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
};
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, EnumWindows, GetClientRect, GetWindow, GetWindowLongPtrW,
    GetWindowPlacement, GetWindowRect, GetWindowTextLengthW, GetWindowThreadProcessId, IsIconic,
    IsWindow, IsWindowVisible, IsZoomed, PostMessageW, RegisterClassExW, SetWindowLongPtrW,
    SetWindowPlacement, SetWindowPos, ShowWindow, GWL_EXSTYLE, GWL_STYLE, GW_OWNER,
    SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOSIZE, SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE,
    SW_SHOWNORMAL, WINDOWPLACEMENT, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_ERASEBKGND,
    WM_NCDESTROY, WNDCLASSEXW, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN, WS_EX_NOACTIVATE,
    WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_POPUP,
    WS_SYSMENU, WS_THICKFRAME,
};

use super::dimensions::picture_dimensions;
use super::displays::{dpi_at, work_area_at};
use super::layout::{
    compute_keyboard_layout, compute_mouse_layout, KeyboardPlacement, PreviewLayout,
};
use super::load_dispatch::HoverPlacement;
use super::never_shown::{self, NeverShown};
use super::pointer::{AvoidRegion, ScreenRegion};
use super::preview_scale::{bitmap_at_display_scale, fit_reduced};
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

/// How long a program that never shows a window is given to show one before the window it has is
/// taken as the picture anyway: a program that honours the hidden launch never shows one, and one
/// that is merely slow to show it is not a program with no window.
const SHOWN_WAIT: Duration = Duration::from_millis(300);

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

/// Where a parked window is put: past the right edge of every display, at a coordinate Windows
/// still accepts. The window is real, visible and drawing there — which is what the Shell needs
/// to compose it — and no user can see it.
const PARKING: i32 = -30000;

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

/// What one launch is: the process it began, the image that program runs under, and the windows
/// and processes that were on the machine before it — which is what tells the program's own
/// window from everything else, and the process it may end from the ones that are the user's.
struct Launch {
    /// The file the program was begun for: the size the box is cut to is the file's own, and it is
    /// read on the worker rather than here (see `base_size`).
    path: PathBuf,
    /// The process this app began.
    pid: u32,
    /// The preset program's own file name, lowercased: what a process the program hands its file
    /// to is running, however it was begun.
    exe: String,
    /// Every top-level window there was before the launch, so that the program's own are the ones
    /// that were not.
    windows_before: Vec<isize>,
    /// Every process there was before the launch, so that one the program started is the only one
    /// this app may end (see `end_the_program`).
    processes_before: Vec<u32>,
    /// Which launch this is, so that a worker begun for an earlier one stands down.
    id: u64,
}

/// The host window, made once and kept off the screen between hovers.
static HOST: AtomicIsize = AtomicIsize::new(0);
/// Whether the host is on the screen.
static SHOWN: AtomicBool = AtomicBool::new(false);
/// The program's window registered into it, while one is.
static SOURCE: AtomicIsize = AtomicIsize::new(0);
/// The process that window belongs to, which is not the launched one where the program is
/// packaged.
static SOURCE_PID: AtomicU32 = AtomicU32::new(0);
/// Whether that process is one the launch made, which is the only kind this app ends.
static SOURCE_IS_NEW: AtomicBool = AtomicBool::new(false);
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

    // **The desktop is read before the program is begun**, so that the windows it puts up are the
    // ones that were not there, and the process it hands the file to is one this app may end.
    let windows_before = top_level_windows();
    let processes_before = engine_processes::running().unwrap_or_default();

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
    let exe_file = Path::new(&exe)
        .file_name()
        .map(|name| name.to_string_lossy().to_lowercase())
        .unwrap_or_default();
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
    let id = LAUNCH.fetch_add(1, Ordering::SeqCst) + 1;
    let launch = Launch {
        path: path.to_path_buf(),
        pid,
        exe: exe_file,
        windows_before,
        processes_before,
        id,
    };
    // The window is carried as a number: `HWND` is a pointer and a pointer is not `Send`, and
    // what the worker does with it is a placement the Shell answers for either way.
    let host = host.0 as isize;
    std::thread::spawn(move || run_it(HWND(host as *mut _), launch, scale));
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
    let launched_pid = program.as_ref().map(|program| program.pid).unwrap_or(0);
    if let Some(program) = program {
        let _ = program.child.kill();
        engine_processes::forget(program.pid);
    }

    // **The window this app asked for is closed, and the process that opened it is ended only
    // where the launch is what made it.** A packaged program's window belongs to a process this
    // app never started, and that process may be one the user had open before the hover: its
    // window is the document this app asked for, so the window goes and the process stays.
    let source = SOURCE.swap(0, Ordering::SeqCst);
    let source_pid = SOURCE_PID.swap(0, Ordering::SeqCst);
    let source_is_new = SOURCE_IS_NEW.swap(false, Ordering::SeqCst);
    if source != 0 {
        let source = HWND(source as *mut _);
        if unsafe { IsWindow(source) }.as_bool() {
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
fn is_current(id: u64) -> bool {
    LAUNCH.load(Ordering::SeqCst) == id
}

/// Find the program's window, keep the box and the picture in step with the hover, and keep the
/// window parked until the hover has been gone for [`LINGER`].
///
/// **One worker per program**, not per hover: the hover can come and go while the program is
/// starting, and what that must cost is a hidden host and a kept window rather than a launch.
fn run_it(host: HWND, launch: Launch, scale: PreviewScale) {
    let Some(source) = wait_for_a_window(&launch) else {
        // A program that made no window is a preview with nothing to show: it goes, and the host
        // is already off the screen.
        if is_current(launch.id) {
            end_the_program();
        }
        return;
    };
    if !is_current(launch.id) {
        return;
    }

    // **Hidden before anything else is done with it.** A program that showed its own window on
    // the desktop is a window the hand is over, and a hover whose pointer left the file is a
    // hover this app is asked to take down (see the module note).
    let _ = unsafe { ShowWindow(source, SW_HIDE) };
    let source_pid = pid_of(source);
    SOURCE.store(source.0 as isize, Ordering::SeqCst);
    SOURCE_PID.store(source_pid, Ordering::SeqCst);
    SOURCE_IS_NEW.store(
        !launch.processes_before.contains(&source_pid),
        Ordering::SeqCst,
    );
    strip_the_frame(source);

    // **The size the box is cut from, read once and never again**: parking writes the parked box
    // into the window's own placement, so a second reading would be the shape of the last preview
    // rather than of the program (see the module note). It is the file's own where the file
    // declares one, and the program's window only where it does not (see `base_size`) — and which
    // of the two it is is also which rule the preset's scale is read by (see `share_for`). A
    // program that will not draw itself at that size is answered by the size it took instead (see
    // `the_box_around`).
    let base = base_size(&launch.path);
    let own = base.unwrap_or_else(|| own_size(source));
    let shape = match base {
        Some(_) => Shape::OfTheFile,
        None => Shape::OfTheProgram,
    };

    let mut processes = process_tree(launch.pid);
    if !processes.contains(&source_pid) {
        processes.push(source_pid);
    }
    let mut tree_read = Instant::now();
    let mut images: HashMap<u32, Option<String>> = HashMap::new();

    let mut box_of_the_host: Option<(i32, i32, i32, i32)> = None;
    let mut registered: isize = 0;
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

        // **Every other window the program puts up is hidden at once.** A splash, a dialog or a
        // second frame is not the picture, and a window of the program's on the desktop is a
        // window under the hand.
        census(&launch, &processes, &mut images, source.0 as isize).take();
        if tree_read.elapsed() >= Duration::from_secs(1) {
            processes = process_tree(launch.pid);
            if !processes.contains(&source_pid) {
                processes.push(source_pid);
            }
            tree_read = Instant::now();
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

        let Some(asked) = layout_for(at, shape, own, scale, follow_cursor()) else {
            hide_the_host();
            std::thread::sleep(KEEP_POLL);
            continue;
        };

        // **The window is put back at the parking corner before it is drawn, and never drawn
        // anywhere else**: a program that restored itself, showed itself again or resized itself is
        // a window on the desktop rather than a picture, and it is hidden before it is moved.
        if !parked(source) {
            park(source, asked.preview_w as i32, asked.preview_h as i32);
        }
        if !parked(source) {
            // **A window that would not stand off the desktop is a picture there is none of**: it
            // is left hidden by `park`, and the hover waits on the next look with nothing up rather
            // than a background card standing in for a picture that is not there.
            hide_the_host();
            std::thread::sleep(KEEP_POLL);
            continue;
        }

        // **The box is the size the window took**, which is read after it has taken it: a program
        // that will not be as small as the box asks is one the picture would otherwise stand
        // letterboxed in (see `the_box_around`).
        let Some(layout) = the_box_around(at, asked, the_size_it_took(source), follow_cursor())
        else {
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
            let mut picture = RECT::default();
            let _ = unsafe { GetClientRect(source, &mut picture) };
            let properties = DWM_THUMBNAIL_PROPERTIES {
                dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE,
                // The box's whole client where the window took the box — which is the box, raised
                // to the size it took (see `the_box_around`) — and its own shape fitted into the
                // client where the placement's room is smaller than the window (see
                // `the_picture_in`).
                rcDestination: the_picture_in(
                    client.right,
                    client.bottom,
                    (picture.right, picture.bottom),
                ),
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

/// What the box's size and shape were read off, which is also what the preset's `scaling` is a
/// share of.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Shape {
    /// The file's own header: a picture's size, which is the picture rule's base.
    OfTheFile,
    /// The program's window, for a file that declares no size of its own: a page in a window
    /// rather than a picture with pixels of its own, which is the document rule's base.
    OfTheProgram,
}

/// The share the box is drawn at, for the base its size was read off.
///
/// **The two rules are the two this app already has.** A picture is a bitmap, so a share of it is
/// a share of its own pixels and a fit is the picture at the size it is — never a photograph
/// enlarged over the display, which is what `bitmap_at_display_scale` states and why `100%` is a
/// preview rather than a screen. A file that declares no size of its own is a page or a listing in
/// a window, drawn at whatever size it is asked for, so a share of it is a share of the room the
/// placement leaves — `fit_reduced`'s rule.
fn share_for(shape: Shape, scale: PreviewScale) -> PreviewScale {
    match shape {
        Shape::OfTheFile => bitmap_at_display_scale(scale),
        Shape::OfTheProgram => fit_reduced(scale),
    }
}

/// The box the preview is drawn in: the placement the tray's settings and the display ask for,
/// at the share of that base the preset's scale names, cut to the base's own shape.
///
/// **The share is a share of the base** (see `share_for`), which is what makes a preset's scale
/// mean what a tray menu's does: `100%` of a picture is the picture at its own size, and `100%` of
/// a file with no size of its own is the room the placement leaves.
fn layout_for(
    at: Where,
    shape: Shape,
    own: (u32, u32),
    scale: PreviewScale,
    follow_cursor: bool,
) -> Option<PreviewLayout> {
    layout_at(at, own, share_for(shape, scale), follow_cursor)
}

/// The placement a box of `own` at `share` comes out at — in the room the display leaves and clear
/// of what the placement keeps off.
///
/// It is the arithmetic behind [`layout_for`], and it is asked directly by the one box whose share
/// is settled already: the size the program's own window took (see `the_box_around`).
fn layout_at(
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

/// The box the hover is drawn in: the one its placement asked for, raised to the size the program's
/// window actually took where the program took a size of its own.
///
/// **A window that will not take the box is the box.** A tall picture makes a narrow box, and a
/// program that will not be narrower than its own floor takes one of its own inside it — the picture
/// then stood in the middle of that box with the host's background in a band above and below it. So
/// the size the window stands at is read off the window (see `the_size_it_took`), and where it is
/// larger than the box the placement is asked again for a box of exactly the window's size at
/// `100%`: the program's own picture fills the hover, which is what the window is. A program that
/// draws smaller than the box it was asked for keeps that box, and `the_picture_in` composes its
/// picture into it.
///
/// It is one rule for every program rather than a list of programs: what the window took is the box
/// it gets.
fn the_box_around(
    at: Where,
    asked: PreviewLayout,
    taken: (u32, u32),
    follow_cursor: bool,
) -> Option<PreviewLayout> {
    if taken.0 <= asked.preview_w && taken.1 <= asked.preview_h {
        return Some(asked);
    }
    layout_at(at, taken, PreviewScale::Percent(100), follow_cursor)
}

/// The shape of the file itself, where the file declares one: the size a picture says it is in its
/// own header.
///
/// **This is what a program's window cannot answer.** A program opens at the size it was last left
/// at — the window Paint was left at yesterday is a wide window whatever the picture in it is — so
/// a shape read off it is the shape of the last file somebody opened with the program, and a tall
/// picture is shown in a wide box. A picture's own bytes say how wide and how tall it is before
/// anything is decoded, and that is the shape the box is cut to (see `picture_dimensions`).
///
/// A file that is not a picture this app measures — a document, an archive, a sound — has no such
/// answer, and the program's window is the only shape there is to go by (see `own_size`).
fn base_size(path: &Path) -> Option<(u32, u32)> {
    picture_dimensions(path).filter(|&(width, height)| width > 0 && height > 0)
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

/// The size the program's window actually took: its client, which with the frame stripped is the
/// picture the Shell composes into the host (see `strip_the_frame`).
///
/// **A program is not obliged to take the box a preview asks for.** Windows 11's Paint answers
/// `WM_GETMINMAXINFO` with a floor of 704x500 and takes anything smaller as 704x500 — measured with
/// its frame stripped, as it is here, and by a `SetWindowPlacement` as well as by a `SetWindowPos` —
/// so the size to go by is the one the window stands at rather than the one it was asked for. It is
/// read off the window itself rather than from a number written down for any one program, and it is
/// read *after* the frame is stripped, because that is the window the preview draws. Nothing at all
/// is answered where the window will not say, and the box is then the one the placement asked for.
fn the_size_it_took(window: HWND) -> (u32, u32) {
    let mut client = RECT::default();
    if unsafe { GetClientRect(window, &mut client) }.is_err() {
        return (0, 0);
    }

    (client.right.max(0) as u32, client.bottom.max(0) as u32)
}

/// The rectangle in the host's client the program's window is drawn in: the window's own client at
/// the size it has where that fits, and its own shape fitted into the client where it does not.
///
/// **The box and the window normally agree**, because the box is raised to the size the window took
/// (see `the_box_around`) and this is then the host's whole client with the picture at its own size.
/// They part where the box the placement leaves is smaller than the window — the room beside the text
/// a hover keeps off, a window whose floor is larger than the display — and a preview that shows the
/// program at its own shape inside the box says more than one that shows nothing. The window is never
/// stretched: the shape it is drawn at is its own, and `background` is what the host paints in the
/// letterbox.
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

/// The box a window parked at this size stands in.
fn the_parked_box(width: i32, height: i32) -> RECT {
    RECT {
        left: PARKING - width,
        top: PARKING,
        right: PARKING,
        bottom: PARKING + height,
    }
}

/// The corner a parked window's box is put at: its top-right, which is the one corner of that box
/// that stands in the same place whatever size the window takes there.
fn at_the_parking_corner(rect: &RECT) -> bool {
    rect.right == PARKING && rect.top == PARKING
}

/// Whether a window stands off the desktop, visibly, at the parking corner — whatever size it took
/// there.
///
/// **The corner is what says a window is parked, not the box it was asked for.** A window that keeps
/// a size of its own keeps it however it is asked (see `the_size_it_took`), so the box the hover
/// draws is the size it took — and where the two still do not agree, the window is one this app *did*
/// move off every display, which is the whole of what parking is for,
/// and `the_picture_in` composes it into the box. A window that is not at the corner is a window
/// standing somewhere a display reaches.
fn parked(window: HWND) -> bool {
    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_err() {
        return false;
    }

    at_the_parking_corner(&rect)
        && unsafe { IsWindowVisible(window) }.as_bool()
        && !unsafe { IsZoomed(window) }.as_bool()
        && !unsafe { IsIconic(window) }.as_bool()
}

/// Restore the window out of any maximized or minimized state, put its top-right corner at the
/// parking place, and show it where no display reaches — **only where it verifiably stands
/// there**.
///
/// **`SetWindowPlacement` is what restores it**, and it is given the parked box as the box to
/// restore to: a `SetWindowPos` on a maximized window only changes the size it *would* have been
/// restored to and leaves the window covering the display it is maximized on, and a restore to the
/// window's own rect would be a full screen for as long as it takes to hide it again. The window
/// is hidden before the move and hidden again after it, so nothing of somebody's program is on a
/// screen at any point of it, and a window that could not be moved is left hidden rather than
/// shown where it stands — a program's frame over the listing is worse than no picture at all.
///
/// **And the corner is asked for again once the window has answered**, because the box is what the
/// window is asked to *take*: one that keeps a size of its own takes it, and its right edge then
/// stands past the parking place by the difference. The second ask moves it by the size it actually
/// took and asks for no size at all (see `at_the_parking_corner`).
fn park(window: HWND, width: i32, height: i32) {
    let parked_box = the_parked_box(width, height);

    let _ = unsafe { ShowWindow(window, SW_HIDE) };

    if unsafe { IsZoomed(window) }.as_bool() || unsafe { IsIconic(window) }.as_bool() {
        let placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            showCmd: SW_SHOWNORMAL.0 as u32,
            rcNormalPosition: parked_box,
            ..Default::default()
        };
        let _ = unsafe { SetWindowPlacement(window, &placement) };
        let _ = unsafe { ShowWindow(window, SW_HIDE) };
    }

    let _ = unsafe {
        SetWindowPos(
            window,
            HWND::default(),
            parked_box.left,
            parked_box.top,
            width,
            height,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };

    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_ok() && !at_the_parking_corner(&rect) {
        let _ = unsafe {
            SetWindowPos(
                window,
                HWND::default(),
                PARKING - (rect.right - rect.left),
                PARKING,
                0,
                0,
                SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE,
            )
        };
    }

    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_err() || !at_the_parking_corner(&rect) {
        // **A window this app could not move is a window it must not show.** It is left hidden for
        // the next look to try again, and a window that is not off the desktop stays off the screen.
        let _ = unsafe { ShowWindow(window, SW_HIDE) };
        return;
    }
    let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
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

/// One look at the desktop: every window the program has put up since it was begun.
///
/// **Each is hidden as it is found** — a program that shows its own window whatever its launch
/// said is one whose window must be off the desktop within a frame of it being on it — and the
/// ones that were showing are recorded, because the window a program actually shows is the one it
/// draws its document in, and a window it keeps hidden is one it has not drawn yet.
struct Census<'a> {
    launch: &'a Launch,
    processes: &'a [u32],
    images: &'a mut HashMap<u32, Option<String>>,
    /// The window this app has already adopted, which is the picture and not to be hidden.
    keep: isize,
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
    keep: isize,
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
        keep,
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
        let owned = unsafe { GetWindow(hwnd, GW_OWNER) }.is_ok_and(|owner| !owner.0.is_null());
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
            // Hidden as it is found, and before anything is decided about it: a window the
            // program showed is a window the hand may be over.
            let _ = unsafe { ShowWindow(hwnd, SW_HIDE) };
        }

        if document {
            self.found.push((area, hwnd));
        }
    }

    /// Whether a process is one the program's: in the tree it was begun in, or running the image
    /// the preset names — which is how a packaged program's own window is found, the activator
    /// having handed the file to a process this app never started.
    fn belongs(&mut self, pid: u32) -> bool {
        if self.processes.contains(&pid) {
            return true;
        }
        image_name_of(pid, self.images).is_some_and(|name| name == self.launch.exe)
    }

    /// Nothing is read out of the census: what it did is the hiding, and what it found is what the
    /// caller asked for (see `wait_for_a_window`).
    fn take(self) {}
}

/// The program's window, waited for within [`WINDOW_WAIT`].
///
/// The window is the one the program **shows** — the window the user would have seen, which is
/// the only one a program can be assumed to draw its document in — and where it shows none (the
/// launch asked it not to) the **largest** titled, top-level, unowned, non-tool window of a
/// document's size that it has (see `THE_SMALLEST_PICTURE`): a program's helper windows are
/// windows too, and the one it draws its file in is the one with a frame to draw in.
fn wait_for_a_window(launch: &Launch) -> Option<HWND> {
    let deadline = Instant::now() + WINDOW_WAIT;
    let began = Instant::now();
    let mut processes = process_tree(launch.pid);
    let mut read = Instant::now();
    let mut images: HashMap<u32, Option<String>> = HashMap::new();

    while Instant::now() < deadline {
        if !is_current(launch.id) {
            return None;
        }

        let look = census(launch, &processes, &mut images, 0);
        if let Some((_, window)) = look.shown {
            return Some(window);
        }
        if began.elapsed() >= SHOWN_WAIT {
            if let Some((_, window)) = look.found.iter().copied().max_by_key(|(area, _)| *area) {
                return Some(window);
            }
        }

        if read.elapsed() >= TREE_POLL {
            processes = process_tree(launch.pid);
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
        let layout = layout_for(
            pointer((0, 0)),
            Shape::OfTheProgram,
            (600, 400),
            PreviewScale::Percent(100),
            true,
        )
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
        let whole = layout_for(
            pointer((0, 0)),
            Shape::OfTheProgram,
            (600, 400),
            PreviewScale::Percent(100),
            true,
        )
        .expect("a display has room for a preview");
        let half = layout_for(
            pointer((0, 0)),
            Shape::OfTheProgram,
            (600, 400),
            PreviewScale::Percent(50),
            true,
        )
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
            Shape::OfTheProgram,
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
    fn a_picture_is_the_picture_at_its_own_size_whatever_the_scale_says() {
        let at_a_hundred = layout_for(
            pointer((0, 0)),
            Shape::OfTheFile,
            (600, 400),
            PreviewScale::Percent(100),
            true,
        )
        .expect("a display has room for a preview");
        assert_eq!(
            (at_a_hundred.preview_w, at_a_hundred.preview_h),
            (600, 400),
            "the base is the picture's own size, so a preview of one is that picture rather than \
             the screen"
        );

        let fitted = layout_for(
            pointer((0, 0)),
            Shape::OfTheFile,
            (600, 400),
            PreviewScale::FitToScreen,
            true,
        )
        .expect("a display has room for a preview");
        assert_eq!(
            (fitted.preview_w, fitted.preview_h),
            (600, 400),
            "and a fit is the picture at the size it is rather than stretched over the display"
        );

        let half = layout_for(
            pointer((0, 0)),
            Shape::OfTheFile,
            (600, 400),
            PreviewScale::Percent(50),
            true,
        )
        .expect("a display has room for a preview");
        assert_eq!(
            (half.preview_w, half.preview_h),
            (300, 200),
            "a percentage below a fit is that share of the picture's own size"
        );
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
    fn a_box_is_raised_to_the_size_the_window_took() {
        // Paint answers `WM_GETMINMAXINFO` with 704x500 and will not go smaller, whatever a small
        // picture asks for, so the box is the size the window stands at rather than the one the
        // placement asked for.
        let asked = layout_for(
            pointer((0, 0)),
            Shape::OfTheFile,
            (80, 60),
            PreviewScale::Percent(100),
            true,
        )
        .expect("a display has room for a preview");
        assert_eq!(
            (asked.preview_w, asked.preview_h),
            (80, 60),
            "the placement asks for the picture's own size first"
        );

        let took = the_box_around(pointer((0, 0)), asked, (704, 500), true)
            .expect("a display has room for the size the window took");
        assert_eq!(
            (took.preview_w, took.preview_h),
            (704, 500),
            "and the box is the smallest one the window will stand in, which is the picture the \
             hover then shows rather than a letterbox around one"
        );

        // An empty box is a window that would not say what it took, which leaves the placement it
        // was asked for.
        let asked_again = layout_for(
            pointer((0, 0)),
            Shape::OfTheFile,
            (80, 60),
            PreviewScale::Percent(100),
            true,
        )
        .expect("a display has room for a preview");
        let asked_at = (asked_again.pos_x, asked_again.pos_y);
        let unsaid = the_box_around(pointer((0, 0)), asked_again, (0, 0), true)
            .expect("a display has room for a preview");
        assert_eq!(
            (
                unsaid.pos_x,
                unsaid.pos_y,
                unsaid.preview_w,
                unsaid.preview_h
            ),
            (asked_at.0, asked_at.1, 80, 60),
            "a window that takes the box leaves the placement as it was"
        );

        // And the case the report was of: a tall picture makes a narrow box, and the window takes
        // its own width at the height it was asked for — the box is that, at the window's own
        // height, rather than the narrow one with the picture's bands of background above and below.
        let tall = layout_for(
            pointer((0, 0)),
            Shape::OfTheFile,
            (40, 120),
            PreviewScale::Percent(100),
            true,
        )
        .expect("a display has room for a preview");
        let took = the_box_around(pointer((0, 0)), tall, (704, 500), true)
            .expect("a display has room for the size the window took");
        assert_eq!(
            (took.preview_w, took.preview_h),
            (704, 500),
            "a window that would not be as narrow as the box is the box"
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
    }
}

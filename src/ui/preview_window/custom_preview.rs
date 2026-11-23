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

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::{Child, Command};
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::{Mutex, Once};
use std::time::{Duration, Instant};

use windows::core::w;
use windows::Win32::Foundation::{
    CloseHandle, BOOL, COLORREF, HWND, LPARAM, LRESULT, RECT, WPARAM,
};
use windows::Win32::Graphics::Dwm::{
    DwmQueryThumbnailSourceSize, DwmRegisterThumbnail, DwmUnregisterThumbnail,
    DwmUpdateThumbnailProperties, DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION,
    DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{
    CreateSolidBrush, DeleteObject, FillRect, GetStockObject, RedrawWindow, HBRUSH, HGDIOBJ,
    NULL_BRUSH, RDW_ALLCHILDREN, RDW_FRAME, RDW_INVALIDATE,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::{
    OpenProcess, QueryFullProcessImageNameW, TerminateProcess, PROCESS_NAME_WIN32,
    PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
};
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, EnumWindows, GetClientRect, GetForegroundWindow,
    GetSystemMetrics, GetWindow, GetWindowLongPtrW, GetWindowPlacement, GetWindowRect,
    GetWindowTextLengthW, GetWindowThreadProcessId, IsIconic, IsWindow, IsWindowVisible, IsZoomed,
    PostMessageW, RegisterClassExW, SendMessageTimeoutW, SetForegroundWindow, SetWindowLongPtrW,
    SetWindowPlacement, SetWindowPos, ShowWindow, GWL_EXSTYLE, GWL_STYLE, GW_OWNER, HWND_BOTTOM,
    SMTO_ABORTIFHUNG, SMTO_NORMAL, SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN,
    SM_YVIRTUALSCREEN, SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOSIZE, SWP_NOZORDER, SW_HIDE,
    SW_SHOWNOACTIVATE, WINDOWPLACEMENT, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_CLOSE,
    WM_ERASEBKGND, WM_NCDESTROY, WM_NULL, WNDCLASSEXW, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN,
    WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_MAXIMIZEBOX,
    WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU, WS_THICKFRAME,
};

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
    custom_preset_for, CustomPreset, CustomScaling, CustomWindowMode, PreviewScale, PreviewType,
    TransparentBackground, DEFAULT_AUDIO_SCALE,
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

/// Fallback off-screen parking coordinate when positive off-screen X is clamped by the target app.
/// Kept well inside i16 (-32768..32767) so WM_MOVE lParam never wraps.
const FALLBACK_PARKING: i32 = -20000;

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
    child: Child,
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
    /// Foreground window before launch (e.g. Explorer), restored if the spawned app steals focus.
    previous_foreground: isize,
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
    let previous_foreground = unsafe { GetForegroundWindow() }.0 as isize;
    let windows_before = top_level_windows();
    let processes_before = engine_processes::running().unwrap_or_default();

    let Some(mut child) = launch(&exe, path) else {
        return false;
    };
    let Some(host) = (unsafe { ensure_host() }) else {
        let _ = child.kill();
        return false;
    };

    BACKGROUND.store(color_of(preset.background).0 as isize, Ordering::SeqCst);
    let pid = child.id();
    let scaling = preset.scaling;
    let window_mode = preset.window_mode;
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
        previous_foreground,
        windows_before,
        processes_before,
        id,
    };
    // The window is carried as a number: `HWND` is a pointer and a pointer is not `Send`, and
    // what the worker does with it is a placement the Shell answers for either way.
    let host = host.0 as isize;
    std::thread::spawn(move || run_it(HWND(host as *mut _), launch, scaling, window_mode));
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
    if let Some(mut program) = program {
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

/// Spawn target normally (no SW_HIDE / CREATE_NO_WINDOW / stdio redirection) so GDI, WinUI,
/// Electron, and UWP apps initialize and paint their surface properly.
fn launch(exe: &str, path: &Path) -> Option<Child> {
    Command::new(exe).arg(path).spawn().ok()
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
fn run_it(host: HWND, launch: Launch, scaling: CustomScaling, window_mode: CustomWindowMode) {
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

    // Wait briefly for target UI thread to finish initial WM_CREATE/WM_SHOWWINDOW/WM_PAINT
    // before parking off-screen (avoids blank white surface in classic Win32 GDI apps).
    let _ = unsafe {
        SendMessageTimeoutW(
            source,
            WM_NULL,
            WPARAM(0),
            LPARAM(0),
            SMTO_ABORTIFHUNG | SMTO_NORMAL,
            250,
            None,
        )
    };
    if !is_current(launch.id) || !unsafe { IsWindow(source) }.as_bool() {
        return;
    }

    let source_pid = pid_of(source);
    SOURCE.store(source.0 as isize, Ordering::SeqCst);
    SOURCE_PID.store(source_pid, Ordering::SeqCst);
    SOURCE_IS_NEW.store(
        !launch.processes_before.contains(&source_pid),
        Ordering::SeqCst,
    );

    // Read intrinsic shape once before parking mutates window placement.
    let own = base_size(&launch.path).unwrap_or_else(|| own_size(source));
    let scale = resolve_scale(&launch.path, scaling);

    // Normal mode keeps target chrome/styles untouched; Borderless strips frame chrome.
    if window_mode == CustomWindowMode::Borderless {
        strip_the_frame(source);
    }

    let mut processes = process_tree(launch.pid);
    if !processes.contains(&source_pid) {
        processes.push(source_pid);
    }
    let mut tree_read = Instant::now();
    let mut images: HashMap<u32, Option<String>> = HashMap::new();

    let mut box_of_the_host: Option<(i32, i32, i32, i32)> = None;
    let mut registered: isize = 0;
    let mut lingering: Option<Instant> = None;
    let mut the_ask: Option<(i32, i32)> = None;

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

        // Hide secondary splash/dialog windows after source is adopted, never source itself.
        census(&launch, &processes, &mut images, source.0 as isize, true).take();
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

        let Some(layout) = layout_for(at, own, scale, follow_cursor(), bypass_placement()) else {
            hide_the_host();
            std::thread::sleep(KEEP_POLL);
            continue;
        };
        let (x, y) = (layout.pos_x, layout.pos_y);
        let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);

        if the_ask.is_none() {
            park(source, width, height);
            restore_foreground(launch.previous_foreground, &processes);
            if parked(source) {
                let ask = the_ask_for((width, height), source_size(source, window_mode));
                the_ask = Some(ask);
                if ask != (width, height) {
                    park(source, ask.0, ask.1);
                }
            }
        }
        if !parked(source) {
            if let Some(ask) = the_ask {
                park(source, ask.0, ask.1);
                restore_foreground(launch.previous_foreground, &processes);
            }
        }
        if !parked(source) {
            hide_the_host();
            std::thread::sleep(KEEP_POLL);
            continue;
        }

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
            let _ = unsafe { DwmUpdateThumbnailProperties(registered, &properties) };

            if !SHOWN.load(Ordering::SeqCst) {
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
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE | DWM_TNP_SOURCECLIENTAREAONLY,
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

/// Positive off-screen origin past the right edge of every monitor, clamped inside i16 range so
/// classic Win32 WM_MOVE lParam coordinates never wrap.
fn parking_origin(vscreen: RECT) -> (i32, i32) {
    let x = (vscreen.right + 1000).clamp(10000, 26000);
    let y = vscreen.top.clamp(-16000, 16000);
    (x, y)
}

fn is_off_screen(rect: &RECT, vscreen: RECT) -> bool {
    rect.left >= vscreen.right
        || rect.right <= vscreen.left
        || rect.top >= vscreen.bottom
        || rect.bottom <= vscreen.top
}

fn parked(window: HWND) -> bool {
    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_err() {
        return false;
    }

    is_off_screen(&rect, virtual_screen())
        && unsafe { IsWindowVisible(window) }.as_bool()
        && !unsafe { IsZoomed(window) }.as_bool()
        && !unsafe { IsIconic(window) }.as_bool()
}

/// Move `window` off-screen WITHOUT calling SW_HIDE (hiding destroys DWM's redirection surface and
/// leaves classic Win32 GDI apps blank white). Falls back to negative off-screen coords if the app
/// clamps positive X.
fn park(window: HWND, width: i32, height: i32) {
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

    let _ = unsafe {
        SetWindowPos(
            window,
            HWND_BOTTOM,
            px,
            py,
            width,
            height,
            SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };

    let mut rect = RECT::default();
    if unsafe { GetWindowRect(window, &mut rect) }.is_ok() && !is_off_screen(&rect, vscreen) {
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

    // Non-blocking repaint trigger so GDI apps redraw at the parked size.
    let _ = unsafe {
        RedrawWindow(
            window,
            None,
            None,
            RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME,
        )
    };
}

/// Restore previous foreground window (e.g. Explorer) if the spawned app stole focus.
fn restore_foreground(previous: isize, processes: &[u32]) {
    if previous == 0 {
        return;
    }
    let prev = HWND(previous as *mut _);
    if !unsafe { IsWindow(prev) }.as_bool() {
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
    /// Hide visible non-`keep` windows only after `source` is adopted (never during discovery).
    hide_others: bool,
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
    hide_others: bool,
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
        hide_others,
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
            if self.hide_others {
                let _ = unsafe { ShowWindow(hwnd, SW_HIDE) };
            }
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

        let look = census(launch, &processes, &mut images, 0, false);
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
    fn thumbnail_properties_sets_source_client_area_flag_for_both_modes() {
        let normal = thumbnail_properties(400, 300, (400, 300), CustomWindowMode::Normal);
        assert_ne!(normal.dwFlags & DWM_TNP_SOURCECLIENTAREAONLY, 0);
        assert!(!normal.fSourceClientAreaOnly.as_bool());

        let borderless = thumbnail_properties(400, 300, (400, 300), CustomWindowMode::Borderless);
        assert_ne!(borderless.dwFlags & DWM_TNP_SOURCECLIENTAREAONLY, 0);
        assert!(borderless.fSourceClientAreaOnly.as_bool());
    }

    #[test]
    fn parking_origin_is_off_screen_and_within_i16_range() {
        let vscreen = RECT {
            left: -1920,
            top: 0,
            right: 3840,
            bottom: 2160,
        };
        let (x, y) = parking_origin(vscreen);
        assert!(x >= vscreen.right);
        assert!(x <= i16::MAX as i32 - 4000);
        assert!(y >= i16::MIN as i32 && y <= i16::MAX as i32);
        let parked_rect = RECT {
            left: x,
            top: y,
            right: x + 800,
            bottom: y + 600,
        };
        assert!(is_off_screen(&parked_rect, vscreen));
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
    }
}

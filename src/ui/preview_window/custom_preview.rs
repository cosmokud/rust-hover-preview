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
//! **Why the program is begun hidden, and un-maximized before it is ever shown.** A program
//! makes its own window and shows it at its own size — and one that remembers being maximized is
//! the whole display from the moment it appears until this app moves it. That is not only a flash
//! of somebody else's window over the listing: it is a window *under the pointer*, so the hook
//! takes the hover down, the launch is killed with it, the pointer is over the file again, and
//! the whole thing repeats — a preview that blinks instead of one that is read. So the launch
//! carries `SW_HIDE` in its `STARTUPINFO` (the road a model's viewer already takes, see
//! `never_shown`), the window is hidden again the moment it is found, and a window that opens
//! maximized is *restored* rather than resized — a `SetWindowPos` on a maximized window changes
//! the size it would have been restored to and leaves it full screen. It is shown only once it
//! stands at the parked box, and it is kept there for as long as the hover is up (see
//! `keep_it_parked`).
//!
//! **Where the box is, and how big.** The placement is the one every other preview takes: the
//! pointer's own (`compute_mouse_layout`) or the focused item's (`compute_keyboard_layout`),
//! which is what reads `Tray → Placement` — the position mode, the pointer standoff and the
//! `Avoid` region — and the room of the display the hover is on. `scaling` is then read the way
//! every other kind's is: a percentage is a share of the size the program's own window opens at,
//! so `100%` is that window at the size it is rather than the whole display, and `Fit to Screen`
//! is the room the placement left. `background` is what the host paints where the program's
//! window is smaller than the box. `volume` is kept in the preset and written back; a program of
//! the user's own is not told what level to play at — there is no portable way to set another
//! process's audio — so it is a setting a future engine-backed preset can use and nothing on this
//! road reads it.
//!
//! **Why the host is a window of the hover's own.** The preview window paints a card and must
//! never take the focus; the host is the same shape as a hovered film's (`hover_film_host`):
//! a borderless, non-activating, topmost window that is shown only once the Shell has the
//! picture composed in it, and taken down with the hover. It is disabled, so a press on it is the
//! hover's press and lands on the listing behind it.
//!
//! **And the pin is not a pin.** Pressing the pin key on one of these does not freeze a picture
//! of the program: it starts the program properly, with its own frame, at its own place on the
//! desktop, and this app's preview goes away — the user asked for the file, not for a picture of
//! it (see [`open_fully`]).

use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::atomic::{AtomicIsize, AtomicU64, Ordering};
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
    CreateWindowExW, DefWindowProcW, DestroyWindow, EnumWindows, GetClientRect, GetWindow,
    GetWindowLongPtrW, GetWindowPlacement, GetWindowRect, GetWindowTextLengthW,
    GetWindowThreadProcessId, IsIconic, IsWindow, IsWindowVisible, IsZoomed, PostMessageW,
    RegisterClassExW, SetWindowLongPtrW, SetWindowPlacement, SetWindowPos, ShowWindow, GWL_EXSTYLE,
    GWL_STYLE, GW_OWNER, SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOZORDER, SW_HIDE,
    SW_SHOWNOACTIVATE, SW_SHOWNORMAL, WINDOWPLACEMENT, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP,
    WM_ERASEBKGND, WM_NCDESTROY, WNDCLASSEXW, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN,
    WS_EX_TOOLWINDOW, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU, WS_THICKFRAME,
    WS_VISIBLE,
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
/// The host's own message: the program made no window to draw, so the host stays off the screen.
const WM_THERE_IS_NOTHING_TO_SHOW: u32 = WM_APP + 42;

/// How long the program is given to make its window, and how often it is asked.
///
/// A program that takes longer than this is one whose preview never appears — a slow launch is
/// not a preview, and the hover has moved on by then. The wait runs on the worker thread, so
/// nothing of this app's is held by it. The cadence is the one a window is *taken down* at: the
/// program shows its own window the moment it has one, and every poll is a poll a window can be
/// hidden on.
const WINDOW_WAIT: Duration = Duration::from_secs(6);
const WINDOW_POLL: Duration = Duration::from_millis(15);

/// How often a parked window is looked at again, and kept at the box if the program moved itself:
/// a window shown again by its own program, or restored by it, is one the Shell would compose at
/// the place the program chose rather than at the hover's.
const KEEP_POLL: Duration = Duration::from_millis(200);

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

/// The host window, while one is up.
static HOST: AtomicIsize = AtomicIsize::new(0);
/// The program's window registered into it, while one is.
static SOURCE: AtomicIsize = AtomicIsize::new(0);
/// The thumbnail registration the Shell is drawing through, while one stands.
static THUMBNAIL: Mutex<isize> = Mutex::new(0);
/// The file on screen, so that the same hover is not a second launch.
static SHOWING: Mutex<Option<PathBuf>> = Mutex::new(None);
/// The program this app started for it, which is ended with the hover.
static CHILD: Mutex<Option<NeverShown>> = Mutex::new(None);
/// The colour the host paints where the program's window does not reach.
static BACKGROUND: AtomicIsize = AtomicIsize::new(0);
/// How many previews have been taken down: a worker that was begun for an earlier one is one
/// whose hover is gone, and it neither places nor shows anything.
static GENERATION: AtomicU64 = AtomicU64::new(0);
static CLASS_REGISTERED: Once = Once::new();

/// Whether a custom preview is on screen.
pub(super) fn is_showing() -> bool {
    SHOWING.lock().map(|held| held.is_some()).unwrap_or(false)
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

    if SHOWING
        .lock()
        .map(|held| held.as_deref() == Some(path))
        .unwrap_or(false)
    {
        return true;
    }

    hide();

    let Some(exe) = preset.exe_path().map(str::to_string) else {
        return false;
    };

    // **The program is begun before its box is known, and never on a screen.** Its window is
    // what says how large `100%` is, so the placement waits for it — and the launch is hidden so
    // that the wait is not a window over the listing (see the module note and `launch`).
    let Some(child) = launch(&exe, path) else {
        return false;
    };
    let pid = child.pid;

    BACKGROUND.store(color_of(preset.background).0 as isize, Ordering::SeqCst);
    let Some(host) = (unsafe { create_host() }) else {
        let _ = child.kill();
        return false;
    };

    if let Ok(mut held) = CHILD.lock() {
        *held = Some(child);
    }
    // **The program is this app's while it draws the hover.** It is put in the run's job so a
    // crash takes it with the app, and written down so a run that never got to end it is
    // answered for by the next one — the same charge every engine of this app's is under
    // (see `engine_processes`). A program opened by the pin is not: that one is the user's,
    // opened the way the listing opens it.
    let _ = engine_processes::adopt(pid);
    engine_processes::record_player(&exe, pid);

    if let Ok(mut held) = SHOWING.lock() {
        *held = Some(path.to_path_buf());
    }

    // Read after the show is written down, so a worker begun for this preview is current exactly
    // when this preview is the one on screen.
    let generation = GENERATION.load(Ordering::SeqCst);
    let scale = preset.scaling;
    // The window is carried as a number: `HWND` is a pointer and a pointer is not `Send`, and
    // what the worker does with it is a placement the Shell answers for either way.
    let host = host.0 as isize;
    std::thread::spawn(move || adopt(HWND(host as *mut _), pid, at, scale, generation));
    true
}

/// Take the preview down: the host window, the registration in it, and the program itself.
pub(super) fn hide() {
    // First, so that a worker still waiting on the program's window stops placing and showing.
    GENERATION.fetch_add(1, Ordering::SeqCst);

    let child = CHILD.lock().ok().and_then(|mut held| held.take());
    if let Some(child) = child {
        let _ = child.kill();
        engine_processes::forget(child.pid);
    }

    let thumbnail = THUMBNAIL
        .lock()
        .map(|mut held| std::mem::replace(&mut *held, 0))
        .unwrap_or(0);
    if thumbnail != 0 {
        let _ = unsafe { DwmUnregisterThumbnail(thumbnail) };
    }

    let host = HOST.swap(0, Ordering::SeqCst);
    if host != 0 {
        let host = HWND(host as *mut _);
        if unsafe { IsWindow(host) }.as_bool() {
            let _ = unsafe { DestroyWindow(host) };
        }
    }
    SOURCE.store(0, Ordering::SeqCst);

    if let Ok(mut held) = SHOWING.lock() {
        *held = None;
    }
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

    hide();
    let _ = Command::new(exe).arg(path).spawn();
    true
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
    never_shown::spawn(&mut command)
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

/// Make the host, off the screen and at no size, until the box the program's window is drawn in
/// is known.
unsafe fn create_host() -> Option<HWND> {
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
                windows::Win32::UI::WindowsAndMessaging::WS_EX_TOOLWINDOW.0
                    | windows::Win32::UI::WindowsAndMessaging::WS_EX_NOACTIVATE.0
                    | windows::Win32::UI::WindowsAndMessaging::WS_EX_TOPMOST.0
                    | windows::Win32::UI::WindowsAndMessaging::WS_EX_TRANSPARENT.0,
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

/// Whether the preview this worker was begun for is still the one on screen.
fn is_current(generation: u64) -> bool {
    GENERATION.load(Ordering::SeqCst) == generation && is_showing()
}

/// Find the program's window, place the host and the window at the box, and register the window
/// into the host: the whole of what makes the hover a live picture of another process.
fn adopt(host: HWND, pid: u32, at: Where, scale: PreviewScale, generation: u64) {
    let Some(source) = wait_for_a_window(pid, generation) else {
        // A program that made no window is a preview with nothing to show: the host is kept off
        // the screen rather than standing empty over the listing, and the hover that ended
        // takes it down the way it takes any preview down.
        if is_current(generation) {
            let _ =
                unsafe { PostMessageW(host, WM_THERE_IS_NOTHING_TO_SHOW, WPARAM(0), LPARAM(0)) };
        }
        return;
    };
    if !is_current(generation) {
        return;
    }

    // **Hidden before anything else is done with it.** A program that showed its own window on
    // the desktop is a window the hand is over, and a hover whose pointer left the file is a
    // hover this app is asked to take down (see the module note).
    let _ = unsafe { ShowWindow(source, SW_HIDE) };

    let Some(layout) = layout_for(at, own_size(source), scale, follow_cursor()) else {
        let _ = unsafe { PostMessageW(host, WM_THERE_IS_NOTHING_TO_SHOW, WPARAM(0), LPARAM(0)) };
        return;
    };
    if !is_current(generation) {
        return;
    }

    let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);
    let _ = unsafe {
        SetWindowPos(
            host,
            HWND::default(),
            layout.pos_x,
            layout.pos_y,
            width,
            height,
            SWP_NOZORDER | SWP_NOACTIVATE,
        )
    };

    // **A frame is not part of the picture.** The title bar, the resize border and the system
    // menu are the program's window around the thing the user asked to see, and a preview is the
    // thing rather than the window — the same answer a film's player gets, which is begun
    // without a frame at all.
    let style = unsafe { GetWindowLongPtrW(source, GWL_STYLE) } as u32;
    let stripped = style
        & !(WS_CAPTION.0
            | WS_THICKFRAME.0
            | WS_SYSMENU.0
            | WS_MINIMIZEBOX.0
            | WS_MAXIMIZEBOX.0
            | WS_CHILD.0);
    let _ = unsafe { SetWindowLongPtrW(source, GWL_STYLE, (stripped | WS_VISIBLE.0) as isize) };

    // Parked past every display and sized to the box: the window is real and drawing, and no
    // user can see it. The Shell composes it into the host from there.
    park(source, width, height);
    SOURCE.store(source.0 as isize, Ordering::SeqCst);

    let Ok(registration) = (unsafe { DwmRegisterThumbnail(host, source) }) else {
        // A registration the Shell refuses is one asked for again by the next hover: a window
        // that has not presented yet is a window with nothing to compose.
        return;
    };
    if let Ok(mut held) = THUMBNAIL.lock() {
        *held = registration;
    }

    let mut client = RECT::default();
    let _ = unsafe { GetClientRect(host, &mut client) };
    let properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE,
        rcDestination: client,
        rcSource: RECT::default(),
        opacity: 255,
        fVisible: true.into(),
        // The program's client area rather than its window: the frame has been stripped above,
        // and a program that draws its own border keeps it — that is part of its picture.
        fSourceClientAreaOnly: true.into(),
    };
    let _ = unsafe { DwmUpdateThumbnailProperties(registration, &properties) };

    // **The host is shown by its own thread, after the picture is in it**: a host shown before
    // the Shell has composed anything is its own black rectangle over the listing.
    let _ = unsafe { PostMessageW(host, WM_THE_PICTURE_IS_READY, WPARAM(0), LPARAM(0)) };

    keep_it_parked(source, width, height, generation);
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

/// Restore the window out of any maximized or minimized state, put it at the parked box at the
/// preview's size, and show it where no display reaches.
///
/// **`SetWindowPlacement` is what restores it**, and it is given the parked box as the box to
/// restore to: a `SetWindowPos` on a maximized window only changes the size it *would* have been
/// restored to and leaves the window covering the display it is maximized on, and a restore to the
/// window's own rect would be a full screen for as long as it takes to hide it again. The window
/// is hidden again at once and shown only where no display reaches.
fn park(window: HWND, width: i32, height: i32) {
    let parked = RECT {
        left: PARKING - width,
        top: PARKING,
        right: PARKING,
        bottom: PARKING + height,
    };
    let placement = WINDOWPLACEMENT {
        length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
        showCmd: SW_SHOWNORMAL.0 as u32,
        rcNormalPosition: parked,
        ..Default::default()
    };
    let _ = unsafe { SetWindowPlacement(window, &placement) };
    let _ = unsafe { ShowWindow(window, SW_HIDE) };

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
    let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
}

/// Keep the parked window parked for as long as the hover is up.
///
/// A program that shows, restores or resizes its own window after this app has placed it is one
/// whose window is no longer the picture: it is on the desktop, or at a size the preview was not
/// laid out for. It is put back at the box the same way it was put there the first time.
fn keep_it_parked(source: HWND, width: i32, height: i32, generation: u64) {
    while is_current(generation) {
        std::thread::sleep(KEEP_POLL);
        if !is_current(generation) {
            return;
        }
        if !unsafe { IsWindow(source) }.as_bool() {
            return;
        }

        let mut rect = RECT::default();
        let _ = unsafe { GetWindowRect(source, &mut rect) };
        let drifted = rect.left != PARKING - width
            || rect.top != PARKING
            || rect.right - rect.left != width
            || rect.bottom - rect.top != height
            || unsafe { IsZoomed(source) }.as_bool()
            || unsafe { IsIconic(source) }.as_bool()
            || !unsafe { IsWindowVisible(source) }.as_bool();
        if drifted {
            park(source, width, height);
        }
    }
}

/// The program's first top-level window with a title, waited for within [`WINDOW_WAIT`].
///
/// It is the window the user would have seen, which is the only window a program can be assumed
/// to draw its document in: a splash screen has no title, and a tool window is not what a program
/// opens a file with. **It need not be visible**: the launch asks for a hidden window, so the one
/// this is looking for is a window that has not been shown at all.
fn wait_for_a_window(pid: u32, generation: u64) -> Option<HWND> {
    let deadline = Instant::now() + WINDOW_WAIT;
    while Instant::now() < deadline {
        if !is_current(generation) {
            return None;
        }
        if let Some(window) = first_window_of(pid) {
            return Some(window);
        }
        std::thread::sleep(WINDOW_POLL);
    }
    None
}

/// The first titled, top-level, non-tool window of `pid`.
fn first_window_of(pid: u32) -> Option<HWND> {
    struct Search {
        pid: u32,
        found: Option<HWND>,
    }

    unsafe extern "system" fn visit(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let search = unsafe { &mut *(lparam.0 as *mut Search) };
        let mut owner = 0u32;
        unsafe { GetWindowThreadProcessId(hwnd, Some(&mut owner)) };
        if owner != search.pid {
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

    let mut search = Search { pid, found: None };
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
}

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
//! **Why the host is a window of the hover's own.** The preview window paints a card and must
//! never take the focus; the host is the same shape as a hovered film's (`hover_film_host`):
//! a borderless, non-activating, topmost window that is shown only once the Shell has the
//! picture composed in it, and taken down with the hover. It is disabled, so a press on it is the
//! hover's press and lands on the listing behind it.
//!
//! **What the preset's three settings mean here.** `scaling` is the share of the display's work
//! area the box is — a foreign window has no native size to take a percentage of, so every
//! percentage is a share of the room, which is what a sound's card means by its own scale.
//! `background` is what the host paints where the program's window is smaller than the box.
//! `volume` is kept in the preset and written back; a program of the user's own is not told what
//! level to play at — there is no portable way to set another process's audio — so it is a
//! setting a future engine-backed preset can use and nothing on this road reads it.
//!
//! **And the pin is not a pin.** Pressing the pin key on one of these does not freeze a picture
//! of the program: it starts the program properly, with its own frame, at its own place on the
//! desktop, and this app's preview goes away — the user asked for the file, not for a picture of
//! it (see [`open_fully`]).

use std::path::{Path, PathBuf};
use std::process::{Child, Command};
use std::sync::atomic::{AtomicIsize, Ordering};
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
    CreateWindowExW, DefWindowProcW, DestroyWindow, EnumWindows, GetClientRect, GetWindowLongPtrW,
    GetWindowTextLengthW, GetWindowThreadProcessId, IsWindow, IsWindowVisible, PostMessageW,
    RegisterClassExW, SetWindowLongPtrW, SetWindowPos, ShowWindow, GWL_STYLE, SWP_FRAMECHANGED,
    SWP_NOACTIVATE, SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE, WINDOW_EX_STYLE, WINDOW_STYLE,
    WM_APP, WM_ERASEBKGND, WM_NCDESTROY, WNDCLASSEXW, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN,
    WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU, WS_THICKFRAME, WS_VISIBLE,
};

use super::displays::work_area_at;
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
/// nothing of this app's is held by it.
const WINDOW_WAIT: Duration = Duration::from_secs(6);
const WINDOW_POLL: Duration = Duration::from_millis(40);

/// Where a parked window is put: past the right edge of every display, at a coordinate Windows
/// still accepts. The window is real, visible and drawing there — which is what the Shell needs
/// to compose it — and no user can see it.
const PARKING: i32 = -30000;

/// The host window, while one is up.
static HOST: AtomicIsize = AtomicIsize::new(0);
/// The program's window registered into it, while one is.
static SOURCE: AtomicIsize = AtomicIsize::new(0);
/// The thumbnail registration the Shell is drawing through, while one stands.
static THUMBNAIL: Mutex<isize> = Mutex::new(0);
/// The file on screen, so that the same hover is not a second launch.
static SHOWING: Mutex<Option<PathBuf>> = Mutex::new(None);
/// The program this app started for it, which is ended with the hover.
static CHILD: Mutex<Option<Child>> = Mutex::new(None);
/// The colour the host paints where the program's window does not reach.
static BACKGROUND: AtomicIsize = AtomicIsize::new(0);
static CLASS_REGISTERED: Once = Once::new();

/// Whether a custom preview is on screen.
pub(super) fn is_showing() -> bool {
    SHOWING.lock().map(|held| held.is_some()).unwrap_or(false)
}

/// Show `path` in the program its preset names, at the box that preset's scale asks for.
///
/// It answers whether a custom preview is what this hover is: `false` means no preset claims the
/// file and the caller goes on to lay out and load it the way it always would.
pub(super) fn show(path: &Path, cursor: (i32, i32)) -> bool {
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

    let (x, y, width, height) = box_for(&preset, cursor);
    let background = color_of(preset.background);
    BACKGROUND.store(background.0 as isize, Ordering::SeqCst);

    let Some(host) = (unsafe { create_host(x, y, width, height) }) else {
        return false;
    };
    if let Ok(mut held) = SHOWING.lock() {
        *held = Some(path.to_path_buf());
    }

    let path = path.to_path_buf();
    let spawned = Command::new(&exe).arg(&path).spawn();
    let Ok(child) = spawned else {
        hide();
        return false;
    };
    let pid = child.id();
    if let Ok(mut held) = CHILD.lock() {
        *held = Some(child);
    }

    // The window is found, framed, parked and registered on a thread of its own: a program that
    // takes a second to make its window must not hold the loop that draws every other preview.
    let host = host.0 as isize;
    std::thread::spawn(move || adopt(HWND(host as *mut _), pid, width, height));
    true
}

/// Take the preview down: the host window, the registration in it, and the program itself.
pub(super) fn hide() {
    let child = CHILD.lock().ok().and_then(|mut held| held.take());
    if let Some(mut child) = child {
        let _ = child.kill();
        let _ = child.wait();
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

/// The box a preset's preview is drawn in, from the display under the pointer and the preset's
/// own scale.
///
/// A foreign window has no native size, so every scale is a share of the display's work area:
/// the fit is the whole of it, and a percentage is that share of it. The box is placed beside
/// the pointer, kept inside the display, and never smaller than a window a program could draw in.
pub(super) fn box_for(preset: &CustomPreset, cursor: (i32, i32)) -> (i32, i32, i32, i32) {
    let room = work_area_at(cursor.0, cursor.1);
    let width = room.right - room.left;
    let height = room.bottom - room.top;

    let share = match preset.scaling {
        PreviewScale::FitToScreen => 1.0,
        PreviewScale::FitToScreenReduced(percent) | PreviewScale::Percent(percent) => {
            percent.clamp(1, 1000) as f32 / 100.0
        }
    }
    .min(1.0);

    let box_width = ((width as f32 * share) as i32).clamp(320, width);
    let box_height = ((height as f32 * share) as i32).clamp(240, height);

    // Beside the pointer rather than under it, the way every other preview is placed, and kept
    // inside the display it was measured on.
    let mut x = cursor.0 + 24;
    let mut y = cursor.1 + 24;
    if x + box_width > room.right {
        x = cursor.0 - 24 - box_width;
    }
    if y + box_height > room.bottom {
        y = cursor.1 - 24 - box_height;
    }
    let x = x.clamp(room.left, (room.right - box_width).max(room.left));
    let y = y.clamp(room.top, (room.bottom - box_height).max(room.top));

    (x, y, box_width, box_height)
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

/// Make the host at the box and leave it off the screen until the picture is composed in it.
unsafe fn create_host(x: i32, y: i32, width: i32, height: i32) -> Option<HWND> {
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
            x,
            y,
            width,
            height,
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

/// Find the program's window, strip its frame, park it at the box's size and register it into
/// the host: the whole of what makes the hover a live picture of another process.
fn adopt(host: HWND, pid: u32, width: i32, height: i32) {
    let Some(source) = wait_for_a_window(pid) else {
        // A program that made no window is a preview with nothing to show: the host is kept off
        // the screen rather than standing empty over the listing, and the hover that ended
        // takes it down the way it takes any preview down.
        let _ = unsafe { PostMessageW(host, WM_THERE_IS_NOTHING_TO_SHOW, WPARAM(0), LPARAM(0)) };
        return;
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
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            PARKING - width,
            PARKING,
            width,
            height,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };
    let _ = unsafe { ShowWindow(source, SW_SHOWNOACTIVATE) };
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
}

/// The program's first visible top-level window with a title, waited for within [`WINDOW_WAIT`].
///
/// It is the window the user would have seen, which is the only window a program can be assumed
/// to draw its document in: a splash screen has no title, and a hidden window is not what a
/// program opens a file with.
fn wait_for_a_window(pid: u32) -> Option<HWND> {
    let deadline = Instant::now() + WINDOW_WAIT;
    while Instant::now() < deadline {
        if let Some(window) = first_window_of(pid) {
            return Some(window);
        }
        std::thread::sleep(WINDOW_POLL);
    }
    None
}

/// The first visible, titled, top-level window of `pid`.
fn first_window_of(pid: u32) -> Option<HWND> {
    struct Search {
        pid: u32,
        found: Option<HWND>,
    }

    unsafe extern "system" fn visit(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let search = unsafe { &mut *(lparam.0 as *mut Search) };
        let mut owner = 0u32;
        unsafe { GetWindowThreadProcessId(hwnd, Some(&mut owner)) };
        if owner == search.pid
            && unsafe { IsWindowVisible(hwnd) }.as_bool()
            && unsafe { GetWindowTextLengthW(hwnd) } > 0
        {
            search.found = Some(hwnd);
            return BOOL(0);
        }
        BOOL(1)
    }

    let mut search = Search { pid, found: None };
    let _ = unsafe { EnumWindows(Some(visit), LPARAM(&mut search as *mut Search as isize)) };
    search.found
}

#[cfg(test)]
mod tests {
    use super::*;

    fn preset(scale: PreviewScale) -> CustomPreset {
        let mut preset = CustomPreset::named("Test");
        preset.exe = "C:/tools/program.exe".into();
        preset.set_extensions_text("rmc");
        preset.scaling = scale;
        preset
    }

    #[test]
    fn the_fit_takes_the_room_and_a_percentage_takes_its_share() {
        let (_, _, width, height) = box_for(&preset(PreviewScale::FitToScreen), (0, 0));
        assert!(width >= 320 && height >= 240);

        let (_, _, half_width, _) = box_for(&preset(PreviewScale::FitToScreenReduced(50)), (0, 0));
        assert!(half_width <= width);
    }

    #[test]
    fn the_box_stays_inside_the_display_it_was_measured_on() {
        let room = work_area_at(0, 0);
        let (x, y, width, height) = box_for(&preset(PreviewScale::FitToScreen), (0, 0));
        assert!(x >= room.left && x + width <= room.right);
        assert!(y >= room.top && y + height <= room.bottom);
    }
}

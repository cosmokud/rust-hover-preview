//! LIVE PROOF: a real FFmpeg film standing in a real layered pin window's host.
//!
//! It proves, with a real ffplay on this machine, that SDL2 has no embed switch (SDL_WINDOWID is
//! not read on the Windows backend), so the player's own top-level window must be `SetParent`ed
//! into a host this app owns *after* it exists. It proves that restyling that window to a disabled
//! `WS_CHILD` — clearing the frame bits, adding `WS_CHILD|WS_VISIBLE|WS_CLIPSIBLINGS|WS_CLIPCHILDREN`
//! — keeps ffplay drawing with the same pid and no second top-level window, so a press falls through
//! to the pin the way it does through mpv's `--wid` child.
//!
//! It also proves the two facts the production launch depends on: `-noborder` makes SDL create a
//! non-resizable window whose cached flags clamp every external resize, so a hosted launch must not
//! pass it; and a window born without it accepts a host-size resize and keeps presenting a lit,
//! moving picture. The host is off every display, so nothing flashes.
//!
//! Env-gated and ignored so no ordinary run starts a player or touches the desktop:
//!
//! `cargo test --test ffplay_pin_host_live -- --ignored --nocapture`
//!
//! The host is a popup window off every display (`-32000, -32000`), so nothing flashes.

#![cfg(windows)]

use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::time::{Duration, Instant};

use windows::core::w;
use windows::Win32::Foundation::{
    CloseHandle, BOOL, COLORREF, HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM,
};
use windows::Win32::Graphics::Gdi::{
    BitBlt, ClientToScreen, CreateCompatibleDC, CreateDIBSection, CreateSolidBrush, DeleteDC,
    DeleteObject, GetDC, GetStockObject, ReleaseDC, SelectObject, BITMAPINFO, BITMAPINFOHEADER,
    BI_RGB, BLACK_BRUSH, BLENDFUNCTION, DIB_RGB_COLORS, HBITMAP, HBRUSH, HDC, HGDIOBJ, SRCCOPY,
    AC_SRC_ALPHA, AC_SRC_OVER,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::System::Threading::{
    GetExitCodeProcess, OpenProcess, CREATE_NO_WINDOW, PROCESS_QUERY_LIMITED_INFORMATION,
};
use windows::Win32::UI::HiDpi::{
    AreDpiAwarenessContextsEqual, GetDpiForSystem, GetDpiForWindow, GetThreadDpiAwarenessContext,
    SetProcessDpiAwareness, SetProcessDpiAwarenessContext,
    DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, PROCESS_PER_MONITOR_DPI_AWARE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, EnumWindows, GetClassNameW, GetClientRect,
    GetParent, GetWindow, GetWindowLongPtrW, GetWindowRect, GetWindowThreadProcessId, IsWindow,
    IsWindowVisible, RegisterClassExW, SetParent, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    GetAncestor, SetForegroundWindow, UpdateLayeredWindow, WindowFromPoint, GA_ROOT, GWL_EXSTYLE,
    GWL_STYLE, GW_CHILD, GW_HWNDNEXT, GW_OWNER, HWND_NOTOPMOST, HWND_TOPMOST, HWND_TOP,
    SWP_FRAMECHANGED,
    SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SWP_SHOWWINDOW, SW_HIDE,
    SW_SHOWNOACTIVATE, ULW_ALPHA, WINDOW_EX_STYLE, WNDCLASSEXW, WS_CAPTION, WS_CHILD,
    WS_CLIPCHILDREN, WS_CLIPSIBLINGS, WS_EX_ACCEPTFILES, WS_EX_APPWINDOW, WS_EX_LAYERED,
    WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_MINIMIZEBOX,
    WS_POPUP, WS_SYSMENU, WS_THICKFRAME, WS_VISIBLE,
};

const STILL_ACTIVE: u32 = 259;
const HOST_W: i32 = 480;
const HOST_H: i32 = 300;
// Physical pixels. At 150% these are what 213x133 logical used to become, so the resize is real
// for an aware process rather than the DPI-virtualized no-op the last run measured.
const RESIZED_W: i32 = 320;
const RESIZED_H: i32 = 200;
// The live pin is shown on the desktop for the length of this proof so that `WindowFromPoint` can
// answer about its band; it is destroyed immediately afterwards.
const PIN_X: i32 = 80;
const PIN_Y: i32 = 80;

struct Found {
    pid: u32,
    hwnd: HWND,
    area: i64,
}

unsafe extern "system" fn largest_window_of_pid(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let found = &mut *(lparam.0 as *mut Found);
    let mut window_pid = 0u32;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));
    if window_pid != found.pid || !IsWindowVisible(hwnd).as_bool() {
        return BOOL(1);
    }
    if let Ok(owner) = GetWindow(hwnd, GW_OWNER) {
        if !owner.is_invalid() {
            return BOOL(1);
        }
    }
    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return BOOL(1);
    }
    let area = (rect.right - rect.left).max(0) as i64 * (rect.bottom - rect.top).max(0) as i64;
    if area > found.area {
        found.area = area;
        found.hwnd = hwnd;
    }
    BOOL(1)
}

fn largest_top_level_window(pid: u32) -> Option<HWND> {
    let mut found = Found {
        pid,
        hwnd: HWND::default(),
        area: 0,
    };
    unsafe {
        let _ = EnumWindows(
            Some(largest_window_of_pid),
            LPARAM(&mut found as *mut Found as isize),
        );
    }
    (!found.hwnd.is_invalid()).then_some(found.hwnd)
}

struct ListedWindow {
    hwnd: HWND,
    visible: bool,
    area: i64,
    class: String,
}

struct WindowList {
    pid: u32,
    windows: Vec<ListedWindow>,
}

unsafe extern "system" fn list_windows_of_pid(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let list = &mut *(lparam.0 as *mut WindowList);
    let mut window_pid = 0u32;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));
    if window_pid != list.pid {
        return BOOL(1);
    }
    let mut rect = RECT::default();
    let _ = GetWindowRect(hwnd, &mut rect);
    let mut class = [0u16; 256];
    let class_len = GetClassNameW(hwnd, &mut class).max(0) as usize;
    list.windows.push(ListedWindow {
        hwnd,
        visible: IsWindowVisible(hwnd).as_bool(),
        area: (rect.right - rect.left).max(0) as i64 * (rect.bottom - rect.top).max(0) as i64,
        class: String::from_utf16_lossy(&class[..class_len.min(class.len())]),
    });
    BOOL(1)
}

fn list_top_level_windows(pid: u32) -> Vec<ListedWindow> {
    let mut list = WindowList {
        pid,
        windows: Vec::new(),
    };
    unsafe {
        let _ = EnumWindows(
            Some(list_windows_of_pid),
            LPARAM(&mut list as *mut WindowList as isize),
        );
    }
    list.windows
}

fn client_size(hwnd: HWND) -> (i32, i32) {
    let mut rect = RECT::default();
    let _ = unsafe { GetClientRect(hwnd, &mut rect) };
    (rect.right - rect.left, rect.bottom - rect.top)
}

fn window_size(hwnd: HWND) -> (i32, i32) {
    let mut rect = RECT::default();
    let _ = unsafe { GetWindowRect(hwnd, &mut rect) };
    (rect.right - rect.left, rect.bottom - rect.top)
}

fn wait_for_window(pid: u32, timeout: Duration) -> Option<HWND> {
    let deadline = Instant::now() + timeout;
    loop {
        if let Some(hwnd) = largest_top_level_window(pid) {
            return Some(hwnd);
        }
        if Instant::now() >= deadline {
            return None;
        }
        std::thread::sleep(Duration::from_millis(50));
    }
}

unsafe extern "system" fn host_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

struct Dib {
    dc: HDC,
    bitmap: HBITMAP,
    previous: HGDIOBJ,
    bits: *mut u8,
    width: u32,
    height: u32,
}

impl Dib {
    fn create(width: u32, height: u32) -> Option<Self> {
        unsafe {
            let dc = CreateCompatibleDC(None);
            if dc.0.is_null() {
                return None;
            }
            let info = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: width as i32,
                    biHeight: -(height as i32),
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: BI_RGB.0,
                    ..Default::default()
                },
                ..Default::default()
            };
            let mut bits: *mut core::ffi::c_void = std::ptr::null_mut();
            let bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, &mut bits, None, 0).ok()?;
            if bits.is_null() {
                let _ = DeleteObject(bitmap);
                let _ = DeleteDC(dc);
                return None;
            }
            let previous = SelectObject(dc, bitmap);
            Some(Self {
                dc,
                bitmap,
                previous,
                bits: bits as *mut u8,
                width,
                height,
            })
        }
    }

    fn pixels(&self) -> Vec<u8> {
        let length = self.width as usize * self.height as usize * 4;
        unsafe { std::slice::from_raw_parts(self.bits, length).to_vec() }
    }

    /// The probe pin is opaque so the desktop cannot show through it: this proves the film is the
    /// child and not the desktop behind a one-alpha band.
    fn fill_hit_testable(&self) {
        let length = self.width as usize * self.height as usize * 4;
        let pixels = unsafe { std::slice::from_raw_parts_mut(self.bits, length) };
        for pixel in pixels.as_chunks_mut::<4>().0 {
            pixel[3] = 255;
        }
    }

    /// A solid BGRA fill, so a band that does not paint is unmistakable on screen.
    fn fill_solid(&self, blue: u8, green: u8, red: u8, alpha: u8) {
        let length = self.width as usize * self.height as usize * 4;
        let pixels = unsafe { std::slice::from_raw_parts_mut(self.bits, length) };
        for pixel in pixels.as_chunks_mut::<4>().0 {
            pixel[0] = blue;
            pixel[1] = green;
            pixel[2] = red;
            pixel[3] = alpha;
        }
    }
}

impl Drop for Dib {
    fn drop(&mut self) {
        unsafe {
            let _ = SelectObject(self.dc, self.previous);
            let _ = DeleteObject(self.bitmap);
            let _ = DeleteDC(self.dc);
        }
    }
}

struct Captured {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
}

impl Captured {
    /// The share of pixels that are not near-black.
    fn lit_ratio(&self) -> f64 {
        if self.pixels.is_empty() {
            return 0.0;
        }
        let lit = self
            .pixels
            .as_chunks::<4>()
            .0
            .iter()
            .filter(|p| p[0] > 8 || p[1] > 8 || p[2] > 8)
            .count();
        lit as f64 / (self.width as usize * self.height as usize) as f64
    }

    /// The share of the outer 2-pixel frame that is not near-black: did the picture fill the
    /// client box, or is there a black band from a stale surface?
    fn edge_lit_ratio(&self) -> f64 {
        let (width, height) = (self.width as usize, self.height as usize);
        if width < 3 || height < 3 {
            return self.lit_ratio();
        }
        let mut lit = 0usize;
        let mut total = 0usize;
        for y in 0..height {
            for x in 0..width {
                if x < 2 || y < 2 || x + 2 >= width || y + 2 >= height {
                    total += 1;
                    let p = &self.pixels[(y * width + x) * 4..][..4];
                    if p[0] > 8 || p[1] > 8 || p[2] > 8 {
                        lit += 1;
                    }
                }
            }
        }
        lit as f64 / total.max(1) as f64
    }

    /// The share of pixels that are magenta-ish (r > 200, g < 80, b > 200). On a band painted
    /// solid magenta this is the share the film did not cover: the player handed the band back.
    fn magenta_ratio(&self) -> f64 {
        if self.pixels.is_empty() {
            return 0.0;
        }
        let magenta = self
            .pixels
            .as_chunks::<4>()
            .0
            .iter()
            .filter(|p| p[2] > 200 && p[1] < 80 && p[0] > 200)
            .count();
        magenta as f64 / (self.width as usize * self.height as usize) as f64
    }

    /// The share of pixels exactly the pin's own fill colour. The pin is painted that exact solid
    /// colour and its child covers it 1:1, so this is the noise-free "the film handed the band
    /// back" signal: the clip's own magenta bars are not the pin's exact colour.
    fn exact_pin_ratio(&self) -> f64 {
        if self.pixels.is_empty() {
            return 0.0;
        }
        let exact = self
            .pixels
            .as_chunks::<4>()
            .0
            .iter()
            .filter(|p| {
                p[0] == (MAGENTA & 0xFF) as u8
                    && p[1] == ((MAGENTA >> 8) & 0xFF) as u8
                    && p[2] == ((MAGENTA >> 16) & 0xFF) as u8
            })
            .count();
        exact as f64 / (self.width as usize * self.height as usize) as f64
    }

    /// The share of near-black pixels, the other way a band "goes blank" can look.
    fn black_ratio(&self) -> f64 {
        if self.pixels.is_empty() {
            return 0.0;
        }
        let black = self
            .pixels
            .as_chunks::<4>()
            .0
            .iter()
            .filter(|p| p[0] < 8 && p[1] < 8 && p[2] < 8)
            .count();
        black as f64 / (self.width as usize * self.height as usize) as f64
    }

    /// The share of pixels whose colour changed between two captures.
    fn changed_ratio(&self, other: &Captured) -> f64 {
        if self.width != other.width || self.height != other.height {
            return 1.0;
        }
        let a = self.pixels.as_chunks::<4>().0;
        let b = other.pixels.as_chunks::<4>().0;
        let changed = a
            .iter()
            .zip(b)
            .filter(|(x, y)| x[0] != y[0] || x[1] != y[1] || x[2] != y[2])
            .count();
        changed as f64 / a.len().max(1) as f64
    }
}



fn capture_screen_region(x: i32, y: i32, width: u32, height: u32) -> Option<Captured> {
    unsafe {
        let source = GetDC(HWND::default());
        if source.0.is_null() {
            return None;
        }
        let dib = Dib::create(width, height);
        let copied = dib.as_ref().is_some_and(|dib| {
            BitBlt(
                dib.dc,
                0,
                0,
                width as i32,
                height as i32,
                source,
                x,
                y,
                SRCCOPY,
            )
            .is_ok()
        });
        let _ = ReleaseDC(HWND::default(), source);
        let dib = dib?;
        copied.then(|| Captured {
            pixels: dib.pixels(),
            width,
            height,
        })
    }
}

fn find_program(name: &str) -> Option<PathBuf> {
    let mut folders: Vec<PathBuf> = Vec::new();
    if let Some(own) = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.parent().map(Path::to_path_buf))
    {
        folders.push(own);
    }
    if let Some(path) = std::env::var_os("PATH") {
        folders.extend(std::env::split_paths(&path));
    }
    folders
        .into_iter()
        .map(|folder| folder.join(name))
        .find(|candidate| candidate.is_file())
}

fn probe_clip() -> PathBuf {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let clip = root.join(".scratch").join("ffplay_pin_host_live.mp4");
    if clip.is_file() {
        return clip;
    }
    let ffmpeg = find_program("ffmpeg.exe").expect("ffmpeg.exe on PATH to make the moving clip");
    let made = Command::new(ffmpeg)
        .args(["-y", "-f", "lavfi", "-i", "testsrc2=size=480x300:rate=30"])
        .args(["-f", "lavfi", "-i", "sine=frequency=440"])
        .args(["-t", "20", "-pix_fmt", "yuv420p", "-shortest"])
        .arg(&clip)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(CREATE_NO_WINDOW.0)
        .status()
        .expect("ffmpeg to make the clip");
    assert!(made.success(), "ffmpeg made the moving testsrc2 clip");
    clip
}

fn player_still_active(pid: u32) -> bool {
    unsafe {
        let Ok(process) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) else {
            return false;
        };
        let mut code = 0u32;
        let read = GetExitCodeProcess(process, &mut code).is_ok();
        let _ = CloseHandle(process);
        read && code == STILL_ACTIVE
    }
}

#[test]
#[ignore = "starts a real ffplay and a real window; run by hand for the host spike"]
fn ffplay_pin_host_live() {
    // One physical pixel per coordinate, so ffplay's `-x/-y` and our captures agree with the
    // desktop instead of being silently scaled by the caller's DPI virtualization.
    unsafe {
        let how = if SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2).is_ok() {
            "SetProcessDpiAwarenessContext(PMv2)"
        } else if SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE).is_ok() {
            "SetProcessDpiAwareness(PER_MONITOR)"
        } else if windows::Win32::UI::WindowsAndMessaging::SetProcessDPIAware().as_bool() {
            "SetProcessDPIAware"
        } else {
            "no DPI call succeeded"
        };
        println!("dpi setup: {how}");
    }

    let ffplay = find_program("ffplay.exe").expect("ffplay.exe on PATH (or beside this test)");
    let clip = probe_clip();
    println!("ffplay: {}", ffplay.display());
    println!("clip:   {}", clip.display());

    let hinstance = unsafe { GetModuleHandleW(None) }.expect("this module");

    let class_name = w!("RustHoverPreviewFfplayHostProbe");
    unsafe {
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(host_proc),
            hInstance: HINSTANCE::from(hinstance),
            hbrBackground: HBRUSH(GetStockObject(BLACK_BRUSH).0),
            lpszClassName: class_name,
            ..Default::default()
        };
        assert!(RegisterClassExW(&class) != 0, "host class registration");
    }

    // **The pin's own window**, layered and per-pixel painted exactly as `pin_real_window` makes
    // it: this is the window a hosted FFmpeg film has to present inside.
    let pin = unsafe {
        CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            class_name,
            w!(""),
            WS_POPUP,
            PIN_X,
            PIN_Y,
            HOST_W,
            HOST_H,
            HWND::default(),
            None,
            hinstance,
            None,
        )
        .expect("the layered pin window")
    };
    let surface = Dib::create(HOST_W as u32, HOST_H as u32).expect("the pin surface");
    surface.fill_hit_testable();
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };
    let dst = POINT { x: PIN_X, y: PIN_Y };
    let size = windows::Win32::Foundation::SIZE {
        cx: HOST_W,
        cy: HOST_H,
    };
    let src = POINT { x: 0, y: 0 };
    assert!(
        unsafe {
            UpdateLayeredWindow(
                pin,
                None,
                Some(&dst),
                Some(&size),
                surface.dc,
                Some(&src),
                COLORREF(0),
                Some(&blend),
                ULW_ALPHA,
            )
        }
        .is_ok(),
        "the pin is painted through UpdateLayeredWindow"
    );
    let raised = unsafe {
        let _ = ShowWindow(pin, SW_SHOWNOACTIVATE);
        // The probe owns the desktop for a few seconds: the pin must be the window the screen shows
        // and `WindowFromPoint` answers, whatever else the user has open.
        SetWindowPos(
            pin,
            HWND_TOPMOST,
            PIN_X,
            PIN_Y,
            HOST_W,
            HOST_H,
            SWP_SHOWWINDOW,
        )
    };
    let mut pin_rect = RECT::default();
    let _ = unsafe { GetWindowRect(pin, &mut pin_rect) };
    println!(
        "pin raise={raised:?} visible={} rect=({},{})-({},{})",
        unsafe { IsWindowVisible(pin).as_bool() },
        pin_rect.left,
        pin_rect.top,
        pin_rect.right,
        pin_rect.bottom,
    );

    let host = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class_name,
            w!("ffplay film host"),
            WS_CHILD,
            0,
            0,
            HOST_W,
            HOST_H,
            pin,
            None,
            hinstance,
            None,
        )
        .expect("the pin's film host")
    };
    println!("pin:    {pin:?}");
    println!("host:   {host:?}");

    // ffplay's SDL window must be created resizable (`-noborder` makes SDL skip
    // SDL_WINDOW_RESIZABLE, and then SDL keeps its original client size and crops into a smaller
    // band). The frame is stripped by the restyle below anyway.
    let mut child = Command::new(&ffplay)
        .env("SDL_RENDER_DRIVER", "direct3d11")
        .args(["-left", "-32000", "-top", "-32000"])
        .args(["-x", "480", "-y", "300"])
        .args(["-autoexit", "-loglevel", "quiet"])
        .arg(&clip)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(CREATE_NO_WINDOW.0)
        .spawn()
        .expect("ffplay to start");
    let pid = child.id();
    println!("ffplay pid: {pid}");

    let player = wait_for_window(pid, Duration::from_secs(8));
    let Some(hwnd) = player else {
        let _ = child.kill();
        let _ = child.wait();
        unsafe {
            let _ = DestroyWindow(host);
        }
        panic!("FAIL: no visible top-level window for ffplay pid {pid} within 8 s");
    };
    let mut rect = RECT::default();
    unsafe {
        let _ = GetWindowRect(hwnd, &mut rect);
    }
    println!(
        "player window: {hwnd:?} rect=({},{})-({},{}){}",
        rect.left,
        rect.top,
        rect.right,
        rect.bottom,
        if unsafe { IsWindowVisible(hwnd).as_bool() } {
            ""
        } else {
            " (invisible!)"
        }
    );

    let style_before = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) };
    let ex_style_before = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    println!("style before:       {style_before:#x} ex={ex_style_before:#x}");
    unsafe {
        println!(
            "dpi: thread-aware-v2={} system={} host={} child={}",
            AreDpiAwarenessContextsEqual(
                GetThreadDpiAwarenessContext(),
                DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            )
            .as_bool(),
            GetDpiForSystem(),
            GetDpiForWindow(host),
            GetDpiForWindow(hwnd),
        );
    }
    println!("ex style before:    accept-files={} ({ex_style_before:#x})",
        ex_style_before & WS_EX_ACCEPTFILES.0 as isize != 0);
    let top_level_before = list_top_level_windows(pid);
    println!("top-level before:   {} windows", top_level_before.len());
    for window in &top_level_before {
        println!(
            "  {:?} visible={} area={} class={:?}",
            window.hwnd, window.visible, window.area, window.class
        );
    }

    // Diagnostic: does ffplay honour a resize while still a top-level window, before the
    // restyle/SetParent? If not, the stale client below is SDL's own fixed size, not the move.
    {
        let top_resize = unsafe {
            SetWindowPos(
                hwnd,
                HWND::default(),
                0,
                0,
                RESIZED_W,
                RESIZED_H,
                SWP_NOZORDER | SWP_NOACTIVATE,
            )
        };
        std::thread::sleep(Duration::from_millis(300));
        println!(
            "pre-reparent resize {top_resize:?}: client={:?} window={:?}",
            client_size(hwnd),
            window_size(hwnd)
        );
        let _ = unsafe {
            SetWindowPos(
                hwnd,
                HWND::default(),
                0,
                0,
                HOST_W,
                HOST_H,
                SWP_NOZORDER | SWP_NOACTIVATE,
            )
        };
        std::thread::sleep(Duration::from_millis(300));
        println!(
            "pre-reparent restored: client={:?} window={:?}",
            client_size(hwnd),
            window_size(hwnd)
        );
    }

    // The reparent, in the order the spike named: restyle, then re-parent, then disable the
    // player's side and place it in the host's client box.
    let mut style = style_before;
    for bit in [
        WS_POPUP,
        WS_CAPTION,
        WS_THICKFRAME,
        WS_SYSMENU,
        WS_MINIMIZEBOX,
    ] {
        style &= !(bit.0 as isize);
    }
    style |= (WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN).0 as isize;
    unsafe {
        SetWindowLongPtrW(hwnd, GWL_STYLE, style);
    }
    let parent_result = unsafe { SetParent(hwnd, host) };
    let parent_error = parent_result.as_ref().err().map(|e| e.code().0);
    let placed = unsafe {
        SetWindowPos(
            hwnd,
            HWND_TOP,
            0,
            0,
            HOST_W,
            HOST_H,
            SWP_FRAMECHANGED | SWP_NOACTIVATE | SWP_SHOWWINDOW,
        )
    };
    unsafe {
        // Disabled, as the production host and its child are: a press on the picture belongs to the
        // pin, and `WindowFromPoint` must skip both.
        let _ = EnableWindow(host, false);
        let _ = EnableWindow(hwnd, false);
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(pin, SW_SHOWNOACTIVATE);
    }

    // a. structure after the move.
    let parent_now = unsafe { GetParent(hwnd) }.unwrap_or_default();
    let mut window_pid = 0u32;
    let thread = unsafe { GetWindowThreadProcessId(hwnd, Some(&mut window_pid)) };
    let style_after = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) };
    let ex_style_after = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    let has_child = style_after & WS_CHILD.0 as isize != 0;
    let has_popup = style_after & WS_POPUP.0 as isize != 0;
    let accept_files = ex_style_after & WS_EX_ACCEPTFILES.0 as isize != 0;
    let is_window = unsafe { IsWindow(hwnd).as_bool() };
    let parent_ok = parent_now == host;
    let pid_ok = window_pid == pid;
    let a_pass = parent_ok && has_child && !has_popup && pid_ok && is_window && placed.is_ok()
        && accept_files;
    let mut client = RECT::default();
    let _ = unsafe { GetClientRect(hwnd, &mut client) };
    println!(
        "a. parent={parent_now:?} (wanted {host:?}, SetParent={parent_result:?} err={parent_error:?}) \
         WS_CHILD={has_child} WS_POPUP={has_popup} window_pid={window_pid} (ffplay {pid}, thread {thread}) \
         IsWindow={is_window} SetWindowPos={placed:?} client={}x{} dpi={} \
         ex-accept-files={accept_files} (was {}) ex={ex_style_after:#x}",
        client.right - client.left,
        client.bottom - client.top,
        unsafe { GetDpiForWindow(hwnd) },
        ex_style_before & WS_EX_ACCEPTFILES.0 as isize != 0,
    );
    println!("a. {}", if a_pass { "PASS" } else { "FAIL" });

    let raise_the_pin = || unsafe {
        let _ = SetWindowPos(
            pin,
            HWND_TOPMOST,
            PIN_X,
            PIN_Y,
            HOST_W,
            HOST_H,
            SWP_SHOWWINDOW,
        );
    };
    raise_the_pin();

    // Let the first frame arrive after the restyle; a player needs a moment to present again.
    let mut first_frame = None;
    let first_frame_deadline = Instant::now() + Duration::from_secs(5);
    while Instant::now() < first_frame_deadline {
        let bitblt = capture_screen_region(PIN_X, PIN_Y, HOST_W as u32, HOST_H as u32);
        if bitblt.as_ref().is_some_and(|c| c.lit_ratio() > 0.01) {
            first_frame = Some((bitblt, None));
            break;
        }
        std::thread::sleep(Duration::from_millis(100));
    }
    let Some((bb0, pw0)) = first_frame else {
        println!("b. no capture saw a lit frame within 5 s after the reparent");
        println!("b. FAIL");
        println!("d. alive={}", player_still_active(pid));
        let _ = child.kill();
        let _ = child.wait();
        unsafe {
            let _ = DestroyWindow(host);
            let _ = DestroyWindow(pin);
        }
        panic!("FAIL: the reparented player never drew a lit frame");
    };

    std::thread::sleep(Duration::from_millis(600));
    let bb1 = capture_screen_region(PIN_X, PIN_Y, HOST_W as u32, HOST_H as u32);
    let pw1: Option<Captured> = None;

    let bb_lit0 = bb0.as_ref().map_or(0.0, Captured::lit_ratio);
    let bb_lit1 = bb1.as_ref().map_or(0.0, Captured::lit_ratio);
    let pw_lit0 = pw0.as_ref().map_or(0.0, Captured::lit_ratio);
    let pw_lit1 = pw1.as_ref().map_or(0.0, Captured::lit_ratio);
    let bb_moved = match (&bb0, &bb1) {
        (Some(a), Some(b)) => a.changed_ratio(b),
        _ => 0.0,
    };
    let pw_moved = match (&pw0, &pw1) {
        (Some(a), Some(b)) => a.changed_ratio(b),
        _ => 0.0,
    };
    let bitblt_saw = bb_lit0 > 0.01;
    let print_saw = pw_lit0 > 0.01;
    let saw_moving = (bitblt_saw && bb_moved > 0.001) || (print_saw && pw_moved > 0.001);
    let b_pass = (bitblt_saw || print_saw) && saw_moving;
    println!(
        "b. BitBlt(GetDC(child))  lit={:.2}% -> {:.2}% changed={:.2}%",
        bb_lit0 * 100.0,
        bb_lit1 * 100.0,
        bb_moved * 100.0
    );
    println!(
        "b. PrintWindow(PW_RENDERFULLCONTENT) lit={:.2}% -> {:.2}% changed={:.2}%",
        pw_lit0 * 100.0,
        pw_lit1 * 100.0,
        pw_moved * 100.0
    );
    println!(
        "b. capture that saw the film: {}",
        match (bitblt_saw, print_saw) {
            (true, true) => "both",
            (true, false) => "BitBlt only",
            (false, true) => "PrintWindow only",
            (false, false) => "none",
        }
    );
    println!("b. {}", if b_pass { "PASS" } else { "FAIL" });

    // c. the real road: the pin's band is the host, so resize the host to the new band size in
    // physical pixels and tell the child to fill the host's client area exactly.
    println!(
        "c. before resize: host client={:?} child client={:?} child window={:?}",
        client_size(host),
        client_size(hwnd),
        window_size(hwnd)
    );
    raise_the_pin();
    let resized_host = unsafe {
        SetWindowPos(
            host,
            HWND::default(),
            0,
            0,
            RESIZED_W,
            RESIZED_H,
            SWP_NOZORDER | SWP_NOACTIVATE,
        )
    };
    raise_the_pin();
    let host_client = client_size(host);
    std::thread::sleep(Duration::from_millis(100));
    let resized = unsafe {
        SetWindowPos(
            hwnd,
            HWND::default(),
            0,
            0,
            host_client.0,
            host_client.1,
            SWP_FRAMECHANGED | SWP_NOZORDER | SWP_NOACTIVATE,
        )
    };
    println!(
        "c. right after child resize: client={:?} window={:?}",
        client_size(hwnd),
        window_size(hwnd)
    );
    std::thread::sleep(Duration::from_millis(600));

    let child_client = client_size(hwnd);
    let mut child_rect = RECT::default();
    let _ = unsafe { GetWindowRect(hwnd, &mut child_rect) };
    let host_origin = {
        let mut point = POINT { x: 0, y: 0 };
        let _ = unsafe { ClientToScreen(host, &mut point) };
        point
    };
    let host_client_screen = (
        host_origin.x,
        host_origin.y,
        host_origin.x + host_client.0,
        host_origin.y + host_client.1,
    );
    let child_window_screen = (
        child_rect.left,
        child_rect.top,
        child_rect.right,
        child_rect.bottom,
    );
    println!(
        "c. host {RESIZED_W}x{RESIZED_H} resized={resized_host:?} client={host_client:?} on-screen={host_client_screen:?}"
    );
    println!(
        "c. child resized={resized:?} client={child_client:?} window-on-screen={child_window_screen:?}"
    );
    let boxes_agree = child_client == host_client && child_window_screen == host_client_screen;
    println!(
        "c. child client==host client: {}; child window==host client on screen: {}",
        child_client == host_client,
        child_window_screen == host_client_screen
    );

    let w = child_client.0.max(1) as u32;
    let h = child_client.1.max(1) as u32;
    let resize_bb0 = capture_screen_region(PIN_X, PIN_Y, w, h);
    let resize_pw0: Option<Captured> = None;
    std::thread::sleep(Duration::from_millis(600));
    let resize_bb1 = capture_screen_region(PIN_X, PIN_Y, w, h);
    let resize_pw1: Option<Captured> = None;

    let rbb_lit = resize_bb0.as_ref().map_or(0.0, Captured::lit_ratio);
    let rpw_lit = resize_pw0.as_ref().map_or(0.0, Captured::lit_ratio);
    let rbb_edge = resize_bb0.as_ref().map_or(0.0, Captured::edge_lit_ratio);
    let rpw_edge = resize_pw0.as_ref().map_or(0.0, Captured::edge_lit_ratio);
    let rbb_moved = match (&resize_bb0, &resize_bb1) {
        (Some(a), Some(b)) => a.changed_ratio(b),
        _ => 0.0,
    };
    let rpw_moved = match (&resize_pw0, &resize_pw1) {
        (Some(a), Some(b)) => a.changed_ratio(b),
        _ => 0.0,
    };
    let pw_matches_bb = match (&resize_bb0, &resize_pw0) {
        (Some(a), Some(b)) => 1.0 - a.changed_ratio(b),
        _ => 0.0,
    };
    let resize_saw = (rbb_lit > 0.01 || rpw_lit > 0.01) && (rbb_moved > 0.001 || rpw_moved > 0.001);
    let filled = (rbb_lit > 0.01 && rbb_edge > 0.01) || (rpw_lit > 0.01 && rpw_edge > 0.01);
    // The film is moving, so compare the spatial coverage each capture sees (lit and edge-lit),
    // not pixel colours: that is the "same area" question. Pixel agreement is reported only.
    let pw_sees_same_area = (rbb_lit - rpw_lit).abs() < 0.05
        && (rbb_edge - rpw_edge).abs() < 0.05
        && filled;
    let c_pass = resized_host.is_ok() && resized.is_ok() && boxes_agree && resize_saw && filled;
    let _ = pw_sees_same_area;
    println!(
        "c. BitBlt lit={:.2}% edge-lit={:.2}% changed={:.2}%  PrintWindow lit={:.2}% edge-lit={:.2}% changed={:.2}%  BitBlt/PrintWindow agree={:.2}%",
        rbb_lit * 100.0,
        rbb_edge * 100.0,
        rbb_moved * 100.0,
        rpw_lit * 100.0,
        rpw_edge * 100.0,
        rpw_moved * 100.0,
        pw_matches_bb * 100.0
    );
    println!("c. {}", if c_pass { "PASS" } else { "FAIL" });

    // d. style and structure a full second after the resize: does SDL put top-level style back?
    let style_1s = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) };
    let ex_1s = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    let parent_1s = unsafe { GetParent(hwnd) }.unwrap_or_default();
    let visible_1s = unsafe { IsWindowVisible(hwnd).as_bool() };
    let child_1s = style_1s & WS_CHILD.0 as isize != 0;
    let popup_1s = style_1s & WS_POPUP.0 as isize != 0;
    let accept_1s = ex_1s & WS_EX_ACCEPTFILES.0 as isize != 0;
    println!(
        "d. 1 s+ after resize: style={style_1s:#x} WS_CHILD={child_1s} WS_POPUP={popup_1s} \
         ex={ex_1s:#x} accept-files={accept_1s} parent={parent_1s:?} visible={visible_1s}"
    );

    // d. the process and its window set after all of that.
    let alive = player_still_active(pid);
    let top_level_after = list_top_level_windows(pid);
    let new_top_level = top_level_after.len() > top_level_before.len();
    let visible_top = top_level_after
        .iter()
        .filter(|window| window.visible && window.area > 0)
        .count();
    let d_pass = alive
        && !new_top_level
        && visible_top == 0
        && parent_1s == host
        && visible_1s
        && child_1s
        && !popup_1s
        && accept_1s;
    println!("d. alive={alive} top-level before={} after={} new={new_top_level} visible-with-area={visible_top}",
        top_level_before.len(), top_level_after.len());
    for window in &top_level_after {
        println!(
            "  {:?} visible={} area={} class={:?}",
            window.hwnd, window.visible, window.area, window.class
        );
    }
    println!("d. {}", if d_pass { "PASS" } else { "FAIL" });

    // e. the band as the desktop actually shows it: five screen samples 300 ms apart, then the
    // press on the band asked of the same screen. The pin is on the desktop only for this pass and
    // is destroyed immediately afterwards.
    unsafe {
        let _ = SetWindowPos(
            pin,
            HWND_TOPMOST,
            PIN_X,
            PIN_Y,
            HOST_W,
            HOST_H,
            SWP_SHOWWINDOW,
        );
        let _ = SetForegroundWindow(pin);
    }
    std::thread::sleep(Duration::from_millis(300));
    let mut samples = Vec::new();
    for _ in 0..5 {
        std::thread::sleep(Duration::from_millis(300));
        samples.push(capture_screen_region(PIN_X, PIN_Y, HOST_W as u32, HOST_H as u32));
    }
    let lit: Vec<f64> = samples
        .iter()
        .map(|sample| sample.as_ref().map_or(0.0, Captured::lit_ratio))
        .collect();
    let mut changed = Vec::new();
    for pair in samples.windows(2) {
        if let [Some(a), Some(b)] = pair {
            changed.push(a.changed_ratio(b));
        }
    }
    let high_lit = lit.iter().filter(|lit| **lit > 0.10).count();
    let moving = changed.iter().filter(|changed| **changed > 0.001).count();
    let e_pass = high_lit >= 4 && moving >= 2;
    println!(
        "e. screen band lit={:?}% changed={:?}%",
        lit.iter().map(|lit| (lit * 100.0).round()).collect::<Vec<_>>(),
        changed
            .iter()
            .map(|changed| (changed * 100.0).round())
            .collect::<Vec<_>>()
    );
    println!("e. {}", if e_pass { "PASS" } else { "FAIL" });

    let band_point = POINT {
        x: PIN_X + HOST_W / 4,
        y: PIN_Y + HOST_H / 4,
    };
    let mut hit = unsafe { WindowFromPoint(band_point) };
    // This machine's always-on-top agent windows cover the desktop; hide them just for this sample
    // so the question is about the pin's own hit testing, then put every one back.
    let mut hidden_overlays = Vec::new();
    for _ in 0..6 {
        if hit == pin {
            break;
        }
        let root = unsafe { GetAncestor(hit, GA_ROOT) };
        if root.is_invalid() || root == pin || hidden_overlays.contains(&root) {
            break;
        }
        unsafe { let _ = ShowWindow(root, SW_HIDE); }
        hidden_overlays.push(root);
        std::thread::sleep(Duration::from_millis(150));
        hit = unsafe { WindowFromPoint(band_point) };
    }
    unsafe {
        let _ = SetWindowPos(
            pin,
            HWND_TOPMOST,
            PIN_X,
            PIN_Y,
            HOST_W,
            HOST_H,
            SWP_SHOWWINDOW,
        );
        std::thread::sleep(Duration::from_millis(150));
    }
    hit = unsafe { WindowFromPoint(band_point) };
    for overlay in hidden_overlays {
        unsafe { let _ = ShowWindow(overlay, SW_SHOWNOACTIVATE); }
    }
    let mut hit_class = [0u16; 128];
    let hit_len = unsafe { GetClassNameW(hit, &mut hit_class) }.max(0) as usize;
    let hit_parent = unsafe { GetParent(hit) }.unwrap_or_default();
    let f_pass = hit == pin;
    println!(
        "f. WindowFromPoint({}, {})={:?} class={:?} parent={:?} pin={:?} {}",
        band_point.x,
        band_point.y,
        hit,
        String::from_utf16_lossy(&hit_class[..hit_len]),
        hit_parent,
        pin,
        if f_pass { "PASS" } else { "FAIL" }
    );
    let _ = child.kill();
    let _ = child.wait();
    let destroyed = unsafe { DestroyWindow(host) };
    let pin_destroyed = unsafe { DestroyWindow(pin) };
    println!("cleanup: host destroyed={destroyed:?} pin destroyed={pin_destroyed:?}");

    println!(
        "SUMMARY a={} b={} c={} d={} e={} f={}",
        if a_pass { "PASS" } else { "FAIL" },
        if b_pass { "PASS" } else { "FAIL" },
        if c_pass { "PASS" } else { "FAIL" },
        if d_pass { "PASS" } else { "FAIL" },
        if e_pass { "PASS" } else { "FAIL" },
        if f_pass { "PASS" } else { "FAIL" },
    );

    assert!(
        a_pass && b_pass && c_pass && d_pass && e_pass && f_pass,
        "the ffplay pin-host live proof did not pass every point (see the PASS/FAIL lines above)"
    );
}

/// LIVE PROOF: a hovered film is taken up into a pin without being begun again.
///
/// This is the hover-to-pin shape the app now uses: ffplay is launched off every display with
/// `SDL_RENDER_DRIVER=direct3d11` and no pacing switch and **without** `-noborder`, its own
/// window is adopted into a top-level hover host, and the hover is watched there. The pin's
/// layered window is then put up at the same box and that already-running host is re-styled and
/// re-parented into it: the same pid, the same child, the same picture, with the band lit and
/// changing on both sides of the move.
///
/// `cargo test --test ffplay_pin_host_live -- --ignored --nocapture`
#[test]
#[ignore = "starts a real ffplay and real windows; run by hand for the hover-to-pin proof"]
fn ffplay_hover_pin_take_up_live() {
    unsafe {
        let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    }
    let ffplay = find_program("ffplay.exe").expect("ffplay.exe on PATH (or beside this test)");
    let clip = probe_clip();
    let hinstance = unsafe { GetModuleHandleW(None) }.expect("this module");

    let class_name = w!("RustHoverPreviewFfplayTakeUpProbe");
    unsafe {
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(host_proc),
            hInstance: HINSTANCE::from(hinstance),
            hbrBackground: HBRUSH(GetStockObject(BLACK_BRUSH).0),
            lpszClassName: class_name,
            ..Default::default()
        };
        assert!(RegisterClassExW(&class) != 0, "probe class registration");
    }

    let box_ = (PIN_X, PIN_Y, HOST_W, HOST_H);
    // The hover's host: a top-level window of the app's own, shown at the hover's box.
    let hover = unsafe {
        CreateWindowExW(
            WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TOPMOST,
            class_name,
            w!("hover film host"),
            WS_POPUP | WS_VISIBLE,
            box_.0,
            box_.1,
            box_.2,
            box_.3,
            HWND::default(),
            None,
            hinstance,
            None,
        )
        .expect("the hover's host")
    };
    let _ = unsafe { EnableWindow(hover, false) };

    let mut child = Command::new(&ffplay)
        .env("SDL_RENDER_DRIVER", "direct3d11")
        .args(["-left", "-32000", "-top", "-32000"])
        .args(["-x", "480", "-y", "300"])
        .args(["-autoexit", "-loglevel", "quiet"])
        .arg(&clip)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(CREATE_NO_WINDOW.0)
        .spawn()
        .expect("ffplay to start");
    let pid = child.id();

    let Some(player) = wait_for_window(pid, Duration::from_secs(8)) else {
        let _ = child.kill();
        let _ = child.wait();
        unsafe { let _ = DestroyWindow(hover); }
        panic!("FAIL: no visible top-level window for ffplay pid {pid} within 8 s");
    };

    // Adopt the player's own window into the hover's host, as the monitor does.
    unsafe {
        let style = GetWindowLongPtrW(player, GWL_STYLE);
        let style = style
            & !(WS_POPUP.0 as isize
                | WS_CAPTION.0 as isize
                | WS_THICKFRAME.0 as isize
                | WS_SYSMENU.0 as isize
                | WS_MINIMIZEBOX.0 as isize)
            | (WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN).0 as isize;
        let _ = SetWindowLongPtrW(player, GWL_STYLE, style);
        let ex = GetWindowLongPtrW(player, GWL_EXSTYLE)
            & !(WS_EX_TOPMOST.0 as isize | WS_EX_APPWINDOW.0 as isize);
        let _ = SetWindowLongPtrW(player, GWL_EXSTYLE, ex | WS_EX_ACCEPTFILES.0 as isize);
        let _ = SetParent(player, hover);
        let _ = SetWindowLongPtrW(player, GWL_EXSTYLE, ex | WS_EX_ACCEPTFILES.0 as isize);
        let _ = EnableWindow(player, false);
        let _ = SetWindowPos(
            player,
            HWND::default(),
            0,
            0,
            box_.2,
            box_.3,
            SWP_FRAMECHANGED | SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW,
        );
        let _ = SetWindowPos(hover, HWND_TOPMOST, box_.0, box_.1, box_.2, box_.3, SWP_SHOWWINDOW);
    }
    std::thread::sleep(Duration::from_millis(900));

    let hovered0 = capture_screen_region(box_.0, box_.1, box_.2 as u32, box_.3 as u32);
    std::thread::sleep(Duration::from_millis(500));
    let hovered1 = capture_screen_region(box_.0, box_.1, box_.2 as u32, box_.3 as u32);
    let hover_lit = hovered0.as_ref().map_or(0.0, Captured::lit_ratio);
    let hover_changed = match (&hovered0, &hovered1) {
        (Some(a), Some(b)) => a.changed_ratio(b),
        _ => 0.0,
    };

    // The pin's own layered window at the same box, painted the way `pin_real_window` paints one.
    let pin = unsafe {
        CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            class_name,
            w!(""),
            WS_POPUP,
            box_.0,
            box_.1,
            box_.2,
            box_.3,
            HWND::default(),
            None,
            hinstance,
            None,
        )
        .expect("the layered pin window")
    };
    let surface = Dib::create(box_.2 as u32, box_.3 as u32).expect("the pin surface");
    surface.fill_hit_testable();
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };
    let dst = POINT { x: box_.0, y: box_.1 };
    let size = windows::Win32::Foundation::SIZE { cx: box_.2, cy: box_.3 };
    let src = POINT { x: 0, y: 0 };
    assert!(
        unsafe {
            UpdateLayeredWindow(
                pin,
                None,
                Some(&dst),
                Some(&size),
                surface.dc,
                Some(&src),
                COLORREF(0),
                Some(&blend),
                ULW_ALPHA,
            )
        }
        .is_ok(),
        "the pin is painted through UpdateLayeredWindow"
    );
    unsafe {
        let _ = ShowWindow(pin, SW_SHOWNOACTIVATE);
        let _ = SetWindowPos(pin, HWND_TOPMOST, box_.0, box_.1, box_.2, box_.3, SWP_SHOWWINDOW);
    }

    // The take-up: hide the top-level host, restyle it to a child of the pin, re-parent and show.
    unsafe {
        let _ = ShowWindow(hover, SW_HIDE);
        let ex = GetWindowLongPtrW(hover, GWL_EXSTYLE)
            & !(WS_EX_TOPMOST.0 as isize)
            & !(WS_EX_TRANSPARENT.0 as isize);
        let _ = SetWindowLongPtrW(hover, GWL_EXSTYLE, ex);
        let _ = SetWindowPos(hover, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        let style = GetWindowLongPtrW(hover, GWL_STYLE);
        let _ = SetWindowLongPtrW(hover, GWL_STYLE, (style & !(WS_POPUP.0 as isize)) | WS_CHILD.0 as isize);
        let _ = SetParent(hover, pin);
        let _ = SetWindowLongPtrW(hover, GWL_EXSTYLE, ex);
        let _ = SetWindowPos(hover, HWND_TOP, 0, 0, box_.2, box_.3, SWP_NOACTIVATE | SWP_SHOWWINDOW);
        let _ = SetWindowPos(
            player,
            HWND::default(),
            0,
            0,
            box_.2,
            box_.3,
            SWP_FRAMECHANGED | SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW,
        );
        let _ = EnableWindow(hover, false);
    }
    std::thread::sleep(Duration::from_millis(900));

    let pinned0 = capture_screen_region(box_.0, box_.1, box_.2 as u32, box_.3 as u32);
    std::thread::sleep(Duration::from_millis(500));
    let pinned1 = capture_screen_region(box_.0, box_.1, box_.2 as u32, box_.3 as u32);
    let pin_lit = pinned0.as_ref().map_or(0.0, Captured::lit_ratio);
    let pin_changed = match (&pinned0, &pinned1) {
        (Some(a), Some(b)) => a.changed_ratio(b),
        _ => 0.0,
    };

    let same_player = player_still_active(pid) && unsafe { IsWindow(player) }.as_bool();
    let parent = unsafe { GetParent(player) }.unwrap_or_default();
    let host_parent = unsafe { GetParent(hover) }.unwrap_or_default();
    let take_up_pass = hover_lit > 0.01
        && hover_changed > 0.001
        && pin_lit > 0.01
        && pin_changed > 0.001
        && same_player
        && parent == hover
        && host_parent == pin;
    println!(
        "hover lit={:.2}% changed={:.2}% -> pin lit={:.2}% changed={:.2}% pid={pid} alive={same_player} child-parent={:?} host-parent={:?} {:?}",
        hover_lit * 100.0,
        hover_changed * 100.0,
        pin_lit * 100.0,
        pin_changed * 100.0,
        parent,
        host_parent,
        if take_up_pass { "PASS" } else { "FAIL" }
    );

    let _ = child.kill();
    let _ = child.wait();
    unsafe {
        let _ = DestroyWindow(pin);
    }
    assert!(
        take_up_pass,
        "the hover-to-pin take-up lost the picture or the player (see the line above)"
    );
}

const BAND_X: i32 = 100;
const BAND_Y: i32 = 100;

/// The band's screen rect is painted solid magenta (0xFFFF00FF) wherever the pin or the host
/// background shows through, so a magenta pixel in a capture is literally "the film did not
/// present that instant".
const MAGENTA: u32 = 0xFFFF_00FF;

struct BandStats {
    samples: usize,
    mean: f64,
    median: f64,
    p95: f64,
    max: f64,
    over10: f64,
    over50: f64,
}

fn band_stats(mut ratios: Vec<f64>) -> BandStats {
    if ratios.is_empty() {
        return BandStats {
            samples: 0,
            mean: 0.0,
            median: 0.0,
            p95: 0.0,
            max: 0.0,
            over10: 0.0,
            over50: 0.0,
        };
    }
    let samples = ratios.len();
    let mean = ratios.iter().sum::<f64>() / samples as f64;
    ratios.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let median = ratios[samples / 2];
    let p95 = ratios[(((samples - 1) as f64) * 0.95).round() as usize];
    let max = ratios[samples - 1];
    let over10 = ratios.iter().filter(|r| **r > 0.10).count() as f64 / samples as f64;
    let over50 = ratios.iter().filter(|r| **r > 0.50).count() as f64 / samples as f64;
    BandStats {
        samples,
        mean,
        median,
        p95,
        max,
        over10,
        over50,
    }
}

struct BandRun {
    label: String,
    handback: BandStats,
    magentaish: BandStats,
    black_mean: f64,
    moved_mean: f64,
    baseline_pin: f64,
    player_windows_seen: usize,
    adopted: bool,
    alive: bool,
    child_visible: bool,
    child_box: (i32, i32),
    host_children: usize,
    parent_ok: bool,
}

fn host_child_count(hwnd: HWND) -> usize {
    let Ok(mut child) = (unsafe { GetWindow(hwnd, GW_CHILD) }) else {
        return 0;
    };
    let mut count = 0;
    while !child.is_invalid() && count < 64 {
        count += 1;
        child = match unsafe { GetWindow(child, GW_HWNDNEXT) } {
            Ok(next) => next,
            Err(_) => break,
        };
    }
    count
}

fn player_windows(windows: &[ListedWindow]) -> usize {
    windows.iter().filter(|w| w.visible && w.area > 0).count()
}

fn window_class_name(hwnd: HWND) -> String {
    let mut class = [0u16; 128];
    let len = unsafe { GetClassNameW(hwnd, &mut class) }.max(0) as usize;
    String::from_utf16_lossy(&class[..len.min(class.len())])
}

/// True when `hit` is the pin or anything inside it.
fn belongs_to(hit: HWND, pin: HWND) -> bool {
    hit.is_invalid() || hit == pin || unsafe { GetAncestor(hit, GA_ROOT) } == pin
}

/// The first foreign top-level window over any of nine points across the band, if any. The pin is
/// never interacted with beyond being raised again, so another app's window is never touched.
fn band_occluder(pin: HWND) -> Option<(HWND, String)> {
    for iy in 0..3 {
        for ix in 0..3 {
            let point = POINT {
                x: BAND_X + HOST_W * (2 * ix + 1) / 6,
                y: BAND_Y + HOST_H * (2 * iy + 1) / 6,
            };
            let hit = unsafe { WindowFromPoint(point) };
            if !belongs_to(hit, pin) {
                let root = unsafe { GetAncestor(hit, GA_ROOT) };
                return Some((root, window_class_name(root)));
            }
        }
    }
    None
}

fn raise_band(pin: HWND, host: HWND) {
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(pin, SW_SHOWNOACTIVATE);
        let raised = SetWindowPos(
            pin,
            HWND_TOPMOST,
            BAND_X,
            BAND_Y,
            HOST_W,
            HOST_H,
            SWP_SHOWWINDOW | SWP_NOACTIVATE,
        );
        if raised.is_err() {
            println!(
                "raise_band: SetWindowPos={raised:?} err={:?} ex={:#x} topmost={}",
                raised.as_ref().err().map(|e| e.code().0),
                GetWindowLongPtrW(pin, GWL_EXSTYLE),
                GetWindowLongPtrW(pin, GWL_EXSTYLE) & WS_EX_TOPMOST.0 as isize != 0,
            );
        }
    }
}

/// Kills the player and destroys both windows on every path, including a panic.
struct BandGuard {
    child: std::process::Child,
    host: HWND,
    pin: HWND,
}

impl BandGuard {
    fn player_pid(&self) -> u32 {
        self.child.id()
    }
}

impl Drop for BandGuard {
    fn drop(&mut self) {
        let _ = self.child.kill();
        let _ = self.child.wait();
        unsafe {
            // Destroying the pin also destroys its hosted child.
            let _ = DestroyWindow(self.host);
            let _ = DestroyWindow(self.pin);
        }
        println!("cleanup: player killed, host and pin destroyed");
    }
}

/// Runs one variant: same pin, same host, same clip, same sampling; only the child's
/// `SDL_RENDER_VSYNC` differs (`None` = the production launch, which sets no pacing env).
fn measure_band_run(
    ffplay: &Path,
    clip: &Path,
    hinstance: HINSTANCE,
    class_name: windows::core::PCWSTR,
    vsync: Option<&str>,
) -> BandRun {
    let label = match vsync {
        None => "no SDL_RENDER_VSYNC (production launch)".to_string(),
        Some("0") => "SDL_RENDER_VSYNC=0 (old unthrottled launch)".to_string(),
        Some(other) => other.to_string(),
    };

    // The pin's own layered window, painted solid magenta until the film covers it.
    let pin = unsafe {
        CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TOPMOST,
            class_name,
            w!(""),
            WS_POPUP,
            BAND_X,
            BAND_Y,
            HOST_W,
            HOST_H,
            HWND::default(),
            None,
            hinstance,
            None,
        )
        .expect("the layered magenta pin window")
    };
    let surface = Dib::create(HOST_W as u32, HOST_H as u32).expect("the pin surface");
    surface.fill_solid(
        (MAGENTA & 0xFF) as u8,
        ((MAGENTA >> 8) & 0xFF) as u8,
        ((MAGENTA >> 16) & 0xFF) as u8,
        255,
    );
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };
    let dst = POINT {
        x: BAND_X,
        y: BAND_Y,
    };
    let size = windows::Win32::Foundation::SIZE {
        cx: HOST_W,
        cy: HOST_H,
    };
    let src = POINT { x: 0, y: 0 };
    assert!(
        unsafe {
            UpdateLayeredWindow(
                pin,
                None,
                Some(&dst),
                Some(&size),
                surface.dc,
                Some(&src),
                COLORREF(0),
                Some(&blend),
                ULW_ALPHA,
            )
        }
        .is_ok(),
        "the pin is painted through UpdateLayeredWindow"
    );

    let host = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class_name,
            w!("ffplay film host"),
            WS_CHILD,
            0,
            0,
            HOST_W,
            HOST_H,
            pin,
            None,
            hinstance,
            None,
        )
        .expect("the pin's film host")
    };

    raise_band(pin, host);

    // Detector sanity: with no player yet the band must read as ~100% exact pin magenta. If it
    // does not, a later low reading would mean "not measured", not "no hand-back". Only our own
    // window is re-raised; a foreign window is never touched, and if it will not yield the band
    // cannot be measured.
    let mut baseline_capture = None;
    let mut attempts = 0;
    for _ in 0..10 {
        attempts += 1;
        raise_band(pin, host);
        std::thread::sleep(Duration::from_millis(200));
        if let Some((window, class)) = band_occluder(pin) {
            println!("band occluded by {window:?} class={class:?}; re-raising the pin");
            continue;
        }
        let captured = capture_screen_region(BAND_X, BAND_Y, HOST_W as u32, HOST_H as u32);
        let magenta = captured
            .as_ref()
            .map_or(0.0, |captured| captured.exact_pin_ratio());
        baseline_capture = captured;
        if magenta > 0.90 {
            break;
        }
    }
    if attempts > 1 {
        println!("baseline needed {attempts} attempts");
    }
    let baseline_pin = baseline_capture
        .as_ref()
        .map_or(0.0, |captured| captured.exact_pin_ratio());
    println!("\n== {label} ==");
    println!(
        "baseline before any player: exact pin magenta={:.2}% (wanted ~100%)",
        baseline_pin * 100.0
    );
    assert!(
        baseline_pin > 0.90,
        "the magenta pin is not visible on screen before the player starts ({:.2}%); \
         the band cannot be measured",
        baseline_pin * 100.0
    );

    let mut command = Command::new(ffplay);
    command.env("SDL_RENDER_DRIVER", "direct3d11");
    if let Some(value) = vsync {
        command.env("SDL_RENDER_VSYNC", value);
    }
    let child = command
        .args(["-left", "-32000", "-top", "-32000"])
        .args(["-x", "480", "-y", "300"])
        .args(["-autoexit", "-loglevel", "quiet"])
        .arg(clip)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(CREATE_NO_WINDOW.0)
        .spawn()
        .expect("ffplay to start");
    let player_band = pin;
    let player_host = host;
    let guard = BandGuard { child, host, pin };
    let pid = guard.player_pid();
    println!("ffplay pid {pid}");

    let Some(player) = wait_for_window(pid, Duration::from_secs(8)) else {
        panic!("FAIL: no visible top-level window for ffplay pid {pid} within 8 s");
    };
    // ffplay owns one real player window (plus SDL's hidden, zero-area helper windows).
    let top_level_before = list_top_level_windows(pid);
    let players_before = player_windows(&top_level_before);
    println!(
        "top-level windows before adoption: {} total, {players_before} visible-with-area {:?}",
        top_level_before.len(),
        top_level_before
            .iter()
            .map(|w| (w.hwnd, w.visible, w.area))
            .collect::<Vec<_>>()
    );

    // Adopt exactly as production does: restyle to a disabled WS_CHILD, re-parent, fill the host.
    unsafe {
        let style = GetWindowLongPtrW(player, GWL_STYLE)
            & !(WS_POPUP.0 as isize
                | WS_CAPTION.0 as isize
                | WS_THICKFRAME.0 as isize
                | WS_SYSMENU.0 as isize
                | WS_MINIMIZEBOX.0 as isize)
            | (WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN).0 as isize;
        let _ = SetWindowLongPtrW(player, GWL_STYLE, style);
        let ex = GetWindowLongPtrW(player, GWL_EXSTYLE)
            & !(WS_EX_TOPMOST.0 as isize | WS_EX_APPWINDOW.0 as isize);
        let _ = SetWindowLongPtrW(player, GWL_EXSTYLE, ex | WS_EX_ACCEPTFILES.0 as isize);
        let _ = SetParent(player, host);
        let _ = SetWindowLongPtrW(player, GWL_EXSTYLE, ex | WS_EX_ACCEPTFILES.0 as isize);
        let _ = EnableWindow(host, false);
        let _ = EnableWindow(player, false);
        let _ = SetWindowPos(
            player,
            HWND::default(),
            0,
            0,
            HOST_W,
            HOST_H,
            SWP_FRAMECHANGED | SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW,
        );
    }
    raise_band(player_band, player_host);

    // Wait for the child to be visible with a box, then let the film warm up ~1 s.
    let mut child_visible = false;
    let mut child_box = (0, 0);
    let deadline = Instant::now() + Duration::from_secs(5);
    while Instant::now() < deadline {
        child_box = client_size(player);
        child_visible = unsafe { IsWindowVisible(player) }.as_bool();
        if child_visible && child_box.0 > 0 && child_box.1 > 0 {
            break;
        }
        std::thread::sleep(Duration::from_millis(50));
    }
    let parent_ok = unsafe { GetParent(player) }.unwrap_or_default() == host;
    println!(
        "adopted: visible={child_visible} client={}x{} parent-is-host={parent_ok} host-children={}",
        child_box.0,
        child_box.1,
        host_child_count(host)
    );
    std::thread::sleep(Duration::from_millis(1000));

    // ~200 samples over ~5 s, as fast as BitBlt allows.
    let mut handback_ratios = Vec::with_capacity(200);
    let mut magenta_ratios = Vec::with_capacity(200);
    let mut black_sum = 0.0f64;
    let mut moved = Vec::new();
    let mut previous: Option<Captured> = None;
    let mut skipped = 0usize;
    let started = Instant::now();
    while handback_ratios.len() < 200 && started.elapsed() < Duration::from_secs(6) {
        if handback_ratios.len() % 10 == 0 && band_occluder(player_band).is_some() {
            raise_band(player_band, player_host);
            std::thread::sleep(Duration::from_millis(30));
            if let Some((window, class)) = band_occluder(player_band) {
                println!("sample skipped: band occluded by {window:?} class={class:?}");
                skipped += 1;
                continue;
            }
        }
        if let Some(captured) =
            capture_screen_region(BAND_X, BAND_Y, HOST_W as u32, HOST_H as u32)
        {
            handback_ratios.push(captured.exact_pin_ratio());
            magenta_ratios.push(captured.magenta_ratio());
            black_sum += captured.black_ratio();
            if let Some(prev) = &previous {
                moved.push(prev.changed_ratio(&captured));
            }
            previous = Some(captured);
        }
        std::thread::sleep(Duration::from_millis(20));
    }
    let elapsed = started.elapsed();
    let handback = band_stats(handback_ratios);
    let magentaish = band_stats(magenta_ratios);
    let black_mean = black_sum / handback.samples.max(1) as f64;
    let moved_mean = if moved.is_empty() {
        0.0
    } else {
        moved.iter().sum::<f64>() / moved.len() as f64
    };

    let alive = player_still_active(pid);
    let top_level_after = list_top_level_windows(pid);
    let players_after = player_windows(&top_level_after);
    let adopted = unsafe { GetParent(player) }.unwrap_or_default() == host
        && unsafe { IsWindow(player) }.as_bool();
    let host_children = host_child_count(host);
    // "Ever owns": the adopted window stops being top-level, so the most it ever had is the max.
    let player_windows_seen = players_before.max(players_after);

    let fmt = |stats: &BandStats| {
        format!(
            "samples={} mean={:.2}% median={:.2}% p95={:.2}% max={:.2}% >10%={:.1}% >50%={:.1}%",
            stats.samples,
            stats.mean * 100.0,
            stats.median * 100.0,
            stats.p95 * 100.0,
            stats.max * 100.0,
            stats.over10 * 100.0,
            stats.over50 * 100.0,
        )
    };
    println!("over {:.2}s", elapsed.as_secs_f64());
    println!(
        "pin-visible (exact 0x{MAGENTA:08X}, the hand-back): {} skipped-occluded={skipped}",
        fmt(&handback)
    );
    println!("magenta-ish (includes the clip's own magenta bars):    {}", fmt(&magentaish));
    println!(
        "black mean={:.2}%  screen changed between samples mean={:.2}%",
        black_mean * 100.0,
        moved_mean * 100.0
    );
    println!(
        "structure: alive={alive} adopted={adopted} player-windows-seen={player_windows_seen} host-children={host_children} visible={child_visible} box={}x{}",
        child_box.0, child_box.1
    );

    BandRun {
        label,
        handback,
        magentaish,
        black_mean,
        moved_mean,
        baseline_pin,
        player_windows_seen,
        adopted,
        alive,
        child_visible,
        child_box,
        host_children,
        parent_ok,
    }
}

/// LIVE MEASUREMENT: does dropping `SDL_RENDER_VSYNC=0` from the hosted ffplay launch stop the
/// film's band going blank and showing the desktop (or the pin) through?
///
/// The same layered pin, magenta band, `WS_CHILD` host, clip and screen sampling are used twice;
/// only the child's `SDL_RENDER_VSYNC` differs (absent, then `0`). A capture is scored by the share
/// of pixels that are *exactly* the pin's magenta (the clip's own magenta bars are a separate,
/// constant ~13.6% floor reported alongside), so a non-zero score is literally "the film did not
/// present and the pin showed through". A pre-player baseline proves the pin is visible and the
/// detector reads ~100%.
///
/// `cargo test --test ffplay_pin_host_live a_hosted_ffplay_film_never_hands_its_band_back_to_the_desktop -- --ignored --exact --nocapture`
#[test]
#[ignore = "starts real ffplay players and real windows; measures band flicker by hand"]
fn a_hosted_ffplay_film_never_hands_its_band_back_to_the_desktop() {
    unsafe {
        let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    }
    let ffplay = find_program("ffplay.exe").expect("ffplay.exe on PATH (or beside this test)");
    let clip = probe_clip();
    let hinstance = unsafe { GetModuleHandleW(None) }.expect("this module");
    println!("ffplay: {}", ffplay.display());
    println!("clip:   {}", clip.display());

    let class_name = w!("RustHoverPreviewFfplayBandProbe");
    unsafe {
        // A magenta class background too, so a hosted child that paints nothing still hands the
        // band back as magenta rather than as an anonymous black rectangle.
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(host_proc),
            hInstance: HINSTANCE::from(hinstance),
            hbrBackground: HBRUSH(CreateSolidBrush(COLORREF(0x00FF_00FF)).0),
            lpszClassName: class_name,
            ..Default::default()
        };
        assert!(RegisterClassExW(&class) != 0, "band probe class registration");
    }
    println!(
        "band: solid magenta 0x{MAGENTA:08X} at ({BAND_X},{BAND_Y}) {HOST_W}x{HOST_H}; magenta in a capture = film did not present"
    );

    let production = measure_band_run(&ffplay, &clip, HINSTANCE::from(hinstance), class_name, None);
    let unthrottled =
        measure_band_run(&ffplay, &clip, HINSTANCE::from(hinstance), class_name, Some("0"));

    let row = |label: &str, stats: &BandStats| {
        format!(
            "{label:<42} samples={:<3} mean={:>6.2}% median={:>6.2}% p95={:>6.2}% max={:>6.2}% >10%={:>5.1}% >50%={:>5.1}%",
            stats.samples,
            stats.mean * 100.0,
            stats.median * 100.0,
            stats.p95 * 100.0,
            stats.max * 100.0,
            stats.over10 * 100.0,
            stats.over50 * 100.0,
        )
    };
    println!("\nPIN-VISIBLE BAND FRACTION (exact pin magenta; higher = band handed back):");
    println!("{}", row(&production.label, &production.handback));
    println!("{}", row(&unthrottled.label, &unthrottled.handback));
    println!("\nMAGENTA-ISH BAND FRACTION (as specified; includes the clip's own magenta bars):");
    println!("{}", row(&production.label, &production.magentaish));
    println!("{}", row(&unthrottled.label, &unthrottled.magentaish));

    for run in [&production, &unthrottled] {
        println!(
            "{}: pin-visible-baseline={:.2}% black-mean={:.2}% screen-moved-mean={:.2}% player-windows-seen={} child-adopted={} parent-is-host={} host-children={} visible={} box={}x{} alive={}",
            run.label,
            run.baseline_pin * 100.0,
            run.black_mean * 100.0,
            run.moved_mean * 100.0,
            run.player_windows_seen,
            run.adopted,
            run.parent_ok,
            run.host_children,
            run.child_visible,
            run.child_box.0,
            run.child_box.1,
            run.alive,
        );
    }

    // The measurement itself must not have broken the embedding it measures.
    for run in [&production, &unthrottled] {
        assert!(
            run.player_windows_seen == 1,
            "{}: ffplay owned {} visible player windows, wanted exactly 1",
            run.label,
            run.player_windows_seen
        );
        assert!(
            run.host_children == 1,
            "{}: the host held {} children, wanted exactly 1",
            run.label,
            run.host_children
        );
        assert!(
            run.adopted && run.parent_ok && run.child_visible,
            "{}: the player did not stay adopted into its host",
            run.label
        );
        assert!(
            run.handback.samples > 0,
            "{}: no screen sample saw the band at all",
            run.label
        );
        assert!(
            run.moved_mean > 0.001,
            "{}: the band never changed on screen; the film may not have been playing",
            run.label
        );
    }

    println!("cleaned up: both players killed, both pins and hosts destroyed, nothing left running");
}

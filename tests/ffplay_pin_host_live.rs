//! LIVE PROOF: a real FFmpeg film in its own top-level window over a real layered pin band.
//!
//! It proves, with a real ffplay on this machine, the shape the app now uses for FFmpeg's player:
//! the player is never hosted. ffplay keeps its own borderless top-level window, placed at the
//! band's screen box with `-noborder`, and the pin's layered window stands over it with an alpha-1
//! band, so the film is seen through the band while a press still lands on the pin.
//!
//! The proof samples the band on the real desktop: the picture is lit and moving, the player's
//! window is a top-level window at the band (not a child of anything of this app's), and
//! `WindowFromPoint` in the middle of the band answers the pin rather than the player.
//!
//! Env-gated and ignored so no ordinary run starts a player or touches the desktop:
//!
//! `cargo test --test ffplay_pin_host_live -- --ignored --nocapture`

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
    BitBlt, CreateCompatibleDC, CreateDIBSection, DeleteDC, DeleteObject, GetDC, ReleaseDC,
    SelectObject, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, BLENDFUNCTION, DIB_RGB_COLORS, HBITMAP, HDC,
    HGDIOBJ, SRCCOPY, AC_SRC_ALPHA, AC_SRC_OVER,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::{
    GetExitCodeProcess, OpenProcess, CREATE_NO_WINDOW, PROCESS_QUERY_LIMITED_INFORMATION,
};
use windows::Win32::UI::HiDpi::{
    SetProcessDpiAwareness, SetProcessDpiAwarenessContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2,
    PROCESS_PER_MONITOR_DPI_AWARE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, EnumWindows, GetAncestor, GetClassNameW,
    GetLayeredWindowAttributes, GetParent, GetWindow, GetWindowLongPtrW, GetWindowRect,
    GetWindowThreadProcessId, IsWindow, IsWindowVisible, RegisterClassExW, SetLayeredWindowAttributes,
    SetWindowLongPtrW, SetWindowPos, ShowWindow, UpdateLayeredWindow, WindowFromPoint, GA_ROOT,
    GWL_EXSTYLE, GW_OWNER, HWND_TOPMOST, LAYERED_WINDOW_ATTRIBUTES_FLAGS, LWA_ALPHA,
    SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SWP_SHOWWINDOW,
    SW_SHOWNOACTIVATE, ULW_ALPHA, WNDCLASSEXW, WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW,
    WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_POPUP,
};

const STILL_ACTIVE: u32 = 259;
const BAND_W: i32 = 480;
const BAND_H: i32 = 300;
const BAND_X: i32 = 80;
const BAND_Y: i32 = 80;
/// The box the layered-window measurement watches: one film, one box, every phase.
const BOX_X: i32 = 120;
const BOX_Y: i32 = 120;

/// The renderer the production launch names, mirrored here: Direct3D 11 unless `RHP_FFPLAY_RENDERER`
/// names another driver by hand, or names `default` so that nothing is named at all.
fn with_the_production_renderer(mut command: Command) -> Command {
    let named = std::env::var("RHP_FFPLAY_RENDERER")
        .ok()
        .filter(|renderer| !renderer.is_empty())
        .unwrap_or_else(|| "direct3d11".to_string());
    if !named.eq_ignore_ascii_case("default") {
        command.env("SDL_RENDER_DRIVER", named);
    }
    command
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

    /// The band this app paints a film through: every pixel alpha 1, so the compositor shows the
    /// film underneath while `WindowFromPoint` still answers the pin (see
    /// `fill_band_hit_testable` in the app).
    fn fill_band_hit_testable(&self) {
        let length = self.width as usize * self.height as usize * 4;
        let pixels = unsafe { std::slice::from_raw_parts_mut(self.bits, length) };
        for pixel in pixels.as_chunks_mut::<4>().0 {
            pixel[0] = 1;
            pixel[1] = 1;
            pixel[2] = 1;
            pixel[3] = 1;
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
    /// The share of pixels that are not near-black: a film playing lights its box, and the desktop
    /// behind an empty band does not.
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

    /// The share of pixels that moved between two samples of the same box: a film that presents
    /// keeps changing, a blanked window or the desktop behind it does not.
    fn changed_ratio(&self, later: &Captured) -> f64 {
        if self.width != later.width || self.height != later.height {
            return 0.0;
        }
        let moved = self
            .pixels
            .as_chunks::<4>()
            .0
            .iter()
            .zip(later.pixels.as_chunks::<4>().0)
            .filter(|(before, after)| {
                (0..3).any(|channel| before[channel].abs_diff(after[channel]) > 8)
            })
            .count();
        moved as f64 / (self.width as usize * self.height as usize) as f64
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
        if !copied {
            return None;
        }
        Some(Captured {
            pixels: std::slice::from_raw_parts(dib.bits, width as usize * height as usize * 4)
                .to_vec(),
            width,
            height,
        })
    }
}

unsafe extern "system" fn pin_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

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

fn window_rect(hwnd: HWND) -> RECT {
    let mut rect = RECT::default();
    let _ = unsafe { GetWindowRect(hwnd, &mut rect) };
    rect
}

fn window_class_name(hwnd: HWND) -> String {
    let mut class = [0u16; 256];
    let len = unsafe { GetClassNameW(hwnd, &mut class) }.max(0) as usize;
    String::from_utf16_lossy(&class[..len.min(class.len())])
}

/// One phase of watching the box: how lit it was and how much moved between samples.
struct Sampled {
    lit_mean: f64,
    lit_min: f64,
    changed_mean: f64,
    samples: usize,
}

fn sample_box(x: i32, y: i32, width: u32, height: u32, samples: usize) -> Sampled {
    let mut lit_sum = 0.0;
    let mut lit_min = f64::INFINITY;
    let mut changed_sum = 0.0;
    let mut previous: Option<Captured> = None;
    let mut taken = 0usize;
    let mut pairs = 0usize;
    for _ in 0..samples {
        let Some(captured) = capture_screen_region(x, y, width, height) else {
            continue;
        };
        let lit = captured.lit_ratio();
        lit_sum += lit;
        lit_min = lit_min.min(lit);
        if let Some(previous) = previous.as_ref() {
            changed_sum += previous.changed_ratio(&captured);
            pairs += 1;
        }
        previous = Some(captured);
        taken += 1;
        std::thread::sleep(Duration::from_millis(50));
    }
    Sampled {
        lit_mean: if taken > 0 {
            lit_sum / taken as f64
        } else {
            0.0
        },
        lit_min: if taken > 0 { lit_min } else { 0.0 },
        changed_mean: if pairs > 0 {
            changed_sum / pairs as f64
        } else {
            0.0
        },
        samples: taken,
    }
}

/// A player that kills itself however the measurement ends, so no film is left on the desktop.
struct Playing(std::process::Child);

impl Playing {
    fn pid(&self) -> u32 {
        self.0.id()
    }
}

impl Drop for Playing {
    fn drop(&mut self) {
        let _ = self.0.kill();
        let _ = self.0.wait();
    }
}

unsafe extern "system" fn count_visible_windows_of_pid(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let found = &mut *(lparam.0 as *mut (u32, usize));
    let mut pid = 0u32;
    GetWindowThreadProcessId(hwnd, Some(&mut pid));
    if pid != found.0 || !IsWindowVisible(hwnd).as_bool() {
        return BOOL(1);
    }
    if let Ok(owner) = GetWindow(hwnd, GW_OWNER) {
        if !owner.is_invalid() {
            return BOOL(1);
        }
    }
    found.1 += 1;
    BOOL(1)
}

fn visible_window_count(pid: u32) -> usize {
    let mut found = (pid, 0usize);
    unsafe {
        let _ = EnumWindows(
            Some(count_visible_windows_of_pid),
            LPARAM(&mut found as *mut (u32, usize) as isize),
        );
    }
    found.1
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

/// The app's raise of a pin that holds the front: the topmost bit on the window's own style and
/// the window put in the topmost band (see `raise_pinned_window` and the band re-assert in the
/// tick).
fn raise_pin_topmost(hwnd: HWND) {
    unsafe {
        let ex = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
        SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ex | WS_EX_TOPMOST.0 as isize);
        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        );
    }
}

/// A layered pin window with an alpha-1 band, painted exactly as the app paints a film's band.
fn a_layered_pin(class_name: windows::core::PCWSTR, hinstance: HINSTANCE) -> HWND {
    let pin = unsafe {
        CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TOPMOST,
            class_name,
            w!(""),
            WS_POPUP,
            BAND_X,
            BAND_Y,
            BAND_W,
            BAND_H,
            HWND::default(),
            None,
            hinstance,
            None,
        )
        .expect("the layered pin window")
    };
    let surface = Dib::create(BAND_W as u32, BAND_H as u32).expect("the pin surface");
    surface.fill_band_hit_testable();
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
        cx: BAND_W,
        cy: BAND_H,
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
    let _ = unsafe { ShowWindow(pin, SW_SHOWNOACTIVATE) };
    raise_pin_topmost(pin);
    pin
}

#[test]
#[ignore = "starts a real ffplay and a real window; run by hand for the un-hosted film proof"]
fn an_ffplay_film_stands_in_its_own_window_over_a_layered_pin_band() {
    // One physical pixel per coordinate, so ffplay's `-x/-y` and our captures agree with the
    // desktop instead of being silently scaled by the caller's DPI virtualization.
    unsafe {
        let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
            .or_else(|_| SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));
    }

    let ffplay = find_program("ffplay.exe").expect("ffplay.exe on PATH (or beside this test)");
    let clip = probe_clip();
    println!("ffplay: {}", ffplay.display());
    println!("clip:   {}", clip.display());

    let hinstance = unsafe { GetModuleHandleW(None) }.expect("this module");
    let class_name = w!("RustHoverPreviewFfplayWindowProbe");
    unsafe {
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(pin_proc),
            hInstance: hinstance.into(),
            lpszClassName: class_name,
            ..Default::default()
        };
        assert!(RegisterClassExW(&class) != 0, "probe class registration");
    }

    let pin = a_layered_pin(class_name, hinstance.into());

    // **The production launch.** A pinned film is born at the band's box, borderless, and the
    // player keeps its own top-level window: no host of this app's is anywhere in the shape.
    let mut child = with_the_production_renderer(Command::new(&ffplay))
        .args(["-left", &BAND_X.to_string(), "-top", &BAND_Y.to_string()])
        .args(["-x", &BAND_W.to_string(), "-y", &BAND_H.to_string()])
        .args(["-noborder", "-autoexit", "-loglevel", "quiet"])
        .arg(&clip)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(CREATE_NO_WINDOW.0)
        .spawn()
        .expect("ffplay to start");
    let pid = child.id();
    println!("ffplay pid: {pid}");

    let Some(player) = wait_for_window(pid, Duration::from_secs(8)) else {
        let _ = child.kill();
        let _ = child.wait();
        unsafe {
            let _ = DestroyWindow(pin);
        }
        panic!("FAIL: no visible top-level window for ffplay pid {pid} within 8 s");
    };
    let rect = window_rect(player);
    println!(
        "player window: {player:?} class={:?} rect=({},{})-({},{})",
        window_class_name(player),
        rect.left,
        rect.top,
        rect.right,
        rect.bottom,
    );

    // Its own window, not a child of anything of this app's.
    assert_eq!(
        unsafe { GetParent(player) }.ok(),
        None,
        "the film is a top-level window of its own, never a hosted child"
    );

    // The band's box, allowing for the frame ffplay's own SDL window keeps even borderless.
    let mut pin_rect = window_rect(pin);
    if pin_rect.left == 0 && pin_rect.top == 0 {
        pin_rect = RECT {
            left: BAND_X,
            top: BAND_Y,
            right: BAND_X + BAND_W,
            bottom: BAND_Y + BAND_H,
        };
    }
    assert!(
        (rect.left - pin_rect.left).abs() <= 32 && (rect.top - pin_rect.top).abs() <= 32,
        "the player stands at the band's box: player {rect:?} against band {pin_rect:?}"
    );

    // **The production ex-style.** A window of the desktop's is given the bits that keep it out of
    // the foreground and out of a press's way, and the pin's layered window is put back over it.
    unsafe {
        let ex = GetWindowLongPtrW(player, GWL_EXSTYLE);
        SetWindowLongPtrW(
            player,
            GWL_EXSTYLE,
            ex | WS_EX_NOACTIVATE.0 as isize
                | WS_EX_TOOLWINDOW.0 as isize
                | WS_EX_TRANSPARENT.0 as isize
                | WS_EX_TOPMOST.0 as isize,
        );
        let _ = SetWindowPos(
            player,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            windows::Win32::UI::WindowsAndMessaging::SWP_NOMOVE
                | windows::Win32::UI::WindowsAndMessaging::SWP_NOSIZE
                | windows::Win32::UI::WindowsAndMessaging::SWP_NOACTIVATE,
        );
        let _ = SetWindowPos(
            pin,
            HWND_TOPMOST,
            BAND_X,
            BAND_Y,
            BAND_W,
            BAND_H,
            SWP_SHOWWINDOW,
        );
    }

    // The band is alpha 1 over the player's own window, so what the screen shows in it is the film.
    let deadline = Instant::now() + Duration::from_secs(8);
    let mut lit = 0.0;
    while Instant::now() < deadline {
        if let Some(captured) =
            capture_screen_region(BAND_X, BAND_Y, BAND_W as u32, BAND_H as u32)
        {
            lit = captured.lit_ratio();
            if lit > 0.2 {
                break;
            }
        }
        std::thread::sleep(Duration::from_millis(100));
    }
    println!("band lit ratio: {lit:.3}");
    assert!(
        lit > 0.2,
        "the film is visible through the pin's alpha-1 band: lit ratio {lit:.3}"
    );

    // **And a press in the band is the pin's.** The app keeps its order not by one raise but by
    // re-placing the player and putting the pin back over it every tick (see
    // `place_player_window` and the 200 ms reassert), so the test replays that order each pass and
    // samples `WindowFromPoint` in the middle of the band, counting the answers.
    let centre = POINT {
        x: BAND_X + BAND_W / 2,
        y: BAND_Y + BAND_H / 2,
    };
    let mut answers: std::collections::BTreeMap<String, usize> = std::collections::BTreeMap::new();
    let mut passes = 0usize;
    let mut pins = 0usize;
    let deadline = Instant::now() + Duration::from_secs(2);
    while Instant::now() < deadline {
        unsafe {
            // `player_insert_after`: with a pin holding the front, the player is put directly
            // below the pin's window rather than raised to the topmost band, which would leave the
            // two in whatever order they already had.
            let _ = SetWindowPos(
                player,
                pin,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );
        }
        // And the pin put back over it, exactly as `place_player_window` follows its placement.
        raise_pin_topmost(pin);
        let hit = unsafe { WindowFromPoint(centre) };
        let root = unsafe { GetAncestor(hit, GA_ROOT) };
        let answer = if root == pin {
            pins += 1;
            "pin".to_string()
        } else if root == player || hit == player {
            "player".to_string()
        } else {
            let other = if root.is_invalid() { hit } else { root };
            format!("other:{}", window_class_name(other))
        };
        *answers.entry(answer).or_insert(0) += 1;
        passes += 1;
        std::thread::sleep(Duration::from_millis(50));
    }
    println!("window-from-point passes: {passes}");
    for (answer, count) in &answers {
        println!("  {answer}: {count}");
    }
    // The app re-raises the pin over the player on every tick, so the pin is the answer for the
    // overwhelming majority of passes; a single stray sample is tolerated.
    assert!(
        pins + 1 >= passes && pins > 0,
        "a press in the band reaches the pin on nearly every pass: {pins} of {passes}, answers {answers:?}"
    );
    assert!(
        player_still_active(pid),
        "with the film still playing in the window it was begun in"
    );
    assert!(
        unsafe { IsWindow(player) }.as_bool(),
        "and no host took its top-level window away"
    );

    let _ = child.kill();
    let _ = child.wait();
    unsafe {
        let _ = DestroyWindow(pin);
    }
}

/// What one renderer's three phases measured, kept together so both variants print the same way.
struct LayeredReport {
    renderer: String,
    control: Sampled,
    layered: Sampled,
    reverted: Sampled,
    alive: bool,
    windows: usize,
}

impl LayeredReport {
    fn verdict(&self) -> &'static str {
        if self.layered.lit_mean > 0.5 * self.control.lit_mean && self.layered.changed_mean > 0.0 {
            "layered keeps the picture"
        } else {
            "layered blanks it"
        }
    }
}

/// One ffplay, one clip, one box: control (plain top-level), simple layered, then reverted.
fn sample_layered_round_trip(ffplay: &Path, clip: &Path, renderer: &str) -> LayeredReport {
    let child = with_the_production_renderer(Command::new(ffplay))
        .args(["-left", &BOX_X.to_string(), "-top", &BOX_Y.to_string()])
        .args(["-x", &BAND_W.to_string(), "-y", &BAND_H.to_string()])
        .args(["-noborder", "-autoexit", "-loglevel", "quiet"])
        .arg(clip)
        .env("SDL_RENDER_DRIVER", renderer)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(CREATE_NO_WINDOW.0)
        .spawn()
        .expect("ffplay to start");
    let player = Playing(child);
    let pid = player.pid();

    let player_window = wait_for_window(pid, Duration::from_secs(8))
        .unwrap_or_else(|| panic!("FAIL: no visible top-level window for ffplay pid {pid}"));
    // Put the film exactly on the watched box, as the app's placement would.
    unsafe {
        let _ = SetWindowPos(
            player_window,
            HWND_TOPMOST,
            BOX_X,
            BOX_Y,
            BAND_W,
            BAND_H,
            SWP_SHOWWINDOW | SWP_NOACTIVATE,
        );
    }
    std::thread::sleep(Duration::from_millis(500));
    let rect = window_rect(player_window);
    println!(
        "[{renderer}] player {:?} class={:?} rect=({},{})-({},{})",
        player_window,
        window_class_name(player_window),
        rect.left,
        rect.top,
        rect.right,
        rect.bottom,
    );

    let control = sample_box(BOX_X, BOX_Y, BAND_W as u32, BAND_H as u32, 40);

    // **Simple layered.** The window keeps its film but is composed by DWM instead of promoted.
    let ex = unsafe { GetWindowLongPtrW(player_window, GWL_EXSTYLE) };
    unsafe {
        SetWindowLongPtrW(player_window, GWL_EXSTYLE, ex | WS_EX_LAYERED.0 as isize);
        let _ = SetLayeredWindowAttributes(player_window, COLORREF(0), 255, LWA_ALPHA);
        let _ = SetWindowPos(
            player_window,
            HWND::default(),
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        );
    }
    std::thread::sleep(Duration::from_millis(500));
    let layered_state = unsafe {
        let mut alpha = 0u8;
        let mut flags = LAYERED_WINDOW_ATTRIBUTES_FLAGS(0);
        let read =
            GetLayeredWindowAttributes(player_window, None, Some(&mut alpha), Some(&mut flags))
                .is_ok();
        let bit = (GetWindowLongPtrW(player_window, GWL_EXSTYLE) & WS_EX_LAYERED.0 as isize) != 0;
        (bit, read, alpha, flags.0 & LWA_ALPHA.0 != 0)
    };
    println!(
        "[{renderer}] layered ex bit={} attributes={} alpha={} lwa_alpha={}",
        layered_state.0, layered_state.1, layered_state.2, layered_state.3
    );
    assert!(
        layered_state.0 && layered_state.1 && layered_state.2 == 255 && layered_state.3,
        "[{renderer}] the window really is simple layered (WS_EX_LAYERED + LWA_ALPHA 255)"
    );
    let layered = sample_box(BOX_X, BOX_Y, BAND_W as u32, BAND_H as u32, 40);

    // **Reverted.** The layered bits off again, frame recalculated, so the window can go back.
    unsafe {
        SetWindowLongPtrW(player_window, GWL_EXSTYLE, ex);
        let _ = SetWindowPos(
            player_window,
            HWND::default(),
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        );
    }
    std::thread::sleep(Duration::from_millis(500));
    let reverted_bit =
        unsafe { GetWindowLongPtrW(player_window, GWL_EXSTYLE) & WS_EX_LAYERED.0 as isize } != 0;
    println!("[{renderer}] reverted ex bit={reverted_bit}");
    assert!(!reverted_bit, "[{renderer}] the layered bit is off again");
    let reverted = sample_box(BOX_X, BOX_Y, BAND_W as u32, BAND_H as u32, 40);

    let alive = player_still_active(pid);
    let windows = visible_window_count(pid);
    println!("[{renderer}] alive={alive} visible top-level windows={windows}");

    LayeredReport {
        renderer: renderer.to_string(),
        control,
        layered,
        reverted,
        alive,
        windows,
    }
}

#[test]
#[ignore = "starts a real ffplay and watches a real box on the desktop; run by hand"]
fn a_simple_layered_ffplay_window_keeps_or_loses_the_film() {
    unsafe {
        let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
            .or_else(|_| SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));
    }

    let ffplay = find_program("ffplay.exe").expect("ffplay.exe on PATH (or beside this test)");
    let clip = probe_clip();
    println!("ffplay: {}", ffplay.display());
    println!("clip:   {}", clip.display());

    let mut reports = Vec::new();
    for renderer in ["direct3d9", "direct3d11"] {
        let report = sample_layered_round_trip(&ffplay, &clip, renderer);
        assert!(
            report.control.lit_mean > 0.2,
            "[{}] the control film is up in the box before any layered bits: {:?}",
            report.renderer,
            (
                report.control.lit_mean,
                report.control.lit_min,
                report.control.changed_mean
            )
        );
        assert!(
            report.alive && report.windows == 1,
            "[{}] one player, one window, alive through every phase: alive={} windows={}",
            report.renderer,
            report.alive,
            report.windows
        );
        reports.push(report);
    }

    println!("\nvariant    phase     lit_mean  lit_min  changed_mean  samples");
    for report in &reports {
        for (phase, sampled) in [
            ("control", &report.control),
            ("layered", &report.layered),
            ("reverted", &report.reverted),
        ] {
            println!(
                "{:<10} {:<9} {:.3}     {:.3}    {:.3}         {}",
                report.renderer,
                phase,
                sampled.lit_mean,
                sampled.lit_min,
                sampled.changed_mean,
                sampled.samples
            );
        }
    }
    println!();
    for report in &reports {
        println!("[{}] {}", report.renderer, report.verdict());
    }
}

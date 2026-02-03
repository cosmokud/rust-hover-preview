use crate::{startup, CONFIG, RUNNING};
use std::os::windows::ffi::OsStrExt;
use std::sync::atomic::Ordering;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::{
    Shell_NotifyIconW, ShellExecuteW, NIF_ICON, NIF_MESSAGE, NIF_TIP, NIM_ADD, NIM_DELETE,
    NOTIFYICONDATAW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DispatchMessageW,
    GetCursorPos, LoadImageW, PeekMessageW, PostQuitMessage, RegisterClassExW, RegisterWindowMessageW,
    SetForegroundWindow, TrackPopupMenu, TranslateMessage, CS_HREDRAW, CS_VREDRAW, HICON, IMAGE_ICON,
    LR_DEFAULTSIZE, LR_SHARED, MF_CHECKED, MF_POPUP, MF_STRING, MF_UNCHECKED, MSG, PM_REMOVE,
    SW_SHOWNORMAL, TPM_BOTTOMALIGN, TPM_LEFTALIGN, WM_COMMAND, WM_DESTROY, WM_LBUTTONUP, WM_RBUTTONUP,
    WM_USER, WNDCLASSEXW, WS_EX_TOOLWINDOW, WS_POPUP,
};

const WM_TRAYICON: u32 = WM_USER + 1;
const ID_TRAY_EXIT: u16 = 1001;
const ID_TRAY_STARTUP: u16 = 1002;
const ID_TRAY_ENABLE: u16 = 1003;
const ID_TRAY_VOLUME_MAX: u16 = 1010;      // 100%
const ID_TRAY_VOLUME_HIGH: u16 = 1011;     // 80%
const ID_TRAY_VOLUME_MEDIUM: u16 = 1012;   // 50%
const ID_TRAY_VOLUME_LOW: u16 = 1013;      // 25%
const ID_TRAY_VOLUME_VERY_LOW: u16 = 1014; // 10%
const ID_TRAY_VOLUME_MUTE: u16 = 1015;     // 0%
const ID_TRAY_POSITION_FOLLOW: u16 = 1020; // Follow cursor
const ID_TRAY_POSITION_BEST: u16 = 1021;   // Best position
const ID_TRAY_DELAY_INSTANT: u16 = 1030;   // 0ms
const ID_TRAY_DELAY_VERY_FAST: u16 = 1031; // 200ms
const ID_TRAY_DELAY_MEDIUM: u16 = 1032;    // 500ms
const ID_TRAY_DELAY_SLOW: u16 = 1033;      // 1000ms
const ID_TRAY_OPEN_CONFIG: u16 = 1040;

const TRAY_CLASS: PCWSTR = w!("RustHoverPreviewTrayClass");

static mut TRAY_HWND: HWND = HWND(std::ptr::null_mut());
static mut TASKBAR_CREATED: u32 = 0;

unsafe extern "system" fn tray_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        _ if TASKBAR_CREATED != 0 && msg == TASKBAR_CREATED => {
            // Explorer (taskbar) restarted; re-add tray icon
            remove_tray_icon(hwnd);
            let _ = add_tray_icon(hwnd);
            LRESULT(0)
        }
        WM_TRAYICON => {
            let event = lparam.0 as u32;
            if event == WM_RBUTTONUP || event == WM_LBUTTONUP {
                show_context_menu(hwnd);
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd = (wparam.0 & 0xFFFF) as u16;
            match cmd {
                ID_TRAY_EXIT => {
                    RUNNING.store(false, Ordering::SeqCst);
                    PostQuitMessage(0);
                }
                ID_TRAY_STARTUP => {
                    toggle_startup();
                }
                ID_TRAY_ENABLE => {
                    toggle_preview_enabled();
                }
                ID_TRAY_VOLUME_MAX => set_volume(100),
                ID_TRAY_VOLUME_HIGH => set_volume(80),
                ID_TRAY_VOLUME_MEDIUM => set_volume(50),
                ID_TRAY_VOLUME_LOW => set_volume(25),
                ID_TRAY_VOLUME_VERY_LOW => set_volume(10),
                ID_TRAY_VOLUME_MUTE => set_volume(0),
                ID_TRAY_POSITION_FOLLOW => set_follow_cursor(true),
                ID_TRAY_POSITION_BEST => set_follow_cursor(false),
                ID_TRAY_DELAY_INSTANT => set_hover_delay(0),
                ID_TRAY_DELAY_VERY_FAST => set_hover_delay(200),
                ID_TRAY_DELAY_MEDIUM => set_hover_delay(500),
                ID_TRAY_DELAY_SLOW => set_hover_delay(1000),
                ID_TRAY_OPEN_CONFIG => open_config_file(),
                _ => {}
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            remove_tray_icon(hwnd);
            PostQuitMessage(0);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn show_context_menu(hwnd: HWND) {
    let menu = CreatePopupMenu().unwrap();

    // Add "Enable Preview" with checkmark
    let preview_enabled = CONFIG.lock().map(|c| c.preview_enabled).unwrap_or(true);
    let enable_flags = MF_STRING | if preview_enabled { MF_CHECKED } else { MF_UNCHECKED };
    let _ = AppendMenuW(menu, enable_flags, ID_TRAY_ENABLE as usize, w!("Enable Preview"));

    // Add Preview Delay submenu
    let hover_delay_ms = CONFIG.lock().map(|c| c.hover_delay_ms).unwrap_or(0);
    let delay_menu = CreatePopupMenu().unwrap();

    let delay_flag = |delay: u64| MF_STRING | if hover_delay_ms == delay { MF_CHECKED } else { MF_UNCHECKED };
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(0),
        ID_TRAY_DELAY_INSTANT as usize,
        w!("Instant (0 ms)"),
    );
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(200),
        ID_TRAY_DELAY_VERY_FAST as usize,
        w!("Very Fast (200 ms)"),
    );
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(500),
        ID_TRAY_DELAY_MEDIUM as usize,
        w!("Medium (500 ms)"),
    );
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(1000),
        ID_TRAY_DELAY_SLOW as usize,
        w!("Slow (1000 ms)"),
    );

    let _ = AppendMenuW(menu, MF_STRING | MF_POPUP, delay_menu.0 as usize, w!("Preview Delay"));

    // Add Volume submenu
    let current_volume = CONFIG.lock().map(|c| c.video_volume).unwrap_or(0);
    let volume_menu = CreatePopupMenu().unwrap();
    
    let vol_flag = |vol: u32| MF_STRING | if current_volume == vol { MF_CHECKED } else { MF_UNCHECKED };
    let _ = AppendMenuW(volume_menu, vol_flag(100), ID_TRAY_VOLUME_MAX as usize, w!("Max (100%)"));
    let _ = AppendMenuW(volume_menu, vol_flag(80), ID_TRAY_VOLUME_HIGH as usize, w!("High (80%)"));
    let _ = AppendMenuW(volume_menu, vol_flag(50), ID_TRAY_VOLUME_MEDIUM as usize, w!("Medium (50%)"));
    let _ = AppendMenuW(volume_menu, vol_flag(25), ID_TRAY_VOLUME_LOW as usize, w!("Low (25%)"));
    let _ = AppendMenuW(volume_menu, vol_flag(10), ID_TRAY_VOLUME_VERY_LOW as usize, w!("Very Low (10%)"));
    let _ = AppendMenuW(volume_menu, vol_flag(0), ID_TRAY_VOLUME_MUTE as usize, w!("Mute (0%)"));
    
    let _ = AppendMenuW(menu, MF_STRING | MF_POPUP, volume_menu.0 as usize, w!("Video Volume"));

    // Add Cursor Position submenu
    let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(false);
    let position_menu = CreatePopupMenu().unwrap();
    
    let pos_flag = |follow: bool| MF_STRING | if follow_cursor == follow { MF_CHECKED } else { MF_UNCHECKED };
    let _ = AppendMenuW(position_menu, pos_flag(true), ID_TRAY_POSITION_FOLLOW as usize, w!("Follow Cursor"));
    let _ = AppendMenuW(position_menu, pos_flag(false), ID_TRAY_POSITION_BEST as usize, w!("Best Position"));
    
    let _ = AppendMenuW(menu, MF_STRING | MF_POPUP, position_menu.0 as usize, w!("Preview Position"));

    // Add "Run at Startup" with checkmark
    let startup_enabled = CONFIG.lock().map(|c| c.run_at_startup).unwrap_or(false);
    let flags = MF_STRING | if startup_enabled { MF_CHECKED } else { MF_UNCHECKED };
    let _ = AppendMenuW(menu, flags, ID_TRAY_STARTUP as usize, w!("Run at Startup"));

    // Add "Edit Config.ini"
    let _ = AppendMenuW(menu, MF_STRING, ID_TRAY_OPEN_CONFIG as usize, w!("Edit Config.ini"));

    // Add Exit
    let _ = AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT as usize, w!("Exit"));

    // Get cursor position and show menu
    let mut pt = windows::Win32::Foundation::POINT::default();
    let _ = GetCursorPos(&mut pt);

    let _ = SetForegroundWindow(hwnd).ok();
    let _ = TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_BOTTOMALIGN, pt.x, pt.y, 0, hwnd, None).ok();
    let _ = DestroyMenu(menu);
}

fn toggle_startup() {
    if let Ok(mut config) = CONFIG.lock() {
        config.run_at_startup = !config.run_at_startup;
        config.save();

        if config.run_at_startup {
            startup::enable_startup();
        } else {
            startup::disable_startup();
        }
    }
}

fn toggle_preview_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.preview_enabled = !config.preview_enabled;
        config.save();
    }
}

fn set_volume(volume: u32) {
    if let Ok(mut config) = CONFIG.lock() {
        config.video_volume = volume;
        config.save();
    }
}

fn set_follow_cursor(follow: bool) {
    if let Ok(mut config) = CONFIG.lock() {
        config.follow_cursor = follow;
        config.save();
    }
}

fn set_hover_delay(hover_delay_ms: u64) {
    if let Ok(mut config) = CONFIG.lock() {
        config.hover_delay_ms = hover_delay_ms;
        config.save();
    }
}

fn open_config_file() {
    if let Ok(config) = CONFIG.lock() {
        config.save();
    }

    if let Some(path) = crate::config::AppConfig::config_path() {
        let wide_path: Vec<u16> = path
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();

        unsafe {
            let _ = ShellExecuteW(
                HWND(std::ptr::null_mut()),
                w!("open"),
                PCWSTR(wide_path.as_ptr()),
                PCWSTR(std::ptr::null()),
                PCWSTR(std::ptr::null()),
                SW_SHOWNORMAL,
            );
        }
    }
}

unsafe fn add_tray_icon(hwnd: HWND) -> bool {
    // Load the embedded icon resource (assets/icon.ico compiled via build.rs)
    let hicon = if let Ok(hmodule) = GetModuleHandleW(None) {
        let hinstance = HINSTANCE(hmodule.0);
        match LoadImageW(
            hinstance,
            PCWSTR(1 as *const u16),
            IMAGE_ICON,
            0,
            0,
            LR_DEFAULTSIZE | LR_SHARED,
        ) {
            Ok(h) => HICON(h.0),
            Err(_) => HICON::default(),
        }
    } else {
        HICON::default()
    };
    
    // Fallback to system icon if custom icon failed
    let hicon = if hicon.0.is_null() {
        match LoadImageW(
            None,
            PCWSTR(32512 as *const u16), // IDI_APPLICATION
            IMAGE_ICON,
            0,
            0,
            LR_DEFAULTSIZE | LR_SHARED,
        ) {
            Ok(h) => HICON(h.0),
            Err(_) => HICON::default(),
        }
    } else {
        hicon
    };

    let mut nid = NOTIFYICONDATAW {
        cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
        hWnd: hwnd,
        uID: 1,
        uFlags: NIF_ICON | NIF_MESSAGE | NIF_TIP,
        uCallbackMessage: WM_TRAYICON,
        hIcon: hicon,
        ..Default::default()
    };

    // Set tooltip
    let tip = "Hover Preview";
    let tip_wide: Vec<u16> = tip.encode_utf16().chain(std::iter::once(0)).collect();
    let len = tip_wide.len().min(nid.szTip.len());
    nid.szTip[..len].copy_from_slice(&tip_wide[..len]);

    Shell_NotifyIconW(NIM_ADD, &nid).as_bool()
}

unsafe fn remove_tray_icon(hwnd: HWND) {
    let nid = NOTIFYICONDATAW {
        cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
        hWnd: hwnd,
        uID: 1,
        ..Default::default()
    };
    let _ = Shell_NotifyIconW(NIM_DELETE, &nid);
}

pub fn run_tray() {
    unsafe {
        let hinstance = GetModuleHandleW(None).unwrap();

        // Register window class
        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(tray_window_proc),
            hInstance: hinstance.into(),
            lpszClassName: TRAY_CLASS,
            ..Default::default()
        };

        RegisterClassExW(&wc);

        // Register TaskbarCreated message to detect Explorer restarts
        TASKBAR_CREATED = RegisterWindowMessageW(w!("TaskbarCreated"));

        // Create hidden window for tray messages
        let hwnd = CreateWindowExW(
            WS_EX_TOOLWINDOW,
            TRAY_CLASS,
            w!("Hover Preview Tray"),
            WS_POPUP,
            0,
            0,
            0,
            0,
            None,
            None,
            hinstance,
            None,
        );

        let hwnd = match hwnd {
            Ok(h) => h,
            Err(e) => {
                eprintln!("Failed to create tray window: {:?}", e);
                return;
            }
        };

        TRAY_HWND = hwnd;

        // Add tray icon (retry briefly in case Explorer isn't ready yet)
        let mut added = add_tray_icon(hwnd);
        if !added {
            let mut retries = 20;
            while !added && retries > 0 && RUNNING.load(Ordering::SeqCst) {
                std::thread::sleep(std::time::Duration::from_millis(500));
                added = add_tray_icon(hwnd);
                retries -= 1;
            }
        }

        if !added {
            eprintln!("Failed to add tray icon after retries; exiting.");
            RUNNING.store(false, Ordering::SeqCst);
            remove_tray_icon(hwnd);
            return;
        }

        // Message loop
        let mut msg = MSG::default();
        while RUNNING.load(Ordering::SeqCst) {
            while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                if msg.message == windows::Win32::UI::WindowsAndMessaging::WM_QUIT {
                    RUNNING.store(false, Ordering::SeqCst);
                    break;
                }
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            std::thread::sleep(std::time::Duration::from_millis(10));
        }

        // Cleanup
        remove_tray_icon(hwnd);
    }
}

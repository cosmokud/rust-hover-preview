//! Companion Win32 hook DLL for `CustomCaptureTechnique::InProcessHook`.
//!
//! Runs inside the target preview process via thread-specific `WH_CBT` + `WH_CALLWNDPROC` hooks,
//! applying legal in-process `DWMWA_CLOAK` and synchronous `WM_WINDOWPOSCHANGING` rect locking.
//! Zero third-party dependencies; 100% MIT-licensed.

#![allow(clippy::missing_safety_doc)]

use core::cell::Cell;
use core::ffi::c_void;
use core::ptr;
use core::sync::atomic::{AtomicI32, AtomicIsize, AtomicPtr, AtomicU32, AtomicU8, Ordering};

pub const HOOK_MAGIC: u32 = 0x5248_5048;
pub const HOOK_VERSION: u32 = 1;
pub const CMD_UNCLOAK: usize = 1;
pub const CMD_CLOAK: usize = 2;
pub const CMD_DETACH: usize = 3;
pub const CMD_ENFORCE_NOW: usize = 4;

pub const ZORDER_BOTTOM: u8 = 0;
pub const ZORDER_TOPMOST: u8 = 1;
pub const ZORDER_KEEP: u8 = 2;

const FORCED_OUT_OF_PROC: u8 = 1;
const FORCED_HOOK: u8 = 2;

const DLL_PROCESS_DETACH: u32 = 0;

const FILE_MAP_ALL_ACCESS: u32 = 0x000F_001F;

const HCBT_MOVESIZE: i32 = 0;
const HCBT_MINMAX: i32 = 1;
const HCBT_CREATEWND: i32 = 3;
const HCBT_ACTIVATE: i32 = 5;

const SW_SHOWMINIMIZED: u32 = 2;
const SW_MAXIMIZE: u32 = 3;
const SW_SHOWNOACTIVATE: u32 = 4;
const SW_MINIMIZE: u32 = 6;
const SW_FORCEMINIMIZE: u32 = 11;

const WM_CREATE: u32 = 0x0001;
const WM_SHOWWINDOW: u32 = 0x0018;
const WM_MOUSEACTIVATE: u32 = 0x0021;
const WM_GETMINMAXINFO: u32 = 0x0024;
const MA_NOACTIVATE: isize = 3;
const WM_WINDOWPOSCHANGING: u32 = 0x0046;
const WM_WINDOWPOSCHANGED: u32 = 0x0047;
const WM_NCCREATE: u32 = 0x0081;
const WM_NCDESTROY: u32 = 0x0082;
const WM_SYSCOMMAND: u32 = 0x0112;
const WM_DPICHANGED: u32 = 0x02E0;

const SC_SIZE: usize = 0xF000;
const SC_MOVE: usize = 0xF010;
const SC_MINIMIZE: usize = 0xF020;
const SC_MAXIMIZE: usize = 0xF030;
const SC_RESTORE: usize = 0xF120;

const SUBCLASS_ID: usize = 0x5248_5048;

const GWL_STYLE: i32 = -16;
const GWL_EXSTYLE: i32 = -20;
const GW_OWNER: u32 = 4;

const WS_CAPTION: u32 = 0x00C0_0000;
const WS_MAXIMIZE: u32 = 0x0100_0000;
const WS_MINIMIZE: u32 = 0x2000_0000;
const WS_CHILD: u32 = 0x4000_0000;
const WS_EX_TOOLWINDOW: u32 = 0x0000_0080;

const HWND_BOTTOM: isize = 1;
const HWND_TOPMOST: isize = -1;
const SWP_NOSIZE: u32 = 0x0001;
const SWP_NOMOVE: u32 = 0x0002;
const SWP_NOZORDER: u32 = 0x0004;
const SWP_NOACTIVATE: u32 = 0x0010;
const SWP_SHOWWINDOW: u32 = 0x0040;
const SWP_NOOWNERZORDER: u32 = 0x0200;

const DWMWA_CLOAK: u32 = 13;

#[repr(C)]
pub struct HookSharedState {
    pub magic: u32,
    pub version: u32,
    pub target_pid: u32,
    pub uncloak_msg: u32,
    pub cloak_enabled: AtomicU32,
    pub rect_locked: AtomicU32,
    pub x: AtomicI32,
    pub y: AtomicI32,
    pub w: AtomicI32,
    pub h: AtomicI32,
    pub captured_hwnd: AtomicIsize,
    pub hook_hits: AtomicU32,
    pub forced_placement: AtomicU8,
    pub z_order_mode: AtomicU8,
}

#[repr(C)]
struct Point {
    x: i32,
    y: i32,
}

#[repr(C)]
struct MinMaxInfo {
    pt_reserved: Point,
    pt_max_size: Point,
    pt_max_position: Point,
    pt_min_track_size: Point,
    pt_max_track_size: Point,
}

#[repr(C)]
struct WindowPlacement {
    length: u32,
    flags: u32,
    show_cmd: u32,
    pt_min_position: Point,
    pt_max_position: Point,
    rc_normal_position: HookRect,
}

#[repr(C)]
struct HookRect {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
}

#[repr(C)]
struct CreateStructW {
    lp_create_params: *mut c_void,
    h_instance: isize,
    h_menu: isize,
    hwnd_parent: isize,
    cy: i32,
    cx: i32,
    y: i32,
    x: i32,
    style: i32,
    lpsz_name: *const u16,
    lpsz_class: *const u16,
    dw_ex_style: u32,
}

#[repr(C)]
struct CbtCreateWndW {
    lpcs: *mut CreateStructW,
    hwnd_insert_after: isize,
}

#[repr(C)]
struct CwpStruct {
    lparam: isize,
    wparam: usize,
    message: u32,
    hwnd: isize,
}

#[repr(C)]
struct WindowPos {
    hwnd: isize,
    hwnd_insert_after: isize,
    x: i32,
    y: i32,
    cx: i32,
    cy: i32,
    flags: u32,
}

#[link(name = "kernel32")]
extern "system" {
    fn GetCurrentProcessId() -> u32;
    fn OpenFileMappingW(
        dw_desired_access: u32,
        b_inherit_handle: i32,
        lp_name: *const u16,
    ) -> isize;
    fn MapViewOfFile(
        h_file_mapping_object: isize,
        dw_desired_access: u32,
        dw_file_offset_high: u32,
        dw_file_offset_low: u32,
        dw_number_of_bytes_to_map: usize,
    ) -> *mut c_void;
    fn UnmapViewOfFile(lp_base_address: *const c_void) -> i32;
    fn CloseHandle(h_object: isize) -> i32;
}

#[link(name = "user32")]
extern "system" {
    fn CallNextHookEx(hhk: isize, n_code: i32, wparam: usize, lparam: isize) -> isize;
    fn GetParent(hwnd: isize) -> isize;
    fn GetWindow(hwnd: isize, u_cmd: u32) -> isize;
    fn GetWindowLongPtrW(hwnd: isize, n_index: i32) -> isize;
    fn GetWindowRect(hwnd: isize, lp_rect: *mut HookRect) -> i32;
    fn IsIconic(hwnd: isize) -> i32;
    fn IsWindowVisible(hwnd: isize) -> i32;
    fn IsZoomed(hwnd: isize) -> i32;
    fn SetWindowPlacement(hwnd: isize, lpwndpl: *const WindowPlacement) -> i32;
    fn SetWindowPos(
        hwnd: isize,
        hwnd_insert_after: isize,
        x: i32,
        y: i32,
        cx: i32,
        cy: i32,
        u_flags: u32,
    ) -> i32;
}

#[link(name = "dwmapi")]
extern "system" {
    fn DwmSetWindowAttribute(
        hwnd: isize,
        dw_attribute: u32,
        pv_attribute: *const c_void,
        cb_attribute: u32,
    ) -> i32;
}

type SubclassProc = unsafe extern "system" fn(isize, u32, usize, isize, usize, usize) -> isize;

#[link(name = "comctl32")]
extern "system" {
    fn SetWindowSubclass(
        hwnd: isize,
        pfn_subclass: SubclassProc,
        u_id_subclass: usize,
        dw_ref_data: usize,
    ) -> i32;
    fn RemoveWindowSubclass(hwnd: isize, pfn_subclass: SubclassProc, u_id_subclass: usize) -> i32;
    fn DefSubclassProc(hwnd: isize, u_msg: u32, wparam: usize, lparam: isize) -> isize;
}

// 0 = uninit, 1 = initializing, 2 = initialized
static SHM_INIT: AtomicU32 = AtomicU32::new(0);
static SHM_HANDLE: AtomicIsize = AtomicIsize::new(0);
static SHM_PTR: AtomicPtr<HookSharedState> = AtomicPtr::new(ptr::null_mut());

/// Format `Local\RhpCustomPreviewHook_<pid>\0` into a stack UTF-16 buffer (no heap alloc).
fn format_shm_name(pid: u32, buf: &mut [u16; 64]) {
    const PREFIX: &[u8] = b"Local\\RhpCustomPreviewHook_";
    let mut idx = 0usize;
    for &b in PREFIX {
        buf[idx] = u16::from(b);
        idx += 1;
    }
    let mut digits = [0u8; 10];
    let mut n = pid;
    let mut dlen = 0usize;
    if n == 0 {
        digits[0] = b'0';
        dlen = 1;
    } else {
        while n > 0 {
            digits[dlen] = b'0' + (n % 10) as u8;
            n /= 10;
            dlen += 1;
        }
    }
    while dlen > 0 {
        dlen -= 1;
        buf[idx] = u16::from(digits[dlen]);
        idx += 1;
    }
    buf[idx] = 0;
}

/// Map per-PID shared state once; returns `None` immediately for any non-target process.
#[inline]
unsafe fn shared_state() -> Option<&'static HookSharedState> {
    let state = SHM_INIT.load(Ordering::Acquire);
    if state == 2 {
        let ptr = SHM_PTR.load(Ordering::Acquire);
        return (!ptr.is_null()).then(|| unsafe { &*ptr });
    }
    if SHM_INIT
        .compare_exchange(0, 1, Ordering::AcqRel, Ordering::Acquire)
        .is_ok()
    {
        let pid = unsafe { GetCurrentProcessId() };
        let mut name = [0u16; 64];
        format_shm_name(pid, &mut name);
        let handle = unsafe { OpenFileMappingW(FILE_MAP_ALL_ACCESS, 0, name.as_ptr()) };
        if handle != 0 {
            let view = unsafe {
                MapViewOfFile(
                    handle,
                    FILE_MAP_ALL_ACCESS,
                    0,
                    0,
                    core::mem::size_of::<HookSharedState>(),
                )
            } as *mut HookSharedState;
            if !view.is_null() {
                let s = unsafe { &*view };
                if s.magic == HOOK_MAGIC && s.version == HOOK_VERSION && s.target_pid == pid {
                    SHM_HANDLE.store(handle, Ordering::Release);
                    SHM_PTR.store(view, Ordering::Release);
                } else {
                    let _ = unsafe { UnmapViewOfFile(view as *const c_void) };
                    let _ = unsafe { CloseHandle(handle) };
                }
            } else {
                let _ = unsafe { CloseHandle(handle) };
            }
        }
        SHM_INIT.store(2, Ordering::Release);
    }
    if SHM_INIT.load(Ordering::Acquire) == 2 {
        let ptr = SHM_PTR.load(Ordering::Acquire);
        return (!ptr.is_null()).then(|| unsafe { &*ptr });
    }
    None
}

std::thread_local! {
    static IN_ENFORCE: Cell<bool> = const { Cell::new(false) };
}

struct EnforceGuard;
impl Drop for EnforceGuard {
    fn drop(&mut self) {
        IN_ENFORCE.with(|flag| flag.set(false));
    }
}

#[inline]
unsafe fn is_top_level_for_mode(state: &HookSharedState, hwnd: isize) -> bool {
    if hwnd == 0 || unsafe { GetParent(hwnd) } != 0 {
        return false;
    }
    let owner = unsafe { GetWindow(hwnd, GW_OWNER) };
    if owner != 0 {
        let hook_mode = state.forced_placement.load(Ordering::Acquire) == FORCED_HOOK;
        if !hook_mode || unsafe { IsWindowVisible(owner) } != 0 {
            return false;
        }
    }
    let style = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) } as u32;
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
    (style & WS_CHILD) == 0 && (ex_style & WS_EX_TOOLWINDOW) == 0
}

#[inline]
unsafe fn set_cloak(hwnd: isize, cloak: bool) {
    let val: u32 = u32::from(cloak);
    let _ = unsafe {
        DwmSetWindowAttribute(
            hwnd,
            DWMWA_CLOAK,
            &val as *const u32 as *const c_void,
            core::mem::size_of::<u32>() as u32,
        )
    };
}

#[inline]
unsafe fn is_captured_or_top_level(state: &HookSharedState, hwnd: isize) -> bool {
    state.captured_hwnd.load(Ordering::Acquire) == hwnd
        || unsafe { is_top_level_for_mode(state, hwnd) }
}

#[inline]
unsafe fn apply_locked_window_pos(state: &HookSharedState, wp: *mut WindowPos) {
    if wp.is_null() || state.rect_locked.load(Ordering::Acquire) != 1 {
        return;
    }
    unsafe {
        match state.z_order_mode.load(Ordering::Acquire) {
            ZORDER_TOPMOST => {
                (*wp).hwnd_insert_after = HWND_TOPMOST;
                (*wp).flags &= !SWP_NOZORDER;
            }
            ZORDER_KEEP => {
                (*wp).flags |= SWP_NOZORDER;
            }
            _ => {
                (*wp).hwnd_insert_after = HWND_BOTTOM;
                (*wp).flags &= !SWP_NOZORDER;
            }
        }
        (*wp).flags |= SWP_NOACTIVATE;
        (*wp).x = state.x.load(Ordering::Acquire);
        (*wp).y = state.y.load(Ordering::Acquire);
        (*wp).flags &= !SWP_NOMOVE;
        let w = state.w.load(Ordering::Acquire);
        let h = state.h.load(Ordering::Acquire);
        if w > 0 && h > 0 {
            (*wp).cx = w;
            (*wp).cy = h;
            (*wp).flags &= !SWP_NOSIZE;
        }
    }
}

#[inline]
unsafe fn rect_matches_locked(hwnd: isize, state: &HookSharedState) -> bool {
    let x = state.x.load(Ordering::Acquire);
    let y = state.y.load(Ordering::Acquire);
    let w = state.w.load(Ordering::Acquire);
    let h = state.h.load(Ordering::Acquire);
    let mut rc = HookRect {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
    };
    if unsafe { GetWindowRect(hwnd, &mut rc) } == 0 {
        return false;
    }
    if rc.left != x || rc.top != y {
        return false;
    }
    if w > 0 && h > 0 && ((rc.right - rc.left) != w || (rc.bottom - rc.top) != h) {
        return false;
    }
    (unsafe { IsZoomed(hwnd) }) == 0 && (unsafe { IsIconic(hwnd) }) == 0
}

unsafe fn enforce_locked_rect_now(hwnd: isize, state: &HookSharedState) {
    if hwnd == 0 || state.rect_locked.load(Ordering::Acquire) != 1 {
        return;
    }
    let already = IN_ENFORCE.with(|flag| {
        if flag.get() {
            true
        } else {
            flag.set(true);
            false
        }
    });
    if already {
        return;
    }
    let _guard = EnforceGuard;
    let x = state.x.load(Ordering::Acquire);
    let y = state.y.load(Ordering::Acquire);
    let w = state.w.load(Ordering::Acquire);
    let h = state.h.load(Ordering::Acquire);
    if unsafe { IsZoomed(hwnd) } != 0 || unsafe { IsIconic(hwnd) } != 0 {
        let placement = WindowPlacement {
            length: core::mem::size_of::<WindowPlacement>() as u32,
            flags: 0,
            show_cmd: SW_SHOWNOACTIVATE,
            pt_min_position: Point { x: 0, y: 0 },
            pt_max_position: Point { x: 0, y: 0 },
            rc_normal_position: HookRect {
                left: x,
                top: y,
                right: x + w.max(1),
                bottom: y + h.max(1),
            },
        };
        let _ = unsafe { SetWindowPlacement(hwnd, &placement) };
    }
    let z_mode = state.z_order_mode.load(Ordering::Acquire);
    let insert_after = if z_mode == ZORDER_TOPMOST {
        HWND_TOPMOST
    } else {
        HWND_BOTTOM
    };
    let mut flags = SWP_NOACTIVATE | SWP_NOOWNERZORDER;
    if z_mode == ZORDER_KEEP {
        flags |= SWP_NOZORDER;
    }
    let (cw, ch) = if w > 0 && h > 0 {
        (w, h)
    } else {
        flags |= SWP_NOSIZE;
        (0, 0)
    };
    let _ = unsafe { SetWindowPos(hwnd, insert_after, x, y, cw, ch, flags) };
}

#[no_mangle]
pub unsafe extern "system" fn rhp_cbt_proc(code: i32, wparam: usize, lparam: isize) -> isize {
    if code >= 0 {
        if let Some(state) = unsafe { shared_state() } {
            let forced_mode = state.forced_placement.load(Ordering::Acquire);
            let locked = state.rect_locked.load(Ordering::Acquire) == 1;
            if code == HCBT_CREATEWND {
                let cbt = lparam as *const CbtCreateWndW;
                if !cbt.is_null() {
                    let lpcs = unsafe { (*cbt).lpcs };
                    if !lpcs.is_null() {
                        let style = unsafe { (*lpcs).style } as u32;
                        let ex_style = unsafe { (*lpcs).dw_ex_style };
                        let parent = unsafe { (*lpcs).hwnd_parent };
                        let hook_mode = forced_mode == FORCED_HOOK;
                        let parent_ok =
                            parent == 0 || (hook_mode && unsafe { IsWindowVisible(parent) } == 0);
                        if (style & WS_CHILD) == 0
                            && (ex_style & WS_EX_TOOLWINDOW) == 0
                            && parent_ok
                            && locked
                        {
                            let w = state.w.load(Ordering::Acquire);
                            let h = state.h.load(Ordering::Acquire);
                            unsafe {
                                if hook_mode {
                                    (*lpcs).style = (style & !(WS_MAXIMIZE | WS_MINIMIZE)) as i32;
                                }
                                (*lpcs).x = state.x.load(Ordering::Acquire);
                                (*lpcs).y = state.y.load(Ordering::Acquire);
                                if w > 0 && h > 0 {
                                    (*lpcs).cx = w;
                                    (*lpcs).cy = h;
                                }
                            }
                            state.hook_hits.fetch_add(1, Ordering::Relaxed);
                        }
                    }
                }
            } else if code == HCBT_MINMAX && locked {
                let sw_cmd = lparam as u32 & 0xFFFF;
                let block = match forced_mode {
                    FORCED_OUT_OF_PROC => sw_cmd == SW_MAXIMIZE,
                    FORCED_HOOK => matches!(
                        sw_cmd,
                        SW_MAXIMIZE | SW_SHOWMINIMIZED | SW_MINIMIZE | SW_FORCEMINIMIZE
                    ),
                    _ => false,
                };
                if block && unsafe { is_captured_or_top_level(state, wparam as isize) } {
                    state.hook_hits.fetch_add(1, Ordering::Relaxed);
                    return 1;
                }
            } else if code == HCBT_MOVESIZE
                && matches!(forced_mode, FORCED_OUT_OF_PROC | FORCED_HOOK)
                && locked
                && unsafe { is_captured_or_top_level(state, wparam as isize) }
            {
                let rc = lparam as *mut HookRect;
                let w = state.w.load(Ordering::Acquire);
                let h = state.h.load(Ordering::Acquire);
                if !rc.is_null() && w > 0 && h > 0 {
                    let x = state.x.load(Ordering::Acquire);
                    let y = state.y.load(Ordering::Acquire);
                    unsafe {
                        (*rc).left = x;
                        (*rc).top = y;
                        (*rc).right = x + w;
                        (*rc).bottom = y + h;
                    }
                    state.hook_hits.fetch_add(1, Ordering::Relaxed);
                }
            } else if code == HCBT_ACTIVATE
                && (locked || state.cloak_enabled.load(Ordering::Acquire) == 1)
            {
                state.hook_hits.fetch_add(1, Ordering::Relaxed);
                return 1;
            }
        }
    }
    unsafe { CallNextHookEx(0, code, wparam, lparam) }
}

/// In-process window subclass: `WH_CALLWNDPROC` cannot mutate `WM_WINDOWPOSCHANGING`'s `WINDOWPOS`
/// on Windows, so we enforce rect-lock + cloak directly in the window's WndProc chain.
unsafe extern "system" fn rhp_subclass_proc(
    hwnd: isize,
    msg: u32,
    wparam: usize,
    lparam: isize,
    u_id_subclass: usize,
    _dw_ref_data: usize,
) -> isize {
    if msg == WM_NCDESTROY {
        let _ = unsafe { RemoveWindowSubclass(hwnd, rhp_subclass_proc, u_id_subclass) };
        return unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) };
    }
    if let Some(state) = unsafe { shared_state() } {
        if msg == state.uncloak_msg && state.uncloak_msg != 0 {
            if wparam == CMD_DETACH {
                unsafe { set_cloak(hwnd, false) };
                let _ = unsafe { RemoveWindowSubclass(hwnd, rhp_subclass_proc, u_id_subclass) };
                state.hook_hits.fetch_add(1, Ordering::Relaxed);
                return 0;
            }
            if wparam == CMD_ENFORCE_NOW {
                if !IN_ENFORCE.with(|flag| flag.get())
                    && state.rect_locked.load(Ordering::Acquire) == 1
                    && !unsafe { rect_matches_locked(hwnd, state) }
                {
                    unsafe { enforce_locked_rect_now(hwnd, state) };
                }
                state.hook_hits.fetch_add(1, Ordering::Relaxed);
                return 0;
            }
        }
        let forced_mode = state.forced_placement.load(Ordering::Acquire);
        let locked = state.rect_locked.load(Ordering::Acquire) == 1;
        if msg == WM_MOUSEACTIVATE && locked {
            state.hook_hits.fetch_add(1, Ordering::Relaxed);
            return MA_NOACTIVATE;
        }
        if msg == WM_SYSCOMMAND && locked {
            let cmd = wparam & 0xFFF0;
            let block = match forced_mode {
                FORCED_OUT_OF_PROC => matches!(cmd, SC_SIZE | SC_MOVE | SC_MAXIMIZE),
                FORCED_HOOK => matches!(
                    cmd,
                    SC_SIZE | SC_MOVE | SC_MAXIMIZE | SC_MINIMIZE | SC_RESTORE
                ),
                _ => false,
            };
            if block {
                state.hook_hits.fetch_add(1, Ordering::Relaxed);
                return 0;
            }
        }
        if forced_mode == FORCED_HOOK && locked {
            if msg == WM_GETMINMAXINFO {
                let res = unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) };
                let w = state.w.load(Ordering::Acquire);
                let h = state.h.load(Ordering::Acquire);
                if w > 0 && h > 0 && lparam != 0 {
                    let x = state.x.load(Ordering::Acquire);
                    let y = state.y.load(Ordering::Acquire);
                    let mmi = lparam as *mut MinMaxInfo;
                    unsafe {
                        (*mmi).pt_min_track_size = Point { x: 1, y: 1 };
                        if (*mmi).pt_max_track_size.x < w {
                            (*mmi).pt_max_track_size.x = w;
                        }
                        if (*mmi).pt_max_track_size.y < h {
                            (*mmi).pt_max_track_size.y = h;
                        }
                        (*mmi).pt_max_size = Point { x: w, y: h };
                        (*mmi).pt_max_position = Point { x, y };
                    }
                    state.hook_hits.fetch_add(1, Ordering::Relaxed);
                    return 0;
                }
                return res;
            }
            if msg == WM_WINDOWPOSCHANGING {
                if state.cloak_enabled.load(Ordering::Acquire) == 1 {
                    unsafe { set_cloak(hwnd, true) };
                }
                let res = unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) };
                unsafe { apply_locked_window_pos(state, lparam as *mut WindowPos) };
                return res;
            }
            if msg == WM_WINDOWPOSCHANGED {
                let res = unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) };
                if !IN_ENFORCE.with(|flag| flag.get()) {
                    let w = state.w.load(Ordering::Acquire);
                    let h = state.h.load(Ordering::Acquire);
                    if w > 0 && h > 0 && !unsafe { rect_matches_locked(hwnd, state) } {
                        unsafe { enforce_locked_rect_now(hwnd, state) };
                    }
                }
                return res;
            }
            if msg == WM_SHOWWINDOW && wparam != 0 && !IN_ENFORCE.with(|flag| flag.get()) {
                unsafe { enforce_locked_rect_now(hwnd, state) };
            }
            if msg == WM_DPICHANGED {
                unsafe { enforce_locked_rect_now(hwnd, state) };
                state.hook_hits.fetch_add(1, Ordering::Relaxed);
                return 0;
            }
        } else if msg == WM_WINDOWPOSCHANGING {
            if state.cloak_enabled.load(Ordering::Acquire) == 1 {
                unsafe { set_cloak(hwnd, true) };
            }
            unsafe { apply_locked_window_pos(state, lparam as *mut WindowPos) };
        }
    }
    unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) }
}

#[no_mangle]
pub unsafe extern "system" fn rhp_callwnd_proc(code: i32, wparam: usize, lparam: isize) -> isize {
    if code >= 0 {
        if let Some(state) = unsafe { shared_state() } {
            let cwp = lparam as *const CwpStruct;
            if !cwp.is_null() {
                let hwnd = unsafe { (*cwp).hwnd };
                let msg = unsafe { (*cwp).message };
                if hwnd != 0 {
                    // Handle hook control message even if Borderless/DirectWindow added WS_EX_TOOLWINDOW.
                    if msg == state.uncloak_msg && state.uncloak_msg != 0 {
                        let cmd = unsafe { (*cwp).wparam };
                        if cmd == CMD_UNCLOAK {
                            unsafe { set_cloak(hwnd, false) };
                            state.hook_hits.fetch_add(1, Ordering::Relaxed);
                        } else if cmd == CMD_CLOAK {
                            unsafe { set_cloak(hwnd, true) };
                            state.hook_hits.fetch_add(1, Ordering::Relaxed);
                        } else if cmd == CMD_DETACH {
                            unsafe { set_cloak(hwnd, false) };
                            let _ = unsafe {
                                RemoveWindowSubclass(hwnd, rhp_subclass_proc, SUBCLASS_ID)
                            };
                            state.hook_hits.fetch_add(1, Ordering::Relaxed);
                        } else if cmd == CMD_ENFORCE_NOW {
                            if state.rect_locked.load(Ordering::Acquire) == 1 {
                                let _ = unsafe {
                                    SetWindowSubclass(hwnd, rhp_subclass_proc, SUBCLASS_ID, 0)
                                };
                                unsafe { enforce_locked_rect_now(hwnd, state) };
                            }
                            state.hook_hits.fetch_add(1, Ordering::Relaxed);
                        }
                    } else if matches!(
                        msg,
                        WM_NCCREATE | WM_CREATE | WM_SHOWWINDOW | WM_WINDOWPOSCHANGING
                    ) && unsafe { is_captured_or_top_level(state, hwnd) }
                    {
                        let wp = if msg == WM_WINDOWPOSCHANGING {
                            (unsafe { (*cwp).lparam }) as *mut WindowPos
                        } else {
                            ptr::null_mut()
                        };
                        let cloak = state.cloak_enabled.load(Ordering::Acquire) == 1;
                        let hook_mode =
                            state.forced_placement.load(Ordering::Acquire) == FORCED_HOOK;
                        if cloak {
                            unsafe { set_cloak(hwnd, true) };
                        }
                        if cloak || hook_mode {
                            // Only latch main-window candidates; never 0x0 hidden helpers on WM_NCCREATE/WM_CREATE.
                            let is_main_candidate = (msg == WM_SHOWWINDOW
                                && unsafe { (*cwp).wparam } != 0)
                                || (!wp.is_null()
                                    && ((unsafe { (*wp).flags } & SWP_SHOWWINDOW) != 0
                                        || (unsafe { (*wp).cx } >= 64
                                            && unsafe { (*wp).cy } >= 64
                                            && (unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) }
                                                as u32
                                                & WS_CAPTION)
                                                != 0)));
                            if is_main_candidate {
                                let _ = state.captured_hwnd.compare_exchange(
                                    0,
                                    hwnd,
                                    Ordering::AcqRel,
                                    Ordering::Acquire,
                                );
                            }
                            state.hook_hits.fetch_add(1, Ordering::Relaxed);
                        }
                        if state.rect_locked.load(Ordering::Acquire) == 1 {
                            let _ = unsafe {
                                SetWindowSubclass(hwnd, rhp_subclass_proc, SUBCLASS_ID, 0)
                            };
                        }
                        unsafe { apply_locked_window_pos(state, wp) };
                    }
                }
            }
        }
    }
    unsafe { CallNextHookEx(0, code, wparam, lparam) }
}

#[no_mangle]
pub unsafe extern "system" fn DllMain(
    _hinst_dll: isize,
    fdw_reason: u32,
    lpv_reserved: *mut c_void,
) -> i32 {
    if fdw_reason == DLL_PROCESS_DETACH && lpv_reserved.is_null() {
        let view = SHM_PTR.swap(ptr::null_mut(), Ordering::AcqRel);
        if !view.is_null() {
            let _ = unsafe { UnmapViewOfFile(view as *const c_void) };
        }
        let handle = SHM_HANDLE.swap(0, Ordering::AcqRel);
        if handle != 0 {
            let _ = unsafe { CloseHandle(handle) };
        }
        SHM_INIT.store(0, Ordering::Release);
    }
    1
}

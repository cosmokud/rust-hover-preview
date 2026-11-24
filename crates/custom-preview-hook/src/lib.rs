//! Companion Win32 hook DLL for `CustomCaptureTechnique::InProcessHook`.
//!
//! Runs inside the target preview process via thread-specific `WH_CBT` + `WH_CALLWNDPROC` hooks,
//! applying legal in-process `DWMWA_CLOAK` and synchronous `WM_WINDOWPOSCHANGING` rect locking.
//! Zero third-party dependencies; 100% MIT-licensed.

#![allow(clippy::missing_safety_doc)]

use core::ffi::c_void;
use core::ptr;
use core::sync::atomic::{AtomicI32, AtomicIsize, AtomicPtr, AtomicU32, AtomicU8, Ordering};

pub const HOOK_MAGIC: u32 = 0x5248_5048;
pub const HOOK_VERSION: u32 = 1;
pub const CMD_UNCLOAK: usize = 1;
pub const CMD_CLOAK: usize = 2;
pub const CMD_DETACH: usize = 3;

const DLL_PROCESS_DETACH: u32 = 0;

const FILE_MAP_ALL_ACCESS: u32 = 0x000F_001F;

const HCBT_MOVESIZE: i32 = 0;
const HCBT_MINMAX: i32 = 1;
const HCBT_CREATEWND: i32 = 3;
const HCBT_ACTIVATE: i32 = 5;

const SW_MAXIMIZE: u32 = 3;

const WM_CREATE: u32 = 0x0001;
const WM_SHOWWINDOW: u32 = 0x0018;
const WM_WINDOWPOSCHANGING: u32 = 0x0046;
const WM_NCCREATE: u32 = 0x0081;
const WM_NCDESTROY: u32 = 0x0082;
const WM_SYSCOMMAND: u32 = 0x0112;

const SC_SIZE: usize = 0xF000;
const SC_MOVE: usize = 0xF010;
const SC_MAXIMIZE: usize = 0xF030;

const SUBCLASS_ID: usize = 0x5248_5048;

const GWL_STYLE: i32 = -16;
const GWL_EXSTYLE: i32 = -20;
const GW_OWNER: u32 = 4;

const WS_CAPTION: u32 = 0x00C0_0000;
const WS_CHILD: u32 = 0x4000_0000;
const WS_EX_TOOLWINDOW: u32 = 0x0000_0080;

const HWND_BOTTOM: isize = 1;
const SWP_NOSIZE: u32 = 0x0001;
const SWP_NOMOVE: u32 = 0x0002;
const SWP_NOZORDER: u32 = 0x0004;
const SWP_NOACTIVATE: u32 = 0x0010;
const SWP_SHOWWINDOW: u32 = 0x0040;

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

#[inline]
unsafe fn is_top_level_non_tool(hwnd: isize) -> bool {
    if hwnd == 0 || unsafe { GetParent(hwnd) } != 0 || unsafe { GetWindow(hwnd, GW_OWNER) } != 0 {
        return false;
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
    state.captured_hwnd.load(Ordering::Acquire) == hwnd || unsafe { is_top_level_non_tool(hwnd) }
}

#[inline]
unsafe fn apply_locked_window_pos(state: &HookSharedState, wp: *mut WindowPos) {
    if wp.is_null() || state.rect_locked.load(Ordering::Acquire) != 1 {
        return;
    }
    unsafe {
        (*wp).hwnd_insert_after = HWND_BOTTOM;
        (*wp).flags = ((*wp).flags & !SWP_NOZORDER) | SWP_NOACTIVATE;
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

#[no_mangle]
pub unsafe extern "system" fn rhp_cbt_proc(code: i32, wparam: usize, lparam: isize) -> isize {
    if code >= 0 {
        if let Some(state) = unsafe { shared_state() } {
            if code == HCBT_CREATEWND {
                let cbt = lparam as *const CbtCreateWndW;
                if !cbt.is_null() {
                    let lpcs = unsafe { (*cbt).lpcs };
                    if !lpcs.is_null() {
                        let style = unsafe { (*lpcs).style } as u32;
                        let ex_style = unsafe { (*lpcs).dw_ex_style };
                        let parent = unsafe { (*lpcs).hwnd_parent };
                        if (style & WS_CHILD) == 0
                            && (ex_style & WS_EX_TOOLWINDOW) == 0
                            && parent == 0
                            && state.rect_locked.load(Ordering::Acquire) == 1
                        {
                            let w = state.w.load(Ordering::Acquire);
                            let h = state.h.load(Ordering::Acquire);
                            unsafe {
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
            } else if code == HCBT_MINMAX
                && state.forced_placement.load(Ordering::Acquire) == 1
                && state.rect_locked.load(Ordering::Acquire) == 1
                && (lparam as u32 & 0xFFFF) == SW_MAXIMIZE
                && unsafe { is_captured_or_top_level(state, wparam as isize) }
            {
                state.hook_hits.fetch_add(1, Ordering::Relaxed);
                return 1;
            } else if code == HCBT_MOVESIZE
                && state.forced_placement.load(Ordering::Acquire) == 1
                && state.rect_locked.load(Ordering::Acquire) == 1
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
            } else if code == HCBT_ACTIVATE && state.cloak_enabled.load(Ordering::Acquire) == 1 {
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
        if msg == state.uncloak_msg && state.uncloak_msg != 0 && wparam == CMD_DETACH {
            unsafe { set_cloak(hwnd, false) };
            let _ = unsafe { RemoveWindowSubclass(hwnd, rhp_subclass_proc, u_id_subclass) };
            state.hook_hits.fetch_add(1, Ordering::Relaxed);
            return 0;
        }
        if msg == WM_SYSCOMMAND
            && state.forced_placement.load(Ordering::Acquire) == 1
            && state.rect_locked.load(Ordering::Acquire) == 1
        {
            let cmd = wparam & 0xFFF0;
            if matches!(cmd, SC_SIZE | SC_MOVE | SC_MAXIMIZE) {
                state.hook_hits.fetch_add(1, Ordering::Relaxed);
                return 0;
            }
        }
        if msg == WM_WINDOWPOSCHANGING {
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
                    // Handle cloak control message even if Borderless mode added WS_EX_TOOLWINDOW.
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
                        }
                    } else if matches!(
                        msg,
                        WM_NCCREATE | WM_CREATE | WM_SHOWWINDOW | WM_WINDOWPOSCHANGING
                    ) {
                        let is_captured = state.captured_hwnd.load(Ordering::Acquire) == hwnd;
                        if is_captured || unsafe { is_top_level_non_tool(hwnd) } {
                            let wp = if msg == WM_WINDOWPOSCHANGING {
                                (unsafe { (*cwp).lparam }) as *mut WindowPos
                            } else {
                                ptr::null_mut()
                            };
                            if state.cloak_enabled.load(Ordering::Acquire) == 1 {
                                unsafe { set_cloak(hwnd, true) };
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

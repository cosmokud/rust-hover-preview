//! The `Custom Preview` window: frameless audio-card-styled editor for custom preview presets.
//!
//! Runs on its own thread (opened from tray `WM_COMMAND`), single instance via `OPEN`/`WINDOW`.
//! Uses native Win32 child controls (`EDIT`, `LISTBOX`, `COMBOBOX`, `msctls_trackbar32`,
//! owner-drawn `BUTTON`) painted with the active `TextTheme` palette and live theme switching.

use std::sync::atomic::{AtomicBool, AtomicIsize, Ordering};

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Dwm::{
    DwmExtendFrameIntoClientArea, DwmSetWindowAttribute, DWMWA_USE_IMMERSIVE_DARK_MODE,
    DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_ROUND,
};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, CreateFontW, CreatePen, CreateSolidBrush, DeleteObject, DrawTextW, EndPaint,
    FillRect, GetStockObject, InvalidateRect, LineTo, MoveToEx, RedrawWindow, RoundRect,
    ScreenToClient, SelectObject, SetBkColor, SetBkMode, SetTextColor, UpdateWindow,
    CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET, DEFAULT_PITCH, DT_CENTER, DT_LEFT, DT_NOPREFIX,
    DT_SINGLELINE, DT_VCENTER, FW_NORMAL, FW_SEMIBOLD, HBRUSH, HDC, HFONT, HGDIOBJ, NULL_BRUSH,
    OUT_DEFAULT_PRECIS, PAINTSTRUCT, PROOF_QUALITY, PS_SOLID, RDW_ALLCHILDREN, RDW_ERASE,
    RDW_INVALIDATE, RDW_UPDATENOW, TRANSPARENT,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{
    DRAWITEMSTRUCT, EM_GETSEL, EM_REPLACESEL, EM_SETLIMITTEXT, EM_SETMARGINS, EM_SETSEL, MARGINS,
    MEASUREITEMSTRUCT, ODS_DISABLED, ODS_SELECTED, TBM_SETPOS, TBM_SETRANGE,
};
use windows::Win32::UI::HiDpi::{GetDpiForWindow, GetSystemMetricsForDpi};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetDoubleClickTime, GetFocus, GetKeyState, ReleaseCapture, SetCapture, SetFocus,
    TrackMouseEvent, TME_LEAVE, TRACKMOUSEEVENT, VK_A, VK_BACK, VK_CONTROL, VK_DOWN, VK_ESCAPE,
    VK_MENU, VK_RETURN, VK_SHIFT, VK_TAB, VK_UP,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CallWindowProcW, CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW,
    GetClientRect, GetDlgItem, GetMessageTime, GetMessageW, GetNextDlgTabItem, GetSystemMetrics,
    GetWindowLongPtrW, GetWindowRect, GetWindowTextLengthW, GetWindowTextW, IsWindow, IsZoomed,
    KillTimer, LoadCursorW, PostMessageW, PostQuitMessage, RegisterClassExW, SendMessageW,
    SetForegroundWindow, SetTimer, SetWindowLongPtrW, SetWindowPos, SetWindowTextW, ShowWindow,
    TranslateMessage, BS_OWNERDRAW, CBS_DROPDOWNLIST, CBS_HASSTRINGS, CBS_OWNERDRAWFIXED,
    CB_ADDSTRING, CB_GETCURSEL, CB_GETLBTEXT, CB_GETLBTEXTLEN, CB_SETCURSEL, CS_HREDRAW,
    CS_VREDRAW, CW_USEDEFAULT, EN_CHANGE, ES_AUTOHSCROLL, GWLP_USERDATA, GWLP_WNDPROC, HMENU,
    HTBOTTOM, HTBOTTOMLEFT, HTBOTTOMRIGHT, HTCAPTION, HTCLIENT, HTLEFT, HTRIGHT, HTTOP, HTTOPLEFT,
    HTTOPRIGHT, IDC_ARROW, LBN_SELCHANGE, LBS_HASSTRINGS, LBS_NOTIFY, LBS_OWNERDRAWFIXED,
    LB_ADDSTRING, LB_GETCURSEL, LB_GETTEXT, LB_GETTEXTLEN, LB_RESETCONTENT, LB_SETCURSEL,
    MINMAXINFO, MSG, NCCALCSIZE_PARAMS, SM_CXDOUBLECLK, SM_CXPADDEDBORDER, SM_CXSIZEFRAME,
    SM_CYDOUBLECLK, SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER,
    SW_MAXIMIZE, SW_MINIMIZE, SW_RESTORE, SW_SHOW, SW_SHOWNORMAL, WINDOW_EX_STYLE, WINDOW_STYLE,
    WM_APP, WM_CANCELMODE, WM_CAPTURECHANGED, WM_CHAR, WM_CLOSE, WM_COMMAND, WM_CREATE,
    WM_CTLCOLORBTN, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX, WM_CTLCOLORSTATIC, WM_DESTROY,
    WM_DRAWITEM, WM_ERASEBKGND, WM_GETMINMAXINFO, WM_HSCROLL, WM_KEYDOWN, WM_KILLFOCUS,
    WM_LBUTTONDBLCLK, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MEASUREITEM, WM_MOUSEMOVE, WM_NCCALCSIZE,
    WM_NCDESTROY, WM_NCHITTEST, WM_PAINT, WM_SETFOCUS, WM_SETFONT, WM_SIZE, WM_TIMER, WNDCLASSEXW,
    WNDPROC, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_OVERLAPPED,
    WS_SYSMENU, WS_TABSTOP, WS_THICKFRAME, WS_VISIBLE, WS_VSCROLL,
};

use crate::config::config::{
    CustomPreset, CustomScaling, CustomWindowMode, PreviewScale, TextTheme, TransparentBackground,
    MAX_PRESETS,
};
use crate::text::{text_paint, text_theme};
use crate::CONFIG;

const CLASS_NAME: PCWSTR = w!("RustHoverPreviewCustomPreview");

static OPEN: AtomicBool = AtomicBool::new(false);
static WINDOW: AtomicIsize = AtomicIsize::new(0);

const TBM_GETPOS: u32 = 0x0408;
const WM_MOUSELEAVE: u32 = 0x02A3;
const EC_LEFTMARGIN: usize = 0x0001;
const EC_RIGHTMARGIN: usize = 0x0002;

const WM_BRING_FORWARD: u32 = WM_APP + 17;
const WM_THEME_CHANGED: u32 = WM_APP + 18;
const WM_STEP_NUMERIC: u32 = WM_APP + 19;

const TITLEBAR_LOGICAL_H: i32 = 40;
const TITLEBAR_BTN_W: i32 = 44;
const RESIZE_BORDER_LOGICAL: i32 = 6;
const STEP_DELAY_MS: u32 = 380;
const STEP_REPEAT_MS: u32 = 55;
const STEP_TIMER_INITIAL: usize = 1;
const STEP_TIMER_REPEAT: usize = 2;

const SCALING_MIN: u32 = 1;
const SCALING_MAX: u32 = 1000;
const VOLUME_MIN: u32 = 0;
const VOLUME_MAX: u32 = 100;

const ID_SEARCH: i32 = 100;
const ID_ADD: i32 = 101;
const ID_LIST: i32 = 102;
const ID_NAME: i32 = 110;
const ID_EXTENSIONS: i32 = 111;
const ID_EXE: i32 = 112;
const ID_BROWSE: i32 = 113;
const ID_SCALING: i32 = 114;
const ID_SCALING_PERCENT: i32 = 115;
const ID_BACKGROUND: i32 = 116;
const ID_VOLUME: i32 = 117;
const ID_VOLUME_VALUE: i32 = 118;
const ID_WINDOW_MODE: i32 = 119;
const ID_SAVE: i32 = 120;
const ID_DELETE: i32 = 121;
const ID_CLOSE_BUTTON: i32 = 122;
const ID_SCALING_UP: i32 = 123;
const ID_SCALING_DOWN: i32 = 124;
const ID_VOLUME_UP: i32 = 125;
const ID_VOLUME_DOWN: i32 = 126;
const ID_STATUS: i32 = 130;
const ID_LABEL_BASE: i32 = 140;

const EDIT_IDS: [i32; 6] = [
    ID_SEARCH,
    ID_NAME,
    ID_EXTENSIONS,
    ID_EXE,
    ID_SCALING_PERCENT,
    ID_VOLUME_VALUE,
];

const SCALE_CHOICES: [(&str, CustomScaling); 5] = [
    ("Follow Tray -> Scaling", CustomScaling::FollowTray),
    (
        "Fit to screen",
        CustomScaling::Override(PreviewScale::FitToScreen),
    ),
    (
        "Screen 100%",
        CustomScaling::Override(PreviewScale::FitToScreenReduced(100)),
    ),
    (
        "Screen 75%",
        CustomScaling::Override(PreviewScale::FitToScreenReduced(75)),
    ),
    (
        "Screen 50%",
        CustomScaling::Override(PreviewScale::FitToScreenReduced(50)),
    ),
];

const WINDOW_MODES: [CustomWindowMode; 2] =
    [CustomWindowMode::Normal, CustomWindowMode::Borderless];

fn window_mode_label(mode: CustomWindowMode) -> &'static str {
    match mode {
        CustomWindowMode::Normal => "Normal",
        CustomWindowMode::Borderless => "Borderless",
    }
}

const BACKGROUNDS: [TransparentBackground; 4] = [
    TransparentBackground::Transparent,
    TransparentBackground::Black,
    TransparentBackground::White,
    TransparentBackground::Checkerboard,
];

fn background_label(background: TransparentBackground) -> &'static str {
    match background {
        TransparentBackground::Transparent => "Transparent",
        TransparentBackground::Black => "Black",
        TransparentBackground::White => "White",
        TransparentBackground::Checkerboard => "Checkerboard",
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum TitlebarButton {
    Minimize,
    Maximize,
    Close,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Palette {
    page: COLORREF,
    titlebar: COLORREF,
    card: COLORREF,
    field: COLORREF,
    border: COLORREF,
    ink: COLORREF,
    muted: COLORREF,
    accent: COLORREF,
    accent_text: COLORREF,
    number: COLORREF,
    button_bg: COLORREF,
    button_hover: COLORREF,
    button_press: COLORREF,
    danger: COLORREF,
    danger_bg: COLORREF,
    danger_hover: COLORREF,
    dark: bool,
}

fn rgb(r: u8, g: u8, b: u8) -> COLORREF {
    COLORREF(u32::from(r) | (u32::from(g) << 8) | (u32::from(b) << 16))
}

fn rgb_arr(c: [u8; 3]) -> COLORREF {
    rgb(c[0], c[1], c[2])
}

fn mix(a: COLORREF, b: COLORREF, share: u32) -> COLORREF {
    let part = |shift: u32| -> u8 {
        let one = (a.0 >> shift) & 0xFF;
        let two = (b.0 >> shift) & 0xFF;
        ((one * (100 - share) + two * share) / 100) as u8
    };
    rgb(part(0), part(8), part(16))
}

impl Palette {
    /// Matches audio-card / `ChromePalette` tokens from the active `TextTheme`.
    fn of(theme: TextTheme) -> Self {
        let (bg, fg, accent_arr, number_arr) = match text_theme::loaded(theme) {
            Some(loaded) => {
                let bg = text_paint::rgb(loaded.background());
                let fg = text_paint::rgb(loaded.foreground());
                let accent =
                    text_paint::rgb(loaded.style_for_scopes(&["support.function"]).foreground);
                let number = text_paint::readable(
                    text_paint::rgb(loaded.style_for_scopes(&["constant.numeric"]).foreground),
                    bg,
                );
                (bg, fg, accent, number)
            }
            None => ([250, 250, 250], [56, 58, 66], [1, 132, 188], [152, 104, 1]),
        };

        let dark = text_paint::luminance(bg) <= 0.5;
        let page = rgb_arr(bg);
        let ink = rgb_arr(fg);
        let accent = rgb_arr(accent_arr);
        let number = rgb_arr(number_arr);

        let (titlebar, card, field, border, button_bg, button_hover, button_press) = if dark {
            let card = mix(page, rgb(255, 255, 255), 4);
            (
                mix(page, rgb(0, 0, 0), 10),
                card,
                mix(page, rgb(0, 0, 0), 18),
                mix(page, ink, 18),
                mix(card, ink, 9),
                mix(card, ink, 16),
                mix(card, ink, 24),
            )
        } else {
            let card = mix(page, rgb(0, 0, 0), 3);
            (
                mix(page, rgb(0, 0, 0), 4),
                card,
                rgb(255, 255, 255),
                mix(page, ink, 16),
                mix(card, ink, 5),
                mix(card, ink, 11),
                mix(card, ink, 18),
            )
        };

        let danger = if dark {
            rgb(224, 108, 117)
        } else {
            rgb(201, 59, 69)
        };
        let danger_bg = mix(card, danger, if dark { 16 } else { 12 });
        let danger_hover = rgb(212, 47, 58);
        let accent_text = if text_paint::luminance(accent_arr) > 0.52 {
            rgb(24, 28, 34)
        } else {
            rgb(255, 255, 255)
        };

        Self {
            page,
            titlebar,
            card,
            field,
            border,
            ink,
            muted: mix(ink, page, 45),
            accent,
            accent_text,
            number,
            button_bg,
            button_hover,
            button_press,
            danger,
            danger_bg,
            danger_hover,
            dark,
        }
    }
}

struct LayoutRects {
    titlebar: RECT,
    sidebar_card: RECT,
    editor_card: RECT,
    btn_min: RECT,
    btn_max: RECT,
    btn_close: RECT,
}

impl LayoutRects {
    fn titlebar_button_at(&self, x: i32, y: i32) -> Option<TitlebarButton> {
        if pt_in_rect(self.btn_close, x, y) {
            Some(TitlebarButton::Close)
        } else if pt_in_rect(self.btn_max, x, y) {
            Some(TitlebarButton::Maximize)
        } else if pt_in_rect(self.btn_min, x, y) {
            Some(TitlebarButton::Minimize)
        } else {
            None
        }
    }
}

fn pt_in_rect(r: RECT, x: i32, y: i32) -> bool {
    x >= r.left && x < r.right && y >= r.top && y < r.bottom
}

struct State {
    hwnd: HWND,
    presets: Vec<CustomPreset>,
    selected: Option<usize>,
    shown: Vec<usize>,
    loading: bool,
    palette: Palette,
    font: HFONT,
    font_bold: HFONT,
    font_small: HFONT,
    brush_page: HBRUSH,
    brush_card: HBRUSH,
    brush_field: HBRUSH,
    dpi: i32,
    hovered_tb: Option<TitlebarButton>,
    pressed_tb: Option<TitlebarButton>,
    tracking_mouse: bool,
    rects: LayoutRects,
}

impl Drop for State {
    fn drop(&mut self) {
        unsafe {
            let _ = DeleteObject(HGDIOBJ(self.font.0));
            let _ = DeleteObject(HGDIOBJ(self.font_bold.0));
            let _ = DeleteObject(HGDIOBJ(self.font_small.0));
            let _ = DeleteObject(HGDIOBJ(self.brush_page.0));
            let _ = DeleteObject(HGDIOBJ(self.brush_card.0));
            let _ = DeleteObject(HGDIOBJ(self.brush_field.0));
        }
    }
}

struct EditSubclass {
    orig_proc: WNDPROC,
    parent: HWND,
    id: i32,
    last_click_time: u32,
    last_click_pos: (i32, i32),
    click_count: u8,
}

struct StepperSubclass {
    orig_proc: WNDPROC,
    parent: HWND,
    target_id: i32,
    delta: i32,
    hovered: bool,
    pressed: bool,
    tracking: bool,
}

/// Open the settings window, or bring the existing one forward.
pub(crate) fn open() {
    let existing = WINDOW.load(Ordering::SeqCst);
    if existing != 0 && OPEN.load(Ordering::SeqCst) {
        let hwnd = HWND(existing as *mut _);
        unsafe {
            let _ = ShowWindow(hwnd, SW_SHOWNORMAL);
            let _ = SetForegroundWindow(hwnd);
            let _ = PostMessageW(hwnd, WM_BRING_FORWARD, WPARAM(0), LPARAM(0));
        }
        return;
    }

    if OPEN.swap(true, Ordering::SeqCst) {
        return;
    }

    std::thread::spawn(|| {
        run();
        OPEN.store(false, Ordering::SeqCst);
        WINDOW.store(0, Ordering::SeqCst);
    });
}

/// Notify the open settings window (if any) to rebuild its palette and repaint immediately.
pub(crate) fn notify_theme_changed() {
    let existing = WINDOW.load(Ordering::SeqCst);
    if existing != 0 {
        let hwnd = HWND(existing as *mut _);
        if unsafe { IsWindow(hwnd) }.as_bool() {
            let _ = unsafe { PostMessageW(hwnd, WM_THEME_CHANGED, WPARAM(0), LPARAM(0)) };
        }
    }
}

fn presets_now() -> Vec<CustomPreset> {
    CONFIG
        .lock()
        .map(|config| config.custom_presets.clone())
        .unwrap_or_default()
}

fn theme_now() -> TextTheme {
    CONFIG
        .lock()
        .map(|config| config.theme)
        .unwrap_or(TextTheme::Light)
}

fn persist(state: &State) -> Result<(), String> {
    let mut cleaned: Vec<CustomPreset> = Vec::new();
    for preset in state.presets.iter() {
        let Some(preset) = preset.clone().sanitized() else {
            return Err("A preset needs a name.".to_string());
        };
        if cleaned
            .iter()
            .any(|held| held.name.eq_ignore_ascii_case(&preset.name))
        {
            return Err(format!("Two presets are named \"{}\".", preset.name));
        }
        cleaned.push(preset);
    }

    let Ok(mut config) = CONFIG.lock() else {
        return Err("The configuration is busy; try again.".to_string());
    };
    config.custom_presets = cleaned;
    config.save();
    Ok(())
}

fn run() {
    let Some(hwnd) = create_window() else {
        return;
    };

    unsafe {
        WINDOW.store(hwnd.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = UpdateWindow(hwnd);

        let mut message = MSG::default();
        while GetMessageW(&mut message, HWND::default(), 0, 0).as_bool() {
            // Handle Tab / Shift+Tab cleanly without breaking edit shortcuts or triggering beeps.
            if message.message == WM_KEYDOWN && message.wParam.0 == VK_TAB.0 as usize {
                let ctrl = GetKeyState(VK_CONTROL.0 as i32) < 0;
                let alt = GetKeyState(VK_MENU.0 as i32) < 0;
                if !ctrl && !alt {
                    let shift = GetKeyState(VK_SHIFT.0 as i32) < 0;
                    let focused = GetFocus();
                    if let Ok(next) = GetNextDlgTabItem(hwnd, focused, shift) {
                        if !next.0.is_null() {
                            let _ = SetFocus(next);
                            continue;
                        }
                    }
                }
            }
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

fn create_window() -> Option<HWND> {
    unsafe {
        let hinstance = GetModuleHandleW(None).ok()?;
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(window_proc),
            hInstance: hinstance.into(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hbrBackground: HBRUSH(GetStockObject(NULL_BRUSH).0),
            lpszClassName: CLASS_NAME,
            ..Default::default()
        };
        RegisterClassExW(&class);

        let title = wide("Custom Preview");
        let style = WS_OVERLAPPED
            | WS_CAPTION
            | WS_SYSMENU
            | WS_THICKFRAME
            | WS_MINIMIZEBOX
            | WS_MAXIMIZEBOX
            | WS_CLIPCHILDREN;
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            CLASS_NAME,
            PCWSTR(title.as_ptr()),
            style,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            920,
            620,
            None,
            None,
            HINSTANCE(hinstance.0),
            None,
        )
        .ok()
    }
}

unsafe fn apply_dwm_chrome(hwnd: HWND, dark: bool) {
    let dark_val: i32 = if dark { 1 } else { 0 };
    let _ = unsafe {
        DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            &dark_val as *const _ as *const _,
            std::mem::size_of::<i32>() as u32,
        )
    };
    let corner = DWMWCP_ROUND;
    let _ = unsafe {
        DwmSetWindowAttribute(
            hwnd,
            DWMWA_WINDOW_CORNER_PREFERENCE,
            &corner as *const _ as *const _,
            std::mem::size_of_val(&corner) as u32,
        )
    };
    // 1px top DWM glass margin preserves native Win11 drop shadow & rounded corners on frameless windows.
    let margins = MARGINS {
        cxLeftWidth: 0,
        cxRightWidth: 0,
        cyTopHeight: 1,
        cyBottomHeight: 0,
    };
    let _ = unsafe { DwmExtendFrameIntoClientArea(hwnd, &margins) };
}

unsafe fn apply_theme(state: &mut State) {
    let palette = Palette::of(theme_now());
    state.palette = palette;
    unsafe {
        let _ = DeleteObject(HGDIOBJ(state.brush_page.0));
        let _ = DeleteObject(HGDIOBJ(state.brush_card.0));
        let _ = DeleteObject(HGDIOBJ(state.brush_field.0));
        state.brush_page = CreateSolidBrush(palette.page);
        state.brush_card = CreateSolidBrush(palette.card);
        state.brush_field = CreateSolidBrush(palette.field);
        apply_dwm_chrome(state.hwnd, palette.dark);
        let _ = RedrawWindow(
            state.hwnd,
            None,
            None,
            RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW,
        );
    }
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let state = unsafe { build(hwnd) };
            let dark = state.palette.dark;
            let raw = Box::into_raw(Box::new(state));
            unsafe {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, raw as isize);
                apply_dwm_chrome(hwnd, dark);
                let _ = SetWindowPos(
                    hwnd,
                    HWND::default(),
                    0,
                    0,
                    0,
                    0,
                    SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE,
                );
            }
            LRESULT(0)
        }
        WM_NCCALCSIZE => {
            if wparam.0 == 0 {
                return LRESULT(0);
            }
            if unsafe { IsZoomed(hwnd) }.as_bool() {
                let params = unsafe { &mut *(lparam.0 as *mut NCCALCSIZE_PARAMS) };
                let dpi = (unsafe { GetDpiForWindow(hwnd) }).max(96);
                let frame = unsafe {
                    GetSystemMetricsForDpi(SM_CXSIZEFRAME, dpi)
                        + GetSystemMetricsForDpi(SM_CXPADDEDBORDER, dpi)
                };
                params.rgrc[0].left += frame;
                params.rgrc[0].top += frame;
                params.rgrc[0].right -= frame;
                params.rgrc[0].bottom -= frame;
            }
            LRESULT(0)
        }
        WM_NCHITTEST => {
            let Some(state) = (unsafe { state_of(hwnd) }) else {
                return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            };
            let mut pt = POINT {
                x: (lparam.0 & 0xFFFF) as i16 as i32,
                y: ((lparam.0 >> 16) & 0xFFFF) as i16 as i32,
            };
            let _ = unsafe { ScreenToClient(hwnd, &mut pt) };
            let mut client = RECT::default();
            let _ = unsafe { GetClientRect(hwnd, &mut client) };
            let zoomed = unsafe { IsZoomed(hwnd) }.as_bool();
            let border = scale(RESIZE_BORDER_LOGICAL, state.dpi);

            if !zoomed {
                let left = pt.x < border;
                let right = pt.x >= client.right - border;
                let top = pt.y < border;
                let bottom = pt.y >= client.bottom - border;
                if top && left {
                    return LRESULT(HTTOPLEFT as isize);
                }
                if top && right {
                    return LRESULT(HTTOPRIGHT as isize);
                }
                if bottom && left {
                    return LRESULT(HTBOTTOMLEFT as isize);
                }
                if bottom && right {
                    return LRESULT(HTBOTTOMRIGHT as isize);
                }
                if left {
                    return LRESULT(HTLEFT as isize);
                }
                if right {
                    return LRESULT(HTRIGHT as isize);
                }
                if top {
                    return LRESULT(HTTOP as isize);
                }
                if bottom {
                    return LRESULT(HTBOTTOM as isize);
                }
            }

            if pt.y < state.rects.titlebar.bottom {
                if state.rects.titlebar_button_at(pt.x, pt.y).is_some() {
                    return LRESULT(HTCLIENT as isize);
                }
                return LRESULT(HTCAPTION as isize);
            }
            LRESULT(HTCLIENT as isize)
        }
        WM_MOUSEMOVE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
                let next = state.rects.titlebar_button_at(x, y);
                if next != state.hovered_tb {
                    state.hovered_tb = next;
                    unsafe {
                        let _ = InvalidateRect(hwnd, Some(&state.rects.titlebar), false);
                    }
                }
                if !state.tracking_mouse {
                    state.tracking_mouse = true;
                    let mut tme = TRACKMOUSEEVENT {
                        cbSize: std::mem::size_of::<TRACKMOUSEEVENT>() as u32,
                        dwFlags: TME_LEAVE,
                        hwndTrack: hwnd,
                        dwHoverTime: 0,
                    };
                    let _ = unsafe { TrackMouseEvent(&mut tme) };
                }
            }
            LRESULT(0)
        }
        WM_MOUSELEAVE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                state.tracking_mouse = false;
                if state.hovered_tb.is_some() || state.pressed_tb.is_some() {
                    state.hovered_tb = None;
                    state.pressed_tb = None;
                    unsafe {
                        let _ = InvalidateRect(hwnd, Some(&state.rects.titlebar), false);
                    }
                }
            }
            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
                if let Some(btn) = state.rects.titlebar_button_at(x, y) {
                    state.pressed_tb = Some(btn);
                    unsafe {
                        SetCapture(hwnd);
                        let _ = InvalidateRect(hwnd, Some(&state.rects.titlebar), false);
                    }
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_LBUTTONUP => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                if let Some(pressed) = state.pressed_tb.take() {
                    let _ = unsafe { ReleaseCapture() };
                    let x = (lparam.0 & 0xFFFF) as i16 as i32;
                    let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
                    unsafe {
                        let _ = InvalidateRect(hwnd, Some(&state.rects.titlebar), false);
                    }
                    if state.rects.titlebar_button_at(x, y) == Some(pressed) {
                        match pressed {
                            TitlebarButton::Minimize => {
                                let _ = unsafe { ShowWindow(hwnd, SW_MINIMIZE) };
                            }
                            TitlebarButton::Maximize => {
                                let cmd = if unsafe { IsZoomed(hwnd) }.as_bool() {
                                    SW_RESTORE
                                } else {
                                    SW_MAXIMIZE
                                };
                                let _ = unsafe { ShowWindow(hwnd, cmd) };
                            }
                            TitlebarButton::Close => {
                                let _ =
                                    unsafe { PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) };
                            }
                        }
                    }
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_CAPTURECHANGED | WM_CANCELMODE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                if state.pressed_tb.take().is_some() {
                    unsafe {
                        let _ = InvalidateRect(hwnd, Some(&state.rects.titlebar), false);
                    }
                }
            }
            LRESULT(0)
        }
        WM_SIZE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { lay_out(state) };
            }
            LRESULT(0)
        }
        WM_GETMINMAXINFO => {
            let info = unsafe { &mut *(lparam.0 as *mut MINMAXINFO) };
            let dpi = (unsafe { GetDpiForWindow(hwnd) } as i32).max(96);
            info.ptMinTrackSize.x = scale(680, dpi);
            info.ptMinTrackSize.y = scale(460, dpi);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xFFFF) as i32;
            let notice = ((wparam.0 >> 16) & 0xFFFF) as u32;
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { command(state, id, notice) };
            }
            LRESULT(0)
        }
        WM_STEP_NUMERIC => {
            let target_id = wparam.0 as i32;
            let delta = lparam.0 as i32;
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { step_numeric_field(state, target_id, delta) };
            }
            LRESULT(0)
        }
        WM_HSCROLL => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { volume_moved(state) };
            }
            LRESULT(0)
        }
        WM_MEASUREITEM => {
            let item = unsafe { &mut *(lparam.0 as *mut MEASUREITEMSTRUCT) };
            let dpi = unsafe { state_of(hwnd) }.map_or(96, |state| state.dpi);
            item.itemHeight = scale(26, dpi) as u32;
            LRESULT(1)
        }
        WM_DRAWITEM => {
            let item = unsafe { &*(lparam.0 as *const DRAWITEMSTRUCT) };
            let drawn = unsafe { state_of(hwnd) }.map(|state| unsafe { draw_item(state, item) });
            LRESULT(drawn.unwrap_or(0) as isize)
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLORBTN => {
            let Some(state) = (unsafe { state_of(hwnd) }) else {
                return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            };
            let hdc = HDC(wparam.0 as *mut _);
            let ctl = HWND(lparam.0 as *mut _);
            let status_hwnd = unsafe { GetDlgItem(hwnd, ID_STATUS) }.unwrap_or_default();
            let (bg, brush, ink) = if ctl == status_hwnd {
                (state.palette.page, state.brush_page, state.palette.muted)
            } else {
                (state.palette.card, state.brush_card, state.palette.muted)
            };
            unsafe {
                SetTextColor(hdc, ink);
                SetBkColor(hdc, bg);
            }
            LRESULT(brush.0 as isize)
        }
        WM_CTLCOLOREDIT => {
            let Some(state) = (unsafe { state_of(hwnd) }) else {
                return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            };
            let hdc = HDC(wparam.0 as *mut _);
            unsafe {
                SetTextColor(hdc, state.palette.ink);
                SetBkColor(hdc, state.palette.field);
            }
            LRESULT(state.brush_field.0 as isize)
        }
        WM_CTLCOLORLISTBOX => {
            let Some(state) = (unsafe { state_of(hwnd) }) else {
                return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            };
            let hdc = HDC(wparam.0 as *mut _);
            unsafe {
                SetTextColor(hdc, state.palette.ink);
                SetBkColor(hdc, state.palette.field);
            }
            LRESULT(state.brush_field.0 as isize)
        }
        WM_ERASEBKGND => LRESULT(1),
        WM_PAINT => {
            let mut paint = PAINTSTRUCT::default();
            let hdc = unsafe { BeginPaint(hwnd, &mut paint) };
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { paint_window(state, hdc) };
            }
            let _ = unsafe { EndPaint(hwnd, &paint) };
            LRESULT(0)
        }
        WM_THEME_CHANGED => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { apply_theme(state) };
            }
            LRESULT(0)
        }
        WM_BRING_FORWARD => {
            unsafe {
                let _ = SetForegroundWindow(hwnd);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                let _ = persist(state);
            }
            unsafe {
                let _ = DestroyWindow(hwnd);
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            unsafe { PostQuitMessage(0) };
            LRESULT(0)
        }
        WM_NCDESTROY => {
            WINDOW.store(0, Ordering::SeqCst);
            let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut State;
            if !raw.is_null() {
                unsafe {
                    let _ = SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                    drop(Box::from_raw(raw));
                }
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

unsafe fn paint_window(state: &State, hdc: HDC) {
    let mut client = RECT::default();
    let _ = unsafe { GetClientRect(state.hwnd, &mut client) };
    let p = state.palette;
    let dpi = state.dpi;

    unsafe {
        // Page background.
        FillRect(hdc, &client, state.brush_page);

        // Titlebar band + bottom hairline.
        let tb_brush = CreateSolidBrush(p.titlebar);
        FillRect(hdc, &state.rects.titlebar, tb_brush);
        let _ = DeleteObject(HGDIOBJ(tb_brush.0));

        let border_pen = CreatePen(PS_SOLID, 1, p.border);
        let old_pen = SelectObject(hdc, HGDIOBJ(border_pen.0));
        let _ = MoveToEx(hdc, 0, state.rects.titlebar.bottom - 1, None);
        let _ = LineTo(hdc, client.right, state.rects.titlebar.bottom - 1);

        // Accent dot + title on left (audio-card motif).
        let dot_d = scale(8, dpi);
        let dot_x = scale(16, dpi);
        let dot_y = (state.rects.titlebar.bottom - dot_d) / 2;
        let accent_brush = CreateSolidBrush(p.accent);
        let accent_pen = CreatePen(PS_SOLID, 1, p.accent);
        let _ = SelectObject(hdc, HGDIOBJ(accent_pen.0));
        let old_brush = SelectObject(hdc, HGDIOBJ(accent_brush.0));
        let _ = RoundRect(
            hdc,
            dot_x,
            dot_y,
            dot_x + dot_d,
            dot_y + dot_d,
            dot_d,
            dot_d,
        );
        let _ = SelectObject(hdc, old_brush);
        let _ = SelectObject(hdc, HGDIOBJ(border_pen.0));
        let _ = DeleteObject(HGDIOBJ(accent_brush.0));
        let _ = DeleteObject(HGDIOBJ(accent_pen.0));

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, p.ink);
        let old_font = SelectObject(hdc, HGDIOBJ(state.font_bold.0));
        let mut title_rc = RECT {
            left: dot_x + dot_d + scale(10, dpi),
            top: 0,
            right: state.rects.btn_min.left - scale(8, dpi),
            bottom: state.rects.titlebar.bottom - 1,
        };
        let mut title_w: Vec<u16> = "Custom Preview".encode_utf16().collect();
        DrawTextW(
            hdc,
            &mut title_w,
            &mut title_rc,
            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
        );
        let _ = SelectObject(hdc, old_font);

        // Titlebar window buttons (Minimize, Maximize/Restore, Close).
        paint_titlebar_button(state, hdc, state.rects.btn_min, TitlebarButton::Minimize);
        paint_titlebar_button(state, hdc, state.rects.btn_max, TitlebarButton::Maximize);
        paint_titlebar_button(state, hdc, state.rects.btn_close, TitlebarButton::Close);

        // Elevated sidebar & editor cards with 1px theme border.
        let radius = scale(8, dpi);
        let _ = SelectObject(hdc, HGDIOBJ(border_pen.0));
        let _ = SelectObject(hdc, HGDIOBJ(state.brush_card.0));
        let sc = state.rects.sidebar_card;
        let _ = RoundRect(hdc, sc.left, sc.top, sc.right, sc.bottom, radius, radius);
        let ec = state.rects.editor_card;
        let _ = RoundRect(hdc, ec.left, ec.top, ec.right, ec.bottom, radius, radius);

        // Border & focus accent ring around each EDIT control and LISTBOX.
        let focused = GetFocus();
        let focus_pen = CreatePen(PS_SOLID, 1, p.accent);
        for id in EDIT_IDS.into_iter().chain(std::iter::once(ID_LIST)) {
            let Ok(ctrl) = GetDlgItem(state.hwnd, id) else {
                continue;
            };
            let mut r = RECT::default();
            if GetWindowRect(ctrl, &mut r).is_err() {
                continue;
            }
            let mut tl = POINT {
                x: r.left,
                y: r.top,
            };
            let mut br = POINT {
                x: r.right,
                y: r.bottom,
            };
            let _ = ScreenToClient(state.hwnd, &mut tl);
            let _ = ScreenToClient(state.hwnd, &mut br);
            let is_focused = ctrl == focused;
            let _ = SelectObject(
                hdc,
                HGDIOBJ(if is_focused {
                    focus_pen.0
                } else {
                    border_pen.0
                }),
            );
            let _ = SelectObject(hdc, HGDIOBJ(state.brush_field.0));
            let _ = RoundRect(
                hdc,
                tl.x - 1,
                tl.y - 1,
                br.x + 1,
                br.y + 1,
                scale(4, dpi),
                scale(4, dpi),
            );
        }

        let _ = SelectObject(hdc, old_pen);
        let _ = DeleteObject(HGDIOBJ(border_pen.0));
        let _ = DeleteObject(HGDIOBJ(focus_pen.0));
    }
}

unsafe fn paint_titlebar_button(state: &State, hdc: HDC, rect: RECT, kind: TitlebarButton) {
    let p = state.palette;
    let hovered = state.hovered_tb == Some(kind);
    let pressed = state.pressed_tb == Some(kind);

    let bg = match (kind, pressed, hovered) {
        (TitlebarButton::Close, true, _) => mix(p.danger_hover, rgb(0, 0, 0), 18),
        (TitlebarButton::Close, false, true) => p.danger_hover,
        (_, true, _) => p.button_press,
        (_, false, true) => p.button_hover,
        _ => p.titlebar,
    };
    let fg = if kind == TitlebarButton::Close && (hovered || pressed) {
        rgb(255, 255, 255)
    } else if hovered || pressed {
        p.ink
    } else {
        p.muted
    };

    unsafe {
        let fill_rc = RECT {
            left: rect.left,
            top: rect.top,
            right: rect.right,
            bottom: rect.bottom - 1,
        };
        let brush = CreateSolidBrush(bg);
        FillRect(hdc, &fill_rc, brush);
        let _ = DeleteObject(HGDIOBJ(brush.0));

        let pen = CreatePen(PS_SOLID, 1, fg);
        let old_pen = SelectObject(hdc, HGDIOBJ(pen.0));
        let old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));

        let cx = (rect.left + rect.right) / 2;
        let cy = (rect.top + rect.bottom - 1) / 2;
        let half = scale(5, state.dpi);

        match kind {
            TitlebarButton::Minimize => {
                let _ = MoveToEx(hdc, cx - half, cy, None);
                let _ = LineTo(hdc, cx + half + 1, cy);
            }
            TitlebarButton::Maximize => {
                if IsZoomed(state.hwnd).as_bool() {
                    let s = half - 1;
                    // Back square (top-right offset).
                    let _ = MoveToEx(hdc, cx - s + 2, cy - s - 2, None);
                    let _ = LineTo(hdc, cx + s + 2, cy - s - 2);
                    let _ = LineTo(hdc, cx + s + 2, cy + s - 2);
                    // Front square.
                    let _ = MoveToEx(hdc, cx - s, cy - s, None);
                    let _ = LineTo(hdc, cx + s, cy - s);
                    let _ = LineTo(hdc, cx + s, cy + s);
                    let _ = LineTo(hdc, cx - s, cy + s);
                    let _ = LineTo(hdc, cx - s, cy - s);
                } else {
                    let _ = MoveToEx(hdc, cx - half, cy - half, None);
                    let _ = LineTo(hdc, cx + half, cy - half);
                    let _ = LineTo(hdc, cx + half, cy + half);
                    let _ = LineTo(hdc, cx - half, cy + half);
                    let _ = LineTo(hdc, cx - half, cy - half);
                }
            }
            TitlebarButton::Close => {
                let _ = MoveToEx(hdc, cx - half, cy - half, None);
                let _ = LineTo(hdc, cx + half + 1, cy + half + 1);
                let _ = MoveToEx(hdc, cx + half, cy - half, None);
                let _ = LineTo(hdc, cx - half - 1, cy + half + 1);
            }
        }

        let _ = SelectObject(hdc, old_brush);
        let _ = SelectObject(hdc, old_pen);
        let _ = DeleteObject(HGDIOBJ(pen.0));
    }
}

unsafe fn state_of(hwnd: HWND) -> Option<&'static mut State> {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut State;
    unsafe { raw.as_mut() }
}

fn scale(value: i32, dpi: i32) -> i32 {
    (value * dpi + 48) / 96
}

fn child_style(extra: u32) -> WINDOW_STYLE {
    WINDOW_STYLE(WS_CHILD.0 | WS_VISIBLE.0 | extra)
}

unsafe fn make_font(dpi: i32, size: i32, weight: i32) -> HFONT {
    let face = wide("Segoe UI");
    unsafe {
        CreateFontW(
            -scale(size, dpi),
            0,
            0,
            0,
            weight,
            0,
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            PROOF_QUALITY.0 as u32,
            DEFAULT_PITCH.0 as u32,
            PCWSTR(face.as_ptr()),
        )
    }
}

unsafe fn build(hwnd: HWND) -> State {
    let dpi = (unsafe { GetDpiForWindow(hwnd) } as i32).max(96);
    let palette = Palette::of(theme_now());

    let font = unsafe { make_font(dpi, 14, FW_NORMAL.0 as i32) };
    let font_bold = unsafe { make_font(dpi, 14, FW_SEMIBOLD.0 as i32) };
    let font_small = unsafe { make_font(dpi, 10, FW_NORMAL.0 as i32) };

    let mut state = State {
        hwnd,
        presets: presets_now(),
        selected: None,
        shown: Vec::new(),
        loading: false,
        palette,
        font,
        font_bold,
        font_small,
        brush_page: unsafe { CreateSolidBrush(palette.page) },
        brush_card: unsafe { CreateSolidBrush(palette.card) },
        brush_field: unsafe { CreateSolidBrush(palette.field) },
        dpi,
        hovered_tb: None,
        pressed_tb: None,
        tracking_mouse: false,
        rects: LayoutRects {
            titlebar: RECT::default(),
            sidebar_card: RECT::default(),
            editor_card: RECT::default(),
            btn_min: RECT::default(),
            btn_max: RECT::default(),
            btn_close: RECT::default(),
        },
    };

    let mut controls: Vec<HWND> = Vec::new();
    let edit_style = child_style(WS_TABSTOP.0 | ES_AUTOHSCROLL as u32);

    // Sidebar: search, Add, preset list.
    controls.push(unsafe { child(hwnd, w!("EDIT"), "", edit_style, ID_SEARCH) });
    controls.push(unsafe { button(hwnd, "Add", ID_ADD) });
    controls.push(unsafe {
        child(
            hwnd,
            w!("LISTBOX"),
            "",
            child_style(
                WS_TABSTOP.0
                    | WS_VSCROLL.0
                    | LBS_NOTIFY as u32
                    | LBS_OWNERDRAWFIXED as u32
                    | LBS_HASSTRINGS as u32,
            ),
            ID_LIST,
        )
    });

    // Editor pane: labels & inputs.
    for index in 0..7 {
        controls.push(unsafe {
            child(
                hwnd,
                w!("STATIC"),
                "",
                child_style(0),
                ID_LABEL_BASE + index,
            )
        });
    }
    controls.push(unsafe { child(hwnd, w!("EDIT"), "", edit_style, ID_NAME) });
    controls.push(unsafe { child(hwnd, w!("EDIT"), "", edit_style, ID_EXTENSIONS) });
    controls.push(unsafe { child(hwnd, w!("EDIT"), "", edit_style, ID_EXE) });
    controls.push(unsafe { button(hwnd, "Browse…", ID_BROWSE) });
    controls.push(unsafe {
        child(
            hwnd,
            w!("COMBOBOX"),
            "",
            child_style(
                WS_TABSTOP.0
                    | WS_VSCROLL.0
                    | CBS_DROPDOWNLIST as u32
                    | CBS_OWNERDRAWFIXED as u32
                    | CBS_HASSTRINGS as u32,
            ),
            ID_WINDOW_MODE,
        )
    });
    controls.push(unsafe {
        child(
            hwnd,
            w!("COMBOBOX"),
            "",
            child_style(
                WS_TABSTOP.0
                    | WS_VSCROLL.0
                    | CBS_DROPDOWNLIST as u32
                    | CBS_OWNERDRAWFIXED as u32
                    | CBS_HASSTRINGS as u32,
            ),
            ID_SCALING,
        )
    });
    controls.push(unsafe { child(hwnd, w!("EDIT"), "", edit_style, ID_SCALING_PERCENT) });
    let scaling_up = unsafe { stepper_button(hwnd, "▲", ID_SCALING_UP, ID_SCALING_PERCENT, 1) };
    let scaling_dn =
        unsafe { stepper_button(hwnd, "▼", ID_SCALING_DOWN, ID_SCALING_PERCENT, -1) };
    controls.push(unsafe {
        child(
            hwnd,
            w!("COMBOBOX"),
            "",
            child_style(
                WS_TABSTOP.0
                    | WS_VSCROLL.0
                    | CBS_DROPDOWNLIST as u32
                    | CBS_OWNERDRAWFIXED as u32
                    | CBS_HASSTRINGS as u32,
            ),
            ID_BACKGROUND,
        )
    });
    controls.push(unsafe {
        child(
            hwnd,
            w!("msctls_trackbar32"),
            "",
            child_style(WS_TABSTOP.0),
            ID_VOLUME,
        )
    });
    controls.push(unsafe { child(hwnd, w!("EDIT"), "", edit_style, ID_VOLUME_VALUE) });
    let vol_up = unsafe { stepper_button(hwnd, "▲", ID_VOLUME_UP, ID_VOLUME_VALUE, 1) };
    let vol_dn = unsafe { stepper_button(hwnd, "▼", ID_VOLUME_DOWN, ID_VOLUME_VALUE, -1) };

    // Footer: status & actions.
    controls.push(unsafe { child(hwnd, w!("STATIC"), "", child_style(0), ID_STATUS) });
    controls.push(unsafe { button(hwnd, "Save", ID_SAVE) });
    controls.push(unsafe { button(hwnd, "Delete", ID_DELETE) });
    controls.push(unsafe { button(hwnd, "Close", ID_CLOSE_BUTTON) });

    for control in controls {
        if !control.0.is_null() {
            unsafe {
                SendMessageW(
                    control,
                    WM_SETFONT,
                    WPARAM(state.font.0 as usize),
                    LPARAM(1),
                );
            }
        }
    }
    for stepper in [scaling_up, scaling_dn, vol_up, vol_dn] {
        if !stepper.0.is_null() {
            unsafe {
                SendMessageW(
                    stepper,
                    WM_SETFONT,
                    WPARAM(state.font_small.0 as usize),
                    LPARAM(1),
                );
            }
        }
    }

    // Configure every EDIT: generous text limits, inner margins, and keyboard/triple-click subclass.
    let edit_limits = [
        (ID_SEARCH, 32767usize),
        (ID_NAME, 48),
        (ID_EXTENSIONS, 32767),
        (ID_EXE, 32767),
        (ID_SCALING_PERCENT, 4),
        (ID_VOLUME_VALUE, 3),
    ];
    let margin = scale(6, dpi) as usize;
    let margin_lparam = LPARAM(((margin << 16) | (margin & 0xFFFF)) as isize);
    for (id, limit) in edit_limits {
        if let Ok(edit) = unsafe { GetDlgItem(hwnd, id) } {
            unsafe {
                SendMessageW(edit, EM_SETLIMITTEXT, WPARAM(limit), LPARAM(0));
                SendMessageW(
                    edit,
                    EM_SETMARGINS,
                    WPARAM(EC_LEFTMARGIN | EC_RIGHTMARGIN),
                    margin_lparam,
                );
                install_edit_subclass(edit, hwnd, id);
            }
        }
    }

    let window_mode_combo = unsafe { GetDlgItem(hwnd, ID_WINDOW_MODE) }.unwrap_or_default();
    for mode in WINDOW_MODES {
        unsafe { combo_add(window_mode_combo, window_mode_label(mode)) };
    }

    let scaling = unsafe { GetDlgItem(hwnd, ID_SCALING) }.unwrap_or_default();
    for (label, _) in SCALE_CHOICES {
        unsafe { combo_add(scaling, label) };
    }
    unsafe { combo_add(scaling, "Custom…") };

    let background_combo = unsafe { GetDlgItem(hwnd, ID_BACKGROUND) }.unwrap_or_default();
    for background in BACKGROUNDS {
        unsafe { combo_add(background_combo, background_label(background)) };
    }

    let volume = unsafe { GetDlgItem(hwnd, ID_VOLUME) }.unwrap_or_default();
    unsafe {
        SendMessageW(volume, TBM_SETRANGE, WPARAM(1), LPARAM(0x640000));
    }

    if !state.presets.is_empty() {
        state.selected = Some(0);
    }
    unsafe { lay_out(&mut state) };
    unsafe { rebuild_list(&mut state) };
    unsafe { fill_fields(&mut state) };
    state
}

unsafe fn install_edit_subclass(edit: HWND, parent: HWND, id: i32) {
    let orig = unsafe {
        SetWindowLongPtrW(
            edit,
            GWLP_WNDPROC,
            edit_subclass_proc as *const () as usize as isize,
        )
    };
    let orig_proc: WNDPROC = unsafe { std::mem::transmute(orig) };
    let data = Box::into_raw(Box::new(EditSubclass {
        orig_proc,
        parent,
        id,
        last_click_time: 0,
        last_click_pos: (0, 0),
        click_count: 0,
    }));
    unsafe { SetWindowLongPtrW(edit, GWLP_USERDATA, data as isize) };
}

unsafe extern "system" fn edit_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut EditSubclass;
    let Some(sub) = (unsafe { raw.as_mut() }) else {
        return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
    };
    let orig_proc = sub.orig_proc;

    match msg {
        WM_KEYDOWN => {
            let vk = wparam.0 as u16;
            let ctrl = unsafe { GetKeyState(VK_CONTROL.0 as i32) } < 0;
            let alt = unsafe { GetKeyState(VK_MENU.0 as i32) } < 0;

            if ctrl && !alt && vk == VK_A.0 {
                unsafe { SendMessageW(hwnd, EM_SETSEL, WPARAM(0), LPARAM(-1)) };
                return LRESULT(0);
            }
            if ctrl && !alt && vk == VK_BACK.0 {
                unsafe { edit_delete_prev_word(hwnd) };
                return LRESULT(0);
            }
            if !ctrl && !alt && (sub.id == ID_SCALING_PERCENT || sub.id == ID_VOLUME_VALUE) {
                if vk == VK_UP.0 {
                    unsafe {
                        SendMessageW(
                            sub.parent,
                            WM_STEP_NUMERIC,
                            WPARAM(sub.id as usize),
                            LPARAM(1),
                        );
                    }
                    return LRESULT(0);
                }
                if vk == VK_DOWN.0 {
                    unsafe {
                        SendMessageW(
                            sub.parent,
                            WM_STEP_NUMERIC,
                            WPARAM(sub.id as usize),
                            LPARAM(-1),
                        );
                    }
                    return LRESULT(0);
                }
            }
            if vk == VK_RETURN.0 || vk == VK_ESCAPE.0 {
                return LRESULT(0);
            }
        }
        WM_CHAR => {
            // Suppress single-line EDIT control beeps for Ctrl+A (0x01), Ctrl+BS (0x7F), Tab, Enter, Esc.
            let ch = wparam.0 as u32;
            if matches!(ch, 0x01 | 0x09 | 0x0D | 0x1B | 0x7F) {
                return LRESULT(0);
            }
        }
        WM_LBUTTONDOWN | WM_LBUTTONDBLCLK => {
            let now = unsafe { GetMessageTime() } as u32;
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
            let dbl_time = unsafe { GetDoubleClickTime() };
            let dx_max = unsafe { GetSystemMetrics(SM_CXDOUBLECLK) }.max(4);
            let dy_max = unsafe { GetSystemMetrics(SM_CYDOUBLECLK) }.max(4);

            let within = sub.click_count > 0
                && now.wrapping_sub(sub.last_click_time) <= dbl_time
                && (x - sub.last_click_pos.0).abs() <= dx_max
                && (y - sub.last_click_pos.1).abs() <= dy_max;

            sub.click_count = if within { sub.click_count + 1 } else { 1 };
            sub.last_click_time = now;
            sub.last_click_pos = (x, y);

            if sub.click_count >= 3 {
                sub.click_count = 0;
                unsafe {
                    let _ = SetFocus(hwnd);
                    SendMessageW(hwnd, EM_SETSEL, WPARAM(0), LPARAM(-1));
                }
                return LRESULT(0);
            }
        }
        WM_SETFOCUS | WM_KILLFOCUS => {
            let res = unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) };
            unsafe {
                let _ = InvalidateRect(sub.parent, None, false);
            }
            return res;
        }
        WM_NCDESTROY => {
            unsafe {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                drop(Box::from_raw(raw));
            }
            return unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) };
        }
        _ => {}
    }

    unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) }
}

unsafe fn edit_delete_prev_word(edit: HWND) {
    let mut start = 0u32;
    let mut end = 0u32;
    unsafe {
        SendMessageW(
            edit,
            EM_GETSEL,
            WPARAM(&mut start as *mut u32 as usize),
            LPARAM(&mut end as *mut u32 as isize),
        );
    }
    let length = unsafe { GetWindowTextLengthW(edit) }.max(0) as usize;
    let mut buf = vec![0u16; length + 1];
    let read = unsafe { GetWindowTextW(edit, &mut buf) }.max(0) as usize;
    let units = &buf[..read];
    let (del_start, del_end) = prev_word_delete_range_utf16(units, start as usize, end as usize);
    if del_start < del_end {
        let empty = [0u16];
        unsafe {
            SendMessageW(edit, EM_SETSEL, WPARAM(del_start), LPARAM(del_end as isize));
            SendMessageW(
                edit,
                EM_REPLACESEL,
                WPARAM(1),
                LPARAM(empty.as_ptr() as isize),
            );
        }
    }
}

/// Computes UTF-16 `[start, end)` range to delete on `Ctrl+Backspace`.
fn prev_word_delete_range_utf16(units: &[u16], sel_start: usize, sel_end: usize) -> (usize, usize) {
    let len = units.len();
    let s = sel_start.min(len);
    let e = sel_end.min(len);
    if s < e {
        return (s, e);
    }
    if s == 0 {
        return (0, 0);
    }
    let is_space = |u: u16| matches!(u, 0x20 | 0x09);
    let is_word =
        |u: u16| char::from_u32(u as u32).is_some_and(|c| c.is_alphanumeric() || c == '_');
    let mut i = s;
    while i > 0 && is_space(units[i - 1]) {
        i -= 1;
    }
    if i > 0 {
        let word = is_word(units[i - 1]);
        while i > 0 && !is_space(units[i - 1]) && is_word(units[i - 1]) == word {
            i -= 1;
        }
    }
    (i, s)
}

unsafe fn stepper_button(hwnd: HWND, glyph: &str, id: i32, target_id: i32, delta: i32) -> HWND {
    let btn = unsafe {
        child(
            hwnd,
            w!("BUTTON"),
            glyph,
            child_style(BS_OWNERDRAW as u32),
            id,
        )
    };
    if !btn.0.is_null() {
        let orig = unsafe {
            SetWindowLongPtrW(
                btn,
                GWLP_WNDPROC,
                stepper_subclass_proc as *const () as usize as isize,
            )
        };
        let orig_proc: WNDPROC = unsafe { std::mem::transmute(orig) };
        let data = Box::into_raw(Box::new(StepperSubclass {
            orig_proc,
            parent: hwnd,
            target_id,
            delta,
            hovered: false,
            pressed: false,
            tracking: false,
        }));
        unsafe { SetWindowLongPtrW(btn, GWLP_USERDATA, data as isize) };
    }
    btn
}

unsafe extern "system" fn stepper_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut StepperSubclass;
    let Some(sub) = (unsafe { raw.as_mut() }) else {
        return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
    };
    let orig_proc = sub.orig_proc;

    match msg {
        WM_LBUTTONDOWN | WM_LBUTTONDBLCLK => {
            sub.pressed = true;
            unsafe {
                SetCapture(hwnd);
                SendMessageW(
                    sub.parent,
                    WM_STEP_NUMERIC,
                    WPARAM(sub.target_id as usize),
                    LPARAM(sub.delta as isize),
                );
                let _ = KillTimer(hwnd, STEP_TIMER_INITIAL);
                let _ = KillTimer(hwnd, STEP_TIMER_REPEAT);
                let _ = SetTimer(hwnd, STEP_TIMER_INITIAL, STEP_DELAY_MS, None);
                let _ = InvalidateRect(hwnd, None, false);
            }
            return LRESULT(0);
        }
        WM_TIMER => {
            if wparam.0 == STEP_TIMER_INITIAL {
                unsafe {
                    let _ = KillTimer(hwnd, STEP_TIMER_INITIAL);
                    SendMessageW(
                        sub.parent,
                        WM_STEP_NUMERIC,
                        WPARAM(sub.target_id as usize),
                        LPARAM(sub.delta as isize),
                    );
                    let _ = SetTimer(hwnd, STEP_TIMER_REPEAT, STEP_REPEAT_MS, None);
                }
                return LRESULT(0);
            }
            if wparam.0 == STEP_TIMER_REPEAT {
                unsafe {
                    SendMessageW(
                        sub.parent,
                        WM_STEP_NUMERIC,
                        WPARAM(sub.target_id as usize),
                        LPARAM(sub.delta as isize),
                    );
                }
                return LRESULT(0);
            }
        }
        WM_LBUTTONUP => {
            sub.pressed = false;
            unsafe {
                let _ = KillTimer(hwnd, STEP_TIMER_INITIAL);
                let _ = KillTimer(hwnd, STEP_TIMER_REPEAT);
                let _ = ReleaseCapture();
                let _ = InvalidateRect(hwnd, None, false);
            }
            return LRESULT(0);
        }
        WM_CAPTURECHANGED | WM_CANCELMODE => {
            sub.pressed = false;
            unsafe {
                let _ = KillTimer(hwnd, STEP_TIMER_INITIAL);
                let _ = KillTimer(hwnd, STEP_TIMER_REPEAT);
                let _ = InvalidateRect(hwnd, None, false);
            }
            return LRESULT(0);
        }
        WM_MOUSEMOVE => {
            if !sub.hovered {
                sub.hovered = true;
                unsafe {
                    let _ = InvalidateRect(hwnd, None, false);
                }
            }
            if !sub.tracking {
                sub.tracking = true;
                let mut tme = TRACKMOUSEEVENT {
                    cbSize: std::mem::size_of::<TRACKMOUSEEVENT>() as u32,
                    dwFlags: TME_LEAVE,
                    hwndTrack: hwnd,
                    dwHoverTime: 0,
                };
                let _ = unsafe { TrackMouseEvent(&mut tme) };
            }
        }
        WM_MOUSELEAVE => {
            sub.tracking = false;
            if sub.hovered {
                sub.hovered = false;
                unsafe {
                    let _ = InvalidateRect(hwnd, None, false);
                }
            }
        }
        WM_NCDESTROY => {
            unsafe {
                let _ = KillTimer(hwnd, STEP_TIMER_INITIAL);
                let _ = KillTimer(hwnd, STEP_TIMER_REPEAT);
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                drop(Box::from_raw(raw));
            }
            return unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) };
        }
        _ => {}
    }

    unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) }
}

fn step_scaling_percent(current: u32, delta: i32) -> u32 {
    (current as i64 + delta as i64).clamp(SCALING_MIN as i64, SCALING_MAX as i64) as u32
}

fn step_volume(current: u32, delta: i32) -> u32 {
    (current as i64 + delta as i64).clamp(VOLUME_MIN as i64, VOLUME_MAX as i64) as u32
}

unsafe fn step_numeric_field(state: &mut State, target_id: i32, delta: i32) {
    let Some(index) = state.selected else {
        return;
    };
    match target_id {
        ID_SCALING_PERCENT => {
            let current = unsafe { text_of(state.hwnd, ID_SCALING_PERCENT) }
                .trim()
                .trim_end_matches('%')
                .trim()
                .parse::<u32>()
                .unwrap_or(100);
            let next = step_scaling_percent(current, delta);
            if let Some(preset) = state.presets.get_mut(index) {
                preset.scaling = CustomScaling::Override(PreviewScale::Percent(next));
            }
            state.loading = true;
            unsafe { set_text(state.hwnd, ID_SCALING_PERCENT, &next.to_string()) };
            state.loading = false;
        }
        ID_VOLUME_VALUE => {
            let current = unsafe { text_of(state.hwnd, ID_VOLUME_VALUE) }
                .trim()
                .trim_end_matches('%')
                .trim()
                .parse::<u32>()
                .unwrap_or_else(|_| state.presets.get(index).map_or(100, |p| p.volume));
            let next = step_volume(current, delta);
            if let Some(preset) = state.presets.get_mut(index) {
                preset.volume = next;
            }
            state.loading = true;
            unsafe {
                set_text(state.hwnd, ID_VOLUME_VALUE, &next.to_string());
                let trackbar = GetDlgItem(state.hwnd, ID_VOLUME).unwrap_or_default();
                SendMessageW(trackbar, TBM_SETPOS, WPARAM(1), LPARAM(next as isize));
            }
            state.loading = false;
        }
        _ => {}
    }
}

unsafe fn child(hwnd: HWND, class: PCWSTR, text: &str, style: WINDOW_STYLE, id: i32) -> HWND {
    let text = wide(text);
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class,
            PCWSTR(text.as_ptr()),
            style,
            0,
            0,
            10,
            10,
            hwnd,
            HMENU(id as isize as *mut _),
            None,
            None,
        )
    }
    .unwrap_or_default()
}

unsafe fn button(hwnd: HWND, text: &str, id: i32) -> HWND {
    unsafe {
        child(
            hwnd,
            w!("BUTTON"),
            text,
            child_style(WS_TABSTOP.0 | BS_OWNERDRAW as u32),
            id,
        )
    }
}

unsafe fn combo_add(combo: HWND, text: &str) {
    let text = wide(text);
    unsafe {
        SendMessageW(
            combo,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(text.as_ptr() as isize),
        );
    }
}

fn combo_drop_height(dpi: i32) -> i32 {
    scale(26, dpi) * 6
}

fn compute_layout_rects(width: i32, height: i32, dpi: i32) -> LayoutRects {
    let tb_h = scale(TITLEBAR_LOGICAL_H, dpi);
    let btn_w = scale(TITLEBAR_BTN_W, dpi);
    let pad = scale(14, dpi);
    let card_gap = scale(12, dpi);
    let footer = scale(52, dpi);
    let sidebar_w = scale(264, dpi);

    let btn_close = RECT {
        left: width - btn_w,
        top: 0,
        right: width,
        bottom: tb_h,
    };
    let btn_max = RECT {
        left: btn_close.left - btn_w,
        top: 0,
        right: btn_close.left,
        bottom: tb_h,
    };
    let btn_min = RECT {
        left: btn_max.left - btn_w,
        top: 0,
        right: btn_max.left,
        bottom: tb_h,
    };

    let content_top = tb_h + pad;
    let content_bottom = (height - footer).max(content_top + scale(120, dpi));

    let sidebar_card = RECT {
        left: pad,
        top: content_top,
        right: pad + sidebar_w,
        bottom: content_bottom,
    };
    let editor_card = RECT {
        left: sidebar_card.right + card_gap,
        top: content_top,
        right: (width - pad).max(sidebar_card.right + card_gap + scale(260, dpi)),
        bottom: content_bottom,
    };

    LayoutRects {
        titlebar: RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: tb_h,
        },
        sidebar_card,
        editor_card,
        btn_min,
        btn_max,
        btn_close,
    }
}

unsafe fn lay_out(state: &mut State) {
    let dpi = state.dpi;
    let mut client = RECT::default();
    let _ = unsafe { GetClientRect(state.hwnd, &mut client) };
    let width = client.right;
    let height = client.bottom;

    state.rects = compute_layout_rects(width, height, dpi);

    let pad = scale(14, dpi);
    let inner = scale(12, dpi);
    let row = scale(30, dpi);
    let gap = scale(10, dpi);
    let label_w = scale(92, dpi);
    let stepper_w = scale(22, dpi);
    let half_row = row / 2;

    let place = |id: i32, x: i32, y: i32, w: i32, h: i32| {
        if let Ok(control) = unsafe { GetDlgItem(state.hwnd, id) } {
            let _ = unsafe { SetWindowPos(control, HWND::default(), x, y, w, h, SWP_NOZORDER) };
        }
    };

    // Sidebar card interior.
    let sc = state.rects.sidebar_card;
    let s_left = sc.left + inner;
    let s_top = sc.top + inner;
    let s_w = (sc.right - sc.left - inner * 2).max(scale(100, dpi));
    let add_w = scale(50, dpi);
    let search_w = (s_w - add_w - scale(8, dpi)).max(scale(60, dpi));
    place(ID_SEARCH, s_left, s_top, search_w, row);
    place(ID_ADD, s_left + search_w + scale(8, dpi), s_top, add_w, row);
    place(
        ID_LIST,
        s_left,
        s_top + row + gap,
        s_w,
        (sc.bottom - inner - (s_top + row + gap)).max(row),
    );

    // Editor card interior.
    let ec = state.rects.editor_card;
    let e_left = ec.left + inner + scale(4, dpi);
    let field_left = e_left + label_w;
    let field_w = (ec.right - inner - scale(4, dpi) - field_left).max(scale(140, dpi));

    let labels = [
        "Name",
        "Extensions",
        "Program",
        "Window",
        "Scaling",
        "Background",
        "Volume",
    ];
    let mut y = ec.top + inner + scale(4, dpi);
    for (index, label) in labels.iter().enumerate() {
        unsafe { set_text(state.hwnd, ID_LABEL_BASE + index as i32, label) };
        place(
            ID_LABEL_BASE + index as i32,
            e_left,
            y + scale(6, dpi),
            label_w - scale(6, dpi),
            row - scale(6, dpi),
        );
        y += row + gap;
    }

    let mut y = ec.top + inner + scale(4, dpi);
    place(ID_NAME, field_left, y, field_w, row);
    y += row + gap;
    place(ID_EXTENSIONS, field_left, y, field_w, row);
    y += row + gap;

    let browse_w = scale(86, dpi);
    place(
        ID_EXE,
        field_left,
        y,
        (field_w - browse_w - scale(8, dpi)).max(scale(80, dpi)),
        row,
    );
    place(ID_BROWSE, field_left + field_w - browse_w, y, browse_w, row);
    y += row + gap;

    place(
        ID_WINDOW_MODE,
        field_left,
        y,
        field_w,
        row + combo_drop_height(dpi),
    );
    y += row + gap;

    // Scaling dropdown + numeric percent edit + stacked ▲/▼ stepper buttons.
    let pct_edit_w = scale(62, dpi);
    let pct_group_w = pct_edit_w + scale(2, dpi) + stepper_w;
    let scaling_combo_w = (field_w - pct_group_w - scale(8, dpi)).max(scale(80, dpi));
    let pct_x = field_left + field_w - pct_group_w;
    place(
        ID_SCALING,
        field_left,
        y,
        scaling_combo_w,
        row + combo_drop_height(dpi),
    );
    place(ID_SCALING_PERCENT, pct_x, y, pct_edit_w, row);
    place(
        ID_SCALING_UP,
        pct_x + pct_edit_w + scale(2, dpi),
        y,
        stepper_w,
        half_row,
    );
    place(
        ID_SCALING_DOWN,
        pct_x + pct_edit_w + scale(2, dpi),
        y + half_row,
        stepper_w,
        row - half_row,
    );
    y += row + gap;

    place(
        ID_BACKGROUND,
        field_left,
        y,
        field_w,
        row + combo_drop_height(dpi),
    );
    y += row + gap;

    // Volume slider + numeric edit + stacked ▲/▼ stepper buttons.
    let vol_edit_w = scale(54, dpi);
    let vol_group_w = vol_edit_w + scale(2, dpi) + stepper_w;
    let vol_slider_w = (field_w - vol_group_w - scale(8, dpi)).max(scale(80, dpi));
    let vol_x = field_left + field_w - vol_group_w;
    place(ID_VOLUME, field_left, y, vol_slider_w, row);
    place(ID_VOLUME_VALUE, vol_x, y, vol_edit_w, row);
    place(
        ID_VOLUME_UP,
        vol_x + vol_edit_w + scale(2, dpi),
        y,
        stepper_w,
        half_row,
    );
    place(
        ID_VOLUME_DOWN,
        vol_x + vol_edit_w + scale(2, dpi),
        y + half_row,
        stepper_w,
        row - half_row,
    );

    // Footer row.
    let btn_w = scale(86, dpi);
    let btn_h = scale(30, dpi);
    let btn_y = height - pad - btn_h + scale(2, dpi);
    place(ID_CLOSE_BUTTON, width - pad - btn_w, btn_y, btn_w, btn_h);
    place(
        ID_DELETE,
        width - pad - btn_w * 2 - scale(8, dpi),
        btn_y,
        btn_w,
        btn_h,
    );
    place(
        ID_SAVE,
        width - pad - btn_w * 3 - scale(16, dpi),
        btn_y,
        btn_w,
        btn_h,
    );
    place(
        ID_STATUS,
        pad + scale(4, dpi),
        btn_y + scale(6, dpi),
        (width - pad * 2 - btn_w * 3 - scale(28, dpi)).max(scale(80, dpi)),
        btn_h - scale(6, dpi),
    );

    unsafe {
        let _ = InvalidateRect(state.hwnd, None, true);
    }
}

unsafe fn command(state: &mut State, id: i32, notice: u32) {
    match id {
        ID_SEARCH => {
            if notice == EN_CHANGE {
                unsafe { rebuild_list(state) };
            }
        }
        ID_LIST => {
            if notice == LBN_SELCHANGE {
                unsafe { select_from_list(state) };
            }
        }
        ID_ADD => unsafe { add_preset(state) },
        ID_NAME | ID_EXTENSIONS | ID_EXE => {
            if notice == EN_CHANGE {
                unsafe { fields_changed(state) };
            }
        }
        ID_WINDOW_MODE => {
            if notice == windows::Win32::UI::WindowsAndMessaging::CBN_SELCHANGE {
                unsafe { window_mode_chosen(state) };
            }
        }
        ID_SCALING => {
            if notice == windows::Win32::UI::WindowsAndMessaging::CBN_SELCHANGE {
                unsafe { scaling_chosen(state) };
            }
        }
        ID_SCALING_PERCENT => {
            if notice == EN_CHANGE {
                unsafe { percent_typed(state) };
            }
        }
        ID_BACKGROUND => {
            if notice == windows::Win32::UI::WindowsAndMessaging::CBN_SELCHANGE {
                unsafe { background_chosen(state) };
            }
        }
        ID_VOLUME_VALUE => {
            if notice == EN_CHANGE {
                unsafe { volume_typed(state) };
            }
        }
        ID_BROWSE => unsafe { browse(state) },
        ID_SAVE => {
            let answer = persist(state);
            let line = answer
                .as_ref()
                .err()
                .map(String::as_str)
                .unwrap_or("Saved.");
            unsafe { set_status(state, line) };
        }
        ID_DELETE => unsafe { delete_selected(state) },
        ID_CLOSE_BUTTON => {
            let _ = persist(state);
            unsafe {
                let _ = DestroyWindow(state.hwnd);
            }
        }
        _ => {}
    }
}

unsafe fn set_scaling_percent_enabled(hwnd: HWND, enabled: bool) {
    for id in [ID_SCALING_PERCENT, ID_SCALING_UP, ID_SCALING_DOWN] {
        if let Ok(ctrl) = unsafe { GetDlgItem(hwnd, id) } {
            let _ = unsafe { EnableWindow(ctrl, enabled) };
        }
    }
}

unsafe fn fill_fields(state: &mut State) {
    let Some(preset) = state.selected.and_then(|index| state.presets.get(index)) else {
        state.loading = true;
        for id in [
            ID_NAME,
            ID_EXTENSIONS,
            ID_EXE,
            ID_SCALING_PERCENT,
            ID_VOLUME_VALUE,
        ] {
            unsafe { set_text(state.hwnd, id, "") };
        }
        for id in [ID_WINDOW_MODE, ID_SCALING, ID_BACKGROUND] {
            let combo = unsafe { GetDlgItem(state.hwnd, id) }.unwrap_or_default();
            unsafe { SendMessageW(combo, CB_SETCURSEL, WPARAM(usize::MAX), LPARAM(0)) };
        }
        state.loading = false;
        unsafe { set_fields_enabled(state, false) };
        return;
    };
    let name = preset.name.clone();
    let extensions = preset.extensions_text();
    let exe = preset.exe.clone();
    let window_mode = preset.window_mode;
    let scaling = preset.scaling;
    let background = preset.background;
    let volume = preset.volume;

    state.loading = true;
    unsafe {
        set_text(state.hwnd, ID_NAME, &name);
        set_text(state.hwnd, ID_EXTENSIONS, &extensions);
        set_text(state.hwnd, ID_EXE, &exe);

        let mode_combo = GetDlgItem(state.hwnd, ID_WINDOW_MODE).unwrap_or_default();
        let mode_index = WINDOW_MODES
            .iter()
            .position(|held| *held == window_mode)
            .unwrap_or(0) as i32;
        SendMessageW(
            mode_combo,
            CB_SETCURSEL,
            WPARAM(mode_index as usize),
            LPARAM(0),
        );

        let combo = GetDlgItem(state.hwnd, ID_SCALING).unwrap_or_default();
        let choice = SCALE_CHOICES
            .iter()
            .position(|(_, value)| *value == scaling)
            .unwrap_or(SCALE_CHOICES.len()) as i32;
        SendMessageW(combo, CB_SETCURSEL, WPARAM(choice as usize), LPARAM(0));

        let (percent_text, percent_enabled) = match scaling {
            CustomScaling::FollowTray => (String::new(), false),
            CustomScaling::Override(PreviewScale::FitToScreen) => ("100".to_string(), false),
            CustomScaling::Override(
                PreviewScale::FitToScreenReduced(percent) | PreviewScale::Percent(percent),
            ) => (percent.to_string(), true),
        };
        set_text(state.hwnd, ID_SCALING_PERCENT, &percent_text);
        set_scaling_percent_enabled(state.hwnd, percent_enabled);

        let background_combo = GetDlgItem(state.hwnd, ID_BACKGROUND).unwrap_or_default();
        let index = BACKGROUNDS
            .iter()
            .position(|held| *held == background)
            .unwrap_or(1) as i32;
        SendMessageW(
            background_combo,
            CB_SETCURSEL,
            WPARAM(index as usize),
            LPARAM(0),
        );

        let trackbar = GetDlgItem(state.hwnd, ID_VOLUME).unwrap_or_default();
        SendMessageW(trackbar, TBM_SETPOS, WPARAM(1), LPARAM(volume as isize));
        set_text(state.hwnd, ID_VOLUME_VALUE, &volume.to_string());
    }
    state.loading = false;
    unsafe { set_fields_enabled(state, true) };
}

unsafe fn set_fields_enabled(state: &State, enabled: bool) {
    for id in [
        ID_NAME,
        ID_EXTENSIONS,
        ID_EXE,
        ID_BROWSE,
        ID_WINDOW_MODE,
        ID_SCALING,
        ID_BACKGROUND,
        ID_VOLUME,
        ID_VOLUME_VALUE,
        ID_VOLUME_UP,
        ID_VOLUME_DOWN,
        ID_DELETE,
    ] {
        if let Ok(control) = unsafe { GetDlgItem(state.hwnd, id) } {
            let _ = unsafe { EnableWindow(control, enabled) };
        }
    }
    if !enabled {
        unsafe { set_scaling_percent_enabled(state.hwnd, false) };
    }
}

fn shown_presets(presets: &[CustomPreset], search: &str) -> Vec<usize> {
    let needle = search.trim().to_lowercase();
    presets
        .iter()
        .enumerate()
        .filter(|(_, preset)| needle.is_empty() || preset.name.to_lowercase().contains(&needle))
        .map(|(index, _)| index)
        .collect()
}

unsafe fn rebuild_list(state: &mut State) {
    let search = unsafe { text_of(state.hwnd, ID_SEARCH) };
    state.shown = shown_presets(&state.presets, &search);

    let list = unsafe { GetDlgItem(state.hwnd, ID_LIST) }.unwrap_or_default();
    unsafe {
        SendMessageW(list, LB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for index in &state.shown {
            let name = wide(&state.presets[*index].name);
            SendMessageW(
                list,
                LB_ADDSTRING,
                WPARAM(0),
                LPARAM(name.as_ptr() as isize),
            );
        }
        match state
            .selected
            .and_then(|selected| state.shown.iter().position(|held| *held == selected))
        {
            Some(position) => {
                SendMessageW(list, LB_SETCURSEL, WPARAM(position), LPARAM(0));
            }
            None => {
                SendMessageW(list, LB_SETCURSEL, WPARAM(usize::MAX), LPARAM(0));
            }
        }
    }
}

unsafe fn select_from_list(state: &mut State) {
    let list = unsafe { GetDlgItem(state.hwnd, ID_LIST) }.unwrap_or_default();
    let choice = unsafe { SendMessageW(list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if choice < 0 {
        return;
    }
    let Some(index) = state.shown.get(choice as usize).copied() else {
        return;
    };
    state.selected = Some(index);
    unsafe { fill_fields(state) };
    unsafe { set_status(state, "") };
}

unsafe fn add_preset(state: &mut State) {
    if state.presets.len() >= MAX_PRESETS {
        unsafe {
            set_status(
                state,
                &format!("{MAX_PRESETS} presets is the most this window holds."),
            )
        };
        return;
    }

    let mut number = state.presets.len() + 1;
    let mut name = format!("New Preset {number}");
    while state
        .presets
        .iter()
        .any(|preset| preset.name.eq_ignore_ascii_case(&name))
    {
        number += 1;
        name = format!("New Preset {number}");
    }

    state.presets.push(CustomPreset::named(&name));
    state.selected = Some(state.presets.len() - 1);
    unsafe { set_text(state.hwnd, ID_SEARCH, "") };
    unsafe { rebuild_list(state) };
    unsafe { fill_fields(state) };
    if let Ok(edit) = unsafe { GetDlgItem(state.hwnd, ID_NAME) } {
        unsafe {
            let _ = SetForegroundWindow(state.hwnd);
            let _ = SetFocus(edit);
            SendMessageW(edit, EM_SETSEL, WPARAM(0), LPARAM(-1));
        }
    }
    unsafe { set_status(state, "") };
}

unsafe fn fields_changed(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };

    let name = unsafe { text_of(state.hwnd, ID_NAME) };
    let extensions = unsafe { text_of(state.hwnd, ID_EXTENSIONS) };
    let exe = unsafe { text_of(state.hwnd, ID_EXE) };

    let Some(preset) = state.presets.get_mut(index) else {
        return;
    };
    preset.name = name;
    preset.exe = exe;
    preset.set_extensions_text(&extensions);

    unsafe { rebuild_list(state) };
}

unsafe fn window_mode_chosen(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    let combo = unsafe { GetDlgItem(state.hwnd, ID_WINDOW_MODE) }.unwrap_or_default();
    let choice = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    let Some(mode) = usize::try_from(choice)
        .ok()
        .and_then(|choice| WINDOW_MODES.get(choice).copied())
    else {
        return;
    };
    if let Some(preset) = state.presets.get_mut(index) {
        preset.window_mode = mode;
    }
}

unsafe fn scaling_chosen(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    let combo = unsafe { GetDlgItem(state.hwnd, ID_SCALING) }.unwrap_or_default();
    let choice = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if choice < 0 {
        return;
    }

    let Some((_, chosen)) = SCALE_CHOICES.get(choice as usize).copied() else {
        unsafe { set_scaling_percent_enabled(state.hwnd, true) };
        if let Ok(edit) = unsafe { GetDlgItem(state.hwnd, ID_SCALING_PERCENT) } {
            let _ = unsafe { SetFocus(edit) };
        }
        return;
    };

    if let Some(preset) = state.presets.get_mut(index) {
        preset.scaling = chosen;
    }

    let (percent_text, percent_enabled) = match chosen {
        CustomScaling::FollowTray => (String::new(), false),
        CustomScaling::Override(PreviewScale::FitToScreen) => ("100".to_string(), false),
        CustomScaling::Override(
            PreviewScale::FitToScreenReduced(percent) | PreviewScale::Percent(percent),
        ) => (percent.to_string(), true),
    };
    state.loading = true;
    unsafe { set_text(state.hwnd, ID_SCALING_PERCENT, &percent_text) };
    state.loading = false;
    unsafe { set_scaling_percent_enabled(state.hwnd, percent_enabled) };
}

unsafe fn percent_typed(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    let text = unsafe { text_of(state.hwnd, ID_SCALING_PERCENT) };
    let Ok(percent) = text.trim().trim_end_matches('%').trim().parse::<u32>() else {
        return;
    };
    if let Some(preset) = state.presets.get_mut(index) {
        preset.scaling = CustomScaling::Override(PreviewScale::Percent(
            percent.clamp(SCALING_MIN, SCALING_MAX),
        ));
    }
}

unsafe fn background_chosen(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    let combo = unsafe { GetDlgItem(state.hwnd, ID_BACKGROUND) }.unwrap_or_default();
    let choice = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    let Some(background) = usize::try_from(choice)
        .ok()
        .and_then(|choice| BACKGROUNDS.get(choice).copied())
    else {
        return;
    };
    if let Some(preset) = state.presets.get_mut(index) {
        preset.background = background;
    }
}

unsafe fn volume_moved(state: &mut State) {
    let Some(index) = state.selected else {
        return;
    };
    let trackbar = unsafe { GetDlgItem(state.hwnd, ID_VOLUME) }.unwrap_or_default();
    let position = unsafe { SendMessageW(trackbar, TBM_GETPOS, WPARAM(0), LPARAM(0)).0 }
        .clamp(VOLUME_MIN as isize, VOLUME_MAX as isize) as u32;

    state.loading = true;
    unsafe { set_text(state.hwnd, ID_VOLUME_VALUE, &position.to_string()) };
    state.loading = false;

    if let Some(preset) = state.presets.get_mut(index) {
        preset.volume = position;
    }
}

unsafe fn volume_typed(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    let text = unsafe { text_of(state.hwnd, ID_VOLUME_VALUE) };
    let Ok(volume) = text.trim().trim_end_matches('%').trim().parse::<u32>() else {
        return;
    };
    let volume = volume.min(VOLUME_MAX);
    if let Some(preset) = state.presets.get_mut(index) {
        preset.volume = volume;
    }
    let trackbar = unsafe { GetDlgItem(state.hwnd, ID_VOLUME) }.unwrap_or_default();
    unsafe {
        SendMessageW(trackbar, TBM_SETPOS, WPARAM(1), LPARAM(volume as isize));
    }
}

unsafe fn delete_selected(state: &mut State) {
    let Some(index) = state.selected else {
        return;
    };
    if index >= state.presets.len() {
        return;
    }
    state.presets.remove(index);
    state.selected = if state.presets.is_empty() {
        None
    } else {
        Some(index.min(state.presets.len() - 1))
    };
    unsafe { rebuild_list(state) };
    unsafe { fill_fields(state) };
    match persist(state) {
        Ok(()) => unsafe { set_status(state, "Deleted.") },
        Err(reason) => unsafe { set_status(state, &reason) },
    }
}

unsafe fn browse(state: &mut State) {
    let Some(index) = state.selected else {
        return;
    };
    let mut buffer = vec![0u16; 32 * 1024];
    if let Some(exe) = state.presets.get(index).map(|preset| preset.exe.clone()) {
        for (slot, unit) in buffer.iter_mut().zip(exe.encode_utf16()) {
            *slot = unit;
        }
    }

    let filter: Vec<u16> = "Programs (*.exe)\0*.exe\0All files (*.*)\0*.*\0\0"
        .encode_utf16()
        .collect();
    let title = wide("Choose the program a custom preview runs");

    let mut dialog = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: state.hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: windows::core::PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        lpstrTitle: PCWSTR(title.as_ptr()),
        Flags: OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        ..Default::default()
    };

    if unsafe { GetOpenFileNameW(&mut dialog) }.as_bool() {
        let length = buffer.iter().position(|unit| *unit == 0).unwrap_or(0);
        let chosen = String::from_utf16_lossy(&buffer[..length]);
        if let Some(preset) = state.presets.get_mut(index) {
            preset.exe = chosen.clone();
        }
        state.loading = true;
        unsafe { set_text(state.hwnd, ID_EXE, &chosen) };
        state.loading = false;
    }
}

unsafe fn set_text(hwnd: HWND, id: i32, text: &str) {
    if let Ok(control) = unsafe { GetDlgItem(hwnd, id) } {
        let text = wide(text);
        let _ = unsafe { SetWindowTextW(control, PCWSTR(text.as_ptr())) };
    }
}

unsafe fn text_of(hwnd: HWND, id: i32) -> String {
    let Ok(control) = (unsafe { GetDlgItem(hwnd, id) }) else {
        return String::new();
    };
    let length = unsafe { GetWindowTextLengthW(control) }.max(0) as usize;
    let mut buffer = vec![0u16; length + 1];
    let read = unsafe { GetWindowTextW(control, &mut buffer) }.max(0) as usize;
    String::from_utf16_lossy(&buffer[..read])
}

unsafe fn set_status(state: &State, text: &str) {
    unsafe { set_text(state.hwnd, ID_STATUS, text) };
}

fn is_stepper(id: u32) -> bool {
    id == ID_SCALING_UP as u32
        || id == ID_SCALING_DOWN as u32
        || id == ID_VOLUME_UP as u32
        || id == ID_VOLUME_DOWN as u32
}

fn is_button(id: u32) -> bool {
    id == ID_ADD as u32
        || id == ID_SAVE as u32
        || id == ID_DELETE as u32
        || id == ID_BROWSE as u32
        || id == ID_CLOSE_BUTTON as u32
        || is_stepper(id)
}

unsafe fn draw_item(state: &State, item: &DRAWITEMSTRUCT) -> usize {
    let hdc = item.hDC;
    let p = state.palette;
    let selected = item.itemState.0 & ODS_SELECTED.0 != 0;
    let disabled = item.itemState.0 & ODS_DISABLED.0 != 0;
    let dpi = state.dpi;

    let (fill, border_col, text_col) = if is_stepper(item.CtlID) {
        let raw =
            unsafe { GetWindowLongPtrW(item.hwndItem, GWLP_USERDATA) } as *const StepperSubclass;
        let (hov, prs) = unsafe { raw.as_ref() }.map_or((false, false), |s| (s.hovered, s.pressed));
        let fill = if disabled {
            p.card
        } else if prs || selected {
            p.accent
        } else if hov {
            p.button_hover
        } else {
            p.button_bg
        };
        let fg = if disabled {
            p.muted
        } else if prs || selected {
            p.accent_text
        } else {
            p.ink
        };
        (fill, p.border, fg)
    } else if item.CtlID == ID_SAVE as u32 {
        let fill = if disabled {
            p.button_bg
        } else if selected {
            mix(p.accent, p.page, 20)
        } else {
            p.accent
        };
        (
            fill,
            if disabled { p.border } else { p.accent },
            if disabled { p.muted } else { p.accent_text },
        )
    } else if item.CtlID == ID_DELETE as u32 {
        let fill = if disabled {
            p.button_bg
        } else if selected {
            p.danger
        } else {
            p.danger_bg
        };
        let fg = if disabled {
            p.muted
        } else if selected {
            rgb(255, 255, 255)
        } else {
            p.danger
        };
        (fill, if disabled { p.border } else { p.danger }, fg)
    } else if is_button(item.CtlID) {
        let fill = if disabled {
            p.card
        } else if selected {
            p.button_press
        } else {
            p.button_bg
        };
        (fill, p.border, if disabled { p.muted } else { p.ink })
    } else if item.CtlID == ID_LIST as u32 && selected {
        (mix(p.field, p.accent, 22), p.accent, p.ink)
    } else {
        (p.field, p.border, if disabled { p.muted } else { p.ink })
    };

    unsafe {
        if is_button(item.CtlID) {
            let parent_brush = if item.CtlID == ID_SAVE as u32
                || item.CtlID == ID_DELETE as u32
                || item.CtlID == ID_CLOSE_BUTTON as u32
            {
                state.brush_page
            } else {
                state.brush_card
            };
            FillRect(hdc, &item.rcItem, parent_brush);

            let r = item.rcItem;
            let radius = if is_stepper(item.CtlID) {
                scale(3, dpi)
            } else {
                scale(6, dpi)
            };
            let brush = CreateSolidBrush(fill);
            let pen = CreatePen(PS_SOLID, 1, border_col);
            let old_b = SelectObject(hdc, HGDIOBJ(brush.0));
            let old_p = SelectObject(hdc, HGDIOBJ(pen.0));
            let _ = RoundRect(hdc, r.left, r.top, r.right, r.bottom, radius, radius);
            let _ = SelectObject(hdc, old_b);
            let _ = SelectObject(hdc, old_p);
            let _ = DeleteObject(HGDIOBJ(brush.0));
            let _ = DeleteObject(HGDIOBJ(pen.0));
        } else {
            let brush = CreateSolidBrush(fill);
            FillRect(hdc, &item.rcItem, brush);
            let _ = DeleteObject(HGDIOBJ(brush.0));
            if item.CtlID == ID_LIST as u32 && selected {
                // Left accent indicator bar on selected preset tab.
                let bar = RECT {
                    left: item.rcItem.left,
                    top: item.rcItem.top + scale(3, dpi),
                    right: item.rcItem.left + scale(3, dpi),
                    bottom: item.rcItem.bottom - scale(3, dpi),
                };
                let bar_brush = CreateSolidBrush(p.accent);
                FillRect(hdc, &bar, bar_brush);
                let _ = DeleteObject(HGDIOBJ(bar_brush.0));
            }
        }

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, text_col);
        let font = if is_stepper(item.CtlID) {
            state.font_small
        } else if item.CtlID == ID_SAVE as u32 {
            state.font_bold
        } else {
            state.font
        };
        let _ = SelectObject(hdc, HGDIOBJ(font.0));

        let text = match item.CtlID {
            id if id == ID_LIST as u32 => {
                let list = item.hwndItem;
                let length =
                    SendMessageW(list, LB_GETTEXTLEN, WPARAM(item.itemID as usize), LPARAM(0))
                        .0
                        .max(0) as usize;
                let mut buffer = vec![0u16; length + 1];
                SendMessageW(
                    list,
                    LB_GETTEXT,
                    WPARAM(item.itemID as usize),
                    LPARAM(buffer.as_mut_ptr() as isize),
                );
                String::from_utf16_lossy(&buffer[..length])
            }
            id if id == ID_WINDOW_MODE as u32
                || id == ID_SCALING as u32
                || id == ID_BACKGROUND as u32 =>
            {
                let combo = item.hwndItem;
                let index = if item.itemID == u32::MAX {
                    SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0
                } else {
                    item.itemID as isize
                };
                if index < 0 {
                    String::new()
                } else {
                    let length =
                        SendMessageW(combo, CB_GETLBTEXTLEN, WPARAM(index as usize), LPARAM(0))
                            .0
                            .max(0) as usize;
                    let mut buffer = vec![0u16; length + 1];
                    SendMessageW(
                        combo,
                        CB_GETLBTEXT,
                        WPARAM(index as usize),
                        LPARAM(buffer.as_mut_ptr() as isize),
                    );
                    String::from_utf16_lossy(&buffer[..length])
                }
            }
            _ => {
                let control = item.hwndItem;
                let length = GetWindowTextLengthW(control).max(0) as usize;
                let mut buffer = vec![0u16; length + 1];
                let read = GetWindowTextW(control, &mut buffer).max(0) as usize;
                String::from_utf16_lossy(&buffer[..read])
            }
        };

        // Empty slice would hand DrawTextW a dangling alignment pointer; return early.
        if text.is_empty() {
            return 1;
        }

        let mut text: Vec<u16> = text.encode_utf16().collect();
        let mut rect = item.rcItem;
        let flags = if is_button(item.CtlID) {
            DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
        } else {
            rect.left += scale(10, dpi);
            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
        };
        DrawTextW(hdc, &mut text, &mut rect, flags);
    }

    1
}

fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::Graphics::Gdi::{GetDC, ReleaseDC};
    use windows::Win32::UI::Controls::{ODA_DRAWENTIRE, ODS_FLAGS, ODT_COMBOBOX};
    use windows::Win32::UI::Input::KeyboardAndMouse::IsWindowEnabled;
    use windows::Win32::UI::WindowsAndMessaging::{
        GetWindowRect, CB_GETDROPPEDCONTROLRECT, CB_SHOWDROPDOWN,
    };

    #[test]
    fn drawing_a_combo_with_nothing_selected_does_not_take_the_process_down() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };

        unsafe {
            let combo = GetDlgItem(hwnd, ID_SCALING).unwrap_or_default();
            SendMessageW(combo, CB_SETCURSEL, WPARAM(usize::MAX), LPARAM(0));

            let hdc = GetDC(hwnd);
            let mut item = DRAWITEMSTRUCT {
                CtlType: ODT_COMBOBOX,
                CtlID: ID_SCALING as u32,
                itemID: u32::MAX,
                itemAction: ODA_DRAWENTIRE,
                itemState: ODS_FLAGS(0),
                hwndItem: combo,
                hDC: hdc,
                rcItem: RECT {
                    left: 0,
                    top: 0,
                    right: 160,
                    bottom: 24,
                },
                itemData: 0,
            };
            let drawn = SendMessageW(
                hwnd,
                WM_DRAWITEM,
                WPARAM(ID_SCALING as usize),
                LPARAM(&mut item as *mut _ as isize),
            );
            assert_eq!(drawn.0, 1, "the empty combo was not drawn");

            ReleaseDC(hwnd, hdc);
            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn a_combo_is_laid_out_with_a_list_to_drop() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };

        unsafe {
            let combo = GetDlgItem(hwnd, ID_SCALING).unwrap_or_default();
            let mut field = RECT::default();
            let _ = GetWindowRect(combo, &mut field);
            SendMessageW(combo, CB_SHOWDROPDOWN, WPARAM(1), LPARAM(0));
            let mut dropped = RECT::default();
            SendMessageW(
                combo,
                CB_GETDROPPEDCONTROLRECT,
                WPARAM(0),
                LPARAM(&mut dropped as *mut _ as isize),
            );
            SendMessageW(combo, CB_SHOWDROPDOWN, WPARAM(0), LPARAM(0));

            let field_height = field.bottom - field.top;
            let dropped_height = dropped.bottom - dropped.top;
            assert!(
                dropped_height >= field_height * 3,
                "the list drops {dropped_height}px high against a {field_height}px field"
            );

            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn with_no_preset_the_fields_are_off() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };

        unsafe {
            let Some(state) = state_of(hwnd) else {
                panic!("the window kept no state");
            };
            if !state.presets.is_empty() {
                let _ = DestroyWindow(hwnd);
                return;
            }

            for id in [
                ID_NAME,
                ID_EXTENSIONS,
                ID_EXE,
                ID_WINDOW_MODE,
                ID_SCALING,
                ID_SCALING_UP,
                ID_SCALING_DOWN,
                ID_BACKGROUND,
                ID_VOLUME_UP,
                ID_VOLUME_DOWN,
                ID_DELETE,
            ] {
                let control = GetDlgItem(hwnd, id).unwrap_or_default();
                assert!(
                    !IsWindowEnabled(control).as_bool(),
                    "control {id} is editable with no preset to edit"
                );
            }

            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn palette_matches_one_dark_pro_and_atom_one_light_tokens() {
        let dark = Palette::of(TextTheme::Dark);
        assert!(dark.dark);
        assert_eq!(dark.page, rgb(0x28, 0x2C, 0x34));
        assert_eq!(dark.ink, rgb(0xAB, 0xB2, 0xBF));
        assert_eq!(dark.accent, rgb(0x56, 0xB6, 0xC2));

        let light = Palette::of(TextTheme::Light);
        assert!(!light.dark);
        assert_eq!(light.page, rgb(0xFA, 0xFA, 0xFA));
        assert_eq!(light.ink, rgb(0x38, 0x3A, 0x42));
        assert_eq!(light.accent, rgb(0x01, 0x84, 0xBC));
    }

    #[test]
    fn numeric_stepping_clamps_to_valid_bounds() {
        assert_eq!(step_scaling_percent(100, 1), 101);
        assert_eq!(step_scaling_percent(1, -1), 1);
        assert_eq!(step_scaling_percent(1000, 1), 1000);

        assert_eq!(step_volume(50, 1), 51);
        assert_eq!(step_volume(0, -1), 0);
        assert_eq!(step_volume(100, 1), 100);
    }

    #[test]
    fn ctrl_backspace_deletes_previous_word_or_selection() {
        let u: Vec<u16> = "foo bar   baz".encode_utf16().collect();
        assert_eq!(prev_word_delete_range_utf16(&u, u.len(), u.len()), (10, 13));
        assert_eq!(prev_word_delete_range_utf16(&u, 10, 10), (4, 10));
        assert_eq!(prev_word_delete_range_utf16(&u, 2, 7), (2, 7));
        assert_eq!(prev_word_delete_range_utf16(&u, 0, 0), (0, 0));
    }

    #[test]
    fn titlebar_button_hit_test_matches_right_edge_buttons() {
        let rects = compute_layout_rects(920, 620, 96);
        assert_eq!(
            rects.titlebar_button_at(rects.btn_close.left + 5, 10),
            Some(TitlebarButton::Close)
        );
        assert_eq!(
            rects.titlebar_button_at(rects.btn_max.left + 5, 10),
            Some(TitlebarButton::Maximize)
        );
        assert_eq!(
            rects.titlebar_button_at(rects.btn_min.left + 5, 10),
            Some(TitlebarButton::Minimize)
        );
        assert_eq!(rects.titlebar_button_at(20, 10), None);
    }
}

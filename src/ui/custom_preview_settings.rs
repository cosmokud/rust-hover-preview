//! The `Custom Preview` window: frameless audio-card-styled editor for custom preview presets.
//!
//! Runs on its own thread (opened from tray `WM_COMMAND`), single instance via `OPEN`/`WINDOW`.
//! Uses native Win32 child controls (`EDIT`, `LISTBOX`, `COMBOBOX`, `msctls_trackbar32`,
//! owner-drawn `BUTTON`) painted with the active `TextTheme` palette and live theme switching.

use std::sync::atomic::{AtomicBool, AtomicIsize, Ordering};
use std::sync::Once;

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Dwm::{
    DwmExtendFrameIntoClientArea, DwmSetWindowAttribute, DWMWA_USE_IMMERSIVE_DARK_MODE,
    DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_ROUND,
};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, BitBlt, CreateCompatibleBitmap, CreateCompatibleDC, CreateDIBSection, CreateFontW,
    CreatePen, CreateSolidBrush, DeleteDC, DeleteObject, DrawTextW, EndPaint, FillRect,
    GdiAlphaBlend, GetDC, GetStockObject, GetTextMetricsW, InvalidateRect, LineTo, MoveToEx,
    RedrawWindow, ReleaseDC, RoundRect, ScreenToClient, SelectObject, SetBkColor, SetBkMode,
    SetBrushOrgEx, SetStretchBltMode, SetTextColor, SetViewportOrgEx, StretchBlt, UpdateWindow,
    AC_SRC_ALPHA, AC_SRC_OVER, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, BLENDFUNCTION,
    CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET, DEFAULT_PITCH, DIB_RGB_COLORS, DT_CENTER,
    DT_END_ELLIPSIS, DT_LEFT, DT_NOPREFIX, DT_SINGLELINE, DT_VCENTER, FW_NORMAL, FW_SEMIBOLD,
    HALFTONE, HBITMAP, HBRUSH, HDC, HFONT, HGDIOBJ, NULL_BRUSH, OUT_DEFAULT_PRECIS, PAINTSTRUCT,
    PROOF_QUALITY, PS_SOLID, RDW_ALLCHILDREN, RDW_ERASE, RDW_INVALIDATE, RDW_UPDATENOW, SRCCOPY,
    TEXTMETRICW, TRANSPARENT,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{
    GetComboBoxInfo, SetWindowTheme, COMBOBOXINFO, DRAWITEMSTRUCT, EM_GETSEL, EM_REPLACESEL,
    EM_SETLIMITTEXT, EM_SETSEL, MARGINS, MEASUREITEMSTRUCT, ODS_DISABLED, ODS_SELECTED,
};
use windows::Win32::UI::HiDpi::{GetDpiForWindow, GetSystemMetricsForDpi};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetDoubleClickTime, GetFocus, GetKeyState, IsWindowEnabled, ReleaseCapture,
    SetCapture, SetFocus, TrackMouseEvent, TME_LEAVE, TRACKMOUSEEVENT, VK_A, VK_BACK, VK_CONTROL,
    VK_DOWN, VK_END, VK_ESCAPE, VK_HOME, VK_LEFT, VK_MENU, VK_RETURN, VK_RIGHT, VK_SHIFT, VK_SPACE,
    VK_TAB, VK_UP,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BeginDeferWindowPos, CallWindowProcW, CreateWindowExW, DefWindowProcW, DeferWindowPos,
    DestroyWindow, DispatchMessageW, EndDeferWindowPos, GetClientRect, GetDlgItem, GetMessageTime,
    GetMessageW, GetNextDlgTabItem, GetSystemMetrics, GetWindowLongPtrW, GetWindowTextLengthW,
    GetWindowTextW, IsWindow, IsZoomed, KillTimer, LoadCursorW, PostMessageW, PostQuitMessage,
    RegisterClassExW, SendMessageW, SetCursor, SetForegroundWindow, SetTimer, SetWindowLongPtrW,
    SetWindowPos, SetWindowTextW, ShowWindow, TranslateMessage, BS_OWNERDRAW, CBN_SELCHANGE,
    CBN_SELENDOK, CBS_DROPDOWNLIST, CBS_HASSTRINGS, CBS_OWNERDRAWFIXED, CB_ADDSTRING, CB_GETCURSEL,
    CB_GETDROPPEDSTATE, CB_GETLBTEXT, CB_GETLBTEXTLEN, CB_SETCURSEL, CW_USEDEFAULT, EN_CHANGE,
    ES_AUTOHSCROLL, GWLP_USERDATA, GWLP_WNDPROC, HMENU, HTBOTTOM, HTBOTTOMLEFT, HTBOTTOMRIGHT,
    HTCAPTION, HTCLIENT, HTLEFT, HTRIGHT, HTTOP, HTTOPLEFT, HTTOPRIGHT, IDC_ARROW, IDC_HAND,
    IDC_SIZEALL, LBN_SELCHANGE, LBS_HASSTRINGS, LBS_NOINTEGRALHEIGHT, LBS_NOTIFY,
    LBS_OWNERDRAWFIXED, LB_ADDSTRING, LB_GETCOUNT, LB_GETCURSEL, LB_GETITEMRECT, LB_GETTEXT,
    LB_GETTEXTLEN, LB_ITEMFROMPOINT, LB_RESETCONTENT, LB_SETCURSEL, MINMAXINFO, MSG,
    NCCALCSIZE_PARAMS, SM_CXDOUBLECLK, SM_CXPADDEDBORDER, SM_CXSIZEFRAME, SM_CYDOUBLECLK,
    SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOCOPYBITS, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER,
    SW_MAXIMIZE, SW_MINIMIZE, SW_RESTORE, SW_SHOW, SW_SHOWNORMAL, WA_INACTIVE, WINDOW_EX_STYLE,
    WINDOW_STYLE, WM_ACTIVATE, WM_APP, WM_CANCELMODE, WM_CAPTURECHANGED, WM_CHAR, WM_CLOSE,
    WM_COMMAND, WM_CREATE, WM_CTLCOLORBTN, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX, WM_CTLCOLORSTATIC,
    WM_DESTROY, WM_DRAWITEM, WM_ENABLE, WM_ERASEBKGND, WM_GETMINMAXINFO, WM_HSCROLL, WM_KEYDOWN,
    WM_KILLFOCUS, WM_LBUTTONDBLCLK, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MEASUREITEM, WM_MOUSEMOVE,
    WM_MOUSEWHEEL, WM_NCCALCSIZE, WM_NCDESTROY, WM_NCHITTEST, WM_PAINT, WM_PRINTCLIENT,
    WM_SETCURSOR, WM_SETFOCUS, WM_SETFONT, WM_SIZE, WM_TIMER, WNDCLASSEXW, WNDCLASS_STYLES,
    WNDPROC, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN, WS_CLIPSIBLINGS, WS_MAXIMIZEBOX,
    WS_MINIMIZEBOX, WS_OVERLAPPED, WS_SYSMENU, WS_TABSTOP, WS_THICKFRAME, WS_VISIBLE, WS_VSCROLL,
};

use crate::config::config::{
    sanitize_program_input, CustomCaptureTechnique, CustomPreset, CustomPromotionTechnique,
    CustomScaling, CustomWindowMode, PreviewScale, TextTheme, MAX_PRESETS,
};
use crate::text::{text_paint, text_theme};
use crate::CONFIG;

const CLASS_NAME: PCWSTR = w!("RustHoverPreviewCustomPreview");
const SLIDER_CLASS: PCWSTR = w!("RustHoverPreviewThemedSlider");
const TOGGLE_CLASS: PCWSTR = w!("RustHoverPreviewThemedToggle");

static OPEN: AtomicBool = AtomicBool::new(false);
static WINDOW: AtomicIsize = AtomicIsize::new(0);
static SLIDER_CLASS_REGISTERED: Once = Once::new();
static TOGGLE_CLASS_REGISTERED: Once = Once::new();

const WM_MOUSELEAVE: u32 = 0x02A3;
const PRF_CLIENT: isize = 0x0004;

const WM_BRING_FORWARD: u32 = WM_APP + 17;
const WM_THEME_CHANGED: u32 = WM_APP + 18;
const WM_STEP_NUMERIC: u32 = WM_APP + 19;
const WM_SLIDER_SET_VALUE: u32 = WM_APP + 20;
const WM_SLIDER_GET_VALUE: u32 = WM_APP + 21;
const WM_TOGGLE_SET_VALUE: u32 = WM_APP + 22;
const WM_TOGGLE_GET_VALUE: u32 = WM_APP + 23;
const WM_TOGGLE_CHANGED: u32 = WM_APP + 24;
/// One step of a preset drag, posted by the list: `wparam` is the drawn row the pointer is over and
/// `lparam` the drawn row the press landed on (read only on the first step of a drag).
///
/// Posted rather than handled in the list, because the reorder rebuilds the box — and doing that
/// from inside the box's own `WM_MOUSEMOVE` is re-entrant mutation of its drag state: an access
/// violation, which ends the process and the tray with it, not a panic.
const WM_PRESET_DRAG_MOVE: u32 = WM_APP + 25;
/// The end of that drag, on the drop or on a capture the list lost some other way.
const WM_PRESET_DRAG_END: u32 = WM_APP + 26;
/// A field has been left, and what it holds is what the user meant: posted by a control when the
/// focus or the pointer leaves it, so the save runs on the parent's own message rather than inside
/// the control's (`WM_PRESET_DRAG_MOVE` says why that matters).
const WM_AUTOSAVE: u32 = WM_APP + 27;
// Posted by list_subclass_proc on row pill click; wparam = drawn row index.
const WM_PRESET_ROW_TOGGLE: u32 = WM_APP + 28;

const TITLEBAR_LOGICAL_H: i32 = 40;
const TITLEBAR_BTN_W: i32 = 44;
const RESIZE_BORDER_LOGICAL: i32 = 6;
/// How far a combo's label is inset inside its item box. Measured so its first glyph lands on the
/// column an `EDIT`'s text starts on: the EDIT's container pad (8) plus its own 1px, less the item
/// box's own 3, less the 2px the combo's font keeps before the glyph.
const COMBO_LABEL_INSET_LOGICAL: i32 = 4;
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
const ID_ENABLED: i32 = 116;
const ID_VOLUME: i32 = 117;
const ID_VOLUME_VALUE: i32 = 118;
const ID_WINDOW_MODE: i32 = 119;
const ID_CAPTURE_TECHNIQUE: i32 = 120;
const ID_DELETE: i32 = 121;
const ID_CLOSE_BUTTON: i32 = 122;
const ID_SCALING_UP: i32 = 123;
const ID_SCALING_DOWN: i32 = 124;
const ID_VOLUME_UP: i32 = 125;
const ID_VOLUME_DOWN: i32 = 126;
const ID_PROMOTION_TECHNIQUE: i32 = 127;
const ID_VOLUME_ENABLED: i32 = 128;
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

const CAPTURE_TECHNIQUES: [CustomCaptureTechnique; 6] = [
    CustomCaptureTechnique::InProcessHook,
    CustomCaptureTechnique::InvisibleLayered,
    CustomCaptureTechnique::BehindHost,
    CustomCaptureTechnique::OffScreen,
    CustomCaptureTechnique::DirectWindow,
    CustomCaptureTechnique::OffscreenStaged,
];

fn capture_technique_label(technique: CustomCaptureTechnique) -> &'static str {
    match technique {
        CustomCaptureTechnique::InProcessHook => "In-Process Hook (Default)",
        CustomCaptureTechnique::InvisibleLayered => "Invisible Layered",
        CustomCaptureTechnique::BehindHost => "Behind Host (On-Screen)",
        CustomCaptureTechnique::OffScreen => "Off-Screen (DWM)",
        CustomCaptureTechnique::DirectWindow => "Direct Window",
        CustomCaptureTechnique::OffscreenStaged => "Off-Screen Staged (-10000 -> Reveal)",
    }
}

const PROMOTION_TECHNIQUES: [CustomPromotionTechnique; 2] = [
    CustomPromotionTechnique::DirectPromotion,
    CustomPromotionTechnique::OpenAsNewProcess,
];

fn promotion_technique_label(technique: CustomPromotionTechnique) -> &'static str {
    match technique {
        CustomPromotionTechnique::DirectPromotion => "Direct Promotion",
        CustomPromotionTechnique::OpenAsNewProcess => "Open as New Process",
    }
}

fn is_combo(id: u32) -> bool {
    id == ID_WINDOW_MODE as u32
        || id == ID_CAPTURE_TECHNIQUE as u32
        || id == ID_PROMOTION_TECHNIQUE as u32
        || id == ID_SCALING as u32
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
    field_disabled: COLORREF,
    border: COLORREF,
    ink: COLORREF,
    label: COLORREF,
    muted: COLORREF,
    accent: COLORREF,
    accent_hover: COLORREF,
    accent_press: COLORREF,
    accent_text: COLORREF,
    number: COLORREF,
    button_bg: COLORREF,
    button_hover: COLORREF,
    button_press: COLORREF,
    button_disabled: COLORREF,
    danger: COLORREF,
    danger_bg: COLORREF,
    danger_text: COLORREF,
    danger_on_fill: COLORREF,
    danger_hover: COLORREF,
    slider_track: COLORREF,
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
                mix(page, ink, 20),
                mix(card, ink, 10),
                mix(card, ink, 18),
                mix(card, ink, 26),
            )
        } else {
            let card = mix(page, rgb(0, 0, 0), 3);
            (
                mix(page, rgb(0, 0, 0), 4),
                card,
                rgb(255, 255, 255),
                mix(page, ink, 18),
                mix(card, ink, 6),
                mix(card, ink, 12),
                mix(card, ink, 19),
            )
        };

        let field_disabled = mix(card, field, 50);
        let button_disabled = mix(card, border, 32);
        let slider_track = mix(card, ink, if dark { 20 } else { 16 });

        // Authentic One Dark Pro red (#E06C75) / Atom One Light red (#E45649).
        let (danger, danger_text, danger_on_fill, danger_bg) = if dark {
            let d = rgb(0xE0, 0x6C, 0x75);
            (d, d, rgb(0x1E, 0x22, 0x2A), mix(card, d, 18))
        } else {
            let d = rgb(0xE4, 0x56, 0x49);
            (
                d,
                rgb(0xC9, 0x3B, 0x32),
                rgb(0xFF, 0xFF, 0xFF),
                mix(card, d, 10),
            )
        };
        let danger_hover = if dark { rgb(0xD8, 0x4B, 0x56) } else { danger };

        let accent_hover = mix(accent, rgb(255, 255, 255), if dark { 14 } else { 12 });
        let accent_press = mix(accent, rgb(0, 0, 0), 16);
        let accent_text = if dark {
            rgb(0x1E, 0x22, 0x2A)
        } else {
            rgb(0xFF, 0xFF, 0xFF)
        };

        let label = if dark {
            rgb(0xD4, 0xDA, 0xE4)
        } else {
            rgb(0x23, 0x26, 0x2E)
        };

        Self {
            page,
            titlebar,
            card,
            field,
            field_disabled,
            border,
            ink,
            label,
            muted: mix(ink, page, 32),
            accent,
            accent_hover,
            accent_press,
            accent_text,
            number,
            button_bg,
            button_hover,
            button_press,
            button_disabled,
            danger,
            danger_bg,
            danger_text,
            danger_on_fill,
            danger_hover,
            slider_track,
            dark,
        }
    }
}

/// Primary CTA button vertical gradient `(top, bottom, border, text)` keeping >= 7:1 text contrast.
fn primary_button_colors(
    p: Palette,
    hovered: bool,
    pressed: bool,
) -> (COLORREF, COLORREF, COLORREF, COLORREF) {
    if p.dark {
        let (top, bottom) = if pressed {
            (mix(p.accent, rgb(255, 255, 255), 8), p.accent)
        } else if hovered {
            (
                mix(p.accent, rgb(255, 255, 255), 24),
                mix(p.accent, rgb(255, 255, 255), 10),
            )
        } else {
            (
                mix(p.accent, rgb(255, 255, 255), 16),
                mix(p.accent, rgb(255, 255, 255), 2),
            )
        };
        (
            top,
            bottom,
            mix(p.accent, rgb(255, 255, 255), 24),
            rgb(0x0B, 0x0F, 0x15),
        )
    } else {
        // Darkened Atom One Light accent keeps >= 7:1 contrast with white text on both stops.
        let (top, bottom) = if pressed {
            (
                mix(p.accent, rgb(0, 0, 0), 40),
                mix(p.accent, rgb(0, 0, 0), 50),
            )
        } else if hovered {
            (
                mix(p.accent, rgb(0, 0, 0), 29),
                mix(p.accent, rgb(0, 0, 0), 38),
            )
        } else {
            (
                mix(p.accent, rgb(0, 0, 0), 31),
                mix(p.accent, rgb(0, 0, 0), 42),
            )
        };
        (top, bottom, mix(p.accent, rgb(0, 0, 0), 46), p.accent_text)
    }
}

const GRADIENT_W: i32 = 128;
const GRADIENT_H: i32 = 96;

#[derive(Clone, Copy, PartialEq, Eq)]
struct GradientKey {
    dark: bool,
    page: COLORREF,
    accent: COLORREF,
    number: COLORREF,
}

struct GradientCache {
    key: GradientKey,
    dib: HBITMAP,
}

impl Drop for GradientCache {
    fn drop(&mut self) {
        if !self.dib.is_invalid() {
            unsafe {
                let _ = DeleteObject(HGDIOBJ(self.dib.0));
            }
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
struct CardSliceKey {
    dark: bool,
    card: COLORREF,
    border: COLORREF,
    dpi: i32,
}

/// Premultiplied-BGRA 9-slice card template (rounded corners + border + top sheen + soft shadow).
struct CardSliceCache {
    key: CardSliceKey,
    dib: HBITMAP,
    tex_w: i32,
    tex_h: i32,
    pad_l: i32,
    pad_t: i32,
    pad_r: i32,
    pad_b: i32,
    corner: i32,
}

impl Drop for CardSliceCache {
    fn drop(&mut self) {
        if !self.dib.is_invalid() {
            unsafe {
                let _ = DeleteObject(HGDIOBJ(self.dib.0));
            }
        }
    }
}

/// Reused window backbuffer DC + bitmap grown lazily on resize to eliminate per-frame GDI allocs.
struct BackbufferCache {
    w: i32,
    h: i32,
    dc: HDC,
    bmp: HBITMAP,
    old_bmp: HGDIOBJ,
}

impl Drop for BackbufferCache {
    fn drop(&mut self) {
        unsafe {
            if !self.dc.is_invalid() {
                if !self.old_bmp.is_invalid() {
                    let _ = SelectObject(self.dc, self.old_bmp);
                }
                let _ = DeleteDC(self.dc);
            }
            if !self.bmp.is_invalid() {
                let _ = DeleteObject(HGDIOBJ(self.bmp.0));
            }
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
    list_rect: RECT,
    edit_boxes: [(i32, RECT); 6],
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

    fn edit_box_at(&self, x: i32, y: i32) -> Option<i32> {
        self.edit_boxes
            .iter()
            .find(|(_, r)| pt_in_rect(*r, x, y))
            .map(|(id, _)| *id)
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
    /// The drawn row holding the preset a drag is carrying, while one is in flight.
    drag: Option<usize>,
    /// A field of the open preset has been edited and not yet written (see `autosave`).
    dirtied: bool,
    loading: bool,
    palette: Palette,
    font: HFONT,
    font_bold: HFONT,
    font_small: HFONT,
    brush_page: HBRUSH,
    brush_card: HBRUSH,
    brush_field: HBRUSH,
    brush_field_disabled: HBRUSH,
    dpi: i32,
    hovered_tb: Option<TitlebarButton>,
    pressed_tb: Option<TitlebarButton>,
    tracking_mouse: bool,
    rects: LayoutRects,
    gradient: Option<GradientCache>,
    card_slice: Option<CardSliceCache>,
    backbuffer: Option<BackbufferCache>,
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
            let _ = DeleteObject(HGDIOBJ(self.brush_field_disabled.0));
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

struct ButtonSubclass {
    orig_proc: WNDPROC,
    hovered: bool,
    tracking: bool,
}

struct ComboSubclass {
    orig_proc: WNDPROC,
    parent: HWND,
    hovered: bool,
    tracking: bool,
}

struct ListSubclass {
    orig_proc: WNDPROC,
    parent: HWND,
    /// Position in the drawn rows a press landed on, while that press may still turn out to be a
    /// click.
    press: Option<usize>,
    /// Whether that press has become a drag of the row it landed on.
    dragging: bool,
    /// The drawn row the drag was last posted for, so one row is one message rather than one a
    /// mouse pixel.
    posted: Option<usize>,
    hovered_pill: Option<usize>,
    tracking: bool,
}

struct SliderState {
    parent: HWND,
    value: u32,
    dragging: bool,
    hovered: bool,
    tracking: bool,
}

struct ToggleState {
    parent: HWND,
    id: i32,
    show_label: bool,
    checked: bool,
    hovered: bool,
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

/// The window's presets as the file may hold them, or the one thing wrong with them.
///
/// It is a plain value rather than a write: a preset with no name and two presets of one name are
/// the two mistakes the window can be in, and either is refused before anything is written — a
/// name being retyped passes through the first on its way to the second, and writing it would put
/// a nameless section in the file.
fn clean(presets: &[CustomPreset]) -> Result<Vec<CustomPreset>, String> {
    let mut cleaned: Vec<CustomPreset> = Vec::new();
    for preset in presets {
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
    Ok(cleaned)
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
            style: WNDCLASS_STYLES(0),
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
            | WS_CLIPCHILDREN
            | WS_CLIPSIBLINGS;
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            CLASS_NAME,
            PCWSTR(title.as_ptr()),
            style,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            920,
            660,
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

unsafe fn apply_combo_list_theme(combo: HWND, dark: bool) {
    let mut info = COMBOBOXINFO {
        cbSize: std::mem::size_of::<COMBOBOXINFO>() as u32,
        ..Default::default()
    };
    if unsafe { GetComboBoxInfo(combo, &mut info) }.is_ok() && !info.hwndList.0.is_null() {
        let theme = if dark {
            w!("DarkMode_Explorer")
        } else {
            w!("")
        };
        let _ = unsafe { SetWindowTheme(info.hwndList, theme, None) };
    }
}

unsafe fn apply_theme(state: &mut State) {
    let palette = Palette::of(theme_now());
    state.palette = palette;
    state.gradient = None;
    state.card_slice = None;
    unsafe {
        let _ = DeleteObject(HGDIOBJ(state.brush_page.0));
        let _ = DeleteObject(HGDIOBJ(state.brush_card.0));
        let _ = DeleteObject(HGDIOBJ(state.brush_field.0));
        let _ = DeleteObject(HGDIOBJ(state.brush_field_disabled.0));
        state.brush_page = CreateSolidBrush(palette.page);
        state.brush_card = CreateSolidBrush(palette.card);
        state.brush_field = CreateSolidBrush(palette.field);
        state.brush_field_disabled = CreateSolidBrush(palette.field_disabled);
        apply_dwm_chrome(state.hwnd, palette.dark);
        for id in [
            ID_WINDOW_MODE,
            ID_CAPTURE_TECHNIQUE,
            ID_PROMOTION_TECHNIQUE,
            ID_SCALING,
        ] {
            if let Ok(combo) = GetDlgItem(state.hwnd, id) {
                apply_combo_list_theme(combo, palette.dark);
            }
        }
        if let Ok(list) = GetDlgItem(state.hwnd, ID_LIST) {
            let theme = if palette.dark {
                w!("DarkMode_Explorer")
            } else {
                w!("")
            };
            let _ = SetWindowTheme(list, theme, None);
        }
        let _ = RedrawWindow(
            state.hwnd,
            None,
            None,
            RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW,
        );
    }
}

/// The editor window's own procedure.
///
/// The whole match runs under `catch_unwind` for the reason the tray's own proc does
/// (`shell::tray::event_loop`): it is `extern "system"`, so a panic unwinding out of it cannot
/// unwind and aborts the process instead — and this window is one thread of the process the tray
/// lives in, so a bug in the editor would take the tray down with it. A panic answers `LRESULT(0)`
/// and leaves the window up.
unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    std::panic::catch_unwind(|| unsafe { editor_proc(hwnd, msg, wparam, lparam) })
        .unwrap_or(LRESULT(0))
}

unsafe fn editor_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
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
                // Clicking anywhere in a padded EDIT container focuses its centered EDIT control.
                if let Some(edit_id) = state.rects.edit_box_at(x, y) {
                    if let Ok(edit) = unsafe { GetDlgItem(hwnd, edit_id) } {
                        if unsafe { IsWindowEnabled(edit) }.as_bool() {
                            let _ = unsafe { SetFocus(edit) };
                            return LRESULT(0);
                        }
                    }
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
            info.ptMinTrackSize.y = scale(540, dpi);
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
        WM_TOGGLE_CHANGED => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                let enabled = wparam.0 != 0;
                if lparam.0 as i32 == ID_VOLUME_ENABLED {
                    unsafe { volume_enabled_toggled(state, enabled) };
                } else {
                    unsafe { enabled_toggled(state, enabled) };
                }
            }
            LRESULT(0)
        }
        WM_PRESET_DRAG_MOVE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { preset_drag_moved(state, wparam.0, lparam.0 as usize) };
            }
            LRESULT(0)
        }
        WM_PRESET_DRAG_END => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                state.drag = None;
                // The list is the order a file is routed by, so a move is written at once.
                state.dirtied = true;
                unsafe { autosave(state) };
            }
            LRESULT(0)
        }
        WM_AUTOSAVE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { autosave(state) };
            }
            LRESULT(0)
        }
        WM_PRESET_ROW_TOGGLE => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { row_pill_toggled(state, wparam.0) };
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
            // Disabled EDIT controls send WM_CTLCOLORSTATIC; match container disabled fill seamlessly.
            let is_edit = EDIT_IDS
                .iter()
                .any(|&id| unsafe { GetDlgItem(hwnd, id) }.is_ok_and(|h| h == ctl));
            let status_hwnd = unsafe { GetDlgItem(hwnd, ID_STATUS) }.unwrap_or_default();
            let (bg, brush, ink) = if is_edit {
                (
                    state.palette.field_disabled,
                    state.brush_field_disabled,
                    state.palette.muted,
                )
            } else if ctl == status_hwnd {
                (state.palette.page, state.brush_page, state.palette.muted)
            } else {
                (state.palette.card, state.brush_card, state.palette.label)
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
            let ctl = HWND(lparam.0 as *mut _);
            let enabled = unsafe { IsWindowEnabled(ctl) }.as_bool();
            let (bg, brush, ink) = if enabled {
                (state.palette.field, state.brush_field, state.palette.ink)
            } else {
                (
                    state.palette.field_disabled,
                    state.brush_field_disabled,
                    state.palette.muted,
                )
            };
            unsafe {
                SetTextColor(hdc, ink);
                SetBkColor(hdc, bg);
            }
            LRESULT(brush.0 as isize)
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
            unsafe {
                if let Some(state) = state_of(hwnd) {
                    paint_window_buffered(state, hwnd);
                } else {
                    let mut ps = PAINTSTRUCT::default();
                    let _ = BeginPaint(hwnd, &mut ps);
                    let _ = EndPaint(hwnd, &ps);
                }
            }
            LRESULT(0)
        }
        WM_THEME_CHANGED => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { apply_theme(state) };
            }
            LRESULT(0)
        }
        WM_ACTIVATE => {
            // The window going inactive is the user done with whatever field held the focus — a
            // control keeps its focus inside an inactive window, so no `WM_KILLFOCUS` comes.
            if (wparam.0 & 0xFFFF) as u32 == WA_INACTIVE {
                if let Some(state) = unsafe { state_of(hwnd) } {
                    unsafe { autosave(state) };
                }
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
                unsafe { autosave(state) };
                // Closing over a preset the file may not hold would lose it with nothing said:
                // the reason is on the status line, and the window stays up to be put right.
                if state.dirtied {
                    return LRESULT(0);
                }
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

/// Offscreen-buffered `WM_PAINT` helper so zero unbuffered fills hit the screen DC.
unsafe fn paint_double_buffered<F: FnOnce(HDC, RECT)>(hwnd: HWND, draw: F) {
    let mut ps = PAINTSTRUCT::default();
    let hdc = unsafe { BeginPaint(hwnd, &mut ps) };
    let mut rc = RECT::default();
    let _ = unsafe { GetClientRect(hwnd, &mut rc) };
    let w = (rc.right - rc.left).max(1);
    let h = (rc.bottom - rc.top).max(1);
    let mem_dc = unsafe { CreateCompatibleDC(hdc) };
    let mem_bmp = unsafe { CreateCompatibleBitmap(hdc, w, h) };
    if !mem_dc.is_invalid() && !mem_bmp.is_invalid() {
        let old_bmp = unsafe { SelectObject(mem_dc, HGDIOBJ(mem_bmp.0)) };
        draw(mem_dc, rc);
        let _ = unsafe { BitBlt(hdc, 0, 0, w, h, mem_dc, 0, 0, SRCCOPY) };
        let _ = unsafe { SelectObject(mem_dc, old_bmp) };
    } else {
        draw(hdc, rc);
    }
    if !mem_bmp.is_invalid() {
        let _ = unsafe { DeleteObject(HGDIOBJ(mem_bmp.0)) };
    }
    if !mem_dc.is_invalid() {
        let _ = unsafe { DeleteDC(mem_dc) };
    }
    let _ = unsafe { EndPaint(hwnd, &ps) };
}

/// Paints main window using a reused compatible backbuffer bitmap (grown lazily on resize).
unsafe fn paint_window_buffered(state: &mut State, hwnd: HWND) {
    let mut ps = PAINTSTRUCT::default();
    let hdc = unsafe { BeginPaint(hwnd, &mut ps) };
    let mut rc = RECT::default();
    let _ = unsafe { GetClientRect(hwnd, &mut rc) };
    let w = (rc.right - rc.left).max(1);
    let h = (rc.bottom - rc.top).max(1);

    let needs_alloc = state
        .backbuffer
        .as_ref()
        .is_none_or(|bb| bb.w < w || bb.h < h || bb.dc.is_invalid() || bb.bmp.is_invalid());
    if needs_alloc {
        // Drop old before allocating new; add small headroom to avoid realloc on every 1px drag.
        state.backbuffer = None;
        let alloc_w = w.saturating_add(128);
        let alloc_h = h.saturating_add(128);
        let mem_dc = unsafe { CreateCompatibleDC(hdc) };
        let mem_bmp = unsafe { CreateCompatibleBitmap(hdc, alloc_w, alloc_h) };
        if !mem_dc.is_invalid() && !mem_bmp.is_invalid() {
            let old_bmp = unsafe { SelectObject(mem_dc, HGDIOBJ(mem_bmp.0)) };
            state.backbuffer = Some(BackbufferCache {
                w: alloc_w,
                h: alloc_h,
                dc: mem_dc,
                bmp: mem_bmp,
                old_bmp,
            });
        } else {
            if !mem_bmp.is_invalid() {
                let _ = unsafe { DeleteObject(HGDIOBJ(mem_bmp.0)) };
            }
            if !mem_dc.is_invalid() {
                let _ = unsafe { DeleteDC(mem_dc) };
            }
        }
    }

    if let Some(bb_dc) = state.backbuffer.as_ref().map(|bb| bb.dc) {
        unsafe { paint_window(state, bb_dc) };
        let _ = unsafe { BitBlt(hdc, 0, 0, w, h, bb_dc, 0, 0, SRCCOPY) };
    } else {
        unsafe { paint_window(state, hdc) };
    }
    let _ = unsafe { EndPaint(hwnd, &ps) };
}

/// Rebuilds the cached `128x96` atmospheric gradient DIB only when theme changes (never on resize).
unsafe fn ensure_gradient_cache(state: &mut State, hdc: HDC) -> Option<HBITMAP> {
    let key = GradientKey {
        dark: state.palette.dark,
        page: state.palette.page,
        accent: state.palette.accent,
        number: state.palette.number,
    };
    if let Some(cache) = &state.gradient {
        if cache.key == key && !cache.dib.is_invalid() {
            return Some(cache.dib);
        }
    }
    let surface = render_gradient_backdrop(GRADIENT_W, GRADIENT_H, state.palette);
    let dib = unsafe { surface.create_dib(hdc) }?;
    state.gradient = Some(GradientCache { key, dib });
    Some(dib)
}

/// Rebuilds the cached premultiplied-BGRA 9-slice card template DIB only on theme/DPI change.
unsafe fn ensure_card_slice_cache(state: &mut State, hdc: HDC) -> bool {
    let key = CardSliceKey {
        dark: state.palette.dark,
        card: state.palette.card,
        border: state.palette.border,
        dpi: state.dpi,
    };
    if let Some(cache) = &state.card_slice {
        if cache.key == key && !cache.dib.is_invalid() {
            return true;
        }
    }
    let (pixels, tex_w, tex_h, pad_l, pad_t, pad_r, pad_b, corner) =
        render_card_9slice_pixels(state.palette, state.dpi);
    let Some(dib) = (unsafe { create_bgra_dib(hdc, tex_w, tex_h, &pixels) }) else {
        return false;
    };
    state.card_slice = Some(CardSliceCache {
        key,
        dib,
        tex_w,
        tex_h,
        pad_l,
        pad_t,
        pad_r,
        pad_b,
        corner,
    });
    true
}

/// Draws an elevated SDF rounded card via 8 tiny 9-slice `GdiAlphaBlend` calls + 1 `FillRect`.
unsafe fn draw_elevated_card_9slice(
    hdc: HDC,
    card: &CardSliceCache,
    rect: RECT,
    brush_card: HBRUSH,
) {
    let cw = rect.right - rect.left;
    let ch = rect.bottom - rect.top;
    if cw <= 0 || ch <= 0 {
        return;
    }
    let src_dc = unsafe { CreateCompatibleDC(hdc) };
    if src_dc.is_invalid() {
        return;
    }
    let old = unsafe { SelectObject(src_dc, HGDIOBJ(card.dib.0)) };
    let bf = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };

    let c = card.corner.min(cw / 2).min(ch / 2);
    let sl = card.pad_l + c;
    let st = card.pad_t + c;
    let sr = card.pad_r + c;
    let sb = card.pad_b + c;
    let sx_r = card.tex_w - sr;
    let sy_b = card.tex_h - sb;
    let sx_mid = sl;
    let sy_mid = st;
    let sw_mid = (card.tex_w - sl - sr).max(1);
    let sh_mid = (card.tex_h - st - sb).max(1);

    let dx0 = rect.left - card.pad_l;
    let dy0 = rect.top - card.pad_t;
    let dx_mid = rect.left + c;
    let dy_mid = rect.top + c;
    let dx_r = rect.right - c;
    let dy_b = rect.bottom - c;
    let dw_mid = (dx_r - dx_mid).max(0);
    let dh_mid = (dy_b - dy_mid).max(0);

    unsafe {
        // 4 corners (exact 1:1)
        let _ = GdiAlphaBlend(hdc, dx0, dy0, sl, st, src_dc, 0, 0, sl, st, bf);
        let _ = GdiAlphaBlend(hdc, dx_r, dy0, sr, st, src_dc, sx_r, 0, sr, st, bf);
        let _ = GdiAlphaBlend(hdc, dx0, dy_b, sl, sb, src_dc, 0, sy_b, sl, sb, bf);
        let _ = GdiAlphaBlend(hdc, dx_r, dy_b, sr, sb, src_dc, sx_r, sy_b, sr, sb, bf);

        // 4 straight edge strips (shadow + 1px border + 1px top sheen)
        if dw_mid > 0 {
            let _ = GdiAlphaBlend(
                hdc, dx_mid, dy0, dw_mid, st, src_dc, sx_mid, 0, sw_mid, st, bf,
            );
            let _ = GdiAlphaBlend(
                hdc, dx_mid, dy_b, dw_mid, sb, src_dc, sx_mid, sy_b, sw_mid, sb, bf,
            );
        }
        if dh_mid > 0 {
            let _ = GdiAlphaBlend(
                hdc, dx0, dy_mid, sl, dh_mid, src_dc, 0, sy_mid, sl, sh_mid, bf,
            );
            let _ = GdiAlphaBlend(
                hdc, dx_r, dy_mid, sr, dh_mid, src_dc, sx_r, sy_mid, sr, sh_mid, bf,
            );
        }
        // Solid flat card interior in one FillRect (zero per-pixel SDF work).
        if dw_mid > 0 && dh_mid > 0 {
            let inner = RECT {
                left: dx_mid,
                top: dy_mid,
                right: dx_r,
                bottom: dy_b,
            };
            FillRect(hdc, &inner, brush_card);
        }

        let _ = SelectObject(src_dc, old);
        let _ = DeleteDC(src_dc);
    }
}

unsafe fn paint_window(state: &mut State, hdc: HDC) {
    let mut client = RECT::default();
    let _ = unsafe { GetClientRect(state.hwnd, &mut client) };
    let w = (client.right - client.left).max(1);
    let h = (client.bottom - client.top).max(1);
    let p = state.palette;
    let dpi = state.dpi;

    unsafe {
        // Stretch cached 128x96 atmospheric gradient DIB via HALFTONE (<0.1ms, never rebuilt on resize).
        if let Some(dib) = ensure_gradient_cache(state, hdc) {
            let src_dc = CreateCompatibleDC(hdc);
            if !src_dc.is_invalid() {
                let old = SelectObject(src_dc, HGDIOBJ(dib.0));
                let _ = SetStretchBltMode(hdc, HALFTONE);
                let _ = SetBrushOrgEx(hdc, 0, 0, None);
                let _ = StretchBlt(
                    hdc, 0, 0, w, h, src_dc, 0, 0, GRADIENT_W, GRADIENT_H, SRCCOPY,
                );
                let _ = SelectObject(src_dc, old);
                let _ = DeleteDC(src_dc);
            }
        } else {
            FillRect(hdc, &client, state.brush_page);
        }

        // Clean titlebar band + 1px bottom divider (no top neon bar or accent dot).
        let tb_end = state.rects.titlebar.bottom.clamp(0, h);
        if tb_end > 1 {
            let tb_col = mix(p.page, p.titlebar, 88);
            let tb_brush = CreateSolidBrush(tb_col);
            let tb_rc = RECT {
                left: 0,
                top: 0,
                right: w,
                bottom: tb_end - 1,
            };
            FillRect(hdc, &tb_rc, tb_brush);
            let _ = DeleteObject(HGDIOBJ(tb_brush.0));
        }
        if tb_end > 0 {
            let div_brush = CreateSolidBrush(p.border);
            let div_rc = RECT {
                left: 0,
                top: tb_end - 1,
                right: w,
                bottom: tb_end,
            };
            FillRect(hdc, &div_rc, div_brush);
            let _ = DeleteObject(HGDIOBJ(div_brush.0));
        }

        // Smooth SDF elevated rounded cards via 9-slice AlphaBlend + FillRect (<0.1ms at 144Hz).
        if ensure_card_slice_cache(state, hdc) {
            if let Some(card) = state.card_slice.as_ref() {
                draw_elevated_card_9slice(hdc, card, state.rects.sidebar_card, state.brush_card);
                draw_elevated_card_9slice(hdc, card, state.rects.editor_card, state.brush_card);
            }
        }

        let border_pen = CreatePen(PS_SOLID, 1, p.border);
        let old_pen = SelectObject(hdc, HGDIOBJ(border_pen.0));

        // Title on left aligned cleanly to pad.
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, p.ink);
        let old_font = SelectObject(hdc, HGDIOBJ(state.font_bold.0));
        let mut title_rc = RECT {
            left: scale(16, dpi),
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

        // Padded containers & focus accent ring around each centered EDIT control.
        let focused = GetFocus();
        let focus_pen = CreatePen(PS_SOLID, 1, p.accent);
        let field_radius = scale(4, dpi);
        for (id, r) in state.rects.edit_boxes {
            if r.right <= r.left || r.bottom <= r.top {
                continue;
            }
            let ctrl = GetDlgItem(state.hwnd, id).unwrap_or_default();
            let enabled = !ctrl.0.is_null() && IsWindowEnabled(ctrl).as_bool();
            let is_focused = enabled && ctrl == focused;
            let _ = SelectObject(
                hdc,
                HGDIOBJ(if is_focused {
                    focus_pen.0
                } else {
                    border_pen.0
                }),
            );
            let _ = SelectObject(
                hdc,
                HGDIOBJ(if enabled {
                    state.brush_field.0
                } else {
                    state.brush_field_disabled.0
                }),
            );
            let _ = RoundRect(
                hdc,
                r.left,
                r.top,
                r.right,
                r.bottom,
                field_radius,
                field_radius,
            );
            if id == ID_SEARCH {
                let icon_sz = scale(14, dpi);
                let icon_x = r.right - scale(7, dpi) - icon_sz;
                let icon_y = (r.top + r.bottom - icon_sz) / 2;
                let bg = if enabled { p.field } else { p.field_disabled };
                let fg = if is_focused {
                    p.muted
                } else {
                    mix(p.muted, p.field, 18)
                };
                let icon = render_search_icon(icon_sz, bg, fg);
                icon.blit_to(hdc, icon_x, icon_y);
            }
        }

        // Border & focus accent ring around LISTBOX (uses lay_out rect to stay locked during resize).
        let lr = state.rects.list_rect;
        if lr.right > lr.left && lr.bottom > lr.top {
            let ctrl = GetDlgItem(state.hwnd, ID_LIST).unwrap_or_default();
            let is_focused = !ctrl.0.is_null() && ctrl == focused;
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
                lr.left - 1,
                lr.top - 1,
                lr.right + 1,
                lr.bottom + 1,
                field_radius,
                field_radius,
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
        (TitlebarButton::Close, true, _) => Some(mix(p.danger_hover, rgb(0, 0, 0), 18)),
        (TitlebarButton::Close, false, true) => Some(p.danger_hover),
        (_, true, _) => Some(p.button_press),
        (_, false, true) => Some(p.button_hover),
        _ => None,
    };
    let fg = if kind == TitlebarButton::Close && (hovered || pressed) {
        rgb(255, 255, 255)
    } else if hovered || pressed {
        p.ink
    } else {
        p.muted
    };

    unsafe {
        if let Some(bg_col) = bg {
            let fill_rc = RECT {
                left: rect.left,
                top: rect.top,
                right: rect.right,
                bottom: rect.bottom - 1,
            };
            let brush = CreateSolidBrush(bg_col);
            FillRect(hdc, &fill_rc, brush);
            let _ = DeleteObject(HGDIOBJ(brush.0));
        }

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
    WINDOW_STYLE(WS_CHILD.0 | WS_VISIBLE.0 | WS_CLIPSIBLINGS.0 | extra)
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
        drag: None,
        dirtied: false,
        loading: false,
        palette,
        font,
        font_bold,
        font_small,
        brush_page: unsafe { CreateSolidBrush(palette.page) },
        brush_card: unsafe { CreateSolidBrush(palette.card) },
        brush_field: unsafe { CreateSolidBrush(palette.field) },
        brush_field_disabled: unsafe { CreateSolidBrush(palette.field_disabled) },
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
            list_rect: RECT::default(),
            edit_boxes: EDIT_IDS.map(|id| (id, RECT::default())),
        },
        gradient: None,
        card_slice: None,
        backbuffer: None,
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
                    | LBS_HASSTRINGS as u32
                    | LBS_NOINTEGRALHEIGHT as u32,
            ),
            ID_LIST,
        )
    });

    // Editor pane: labels & inputs (static label text set once here to avoid WM_SIZE repaints).
    let editor_labels = [
        "Enabled",
        "Name",
        "Extensions",
        "Program",
        "Capture Technique",
        "Promotion Technique",
        "Window",
        "Scaling",
        "Volume",
    ];
    for (index, label) in editor_labels.iter().enumerate() {
        controls.push(unsafe {
            child(
                hwnd,
                w!("STATIC"),
                label,
                child_style(0),
                ID_LABEL_BASE + index as i32,
            )
        });
    }
    controls.push(unsafe { create_themed_toggle(hwnd, ID_ENABLED, true, true) });
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
            ID_CAPTURE_TECHNIQUE,
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
            ID_PROMOTION_TECHNIQUE,
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
    controls.push(unsafe { create_themed_toggle(hwnd, ID_VOLUME_ENABLED, false, false) });
    controls.push(unsafe { create_themed_slider(hwnd, ID_VOLUME) });
    controls.push(unsafe { child(hwnd, w!("EDIT"), "", edit_style, ID_VOLUME_VALUE) });
    let vol_up = unsafe { stepper_button(hwnd, "▲", ID_VOLUME_UP, ID_VOLUME_VALUE, 1) };
    let vol_dn = unsafe { stepper_button(hwnd, "▼", ID_VOLUME_DOWN, ID_VOLUME_VALUE, -1) };

    // Footer: status & actions. There is no `Save` — every field is written when it is left
    // (see `autosave`).
    controls.push(unsafe { child(hwnd, w!("STATIC"), "", child_style(0), ID_STATUS) });
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
    for index in 0..9 {
        if let Ok(lbl) = unsafe { GetDlgItem(hwnd, ID_LABEL_BASE + index) } {
            unsafe {
                SendMessageW(
                    lbl,
                    WM_SETFONT,
                    WPARAM(state.font_bold.0 as usize),
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

    // Configure every EDIT: generous text limits and keyboard/triple-click subclass.
    // Horizontal inset comes from the parent-painted outer container rect.
    let edit_limits = [
        (ID_SEARCH, 32767usize),
        (ID_NAME, 48),
        (ID_EXTENSIONS, 32767),
        (ID_EXE, 32767),
        (ID_SCALING_PERCENT, 4),
        (ID_VOLUME_VALUE, 3),
    ];
    for (id, limit) in edit_limits {
        if let Ok(edit) = unsafe { GetDlgItem(hwnd, id) } {
            unsafe {
                SendMessageW(edit, EM_SETLIMITTEXT, WPARAM(limit), LPARAM(0));
                install_edit_subclass(edit, hwnd, id);
            }
        }
    }

    if let Ok(list) = unsafe { GetDlgItem(hwnd, ID_LIST) } {
        let theme = if palette.dark {
            w!("DarkMode_Explorer")
        } else {
            w!("")
        };
        let _ = unsafe { SetWindowTheme(list, theme, None) };
        unsafe { install_list_subclass(list, hwnd) };
    }

    let window_mode_combo = unsafe { GetDlgItem(hwnd, ID_WINDOW_MODE) }.unwrap_or_default();
    for mode in WINDOW_MODES {
        unsafe { combo_add(window_mode_combo, window_mode_label(mode)) };
    }

    let capture_technique_combo =
        unsafe { GetDlgItem(hwnd, ID_CAPTURE_TECHNIQUE) }.unwrap_or_default();
    for technique in CAPTURE_TECHNIQUES {
        unsafe { combo_add(capture_technique_combo, capture_technique_label(technique)) };
    }

    let promotion_technique_combo =
        unsafe { GetDlgItem(hwnd, ID_PROMOTION_TECHNIQUE) }.unwrap_or_default();
    for technique in PROMOTION_TECHNIQUES {
        unsafe {
            combo_add(
                promotion_technique_combo,
                promotion_technique_label(technique),
            )
        };
    }

    let scaling = unsafe { GetDlgItem(hwnd, ID_SCALING) }.unwrap_or_default();
    for (label, _) in SCALE_CHOICES {
        unsafe { combo_add(scaling, label) };
    }
    unsafe { combo_add(scaling, "Custom…") };

    for combo in [
        window_mode_combo,
        capture_technique_combo,
        promotion_technique_combo,
        scaling,
    ] {
        if !combo.0.is_null() {
            unsafe {
                install_combo_subclass(combo, hwnd);
                apply_combo_list_theme(combo, palette.dark);
            }
        }
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
            // The user has moved on to something else, which is the moment the field is written
            // (see `autosave`).
            if msg == WM_KILLFOCUS && sub.id != ID_SEARCH {
                let _ = unsafe { PostMessageW(sub.parent, WM_AUTOSAVE, WPARAM(0), LPARAM(0)) };
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
        WM_ERASEBKGND => return LRESULT(1),
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
                // The steps of one hold are one edit, written when the button comes up
                // (see `autosave`).
                let _ = PostMessageW(sub.parent, WM_AUTOSAVE, WPARAM(0), LPARAM(0));
            }
            return LRESULT(0);
        }
        WM_CAPTURECHANGED | WM_CANCELMODE => {
            sub.pressed = false;
            unsafe {
                let _ = KillTimer(hwnd, STEP_TIMER_INITIAL);
                let _ = KillTimer(hwnd, STEP_TIMER_REPEAT);
                let _ = InvalidateRect(hwnd, None, false);
                let _ = PostMessageW(sub.parent, WM_AUTOSAVE, WPARAM(0), LPARAM(0));
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
    state.dirtied = true;
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
                let slider = GetDlgItem(state.hwnd, ID_VOLUME).unwrap_or_default();
                SendMessageW(
                    slider,
                    WM_SLIDER_SET_VALUE,
                    WPARAM(next as usize),
                    LPARAM(0),
                );
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
    let btn = unsafe {
        child(
            hwnd,
            w!("BUTTON"),
            text,
            child_style(WS_TABSTOP.0 | BS_OWNERDRAW as u32),
            id,
        )
    };
    if !btn.0.is_null() {
        let orig = unsafe {
            SetWindowLongPtrW(
                btn,
                GWLP_WNDPROC,
                button_subclass_proc as *const () as usize as isize,
            )
        };
        let orig_proc: WNDPROC = unsafe { std::mem::transmute(orig) };
        let data = Box::into_raw(Box::new(ButtonSubclass {
            orig_proc,
            hovered: false,
            tracking: false,
        }));
        unsafe { SetWindowLongPtrW(btn, GWLP_USERDATA, data as isize) };
    }
    btn
}

unsafe extern "system" fn button_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut ButtonSubclass;
    let Some(sub) = (unsafe { raw.as_mut() }) else {
        return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
    };
    let orig_proc = sub.orig_proc;

    match msg {
        WM_ERASEBKGND => return LRESULT(1),
        WM_MOUSEMOVE => {
            if !sub.hovered {
                sub.hovered = true;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
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
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
            }
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

unsafe fn install_combo_subclass(combo: HWND, parent: HWND) {
    let orig = unsafe {
        SetWindowLongPtrW(
            combo,
            GWLP_WNDPROC,
            combo_subclass_proc as *const () as usize as isize,
        )
    };
    let orig_proc: WNDPROC = unsafe { std::mem::transmute(orig) };
    let data = Box::into_raw(Box::new(ComboSubclass {
        orig_proc,
        parent,
        hovered: false,
        tracking: false,
    }));
    unsafe { SetWindowLongPtrW(combo, GWLP_USERDATA, data as isize) };
}

unsafe extern "system" fn combo_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut ComboSubclass;
    let Some(sub) = (unsafe { raw.as_mut() }) else {
        return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
    };
    let orig_proc = sub.orig_proc;

    match msg {
        WM_ERASEBKGND => return LRESULT(1),
        WM_PAINT => {
            let (p, dpi, font, card_brush) = match unsafe { state_of(sub.parent) } {
                Some(state) => (state.palette, state.dpi, state.font, Some(state.brush_card)),
                None => (Palette::of(theme_now()), 96, HFONT::default(), None),
            };
            let enabled = unsafe { IsWindowEnabled(hwnd) }.as_bool();
            let focused = enabled && unsafe { GetFocus() } == hwnd;
            let dropped =
                unsafe { SendMessageW(hwnd, CB_GETDROPPEDSTATE, WPARAM(0), LPARAM(0)).0 } != 0;
            let active = enabled && (focused || dropped || sub.hovered);

            unsafe {
                paint_double_buffered(hwnd, |hdc, rc| {
                    if let Some(cb) = card_brush {
                        FillRect(hdc, &rc, cb);
                    } else {
                        let b = CreateSolidBrush(p.card);
                        FillRect(hdc, &rc, b);
                        let _ = DeleteObject(HGDIOBJ(b.0));
                    }

                    let fill_col = if enabled { p.field } else { p.field_disabled };
                    let border_col = if active { p.accent } else { p.border };
                    let fill_brush = CreateSolidBrush(fill_col);
                    let border_pen = CreatePen(PS_SOLID, 1, border_col);
                    let old_b = SelectObject(hdc, HGDIOBJ(fill_brush.0));
                    let old_p = SelectObject(hdc, HGDIOBJ(border_pen.0));
                    let r = scale(4, dpi);
                    let _ = RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, r, r);
                    let _ = SelectObject(hdc, old_b);
                    let _ = SelectObject(hdc, old_p);
                    let _ = DeleteObject(HGDIOBJ(fill_brush.0));
                    let _ = DeleteObject(HGDIOBJ(border_pen.0));

                    // Selected item label.
                    // **Drawn in the combo's own item box, not the client rect.** The item's own
                    // `WM_DRAWITEM` paint starts at the combo's inner margin (x=3), so measuring
                    // the inset from the client edge put this paint 3px left of that one — and
                    // the two swap in and out on mouse-over, which is the label snapping left.
                    let mut info = COMBOBOXINFO {
                        cbSize: std::mem::size_of::<COMBOBOXINFO>() as u32,
                        ..Default::default()
                    };
                    let item = if GetComboBoxInfo(hwnd, &mut info).is_ok()
                        && info.rcItem.right > info.rcItem.left
                    {
                        info.rcItem
                    } else {
                        rc
                    };
                    let sel = SendMessageW(hwnd, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                    if sel >= 0 {
                        let len =
                            SendMessageW(hwnd, CB_GETLBTEXTLEN, WPARAM(sel as usize), LPARAM(0))
                                .0
                                .max(0) as usize;
                        if len > 0 {
                            let mut buf = vec![0u16; len + 1];
                            SendMessageW(
                                hwnd,
                                CB_GETLBTEXT,
                                WPARAM(sel as usize),
                                LPARAM(buf.as_mut_ptr() as isize),
                            );
                            buf.truncate(len);
                            let mut text_rc = combo_label_rect(item, dpi);
                            SetBkMode(hdc, TRANSPARENT);
                            SetTextColor(hdc, if enabled { p.ink } else { p.muted });
                            let old_f = if !font.0.is_null() {
                                Some(SelectObject(hdc, HGDIOBJ(font.0)))
                            } else {
                                None
                            };
                            DrawTextW(
                                hdc,
                                &mut buf,
                                &mut text_rc,
                                DT_LEFT
                                    | DT_VCENTER
                                    | DT_SINGLELINE
                                    | DT_NOPREFIX
                                    | DT_END_ELLIPSIS,
                            );
                            if let Some(of) = old_f {
                                let _ = SelectObject(hdc, of);
                            }
                        }
                    }

                    // Custom chevron on right (replaces white Win32 dropdown button).
                    let chev_col = if !enabled {
                        p.muted
                    } else if active {
                        p.ink
                    } else {
                        p.muted
                    };
                    let cx = rc.right - scale(14, dpi);
                    let cy = (rc.top + rc.bottom) / 2;
                    let hw = scale(4, dpi).max(3);
                    let hh = scale(2, dpi).max(2);
                    let chev_pen = CreatePen(PS_SOLID, scale(1, dpi).max(1), chev_col);
                    let old_p = SelectObject(hdc, HGDIOBJ(chev_pen.0));
                    let _ = MoveToEx(hdc, cx - hw, cy - hh + 1, None);
                    let _ = LineTo(hdc, cx, cy + hh);
                    let _ = LineTo(hdc, cx + hw + 1, cy - hh);
                    let _ = SelectObject(hdc, old_p);
                    let _ = DeleteObject(HGDIOBJ(chev_pen.0));
                });
            }
            return LRESULT(0);
        }
        WM_MOUSEMOVE => {
            if !sub.hovered {
                sub.hovered = true;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
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
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
            }
        }
        WM_SETFOCUS | WM_KILLFOCUS | WM_ENABLE => {
            let res = unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) };
            let _ = unsafe { InvalidateRect(hwnd, None, false) };
            if msg == WM_KILLFOCUS {
                let _ = unsafe { PostMessageW(sub.parent, WM_AUTOSAVE, WPARAM(0), LPARAM(0)) };
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

unsafe fn install_list_subclass(list: HWND, parent: HWND) {
    let orig = unsafe {
        SetWindowLongPtrW(
            list,
            GWLP_WNDPROC,
            list_subclass_proc as *const () as usize as isize,
        )
    };
    let orig_proc: WNDPROC = unsafe { std::mem::transmute(orig) };
    let data = Box::into_raw(Box::new(ListSubclass {
        orig_proc,
        parent,
        press: None,
        dragging: false,
        posted: None,
        hovered_pill: None,
        tracking: false,
    }));
    unsafe { SetWindowLongPtrW(list, GWLP_USERDATA, data as isize) };
}

fn list_row_pill_rect(item_rect: RECT, dpi: u32) -> RECT {
    let dpi = dpi as i32;
    let pill_w = scale(32, dpi);
    let pill_h = scale(17, dpi);
    let right = item_rect.right - scale(8, dpi);
    let left = (right - pill_w).max(item_rect.left + scale(20, dpi));
    let top = (item_rect.top + item_rect.bottom - pill_h) / 2;
    let actual_w = (right - left).max(1);
    RECT {
        left,
        top,
        right: left + actual_w,
        bottom: top + pill_h,
    }
}

unsafe fn invalidate_list_row(list: HWND, row: usize) {
    let mut rc = RECT::default();
    let ok = unsafe {
        SendMessageW(
            list,
            LB_GETITEMRECT,
            WPARAM(row),
            LPARAM(&mut rc as *mut RECT as isize),
        )
        .0
    };
    if ok >= 0 {
        let _ = unsafe { InvalidateRect(list, Some(&rc), false) };
    }
}

unsafe fn list_pill_at_point(list: HWND, parent: HWND, lparam: LPARAM) -> Option<usize> {
    let count = unsafe { SendMessageW(list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0 };
    if count <= 0 {
        return None;
    }
    let (row, outside) = row_at_point(list, lparam);
    if outside || row >= count as usize {
        return None;
    }
    let mut rc = RECT::default();
    let ok = unsafe {
        SendMessageW(
            list,
            LB_GETITEMRECT,
            WPARAM(row),
            LPARAM(&mut rc as *mut RECT as isize),
        )
        .0
    };
    if ok < 0 {
        return None;
    }
    let dpi = unsafe { state_of(parent) }.map_or(96, |s| s.dpi) as u32;
    let pill = list_row_pill_rect(rc, dpi);
    let x = (lparam.0 & 0xFFFF) as i16 as i32;
    let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
    (x >= pill.left && x < pill.right && y >= pill.top && y < pill.bottom).then_some(row)
}

// Double-buffers LISTBOX client area & suppresses unbuffered WM_ERASEBKGND to eliminate resize bounce/flicker.
unsafe extern "system" fn list_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut ListSubclass;
    let Some(sub) = (unsafe { raw.as_mut() }) else {
        return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
    };
    let orig_proc = sub.orig_proc;
    let parent = sub.parent;

    match msg {
        WM_ERASEBKGND => LRESULT(1),
        WM_LBUTTONDOWN | WM_LBUTTONDBLCLK => {
            if let Some(row) = unsafe { list_pill_at_point(hwnd, parent, lparam) } {
                sub.press = None;
                sub.dragging = false;
                sub.posted = None;
                let _ =
                    unsafe { PostMessageW(parent, WM_PRESET_ROW_TOGGLE, WPARAM(row), LPARAM(0)) };
                return LRESULT(0);
            }
            let res = unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) };
            if msg == WM_LBUTTONDOWN {
                // The list's own selection runs first: a press on a row is one that may become a drag,
                // and a press below the last row drags nothing. No `SetCapture` here — the box takes
                // its own on a press, and a second one hands this proc the `WM_CAPTURECHANGED` of the
                // capture it just lost.
                let count = unsafe { SendMessageW(hwnd, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0 };
                let (row, outside) = row_at_point(hwnd, lparam);
                sub.press = (!outside && row < count as usize).then_some(row);
            }
            res
        }
        WM_MOUSEMOVE => {
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
            let res = unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) };
            let next_hover = if sub.dragging {
                None
            } else {
                unsafe { list_pill_at_point(hwnd, parent, lparam) }
            };
            if next_hover != sub.hovered_pill {
                let prev = sub.hovered_pill;
                sub.hovered_pill = next_hover;
                if let Some(r) = prev {
                    unsafe { invalidate_list_row(hwnd, r) };
                }
                if let Some(r) = next_hover {
                    unsafe { invalidate_list_row(hwnd, r) };
                }
            }
            let Some(press) = sub.press else {
                return res;
            };
            // A press whose button came up where this proc never saw it is no press at all: the
            // move message carries the button state, and a drag is only a drag while it is down.
            const MK_LBUTTON: usize = 0x0001;
            if wparam.0 & MK_LBUTTON == 0 {
                sub.press = None;
                sub.dragging = false;
                sub.posted = None;
                return res;
            }
            let count = unsafe { SendMessageW(hwnd, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0 };
            if count <= 0 {
                return res;
            }
            let (row, _) = row_at_point(hwnd, lparam);
            let row = row.min(count as usize - 1);
            if !sub.dragging {
                // A press still on its own row is a click, and it never becomes a move.
                if row == press {
                    return res;
                }
                sub.dragging = true;
                if let Some(prev) = sub.hovered_pill.take() {
                    unsafe { invalidate_list_row(hwnd, prev) };
                }
                let _ = unsafe { SetCursor(LoadCursorW(None, IDC_SIZEALL).unwrap_or_default()) };
            }
            if sub.posted == Some(row) {
                return res;
            }
            sub.posted = Some(row);
            // The reorder is the parent's, and it runs after this message has returned (see
            // `WM_PRESET_DRAG_MOVE`).
            let _ = unsafe {
                PostMessageW(
                    parent,
                    WM_PRESET_DRAG_MOVE,
                    WPARAM(row),
                    LPARAM(press as isize),
                )
            };
            res
        }
        WM_MOUSELEAVE => {
            sub.tracking = false;
            if let Some(prev) = sub.hovered_pill.take() {
                unsafe { invalidate_list_row(hwnd, prev) };
            }
            LRESULT(0)
        }
        WM_SETCURSOR => {
            if !sub.dragging && sub.hovered_pill.is_some() {
                let _ = unsafe { SetCursor(LoadCursorW(None, IDC_HAND).unwrap_or_default()) };
                return LRESULT(1);
            }
            unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) }
        }
        WM_LBUTTONUP => {
            let dragged = sub.dragging;
            sub.press = None;
            sub.dragging = false;
            sub.posted = None;
            if dragged {
                let _ = unsafe { SetCursor(LoadCursorW(None, IDC_ARROW).unwrap_or_default()) };
                let _ = unsafe { PostMessageW(parent, WM_PRESET_DRAG_END, WPARAM(0), LPARAM(0)) };
            }
            unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) }
        }
        WM_CAPTURECHANGED => {
            // The box captures the pointer on a press and gives it back on the release, and this
            // message is the capture it takes handing this proc the one that was there — which is
            // no reason to drop the press. Only a capture lost mid-drag ends the drag.
            if sub.dragging {
                sub.dragging = false;
                sub.posted = None;
                let _ = unsafe { PostMessageW(parent, WM_PRESET_DRAG_END, WPARAM(0), LPARAM(0)) };
            }
            unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) }
        }
        WM_PAINT => {
            unsafe {
                paint_double_buffered(hwnd, |hdc, rc| {
                    if let Some(state) = state_of(parent) {
                        FillRect(hdc, &rc, state.brush_field);
                    }
                    let _ = CallWindowProcW(
                        orig_proc,
                        hwnd,
                        WM_PRINTCLIENT,
                        WPARAM(hdc.0 as usize),
                        LPARAM(PRF_CLIENT),
                    );
                });
            }
            LRESULT(0)
        }
        WM_SIZE => {
            let res = unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) };
            let _ = unsafe { InvalidateRect(hwnd, None, false) };
            res
        }
        WM_NCDESTROY => unsafe {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            drop(Box::from_raw(raw));
            CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam)
        },
        _ => unsafe { CallWindowProcW(orig_proc, hwnd, msg, wparam, lparam) },
    }
}

/// Signed distance to circle centered at `(cx, cy)` with radius `r` (<0 inside, >0 outside).
fn sd_circle(px: f32, py: f32, cx: f32, cy: f32, r: f32) -> f32 {
    (px - cx).hypot(py - cy) - r
}

/// Signed distance to line segment `(ax, ay) -> (bx, by)`.
fn sd_segment(px: f32, py: f32, ax: f32, ay: f32, bx: f32, by: f32) -> f32 {
    let pax = px - ax;
    let pay = py - ay;
    let bax = bx - ax;
    let bay = by - ay;
    let denom = (bax * bax + bay * bay).max(1e-6);
    let h = ((pax * bax + pay * bay) / denom).clamp(0.0, 1.0);
    (pax - bax * h).hypot(pay - bay * h)
}

/// Signed distance to axis-aligned rounded box centered at `(cx, cy)` with half-extents `(hw, hh)` and corner radius `r`.
fn sd_rounded_box(px: f32, py: f32, cx: f32, cy: f32, hw: f32, hh: f32, r: f32) -> f32 {
    let r = r.min(hw).min(hh).max(0.0);
    let qx = (px - cx).abs() - (hw - r);
    let qy = (py - cy).abs() - (hh - r);
    qx.max(0.0).hypot(qy.max(0.0)) + qx.max(qy).min(0.0) - r
}

/// Smooth 1px anti-aliased coverage in `0.0..=1.0` from signed distance `d`.
fn aa_coverage(d: f32) -> f32 {
    (0.5 - d).clamp(0.0, 1.0)
}

fn colorref_rgb(c: COLORREF) -> (u8, u8, u8) {
    (
        (c.0 & 0xFF) as u8,
        ((c.0 >> 8) & 0xFF) as u8,
        ((c.0 >> 16) & 0xFF) as u8,
    )
}

fn pack_bgra(r: u8, g: u8, b: u8) -> u32 {
    u32::from(b) | (u32::from(g) << 8) | (u32::from(r) << 16) | 0xFF00_0000
}

/// 32-bit BGRA top-down DIB canvas for subpixel SDF anti-aliased controls.
struct AaSurface {
    w: i32,
    h: i32,
    pixels: Vec<u32>,
}

impl AaSurface {
    fn new(w: i32, h: i32, bg: COLORREF) -> Self {
        let w = w.max(1);
        let h = h.max(1);
        let (r, g, b) = colorref_rgb(bg);
        Self {
            w,
            h,
            pixels: vec![pack_bgra(r, g, b); (w as usize) * (h as usize)],
        }
    }

    #[inline]
    fn blend(&mut self, x: i32, y: i32, col: COLORREF, alpha: f32) {
        if alpha <= 0.0 || x < 0 || y < 0 || x >= self.w || y >= self.h {
            return;
        }
        let idx = (y as usize) * (self.w as usize) + (x as usize);
        let (sr, sg, sb) = colorref_rgb(col);
        if alpha >= 1.0 {
            self.pixels[idx] = pack_bgra(sr, sg, sb);
            return;
        }
        let dst = self.pixels[idx];
        let db = (dst & 0xFF) as f32;
        let dg = ((dst >> 8) & 0xFF) as f32;
        let dr = ((dst >> 16) & 0xFF) as f32;
        let inv = 1.0 - alpha;
        let r = (f32::from(sr) * alpha + dr * inv + 0.5) as u8;
        let g = (f32::from(sg) * alpha + dg * inv + 0.5) as u8;
        let b = (f32::from(sb) * alpha + db * inv + 0.5) as u8;
        self.pixels[idx] = pack_bgra(r, g, b);
    }

    fn fill_rounded_box(
        &mut self,
        bounds: [f32; 4],
        radius: f32,
        fill: COLORREF,
        border: Option<(COLORREF, f32)>,
    ) {
        let [x0, y0, x1, y1] = bounds;
        let cx = (x0 + x1) * 0.5;
        let cy = (y0 + y1) * 0.5;
        let hw = ((x1 - x0) * 0.5).max(0.0);
        let hh = ((y1 - y0) * 0.5).max(0.0);
        let ix0 = (x0.floor() as i32 - 1).max(0);
        let iy0 = (y0.floor() as i32 - 1).max(0);
        let ix1 = (x1.ceil() as i32 + 1).min(self.w);
        let iy1 = (y1.ceil() as i32 + 1).min(self.h);

        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            for x in ix0..ix1 {
                let px = x as f32 + 0.5;
                let d = sd_rounded_box(px, py, cx, cy, hw, hh, radius);
                let outer = aa_coverage(d);
                if outer <= 0.0 {
                    continue;
                }
                if let Some((bcol, bw)) = border {
                    self.blend(x, y, bcol, outer);
                    let inner = aa_coverage(d + bw);
                    if inner > 0.0 {
                        self.blend(x, y, fill, inner);
                    }
                } else {
                    self.blend(x, y, fill, outer);
                }
            }
        }
    }

    fn fill_soft_circle(&mut self, cx: f32, cy: f32, r: f32, col: COLORREF, max_alpha: f32) {
        let ix0 = ((cx - r).floor() as i32 - 1).max(0);
        let iy0 = ((cy - r).floor() as i32 - 1).max(0);
        let ix1 = ((cx + r).ceil() as i32 + 1).min(self.w);
        let iy1 = ((cy + r).ceil() as i32 + 1).min(self.h);
        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            for x in ix0..ix1 {
                let px = x as f32 + 0.5;
                let dist = (px - cx).hypot(py - cy);
                if dist < r {
                    let t = 1.0 - (dist / r);
                    self.blend(x, y, col, t * t * max_alpha);
                }
            }
        }
    }

    fn fill_circle(
        &mut self,
        cx: f32,
        cy: f32,
        r: f32,
        fill: COLORREF,
        border: Option<(COLORREF, f32)>,
        shadow_alpha: f32,
    ) {
        if shadow_alpha > 0.0 {
            self.fill_soft_circle(cx, cy + 1.0, r + 2.2, rgb(0, 0, 0), shadow_alpha);
        }
        let ix0 = ((cx - r).floor() as i32 - 1).max(0);
        let iy0 = ((cy - r).floor() as i32 - 1).max(0);
        let ix1 = ((cx + r).ceil() as i32 + 1).min(self.w);
        let iy1 = ((cy + r).ceil() as i32 + 1).min(self.h);

        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            for x in ix0..ix1 {
                let px = x as f32 + 0.5;
                let d = sd_circle(px, py, cx, cy, r);
                let outer = aa_coverage(d);
                if outer <= 0.0 {
                    continue;
                }
                if let Some((bcol, bw)) = border {
                    self.blend(x, y, bcol, outer);
                    let inner = aa_coverage(d + bw);
                    if inner > 0.0 {
                        self.blend(x, y, fill, inner);
                    }
                } else {
                    self.blend(x, y, fill, outer);
                }
            }
        }
    }

    fn fill_rounded_box_gradient(
        &mut self,
        bounds: [f32; 4],
        radius: f32,
        top_col: COLORREF,
        bottom_col: COLORREF,
        border: Option<(COLORREF, f32)>,
    ) {
        let [x0, y0, x1, y1] = bounds;
        let cx = (x0 + x1) * 0.5;
        let cy = (y0 + y1) * 0.5;
        let hw = ((x1 - x0) * 0.5).max(0.0);
        let hh = ((y1 - y0) * 0.5).max(0.0);
        let span_y = (y1 - y0).max(1.0);
        let ix0 = (x0.floor() as i32 - 1).max(0);
        let iy0 = (y0.floor() as i32 - 1).max(0);
        let ix1 = (x1.ceil() as i32 + 1).min(self.w);
        let iy1 = (y1.ceil() as i32 + 1).min(self.h);

        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            let t = ((py - y0) / span_y).clamp(0.0, 1.0);
            let fill = mix(top_col, bottom_col, (t * 100.0 + 0.5) as u32);
            for x in ix0..ix1 {
                let px = x as f32 + 0.5;
                let d = sd_rounded_box(px, py, cx, cy, hw, hh, radius);
                let outer = aa_coverage(d);
                if outer <= 0.0 {
                    continue;
                }
                if let Some((bcol, bw)) = border {
                    self.blend(x, y, bcol, outer);
                    let inner = aa_coverage(d + bw);
                    if inner > 0.0 {
                        self.blend(x, y, fill, inner);
                    }
                } else {
                    self.blend(x, y, fill, outer);
                }
            }
        }
    }

    #[cfg(test)]
    fn fill_elevated_card(&mut self, rect: RECT, radius: f32, p: Palette, dpi: i32) {
        if rect.right <= rect.left || rect.bottom <= rect.top {
            return;
        }
        let x0 = rect.left as f32;
        let y0 = rect.top as f32;
        let x1 = rect.right as f32;
        let y1 = rect.bottom as f32;
        let cx = (x0 + x1) * 0.5;
        let cy = (y0 + y1) * 0.5;
        let hw = (x1 - x0) * 0.5;
        let hh = (y1 - y0) * 0.5;

        let shadow_y = scale(2, dpi) as f32;
        let shadow_spread = scale(6, dpi).max(4) as f32;
        let shadow_alpha = if p.dark { 0.18_f32 } else { 0.07_f32 };
        let sheen_alpha = if p.dark { 0.06_f32 } else { 0.65_f32 };

        let sx0 = (x0 - shadow_spread).floor().max(0.0) as i32;
        let sy0 = (y0 - 2.0).floor().max(0.0) as i32;
        let sx1 = (x1 + shadow_spread).ceil().min(self.w as f32) as i32;
        let sy1 = (y1 + shadow_y + shadow_spread).ceil().min(self.h as f32) as i32;

        let (cr, cg, cb) = colorref_rgb(p.card);
        let card_bgra = pack_bgra(cr, cg, cb);
        let corner_band = (radius.ceil() as i32 + 2).max(4);
        let inner_y0 = (rect.top + corner_band).clamp(sy0, sy1);
        let inner_y1 = (rect.bottom - corner_band).clamp(inner_y0, sy1);
        let inner_x0 = (rect.left + corner_band).clamp(sx0, sx1);
        let inner_x1 = (rect.right - corner_band).clamp(inner_x0, sx1);

        for y in sy0..sy1 {
            let py = y as f32 + 0.5;
            let in_mid_y = y >= inner_y0 && y < inner_y1 && inner_x1 > inner_x0;
            let x_ranges = if in_mid_y {
                let row_off = (y as usize) * (self.w as usize);
                self.pixels[row_off + inner_x0 as usize..row_off + inner_x1 as usize]
                    .fill(card_bgra);
                [(sx0, inner_x0), (inner_x1, sx1)]
            } else {
                [(sx0, sx1), (sx1, sx1)]
            };
            for (rx0, rx1) in x_ranges {
                for x in rx0..rx1 {
                    let px = x as f32 + 0.5;
                    let d_card = sd_rounded_box(px, py, cx, cy, hw, hh, radius);
                    if d_card > 0.0 {
                        let d_sh = sd_rounded_box(px, py, cx, cy + shadow_y, hw, hh, radius);
                        if d_sh < shadow_spread {
                            let t = (1.0 - (d_sh.max(0.0) / shadow_spread)).clamp(0.0, 1.0);
                            self.blend(x, y, rgb(0, 0, 0), t * t * shadow_alpha);
                        }
                    }
                    let outer = aa_coverage(d_card);
                    if outer > 0.0 {
                        self.blend(x, y, p.border, outer);
                        let inner = aa_coverage(d_card + 1.0);
                        if inner > 0.0 {
                            self.blend(x, y, p.card, inner);
                            // Subtle 1px top-edge inner sheen inside card boundary.
                            if py >= y0 + 1.0 && py < y0 + 2.2 {
                                self.blend(x, y, rgb(255, 255, 255), inner * sheen_alpha);
                            }
                        }
                    }
                }
            }
        }
    }

    unsafe fn create_dib(&self, dst_dc: HDC) -> Option<HBITMAP> {
        unsafe { create_bgra_dib(dst_dc, self.w, self.h, &self.pixels) }
    }

    unsafe fn blit_to(&self, dst_dc: HDC, dst_x: i32, dst_y: i32) {
        let Some(dib) = (unsafe { self.create_dib(dst_dc) }) else {
            return;
        };
        let src_dc = unsafe { CreateCompatibleDC(dst_dc) };
        if !src_dc.is_invalid() {
            let old = unsafe { SelectObject(src_dc, HGDIOBJ(dib.0)) };
            let _ = unsafe { BitBlt(dst_dc, dst_x, dst_y, self.w, self.h, src_dc, 0, 0, SRCCOPY) };
            let _ = unsafe { SelectObject(src_dc, old) };
            let _ = unsafe { DeleteDC(src_dc) };
        }
        let _ = unsafe { DeleteObject(HGDIOBJ(dib.0)) };
    }
}

unsafe fn create_bgra_dib(dst_dc: HDC, w: i32, h: i32, pixels: &[u32]) -> Option<HBITMAP> {
    let bmi = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: w,
            biHeight: -h,
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0,
            ..Default::default()
        },
        ..Default::default()
    };
    let mut bits: *mut std::ffi::c_void = std::ptr::null_mut();
    let dib = unsafe { CreateDIBSection(dst_dc, &bmi, DIB_RGB_COLORS, &mut bits, None, 0) }.ok()?;
    if bits.is_null() {
        let _ = unsafe { DeleteObject(HGDIOBJ(dib.0)) };
        return None;
    }
    unsafe {
        std::ptr::copy_nonoverlapping(pixels.as_ptr(), bits as *mut u32, pixels.len());
    }
    Some(dib)
}

/// Renders the atmospheric diagonal gradient + soft radial glows into a compact `w x h` surface.
fn render_gradient_backdrop(w: i32, h: i32, p: Palette) -> AaSurface {
    let w = w.max(1);
    let h = h.max(1);
    let mut surface = AaSurface::new(w, h, p.page);

    let (top_col, bot_col) = if p.dark {
        (rgb(0x21, 0x25, 0x2B), p.page)
    } else {
        (rgb(0xF2, 0xF5, 0xFA), p.page)
    };
    let (tr, tg, tb) = colorref_rgb(top_col);
    let (br, bg, bb) = colorref_rgb(bot_col);
    let (ar, ag, ab) = colorref_rgb(p.accent);
    let (nr, ng, nb) = colorref_rgb(p.number);

    let fw = (w - 1).max(1) as f32;
    let fh = (h - 1).max(1) as f32;
    let max_dim = (w.max(h) as f32).max(1.0);
    let glow_cx = fw * 0.88;
    let glow_cy = fh * 0.14;
    let glow_r = max_dim * 0.55;
    let glow_alpha = if p.dark { 0.06_f32 } else { 0.05_f32 };
    let warm_cx = fw * 0.12;
    let warm_cy = fh * 0.88;
    let warm_r = max_dim * 0.50;
    let warm_alpha = if p.dark { 0.03_f32 } else { 0.0_f32 };

    for y in 0..h {
        let fy = y as f32;
        let ty = fy / fh;
        let row_off = (y as usize) * (w as usize);
        for x in 0..w {
            let fx = x as f32;
            let t = (ty * 0.72 + (fx / fw) * 0.28).clamp(0.0, 1.0);
            let mut r = f32::from(tr) + (f32::from(br) - f32::from(tr)) * t;
            let mut g = f32::from(tg) + (f32::from(bg) - f32::from(tg)) * t;
            let mut b = f32::from(tb) + (f32::from(bb) - f32::from(tb)) * t;

            let gd = (fx - glow_cx).hypot(fy - glow_cy);
            if gd < glow_r {
                let k = 1.0 - (gd / glow_r);
                let a = k * k * glow_alpha;
                r += (f32::from(ar) - r) * a;
                g += (f32::from(ag) - g) * a;
                b += (f32::from(ab) - b) * a;
            }
            if warm_alpha > 0.0 {
                let wd = (fx - warm_cx).hypot(fy - warm_cy);
                if wd < warm_r {
                    let k = 1.0 - (wd / warm_r);
                    let a = k * k * warm_alpha;
                    r += (f32::from(nr) - r) * a;
                    g += (f32::from(ng) - g) * a;
                    b += (f32::from(nb) - b) * a;
                }
            }
            surface.pixels[row_off + x as usize] = pack_bgra(
                r.clamp(0.0, 255.0) as u8,
                g.clamp(0.0, 255.0) as u8,
                b.clamp(0.0, 255.0) as u8,
            );
        }
    }
    surface
}

/// Renders a tiny premultiplied-BGRA 9-slice card template (corners + border + top sheen + shadow).
fn render_card_9slice_pixels(
    p: Palette,
    dpi: i32,
) -> (Vec<u32>, i32, i32, i32, i32, i32, i32, i32) {
    let radius = scale(9, dpi) as f32;
    let shadow_y = scale(2, dpi) as f32;
    let shadow_spread = scale(6, dpi).max(4) as f32;
    let shadow_alpha = if p.dark { 0.18_f32 } else { 0.07_f32 };
    let sheen_alpha = if p.dark { 0.06_f32 } else { 0.65_f32 };

    let pad_l = shadow_spread.ceil() as i32;
    let pad_t = 2;
    let pad_r = shadow_spread.ceil() as i32;
    let pad_b = (shadow_y + shadow_spread).ceil() as i32;
    let corner = radius.ceil() as i32 + 2;
    let mid = 4;
    let card_w = corner * 2 + mid;
    let card_h = corner * 2 + mid;
    let tex_w = pad_l + card_w + pad_r;
    let tex_h = pad_t + card_h + pad_b;

    let x0 = pad_l as f32;
    let y0 = pad_t as f32;
    let x1 = (pad_l + card_w) as f32;
    let y1 = (pad_t + card_h) as f32;
    let cx = (x0 + x1) * 0.5;
    let cy = (y0 + y1) * 0.5;
    let hw = (x1 - x0) * 0.5;
    let hh = (y1 - y0) * 0.5;

    let (br, bg, bb) = colorref_rgb(p.border);
    let (cr, cg, cb) = colorref_rgb(p.card);
    let mut pixels = vec![0u32; (tex_w * tex_h) as usize];

    for y in 0..tex_h {
        let py = y as f32 + 0.5;
        let row_off = (y as usize) * (tex_w as usize);
        for x in 0..tex_w {
            let px = x as f32 + 0.5;
            let d_card = sd_rounded_box(px, py, cx, cy, hw, hh, radius);
            // Premultiplied straight-over accumulator (r, g, b in 0..255, a in 0..1).
            let mut pr = 0.0_f32;
            let mut pg = 0.0_f32;
            let mut pb = 0.0_f32;
            let mut pa = 0.0_f32;

            if d_card > 0.0 {
                let d_sh = sd_rounded_box(px, py, cx, cy + shadow_y, hw, hh, radius);
                if d_sh < shadow_spread {
                    let t = (1.0 - (d_sh.max(0.0) / shadow_spread)).clamp(0.0, 1.0);
                    pa = t * t * shadow_alpha;
                }
            }
            let outer = aa_coverage(d_card);
            if outer > 0.0 {
                let inv = 1.0 - outer;
                pr = f32::from(br) * outer + pr * inv;
                pg = f32::from(bg) * outer + pg * inv;
                pb = f32::from(bb) * outer + pb * inv;
                pa = outer + pa * inv;

                let inner = aa_coverage(d_card + 1.0);
                if inner > 0.0 {
                    let inv_i = 1.0 - inner;
                    pr = f32::from(cr) * inner + pr * inv_i;
                    pg = f32::from(cg) * inner + pg * inv_i;
                    pb = f32::from(cb) * inner + pb * inv_i;
                    pa = inner + pa * inv_i;

                    if py >= y0 + 1.0 && py < y0 + 2.2 {
                        let sa = inner * sheen_alpha;
                        let inv_s = 1.0 - sa;
                        pr = 255.0 * sa + pr * inv_s;
                        pg = 255.0 * sa + pg * inv_s;
                        pb = 255.0 * sa + pb * inv_s;
                        pa = sa + pa * inv_s;
                    }
                }
            }
            let a8 = (pa.clamp(0.0, 1.0) * 255.0 + 0.5) as u8;
            let r8 = pr.clamp(0.0, f32::from(a8)) as u8;
            let g8 = pg.clamp(0.0, f32::from(a8)) as u8;
            let b8 = pb.clamp(0.0, f32::from(a8)) as u8;
            pixels[row_off + x as usize] = (u32::from(a8) << 24)
                | (u32::from(r8) << 16)
                | (u32::from(g8) << 8)
                | u32::from(b8);
        }
    }
    (pixels, tex_w, tex_h, pad_l, pad_t, pad_r, pad_b, corner)
}

/// Theme colors `(fill, border, text)` for a preset list row's `"On"` / `"Off"` status pill.
fn status_pill_colors(p: Palette, enabled: bool, hovered: bool) -> (COLORREF, COLORREF, COLORREF) {
    if enabled {
        let fill_share = match (p.dark, hovered) {
            (true, true) => 26,
            (true, false) => 16,
            (false, true) => 20,
            (false, false) => 12,
        };
        let fill = mix(p.field, p.accent, fill_share);
        let border = if hovered { p.accent_hover } else { p.accent };
        let text = if p.dark {
            mix(p.accent, rgb(255, 255, 255), 42)
        } else {
            mix(p.accent, rgb(0, 0, 0), 34)
        };
        (fill, border, text)
    } else {
        let fill = if p.dark {
            if hovered {
                mix(p.field, p.button_hover, 65)
            } else {
                mix(p.field, p.button_bg, 45)
            }
        } else if hovered {
            p.button_hover
        } else {
            p.button_bg
        };
        let border = if hovered {
            mix(p.border, p.ink, 25)
        } else {
            p.border
        };
        let text = if hovered {
            p.ink
        } else {
            mix(p.muted, p.ink, 45)
        };
        (fill, border, text)
    }
}

/// Renders an SDF anti-aliased rounded status pill badge (`"On"` / `"Off"` on preset list rows).
fn render_status_badge_pill(
    w: i32,
    h: i32,
    bg: COLORREF,
    fill: COLORREF,
    border: COLORREF,
) -> AaSurface {
    let w = w.max(1);
    let h = h.max(1);
    let mut surface = AaSurface::new(w, h, bg);
    let radius = ((h as f32) - 1.0).max(1.0) * 0.5;
    surface.fill_rounded_box(
        [0.5, 0.5, w as f32 - 0.5, h as f32 - 0.5],
        radius,
        fill,
        Some((border, 1.0)),
    );
    surface
}

/// Renders Lucide `search` icon (24x24 viewBox, stroke-width 2, round caps):
/// ring at `(11, 11)` r=8 + handle `(16.65, 16.65) -> (21, 21)`.
fn render_search_icon(size: i32, bg: COLORREF, fg: COLORREF) -> AaSurface {
    let size = size.max(1);
    let mut surface = AaSurface::new(size, size, bg);
    let s = size as f32 / 24.0;
    let half_stroke = s;
    for y in 0..size {
        let py = y as f32 + 0.5;
        for x in 0..size {
            let px = x as f32 + 0.5;
            let d_ring = sd_circle(px, py, 11.0 * s, 11.0 * s, 8.0 * s).abs() - half_stroke;
            let d_handle =
                sd_segment(px, py, 16.65 * s, 16.65 * s, 21.0 * s, 21.0 * s) - half_stroke;
            let cov = aa_coverage(d_ring.min(d_handle));
            if cov > 0.0 {
                surface.blend(x, y, fg, cov);
            }
        }
    }
    surface
}

/// Renders the atmospheric window backdrop: diagonal theme gradient, subtle radial accent glow,
/// titlebar band + crisp divider, and SDF elevated rounded cards (no top neon bar or accent dot).
#[cfg(test)]
fn render_backdrop(
    w: i32,
    h: i32,
    p: Palette,
    dpi: i32,
    titlebar_bottom: i32,
    sidebar_card: RECT,
    editor_card: RECT,
) -> AaSurface {
    let w = w.max(1);
    let h = h.max(1);
    let mut surface = render_gradient_backdrop(w, h, p);

    let tb_end = titlebar_bottom.clamp(0, h);
    for y in 0..tb_end.saturating_sub(1) {
        for x in 0..w {
            surface.blend(x, y, p.titlebar, 0.88);
        }
    }
    if tb_end > 0 {
        let div_y = tb_end - 1;
        for x in 0..w {
            surface.blend(x, div_y, p.border, 1.0);
        }
    }

    // SDF anti-aliased rounded cards with soft drop shadow & top sheen.
    let card_radius = scale(9, dpi) as f32;
    surface.fill_elevated_card(sidebar_card, card_radius, p, dpi);
    surface.fill_elevated_card(editor_card, card_radius, p, dpi);

    surface
}

fn slider_value_from_x(x: i32, track_left: i32, track_right: i32) -> u32 {
    let span = (track_right - track_left).max(1);
    let clamped = (x - track_left).clamp(0, span);
    ((clamped as i64 * VOLUME_MAX as i64 + span as i64 / 2) / span as i64)
        .clamp(VOLUME_MIN as i64, VOLUME_MAX as i64) as u32
}

fn slider_thumb_x(value: u32, track_left: i32, track_right: i32) -> i32 {
    let span = (track_right - track_left).max(1);
    let v = value.clamp(VOLUME_MIN, VOLUME_MAX) as i64;
    track_left + ((v * span as i64 + VOLUME_MAX as i64 / 2) / VOLUME_MAX as i64) as i32
}

unsafe fn create_themed_slider(parent: HWND, id: i32) -> HWND {
    SLIDER_CLASS_REGISTERED.call_once(|| {
        let Ok(hinstance) = (unsafe { GetModuleHandleW(None) }) else {
            return;
        };
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: WNDCLASS_STYLES(0),
            lpfnWndProc: Some(slider_proc),
            hInstance: hinstance.into(),
            hCursor: unsafe { LoadCursorW(None, IDC_ARROW) }.unwrap_or_default(),
            hbrBackground: HBRUSH(unsafe { GetStockObject(NULL_BRUSH) }.0),
            lpszClassName: SLIDER_CLASS,
            ..Default::default()
        };
        unsafe { RegisterClassExW(&class) };
    });

    let slider = unsafe { child(parent, SLIDER_CLASS, "", child_style(WS_TABSTOP.0), id) };
    if !slider.0.is_null() {
        let data = Box::into_raw(Box::new(SliderState {
            parent,
            value: 100,
            dragging: false,
            hovered: false,
            tracking: false,
        }));
        unsafe { SetWindowLongPtrW(slider, GWLP_USERDATA, data as isize) };
    }
    slider
}

unsafe extern "system" fn slider_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut SliderState;
    let Some(sl) = (unsafe { raw.as_mut() }) else {
        return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
    };

    let dpi = unsafe { state_of(sl.parent) }.map_or(96, |s| s.dpi);
    let track_bounds = |hwnd: HWND, dpi: i32| -> (RECT, i32, i32) {
        let mut rc = RECT::default();
        let _ = unsafe { GetClientRect(hwnd, &mut rc) };
        let pad = scale(11, dpi);
        let left = pad;
        let right = (rc.right - pad).max(left + 1);
        (rc, left, right)
    };
    let notify_parent = |sl: &mut SliderState, hwnd: HWND, next: u32| {
        let clamped = next.clamp(VOLUME_MIN, VOLUME_MAX);
        if sl.value != clamped {
            sl.value = clamped;
            let _ = unsafe { InvalidateRect(hwnd, None, false) };
            unsafe {
                SendMessageW(sl.parent, WM_HSCROLL, WPARAM(0), LPARAM(hwnd.0 as isize));
            }
        }
    };

    match msg {
        WM_SLIDER_SET_VALUE => {
            let next = (wparam.0 as u32).clamp(VOLUME_MIN, VOLUME_MAX);
            if sl.value != next {
                sl.value = next;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
            }
            LRESULT(0)
        }
        WM_SLIDER_GET_VALUE => LRESULT(sl.value as isize),
        WM_ERASEBKGND => LRESULT(1),
        WM_PAINT => {
            let p = match unsafe { state_of(sl.parent) } {
                Some(state) => state.palette,
                None => Palette::of(theme_now()),
            };
            let (_, track_left, track_right) = track_bounds(hwnd, dpi);
            let enabled = unsafe { IsWindowEnabled(hwnd) }.as_bool();
            let focused = enabled && unsafe { GetFocus() } == hwnd;
            let dpi_s = dpi as f32 / 96.0;
            let track_h = (4.0 * dpi_s).max(3.0);
            let thumb_x = slider_thumb_x(sl.value, track_left, track_right) as f32;
            let base_r = (6.0 * dpi_s).max(5.0);
            let thumb_r = if enabled && (sl.dragging || sl.hovered) {
                base_r + 1.0 * dpi_s
            } else {
                base_r
            };

            unsafe {
                paint_double_buffered(hwnd, |hdc, rc| {
                    let w = (rc.right - rc.left).max(1);
                    let h = (rc.bottom - rc.top).max(1);
                    let cy = h as f32 * 0.5;
                    let half_h = track_h * 0.5;
                    let mut surface = AaSurface::new(w, h, p.card);

                    // Unfilled pill rail.
                    let track_col = if enabled { p.slider_track } else { p.border };
                    surface.fill_rounded_box(
                        [
                            track_left as f32,
                            cy - half_h,
                            track_right as f32,
                            cy + half_h,
                        ],
                        half_h,
                        track_col,
                        None,
                    );

                    // Active filled rail (0..value).
                    if thumb_x > track_left as f32 {
                        let active_col = if enabled { p.accent } else { p.muted };
                        surface.fill_rounded_box(
                            [
                                track_left as f32,
                                cy - half_h,
                                thumb_x.min(track_right as f32),
                                cy + half_h,
                            ],
                            half_h,
                            active_col,
                            None,
                        );
                    }

                    // Soft hover / drag / focus halo behind circular thumb.
                    if enabled && (sl.dragging || sl.hovered || focused) {
                        let halo_col = if focused && !sl.hovered && !sl.dragging {
                            p.ink
                        } else {
                            p.accent
                        };
                        let halo_alpha = if sl.dragging { 0.30 } else { 0.22 };
                        surface.fill_soft_circle(
                            thumb_x,
                            cy,
                            thumb_r + 4.5 * dpi_s,
                            halo_col,
                            halo_alpha,
                        );
                    }

                    // Anti-aliased circular thumb with subtle shadow & crisp ring.
                    let thumb_fill = if !enabled {
                        p.muted
                    } else if sl.dragging || sl.hovered {
                        p.accent_hover
                    } else {
                        p.accent
                    };
                    let thumb_border = if focused { p.ink } else { p.card };
                    let shadow_alpha = if !enabled {
                        0.0
                    } else if p.dark {
                        0.28
                    } else {
                        0.18
                    };
                    surface.fill_circle(
                        thumb_x,
                        cy,
                        thumb_r,
                        thumb_fill,
                        Some((thumb_border, (1.25 * dpi_s).max(1.0))),
                        shadow_alpha,
                    );
                    surface.blit_to(hdc, 0, 0);
                });
            }
            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            if !unsafe { IsWindowEnabled(hwnd) }.as_bool() {
                return LRESULT(0);
            }
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            let (_, track_left, track_right) = track_bounds(hwnd, dpi);
            sl.dragging = true;
            unsafe {
                let _ = SetFocus(hwnd);
                SetCapture(hwnd);
            }
            let next = slider_value_from_x(x, track_left, track_right);
            notify_parent(sl, hwnd, next);
            let _ = unsafe { InvalidateRect(hwnd, None, false) };
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            if sl.dragging {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let (_, track_left, track_right) = track_bounds(hwnd, dpi);
                let next = slider_value_from_x(x, track_left, track_right);
                notify_parent(sl, hwnd, next);
            } else {
                if !sl.hovered {
                    sl.hovered = true;
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
                if !sl.tracking {
                    sl.tracking = true;
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
        WM_LBUTTONUP => {
            if sl.dragging {
                sl.dragging = false;
                let _ = unsafe { ReleaseCapture() };
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
                // The thumb let go is the end of the drag, and the level it was left at is the
                // one the user meant (see `autosave`).
                let _ = unsafe { PostMessageW(sl.parent, WM_AUTOSAVE, WPARAM(0), LPARAM(0)) };
            }
            LRESULT(0)
        }
        WM_CAPTURECHANGED | WM_CANCELMODE => {
            if sl.dragging {
                sl.dragging = false;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
                let _ = unsafe { PostMessageW(sl.parent, WM_AUTOSAVE, WPARAM(0), LPARAM(0)) };
            }
            LRESULT(0)
        }
        WM_MOUSELEAVE => {
            sl.tracking = false;
            if sl.hovered {
                sl.hovered = false;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
            }
            LRESULT(0)
        }
        WM_MOUSEWHEEL => {
            if !unsafe { IsWindowEnabled(hwnd) }.as_bool() {
                return LRESULT(0);
            }
            let delta = ((wparam.0 >> 16) as i16) as i32;
            let step = if delta > 0 { 5 } else { -5 };
            notify_parent(sl, hwnd, step_volume(sl.value, step));
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if !unsafe { IsWindowEnabled(hwnd) }.as_bool() {
                return LRESULT(0);
            }
            let vk = wparam.0 as u16;
            let next = if vk == VK_LEFT.0 || vk == VK_DOWN.0 {
                Some(step_volume(sl.value, -1))
            } else if vk == VK_RIGHT.0 || vk == VK_UP.0 {
                Some(step_volume(sl.value, 1))
            } else if vk == VK_HOME.0 {
                Some(VOLUME_MIN)
            } else if vk == VK_END.0 {
                Some(VOLUME_MAX)
            } else {
                None
            };
            if let Some(v) = next {
                notify_parent(sl, hwnd, v);
                return LRESULT(0);
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_ENABLE | WM_SETFOCUS | WM_KILLFOCUS => {
            let _ = unsafe { InvalidateRect(hwnd, None, false) };
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_NCDESTROY => {
            unsafe {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                drop(Box::from_raw(raw));
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

unsafe fn create_themed_toggle(
    parent: HWND,
    id: i32,
    show_label: bool,
    initial_checked: bool,
) -> HWND {
    TOGGLE_CLASS_REGISTERED.call_once(|| {
        let Ok(hinstance) = (unsafe { GetModuleHandleW(None) }) else {
            return;
        };
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: WNDCLASS_STYLES(0),
            lpfnWndProc: Some(toggle_proc),
            hInstance: hinstance.into(),
            hCursor: unsafe { LoadCursorW(None, IDC_ARROW) }.unwrap_or_default(),
            hbrBackground: HBRUSH(unsafe { GetStockObject(NULL_BRUSH) }.0),
            lpszClassName: TOGGLE_CLASS,
            ..Default::default()
        };
        unsafe { RegisterClassExW(&class) };
    });

    let toggle = unsafe { child(parent, TOGGLE_CLASS, "", child_style(WS_TABSTOP.0), id) };
    if !toggle.0.is_null() {
        let data = Box::into_raw(Box::new(ToggleState {
            parent,
            id,
            show_label,
            checked: initial_checked,
            hovered: false,
            tracking: false,
        }));
        unsafe { SetWindowLongPtrW(toggle, GWLP_USERDATA, data as isize) };
    }
    toggle
}

unsafe extern "system" fn toggle_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut ToggleState;
    let Some(tg) = (unsafe { raw.as_mut() }) else {
        return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
    };

    let flip = |tg: &mut ToggleState, hwnd: HWND| {
        tg.checked = !tg.checked;
        let _ = unsafe { InvalidateRect(hwnd, None, false) };
        unsafe {
            SendMessageW(
                tg.parent,
                WM_TOGGLE_CHANGED,
                WPARAM(usize::from(tg.checked)),
                LPARAM(tg.id as isize),
            );
            // A switch has no half-typed state: the flip is the whole of what the user meant
            // (see `autosave`).
            let _ = PostMessageW(tg.parent, WM_AUTOSAVE, WPARAM(0), LPARAM(0));
        }
    };

    match msg {
        WM_TOGGLE_SET_VALUE => {
            let next = wparam.0 != 0;
            if tg.checked != next {
                tg.checked = next;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
            }
            LRESULT(0)
        }
        WM_TOGGLE_GET_VALUE => LRESULT(isize::from(tg.checked)),
        WM_ERASEBKGND => LRESULT(1),
        WM_PAINT => {
            let (p, dpi, font, card_brush) = match unsafe { state_of(tg.parent) } {
                Some(state) => (state.palette, state.dpi, state.font, Some(state.brush_card)),
                None => (Palette::of(theme_now()), 96, HFONT::default(), None),
            };
            let enabled = unsafe { IsWindowEnabled(hwnd) }.as_bool();
            let focused = enabled && unsafe { GetFocus() } == hwnd;
            let checked = tg.checked;
            let hovered = enabled && tg.hovered;

            unsafe {
                paint_double_buffered(hwnd, |hdc, rc| {
                    if let Some(cb) = card_brush {
                        FillRect(hdc, &rc, cb);
                    } else {
                        let b = CreateSolidBrush(p.card);
                        FillRect(hdc, &rc, b);
                        let _ = DeleteObject(HGDIOBJ(b.0));
                    }

                    let dpi_s = dpi as f32 / 96.0;
                    let track_w = 38.0 * dpi_s;
                    let track_h = 20.0 * dpi_s;
                    let inset = 2.0;
                    let dib_w = (track_w + inset * 2.0).ceil() as i32;
                    let dib_h = (rc.bottom - rc.top).max(1);
                    let track_top = (dib_h as f32 - track_h) * 0.5;
                    let track_bot = track_top + track_h;
                    let track_left = inset;
                    let track_right = track_left + track_w;

                    let (track_fill, track_border) = if !enabled {
                        (p.button_disabled, p.border)
                    } else if checked {
                        let fill = if hovered { p.accent_hover } else { p.accent };
                        (fill, if focused { p.ink } else { fill })
                    } else {
                        let fill = if hovered {
                            p.button_hover
                        } else {
                            p.slider_track
                        };
                        (fill, if focused { p.accent } else { p.border })
                    };

                    let mut surface = AaSurface::new(dib_w, dib_h, p.card);
                    let border_w = (1.0 * dpi_s).max(1.0);
                    surface.fill_rounded_box(
                        [track_left, track_top, track_right, track_bot],
                        track_h * 0.5,
                        track_fill,
                        Some((track_border, border_w)),
                    );

                    // Anti-aliased circular knob inside pill with subtle drop shadow.
                    let pad = (3.0 * dpi_s).max(2.0);
                    let thumb_r = ((track_h - pad * 2.0) * 0.5).max(4.0);
                    let thumb_cx = if checked {
                        track_right - pad - thumb_r
                    } else {
                        track_left + pad + thumb_r
                    };
                    let thumb_cy = (track_top + track_bot) * 0.5;
                    let thumb_fill = if !enabled {
                        p.muted
                    } else if checked {
                        rgb(255, 255, 255)
                    } else {
                        p.ink
                    };
                    let shadow_alpha = if !enabled {
                        0.0
                    } else if p.dark {
                        0.26
                    } else {
                        0.16
                    };
                    surface.fill_circle(
                        thumb_cx,
                        thumb_cy,
                        thumb_r,
                        thumb_fill,
                        None,
                        shadow_alpha,
                    );
                    surface.blit_to(hdc, rc.left, rc.top);

                    if tg.show_label {
                        // Status label ("Enabled" / "Disabled").
                        let label = if checked { "Enabled" } else { "Disabled" };
                        let mut label_w: Vec<u16> = label.encode_utf16().collect();
                        let mut text_rc = RECT {
                            left: rc.left + dib_w + scale(8, dpi),
                            top: rc.top,
                            right: rc.right,
                            bottom: rc.bottom,
                        };
                        SetBkMode(hdc, TRANSPARENT);
                        SetTextColor(
                            hdc,
                            if !enabled || !checked {
                                p.muted
                            } else {
                                p.label
                            },
                        );
                        let old_f = if !font.0.is_null() {
                            Some(SelectObject(hdc, HGDIOBJ(font.0)))
                        } else {
                            None
                        };
                        DrawTextW(
                            hdc,
                            &mut label_w,
                            &mut text_rc,
                            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
                        );
                        if let Some(of) = old_f {
                            let _ = SelectObject(hdc, of);
                        }
                    }
                });
            }
            LRESULT(0)
        }
        WM_LBUTTONDOWN | WM_LBUTTONDBLCLK => {
            if !unsafe { IsWindowEnabled(hwnd) }.as_bool() {
                return LRESULT(0);
            }
            let _ = unsafe { SetFocus(hwnd) };
            flip(tg, hwnd);
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if !unsafe { IsWindowEnabled(hwnd) }.as_bool() {
                return LRESULT(0);
            }
            let vk = wparam.0 as u16;
            if vk == VK_SPACE.0 || vk == VK_RETURN.0 {
                flip(tg, hwnd);
                return LRESULT(0);
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_MOUSEMOVE => {
            if !tg.hovered {
                tg.hovered = true;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
            }
            if !tg.tracking {
                tg.tracking = true;
                let mut tme = TRACKMOUSEEVENT {
                    cbSize: std::mem::size_of::<TRACKMOUSEEVENT>() as u32,
                    dwFlags: TME_LEAVE,
                    hwndTrack: hwnd,
                    dwHoverTime: 0,
                };
                let _ = unsafe { TrackMouseEvent(&mut tme) };
            }
            LRESULT(0)
        }
        WM_MOUSELEAVE => {
            tg.tracking = false;
            if tg.hovered {
                tg.hovered = false;
                let _ = unsafe { InvalidateRect(hwnd, None, false) };
            }
            LRESULT(0)
        }
        WM_ENABLE | WM_SETFOCUS | WM_KILLFOCUS => {
            let _ = unsafe { InvalidateRect(hwnd, None, false) };
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_NCDESTROY => {
            unsafe {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                drop(Box::from_raw(raw));
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

unsafe fn font_text_height(hwnd: HWND, font: HFONT) -> i32 {
    let hdc = unsafe { GetDC(hwnd) };
    if hdc.0.is_null() {
        return 18;
    }
    let old = unsafe { SelectObject(hdc, HGDIOBJ(font.0)) };
    let mut tm = TEXTMETRICW::default();
    let ok = unsafe { GetTextMetricsW(hdc, &mut tm) }.as_bool();
    let _ = unsafe { SelectObject(hdc, old) };
    let _ = unsafe { ReleaseDC(hwnd, hdc) };
    if ok && tm.tmHeight > 0 {
        tm.tmHeight
    } else {
        18
    }
}

/// Vertically centers a borderless single-line EDIT inside its outer container `RECT`.
fn centered_edit_rect(container: RECT, text_height: i32, pad_x: i32) -> RECT {
    let ch = (container.bottom - container.top).max(1);
    let eh = text_height.clamp(1, ch);
    let top = container.top + (ch - eh) / 2;
    let left = container.left + pad_x;
    let right = (container.right - pad_x).max(left + 1);
    RECT {
        left,
        top,
        right,
        bottom: top + eh,
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

/// Where a closed combo's label is drawn, measured from the combo's own item box.
///
/// **Both paints of that label go through this** — the closed combo's own `WM_PAINT` and the
/// item's `WM_DRAWITEM` — because they swap in and out on mouse-over, and measuring one from the
/// client rect (3px wider than the item box on each side) made the label jump on hover.
fn combo_label_rect(item: RECT, dpi: i32) -> RECT {
    let left = item.left + scale(COMBO_LABEL_INSET_LOGICAL, dpi);
    RECT {
        left,
        top: item.top,
        right: (item.right - scale(26, dpi)).max(left),
        bottom: item.bottom,
    }
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
        list_rect: RECT::default(),
        edit_boxes: EDIT_IDS.map(|id| (id, RECT::default())),
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
    let label_w = scale(156, dpi);
    let stepper_w = scale(22, dpi);
    let half_row = row / 2;
    let text_h = unsafe { font_text_height(state.hwnd, state.font) };
    let edit_pad_x = scale(8, dpi);

    let moves = std::cell::RefCell::new(Vec::<(HWND, i32, i32, i32, i32)>::with_capacity(32));
    let place = |id: i32, x: i32, y: i32, w: i32, h: i32| {
        if let Ok(control) = unsafe { GetDlgItem(state.hwnd, id) } {
            moves.borrow_mut().push((control, x, y, w, h));
        }
    };
    let mut edit_boxes = EDIT_IDS.map(|id| (id, RECT::default()));
    let mut place_edit = |id: i32, x: i32, y: i32, w: i32, h: i32| {
        let container = RECT {
            left: x,
            top: y,
            right: x + w,
            bottom: y + h,
        };
        if let Some(slot) = edit_boxes.iter_mut().find(|(eid, _)| *eid == id) {
            slot.1 = container;
        }
        let mut inner_rc = centered_edit_rect(container, text_h, edit_pad_x);
        if id == ID_SEARCH {
            // Reserve room on the right for the Lucide search icon so typed text never overlaps it.
            let reserve = scale(18, dpi);
            inner_rc.right = (inner_rc.right - reserve).max(inner_rc.left + scale(16, dpi));
        }
        place(
            id,
            inner_rc.left,
            inner_rc.top,
            inner_rc.right - inner_rc.left,
            inner_rc.bottom - inner_rc.top,
        );
    };

    // Sidebar card interior.
    let sc = state.rects.sidebar_card;
    let s_left = sc.left + inner;
    let s_top = sc.top + inner;
    let s_w = (sc.right - sc.left - inner * 2).max(scale(100, dpi));
    let add_w = scale(50, dpi);
    let search_w = (s_w - add_w - scale(8, dpi)).max(scale(60, dpi));
    place_edit(ID_SEARCH, s_left, s_top, search_w, row);
    place(ID_ADD, s_left + search_w + scale(8, dpi), s_top, add_w, row);
    let list_top = s_top + row + gap;
    let list_h = (sc.bottom - inner - list_top).max(row);
    state.rects.list_rect = RECT {
        left: s_left,
        top: list_top,
        right: s_left + s_w,
        bottom: list_top + list_h,
    };
    place(ID_LIST, s_left, list_top, s_w, list_h);

    // Editor card interior.
    let ec = state.rects.editor_card;
    let e_left = ec.left + inner + scale(4, dpi);
    let field_left = e_left + label_w;
    let field_w = (ec.right - inner - scale(4, dpi) - field_left).max(scale(140, dpi));

    let mut y = ec.top + inner + scale(4, dpi);
    for index in 0..9 {
        place(
            ID_LABEL_BASE + index,
            e_left,
            y + (row - text_h) / 2,
            label_w - scale(6, dpi),
            text_h + scale(2, dpi),
        );
        y += row + gap;
    }

    let mut y = ec.top + inner + scale(4, dpi);
    place(ID_ENABLED, field_left, y, field_w, row);
    y += row + gap;
    place_edit(ID_NAME, field_left, y, field_w, row);
    y += row + gap;
    place_edit(ID_EXTENSIONS, field_left, y, field_w, row);
    y += row + gap;

    let browse_w = scale(86, dpi);
    place_edit(
        ID_EXE,
        field_left,
        y,
        (field_w - browse_w - scale(8, dpi)).max(scale(80, dpi)),
        row,
    );
    place(ID_BROWSE, field_left + field_w - browse_w, y, browse_w, row);
    y += row + gap;

    place(
        ID_CAPTURE_TECHNIQUE,
        field_left,
        y,
        field_w,
        row + combo_drop_height(dpi),
    );
    y += row + gap;

    place(
        ID_PROMOTION_TECHNIQUE,
        field_left,
        y,
        field_w,
        row + combo_drop_height(dpi),
    );
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
    place_edit(ID_SCALING_PERCENT, pct_x, y, pct_edit_w, row);
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

    // Volume override toggle + slider + numeric edit + stacked ▲/▼ stepper buttons.
    let vol_toggle_w = scale(42, dpi);
    let vol_edit_w = scale(54, dpi);
    let vol_group_w = vol_edit_w + scale(2, dpi) + stepper_w;
    let vol_slider_x = field_left + vol_toggle_w + scale(8, dpi);
    let vol_slider_w =
        (field_w - vol_toggle_w - scale(8, dpi) - vol_group_w - scale(8, dpi)).max(scale(60, dpi));
    let vol_x = field_left + field_w - vol_group_w;
    place(ID_VOLUME_ENABLED, field_left, y, vol_toggle_w, row);
    place(ID_VOLUME, vol_slider_x, y, vol_slider_w, row);
    place_edit(ID_VOLUME_VALUE, vol_x, y, vol_edit_w, row);
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

    state.rects.edit_boxes = edit_boxes;

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
        ID_STATUS,
        pad + scale(4, dpi),
        btn_y + (btn_h - text_h) / 2,
        (width - pad * 2 - btn_w * 2 - scale(20, dpi)).max(scale(80, dpi)),
        text_h + scale(2, dpi),
    );

    // Batch all child moves/resizes into a single atomic DeferWindowPos transaction.
    // SWP_NOCOPYBITS prevents Windows from blitting stale shifted bits before repaint during drag-resize.
    let moves = moves.into_inner();
    let flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS;
    unsafe {
        if let Ok(mut hdwp) = BeginDeferWindowPos(moves.len() as i32) {
            let mut ok = true;
            for &(control, x, y, w, h) in &moves {
                match DeferWindowPos(hdwp, control, HWND::default(), x, y, w, h, flags) {
                    Ok(next) => hdwp = next,
                    Err(_) => {
                        ok = false;
                        break;
                    }
                }
            }
            if ok {
                let _ = EndDeferWindowPos(hdwp);
            } else {
                for &(control, x, y, w, h) in &moves {
                    let _ = SetWindowPos(control, HWND::default(), x, y, w, h, flags);
                }
            }
        } else {
            for &(control, x, y, w, h) in &moves {
                let _ = SetWindowPos(control, HWND::default(), x, y, w, h, flags);
            }
        }
        let _ = InvalidateRect(state.hwnd, None, false);
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
        ID_NAME | ID_EXTENSIONS => {
            if notice == EN_CHANGE {
                unsafe { fields_changed(state) };
            }
        }
        ID_EXE => {
            if notice == EN_CHANGE {
                // **Cleaned here and not only on save**, so the box holds the program the preview
                // runs: `Copy as path` pastes it quoted, which is not a path to anything.
                unsafe { exe_changed(state) };
            }
        }
        ID_WINDOW_MODE | ID_CAPTURE_TECHNIQUE | ID_PROMOTION_TECHNIQUE | ID_SCALING => {
            if notice == CBN_SELCHANGE {
                if id == ID_WINDOW_MODE {
                    unsafe { window_mode_chosen(state) };
                } else if id == ID_CAPTURE_TECHNIQUE {
                    unsafe { capture_technique_chosen(state) };
                } else if id == ID_PROMOTION_TECHNIQUE {
                    unsafe { promotion_technique_chosen(state) };
                } else {
                    unsafe { scaling_chosen(state) };
                }
            }
            // A pick is confirmed by the pick, not by where the focus goes next: a combo keeps
            // the focus after the list closes, so `WM_KILLFOCUS` may be a long way off.
            if notice == CBN_SELENDOK {
                unsafe { autosave(state) };
            }
        }
        ID_SCALING_PERCENT => {
            if notice == EN_CHANGE {
                unsafe { percent_typed(state) };
            }
        }
        ID_VOLUME_VALUE => {
            if notice == EN_CHANGE {
                unsafe { volume_typed(state) };
            }
        }
        ID_BROWSE => unsafe { browse(state) },
        ID_DELETE => unsafe { delete_selected(state) },
        ID_CLOSE_BUTTON => {
            unsafe { autosave(state) };
            if !state.dirtied {
                unsafe {
                    let _ = DestroyWindow(state.hwnd);
                }
            }
        }
        _ => {}
    }
}

/// Everything the window holds, written to the configuration and announced to the rest of the app
/// — the tray's own rows, the hook, and every viewer — after the field the user was in has been
/// left.
///
/// **Left, not typed.** A save per keystroke would be a file write, a generation bump and a
/// `ConfigChanged` wake for every letter of a name, and a half-typed one written to the file that
/// a hover is routed by. The triggers are the ones that mean the user is done with what they were
/// editing: the focus leaving an `EDIT` or a `COMBOBOX`, the pointer released from the slider, the
/// toggle clicked, the drop of a dragged row. A field that has not changed writes nothing.
unsafe fn autosave(state: &mut State) {
    if state.loading || !state.dirtied {
        return;
    }
    match clean(&state.presets) {
        Ok(presets) => {
            // One write, and the announcement the rest of the app reads it on: the tray's rows,
            // the hook's bindings, and every viewer, which reload the file there rather than wait
            // out their watch (see `config::publish_config_change`).
            crate::config::publish_config_change(|config| config.custom_presets = presets, &[]);
            state.dirtied = false;
            unsafe { set_status(state, "") };
        }
        Err(reason) => unsafe { set_status(state, &reason) },
    }
}

unsafe fn set_scaling_percent_enabled(hwnd: HWND, enabled: bool) {
    for id in [ID_SCALING_PERCENT, ID_SCALING_UP, ID_SCALING_DOWN] {
        if let Ok(ctrl) = unsafe { GetDlgItem(hwnd, id) } {
            let _ = unsafe { EnableWindow(ctrl, enabled) };
        }
    }
    let _ = unsafe { InvalidateRect(hwnd, None, false) };
}

unsafe fn set_volume_controls_enabled(hwnd: HWND, enabled: bool) {
    for id in [ID_VOLUME, ID_VOLUME_VALUE, ID_VOLUME_UP, ID_VOLUME_DOWN] {
        if let Ok(ctrl) = unsafe { GetDlgItem(hwnd, id) } {
            let _ = unsafe { EnableWindow(ctrl, enabled) };
        }
    }
    let _ = unsafe { InvalidateRect(hwnd, None, false) };
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
        for id in [
            ID_WINDOW_MODE,
            ID_CAPTURE_TECHNIQUE,
            ID_PROMOTION_TECHNIQUE,
            ID_SCALING,
        ] {
            let combo = unsafe { GetDlgItem(state.hwnd, id) }.unwrap_or_default();
            unsafe { SendMessageW(combo, CB_SETCURSEL, WPARAM(usize::MAX), LPARAM(0)) };
        }
        for id in [ID_ENABLED, ID_VOLUME_ENABLED] {
            if let Ok(toggle) = unsafe { GetDlgItem(state.hwnd, id) } {
                unsafe { SendMessageW(toggle, WM_TOGGLE_SET_VALUE, WPARAM(0), LPARAM(0)) };
            }
        }
        state.loading = false;
        unsafe { set_fields_enabled(state, false) };
        return;
    };
    let name = preset.name.clone();
    let enabled = preset.enabled;
    let extensions = preset.extensions_text();
    let exe = preset.exe.clone();
    let window_mode = preset.window_mode;
    let capture_technique = preset.capture_technique;
    let promotion_technique = preset.promotion_technique;
    let scaling = preset.scaling;
    let volume_enabled = preset.volume_enabled;
    let volume = preset.volume;

    state.loading = true;
    unsafe {
        if let Ok(toggle) = GetDlgItem(state.hwnd, ID_ENABLED) {
            SendMessageW(
                toggle,
                WM_TOGGLE_SET_VALUE,
                WPARAM(usize::from(enabled)),
                LPARAM(0),
            );
        }
        if let Ok(toggle) = GetDlgItem(state.hwnd, ID_VOLUME_ENABLED) {
            SendMessageW(
                toggle,
                WM_TOGGLE_SET_VALUE,
                WPARAM(usize::from(volume_enabled)),
                LPARAM(0),
            );
        }
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

        let technique_combo = GetDlgItem(state.hwnd, ID_CAPTURE_TECHNIQUE).unwrap_or_default();
        let technique_index = CAPTURE_TECHNIQUES
            .iter()
            .position(|held| *held == capture_technique)
            .unwrap_or(0) as i32;
        SendMessageW(
            technique_combo,
            CB_SETCURSEL,
            WPARAM(technique_index as usize),
            LPARAM(0),
        );

        let promotion_combo = GetDlgItem(state.hwnd, ID_PROMOTION_TECHNIQUE).unwrap_or_default();
        let promotion_index = PROMOTION_TECHNIQUES
            .iter()
            .position(|held| *held == promotion_technique)
            .unwrap_or(0) as i32;
        SendMessageW(
            promotion_combo,
            CB_SETCURSEL,
            WPARAM(promotion_index as usize),
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

        let slider = GetDlgItem(state.hwnd, ID_VOLUME).unwrap_or_default();
        SendMessageW(
            slider,
            WM_SLIDER_SET_VALUE,
            WPARAM(volume as usize),
            LPARAM(0),
        );
        set_text(state.hwnd, ID_VOLUME_VALUE, &volume.to_string());
        set_volume_controls_enabled(state.hwnd, volume_enabled);
    }
    state.loading = false;
    unsafe { set_fields_enabled(state, true) };
}

unsafe fn set_fields_enabled(state: &State, enabled: bool) {
    for id in [
        ID_ENABLED,
        ID_NAME,
        ID_EXTENSIONS,
        ID_EXE,
        ID_BROWSE,
        ID_WINDOW_MODE,
        ID_CAPTURE_TECHNIQUE,
        ID_PROMOTION_TECHNIQUE,
        ID_SCALING,
        ID_VOLUME_ENABLED,
        ID_DELETE,
    ] {
        if let Ok(control) = unsafe { GetDlgItem(state.hwnd, id) } {
            let _ = unsafe { EnableWindow(control, enabled) };
        }
    }
    if !enabled {
        unsafe { set_scaling_percent_enabled(state.hwnd, false) };
        unsafe { set_volume_controls_enabled(state.hwnd, false) };
    } else {
        let _ = unsafe { InvalidateRect(state.hwnd, None, false) };
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

/// The row nearest the pointer a mouse message names, and whether the pointer is past the last
/// row — `LB_ITEMFROMPOINT` clamps to the nearest row, which is what a drag past either end of
/// the list needs.
fn row_at_point(list: HWND, lparam: LPARAM) -> (usize, bool) {
    let answer = unsafe { SendMessageW(list, LB_ITEMFROMPOINT, WPARAM(0), lparam) }.0 as u32;
    ((answer & 0xFFFF) as usize, (answer >> 16) & 1 == 1)
}

/// One step of a preset drag: the row the pointer is over takes the preset of the row it left, so
/// the list follows the pointer and the preset under it is the one being carried.
///
/// `press` is the drawn row the press landed on, read only on the first step — after that the row
/// being carried is the one the last step left, which is the pointer's own. Rows are the drawn
/// ones, so a search narrows which presets a drag can reach.
unsafe fn preset_drag_moved(state: &mut State, row: usize, press: usize) {
    let held = state.drag.unwrap_or(press);
    let (Some(from), Some(to)) = (
        state.shown.get(held).copied(),
        state.shown.get(row).copied(),
    ) else {
        state.drag = Some(held);
        return;
    };
    if from != to {
        let preset = state.presets.remove(from);
        state.presets.insert(to, preset);
        state.selected = state
            .selected
            .map(|selected| if selected == from { to } else { selected });
        let search = unsafe { text_of(state.hwnd, ID_SEARCH) };
        state.shown = shown_presets(&state.presets, &search);
        unsafe { rebuild_list(state) };
    }
    let drawn = state
        .shown
        .iter()
        .position(|held| *held == to)
        .unwrap_or(row);
    state.drag = Some(drawn);
    let list = unsafe { GetDlgItem(state.hwnd, ID_LIST) }.unwrap_or_default();
    let _ = unsafe { SendMessageW(list, LB_SETCURSEL, WPARAM(drawn), LPARAM(0)) };
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

    state.dirtied = true;
    unsafe { rebuild_list(state) };
}

/// The program field as it is typed or pasted: cleaned of the quotes and line breaks a paste
/// brings, put back in the box when it changed, then kept on the preset.
unsafe fn exe_changed(state: &mut State) {
    if state.loading {
        return;
    }
    let typed = unsafe { text_of(state.hwnd, ID_EXE) };
    let cleaned = sanitize_program_input(&typed);
    if cleaned != typed {
        let caret = cleaned.len() as isize;
        state.loading = true;
        unsafe {
            set_text(state.hwnd, ID_EXE, &cleaned);
            if let Ok(edit) = GetDlgItem(state.hwnd, ID_EXE) {
                SendMessageW(edit, EM_SETSEL, WPARAM(caret as usize), LPARAM(caret));
            }
        }
        state.loading = false;
    }
    unsafe { fields_changed(state) };
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
    state.dirtied = true;
}

unsafe fn capture_technique_chosen(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    let combo = unsafe { GetDlgItem(state.hwnd, ID_CAPTURE_TECHNIQUE) }.unwrap_or_default();
    let choice = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    let Some(technique) = usize::try_from(choice)
        .ok()
        .and_then(|choice| CAPTURE_TECHNIQUES.get(choice).copied())
    else {
        return;
    };
    if let Some(preset) = state.presets.get_mut(index) {
        preset.capture_technique = technique;
    }
    state.dirtied = true;
}

unsafe fn promotion_technique_chosen(state: &mut State) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    let combo = unsafe { GetDlgItem(state.hwnd, ID_PROMOTION_TECHNIQUE) }.unwrap_or_default();
    let choice = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    let Some(technique) = usize::try_from(choice)
        .ok()
        .and_then(|choice| PROMOTION_TECHNIQUES.get(choice).copied())
    else {
        return;
    };
    if let Some(preset) = state.presets.get_mut(index) {
        preset.promotion_technique = technique;
    }
    state.dirtied = true;
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
    state.dirtied = true;

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
    state.dirtied = true;
}

unsafe fn enabled_toggled(state: &mut State, enabled: bool) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    if let Some(preset) = state.presets.get_mut(index) {
        preset.enabled = enabled;
    }
    state.dirtied = true;
    if let Ok(list) = unsafe { GetDlgItem(state.hwnd, ID_LIST) } {
        let _ = unsafe { InvalidateRect(list, None, false) };
    }
}

unsafe fn row_pill_toggled(state: &mut State, row: usize) {
    let Some(&preset_idx) = state.shown.get(row) else {
        return;
    };
    let Some(preset) = state.presets.get_mut(preset_idx) else {
        return;
    };
    preset.enabled = !preset.enabled;
    let enabled = preset.enabled;
    state.dirtied = true;

    if state.selected == Some(preset_idx) {
        state.loading = true;
        if let Ok(toggle) = unsafe { GetDlgItem(state.hwnd, ID_ENABLED) } {
            unsafe {
                SendMessageW(
                    toggle,
                    WM_TOGGLE_SET_VALUE,
                    WPARAM(usize::from(enabled)),
                    LPARAM(0),
                );
            }
        }
        state.loading = false;
    }

    unsafe { autosave(state) };
    if let Ok(list) = unsafe { GetDlgItem(state.hwnd, ID_LIST) } {
        unsafe { invalidate_list_row(list, row) };
    }
}

unsafe fn volume_enabled_toggled(state: &mut State, enabled: bool) {
    if state.loading {
        return;
    }
    let Some(index) = state.selected else {
        return;
    };
    if let Some(preset) = state.presets.get_mut(index) {
        preset.volume_enabled = enabled;
    }
    state.dirtied = true;
    unsafe { set_volume_controls_enabled(state.hwnd, enabled) };
}

unsafe fn volume_moved(state: &mut State) {
    let Some(index) = state.selected else {
        return;
    };
    let slider = unsafe { GetDlgItem(state.hwnd, ID_VOLUME) }.unwrap_or_default();
    let position = unsafe { SendMessageW(slider, WM_SLIDER_GET_VALUE, WPARAM(0), LPARAM(0)).0 }
        .clamp(VOLUME_MIN as isize, VOLUME_MAX as isize) as u32;

    state.loading = true;
    unsafe { set_text(state.hwnd, ID_VOLUME_VALUE, &position.to_string()) };
    state.loading = false;

    if let Some(preset) = state.presets.get_mut(index) {
        preset.volume = position;
    }
    state.dirtied = true;
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
    state.dirtied = true;
    let slider = unsafe { GetDlgItem(state.hwnd, ID_VOLUME) }.unwrap_or_default();
    unsafe {
        SendMessageW(
            slider,
            WM_SLIDER_SET_VALUE,
            WPARAM(volume as usize),
            LPARAM(0),
        );
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
    state.dirtied = true;
    unsafe { autosave(state) };
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
        state.dirtied = true;
        unsafe { autosave(state) };
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
        || id == ID_DELETE as u32
        || id == ID_BROWSE as u32
        || id == ID_CLOSE_BUTTON as u32
        || is_stepper(id)
}

/// Offscreen-buffered `WM_DRAWITEM` helper so buttons, steppers, and list rows never flicker.
unsafe fn draw_item_buffered<F: FnOnce(HDC)>(dst_dc: HDC, r: RECT, draw: F) {
    let w = (r.right - r.left).max(1);
    let h = (r.bottom - r.top).max(1);
    let mem_dc = unsafe { CreateCompatibleDC(dst_dc) };
    let mem_bmp = unsafe { CreateCompatibleBitmap(dst_dc, w, h) };
    if !mem_dc.is_invalid() && !mem_bmp.is_invalid() {
        let old_bmp = unsafe { SelectObject(mem_dc, HGDIOBJ(mem_bmp.0)) };
        let _ = unsafe { SetViewportOrgEx(mem_dc, -r.left, -r.top, None) };
        draw(mem_dc);
        let _ = unsafe { SetViewportOrgEx(mem_dc, 0, 0, None) };
        let _ = unsafe { BitBlt(dst_dc, r.left, r.top, w, h, mem_dc, 0, 0, SRCCOPY) };
        let _ = unsafe { SelectObject(mem_dc, old_bmp) };
    } else {
        draw(dst_dc);
    }
    if !mem_bmp.is_invalid() {
        let _ = unsafe { DeleteObject(HGDIOBJ(mem_bmp.0)) };
    }
    if !mem_dc.is_invalid() {
        let _ = unsafe { DeleteDC(mem_dc) };
    }
}

unsafe fn draw_item(state: &State, item: &DRAWITEMSTRUCT) -> usize {
    let p = state.palette;
    let selected = item.itemState.0 & ODS_SELECTED.0 != 0;
    let disabled = item.itemState.0 & ODS_DISABLED.0 != 0;
    let dpi = state.dpi;

    let btn_hovered = |hwnd: HWND| -> bool {
        let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *const ButtonSubclass;
        unsafe { raw.as_ref() }.is_some_and(|s| s.hovered)
    };

    let is_primary_add = item.CtlID == ID_ADD as u32;
    let primary_grad = if is_primary_add && !disabled {
        let hov = btn_hovered(item.hwndItem);
        Some(primary_button_colors(p, hov, selected))
    } else {
        None
    };

    let (fill, border_col, text_col) = if let Some((top, _, border, fg)) = primary_grad {
        (top, border, fg)
    } else if is_stepper(item.CtlID) {
        let raw =
            unsafe { GetWindowLongPtrW(item.hwndItem, GWLP_USERDATA) } as *const StepperSubclass;
        let (hov, prs) = unsafe { raw.as_ref() }.map_or((false, false), |s| (s.hovered, s.pressed));
        let fill = if disabled {
            p.button_disabled
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
    } else if item.CtlID == ID_DELETE as u32 {
        let hov = !disabled && btn_hovered(item.hwndItem);
        let fill = if disabled {
            p.button_disabled
        } else if selected {
            mix(p.danger, rgb(0, 0, 0), 14)
        } else if hov {
            p.danger
        } else {
            p.danger_bg
        };
        let fg = if disabled {
            p.muted
        } else if selected || hov {
            p.danger_on_fill
        } else {
            p.danger_text
        };
        (fill, if disabled { p.border } else { p.danger }, fg)
    } else if is_button(item.CtlID) {
        let hov = !disabled && btn_hovered(item.hwndItem);
        let fill = if disabled {
            p.button_disabled
        } else if selected {
            p.button_press
        } else if hov {
            p.button_hover
        } else {
            p.button_bg
        };
        (fill, p.border, if disabled { p.muted } else { p.ink })
    } else if (item.CtlID == ID_LIST as u32 || is_combo(item.CtlID)) && selected {
        (mix(p.field, p.accent, 22), p.accent, p.ink)
    } else {
        (p.field, p.border, if disabled { p.muted } else { p.ink })
    };

    let list_preset_disabled = item.CtlID == ID_LIST as u32
        && state
            .shown
            .get(item.itemID as usize)
            .and_then(|&idx| state.presets.get(idx))
            .is_some_and(|preset| !preset.enabled);
    let text_col = if list_preset_disabled {
        p.muted
    } else {
        text_col
    };

    unsafe {
        draw_item_buffered(item.hDC, item.rcItem, |hdc| {
            if is_button(item.CtlID) {
                let r = item.rcItem;
                let radius = if is_stepper(item.CtlID) {
                    scale(3, dpi)
                } else {
                    scale(6, dpi)
                };
                if let Some((top_col, bot_col, bcol, _)) = primary_grad {
                    let bw = (r.right - r.left).max(1);
                    let bh = (r.bottom - r.top).max(1);
                    let mut surface = AaSurface::new(bw, bh, p.card);
                    surface.fill_rounded_box_gradient(
                        [0.5, 0.5, bw as f32 - 0.5, bh as f32 - 0.5],
                        radius as f32,
                        top_col,
                        bot_col,
                        Some((bcol, 1.0)),
                    );
                    surface.blit_to(hdc, r.left, r.top);
                } else {
                    let parent_brush =
                        if item.CtlID == ID_DELETE as u32 || item.CtlID == ID_CLOSE_BUTTON as u32 {
                            state.brush_page
                        } else {
                            state.brush_card
                        };
                    FillRect(hdc, &item.rcItem, parent_brush);

                    let brush = CreateSolidBrush(fill);
                    let pen = CreatePen(PS_SOLID, 1, border_col);
                    let old_b = SelectObject(hdc, HGDIOBJ(brush.0));
                    let old_p = SelectObject(hdc, HGDIOBJ(pen.0));
                    let _ = RoundRect(hdc, r.left, r.top, r.right, r.bottom, radius, radius);
                    let _ = SelectObject(hdc, old_b);
                    let _ = SelectObject(hdc, old_p);
                    let _ = DeleteObject(HGDIOBJ(brush.0));
                    let _ = DeleteObject(HGDIOBJ(pen.0));
                }
            } else {
                let brush = CreateSolidBrush(fill);
                FillRect(hdc, &item.rcItem, brush);
                let _ = DeleteObject(HGDIOBJ(brush.0));
                if item.CtlID == ID_LIST as u32 && selected {
                    // Smooth 3px rounded pill accent indicator bar on selected preset item.
                    let bar_w = scale(3, dpi).max(3);
                    let bar_x = item.rcItem.left + scale(2, dpi);
                    let bar_top = item.rcItem.top + scale(4, dpi);
                    let bar_bot = (item.rcItem.bottom - scale(4, dpi)).max(bar_top + bar_w);
                    let pill_col = if list_preset_disabled {
                        p.muted
                    } else {
                        p.accent
                    };
                    let pill_brush = CreateSolidBrush(pill_col);
                    let pill_pen = CreatePen(PS_SOLID, 1, pill_col);
                    let old_b = SelectObject(hdc, HGDIOBJ(pill_brush.0));
                    let old_p = SelectObject(hdc, HGDIOBJ(pill_pen.0));
                    let _ = RoundRect(hdc, bar_x, bar_top, bar_x + bar_w, bar_bot, bar_w, bar_w);
                    let _ = SelectObject(hdc, old_b);
                    let _ = SelectObject(hdc, old_p);
                    let _ = DeleteObject(HGDIOBJ(pill_brush.0));
                    let _ = DeleteObject(HGDIOBJ(pill_pen.0));
                }
            }

            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, text_col);
            let font = if is_stepper(item.CtlID) {
                state.font_small
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
                id if is_combo(id) => {
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

            let mut rect = if is_combo(item.CtlID) {
                combo_label_rect(item.rcItem, dpi)
            } else {
                item.rcItem
            };
            if item.CtlID == ID_LIST as u32 {
                let mut pill_rc = list_row_pill_rect(item.rcItem, dpi as u32);
                let pill_w = (pill_rc.right - pill_rc.left).max(1);
                let pill_h = (pill_rc.bottom - pill_rc.top).max(1);
                let raw = GetWindowLongPtrW(item.hwndItem, GWLP_USERDATA) as *const ListSubclass;
                let pill_hovered = raw
                    .as_ref()
                    .is_some_and(|s| s.hovered_pill == Some(item.itemID as usize));
                let enabled = !list_preset_disabled;
                let (pill_fill, pill_border, pill_fg) =
                    status_pill_colors(p, enabled, pill_hovered);
                let badge = render_status_badge_pill(pill_w, pill_h, fill, pill_fill, pill_border);
                badge.blit_to(hdc, pill_rc.left, pill_rc.top);

                let old_f = SelectObject(hdc, HGDIOBJ(state.font_small.0));
                SetTextColor(hdc, pill_fg);
                let mut label_w: Vec<u16> =
                    if enabled { "On" } else { "Off" }.encode_utf16().collect();
                DrawTextW(
                    hdc,
                    &mut label_w,
                    &mut pill_rc,
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
                );
                let _ = SelectObject(hdc, old_f);
                SetTextColor(hdc, text_col);
                rect.right = (pill_rc.left - scale(6, dpi)).max(rect.left + scale(10, dpi));
            }
            // Empty slice would hand DrawTextW a dangling alignment pointer; return early.
            if text.is_empty() {
                return;
            }

            let mut text: Vec<u16> = text.encode_utf16().collect();
            let flags = if is_button(item.CtlID) {
                DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
            } else {
                if !is_combo(item.CtlID) {
                    rect.left += scale(10, dpi);
                }
                DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS
            };
            DrawTextW(hdc, &mut text, &mut rect, flags);
        });
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
        GetWindowRect, PeekMessageW, CB_GETDROPPEDCONTROLRECT, CB_SHOWDROPDOWN, PM_REMOVE,
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
    fn a_combo_label_is_measured_from_its_item_box() {
        // The combo's item box starts 3px inside its client rect; both label paints must place
        // the text from the item, or the label jumps 3px when they swap on mouse-over.
        let item = RECT {
            left: 3,
            top: 3,
            right: 472,
            bottom: 29,
        };
        let label = combo_label_rect(item, 96);
        assert_eq!(label.left, item.left + COMBO_LABEL_INSET_LOGICAL);
        assert_eq!(label.top, item.top);
        assert_eq!(label.bottom, item.bottom);
        assert_eq!(label.right, item.right - 26);
    }

    #[test]
    fn a_combo_label_starts_where_an_edit_starts_its_text() {
        // Same row of the editor, one box at the container's left edge (the combo) and one at its
        // 8px pad (the EDIT). Both first glyphs landed on the same pixel on screen: the combo at
        // item 3 + inset 4 + the 2px its font keeps before the glyph, the EDIT at pad 8 + 1.
        let container = RECT {
            left: 398,
            top: 40,
            right: 890,
            bottom: 70,
        };
        let edit_glyph = centered_edit_rect(container, 18, scale(8, 96)).left + 1;
        let combo_item = RECT {
            left: container.left + scale(3, 96), // the item box's own Win32 margin
            ..container
        };
        assert_eq!(
            combo_label_rect(combo_item, 96).left + 2,
            edit_glyph,
            "a combo's label and an EDIT's text start on different columns"
        );
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
                ID_ENABLED,
                ID_NAME,
                ID_EXTENSIONS,
                ID_EXE,
                ID_WINDOW_MODE,
                ID_CAPTURE_TECHNIQUE,
                ID_PROMOTION_TECHNIQUE,
                ID_SCALING,
                ID_SCALING_UP,
                ID_SCALING_DOWN,
                ID_VOLUME_ENABLED,
                ID_VOLUME,
                ID_VOLUME_VALUE,
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

    /// A field left is a save, and a save writes the configuration the rest of the app reads.
    #[test]
    fn a_field_left_writes_the_preset_and_announces_it() {
        let folder = crate::config::config::ConfigFolderOfItsOwn::take();
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };

        unsafe {
            let Some(state) = state_of(hwnd) else {
                panic!("the window kept no state");
            };
            let mut preset = CustomPreset::named("Painted");
            preset.exe = "C:/mspaint.exe".into();
            preset.set_extensions_text("jpg");
            state.presets = vec![preset];
            state.selected = Some(0);
            state.dirtied = true;
            autosave(state);

            assert!(!state.dirtied, "a saved window has nothing left to write");
            let stored = {
                crate::CONFIG
                    .lock()
                    .map(|config| config.custom_presets.clone())
                    .unwrap_or_default()
            };
            assert_eq!(stored.len(), 1);
            assert_eq!(stored[0].name, "Painted");

            // And it reaches the file, which is what the rest of the app reads.
            let written = std::fs::read_to_string(folder.config_file()).unwrap_or_default();
            assert!(
                written.contains("[custom:painted]"),
                "the preset reached the file: {written}"
            );
            let mut read_back = crate::config::config::AppConfig::default();
            let mut ini = configparser::ini::Ini::new();
            ini.read(written).expect("a file this app wrote");
            crate::config::config::custom_presets_read_for_test(&ini, &mut read_back);
            assert_eq!(
                read_back
                    .custom_presets
                    .iter()
                    .map(|preset| preset.name.as_str())
                    .collect::<Vec<_>>(),
                vec!["Painted"]
            );

            let _ = DestroyWindow(hwnd);
        }
    }

    /// The two states the window can be in that no file may hold: a preset with no name, and two
    /// presets of one name. Both are refused before anything is written, and `autosave` says why
    /// on the status line (`clean`).
    #[test]
    fn an_unnamed_or_a_doubled_preset_is_what_may_not_be_written() {
        let mut good = CustomPreset::named("Painted");
        good.exe = "C:/mspaint.exe".into();
        good.set_extensions_text("jpg");
        assert!(clean(std::slice::from_ref(&good)).is_ok());

        let mut blank = good.clone();
        blank.name = "   ".into();
        assert!(clean(&[blank]).is_err(), "a preset with no name is refused");

        let mut twin = good.clone();
        twin.name = "painted".into();
        assert!(
            clean(&[good, twin]).is_err(),
            "two presets of one name are refused"
        );
    }

    #[test]
    fn toggle_switch_updates_selected_preset_enabled_state() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };

        unsafe {
            let Some(state) = state_of(hwnd) else {
                panic!("the window kept no state");
            };
            state.presets.clear();
            state.presets.push(CustomPreset::named("ToggleTest"));
            state.selected = Some(0);
            rebuild_list(state);
            fill_fields(state);

            let toggle = GetDlgItem(hwnd, ID_ENABLED).unwrap_or_default();
            assert!(IsWindowEnabled(toggle).as_bool());
            assert_eq!(
                SendMessageW(toggle, WM_TOGGLE_GET_VALUE, WPARAM(0), LPARAM(0)).0,
                1
            );
            assert!(state.presets[0].enabled);

            // Simulate Space key toggle -> OFF.
            SendMessageW(toggle, WM_KEYDOWN, WPARAM(VK_SPACE.0 as usize), LPARAM(0));
            assert_eq!(
                SendMessageW(toggle, WM_TOGGLE_GET_VALUE, WPARAM(0), LPARAM(0)).0,
                0
            );
            assert!(!state.presets[0].enabled);

            // Simulate click toggle -> ON.
            SendMessageW(toggle, WM_LBUTTONDOWN, WPARAM(0), LPARAM(0));
            assert_eq!(
                SendMessageW(toggle, WM_TOGGLE_GET_VALUE, WPARAM(0), LPARAM(0)).0,
                1
            );
            assert!(state.presets[0].enabled);

            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn a_drag_over_the_preset_list_reorders_it_and_is_survived() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };
        unsafe {
            let Some(state) = state_of(hwnd) else {
                panic!("the window kept no state");
            };
            state.presets.clear();
            for name in ["Alpha", "Bravo", "Charlie"] {
                let mut preset = CustomPreset::named(name);
                preset.exe = "C:/nope.exe".into();
                preset.set_extensions_text("jpg");
                state.presets.push(preset);
            }
            state.selected = Some(0);
            rebuild_list(state);
            fill_fields(state);

            let list = GetDlgItem(hwnd, ID_LIST).unwrap_or_default();
            let row_h = scale(26, 96);
            let at =
                |row: i32| -> LPARAM { LPARAM(((row * row_h + row_h / 2) << 16 | 5) as isize) };
            let pump = || {
                let mut message = MSG::default();
                while PeekMessageW(&mut message, hwnd, 0, 0, PM_REMOVE).as_bool() {
                    let _ = TranslateMessage(&message);
                    let _ = DispatchMessageW(&message);
                }
            };
            let order = || -> Vec<String> {
                state_of(hwnd)
                    .map(|state| {
                        state
                            .presets
                            .iter()
                            .map(|preset| preset.name.clone())
                            .collect()
                    })
                    .unwrap_or_default()
            };

            SendMessageW(list, WM_LBUTTONDOWN, WPARAM(1), at(0));
            SendMessageW(list, WM_MOUSEMOVE, WPARAM(1), at(1));
            // The reorder is the parent's, and it is posted rather than run inside the box's own
            // mouse message: only the queue makes it happen, and it must not have happened yet.
            assert_eq!(
                order(),
                vec!["Alpha", "Bravo", "Charlie"],
                "not before the queue"
            );
            pump();
            assert_eq!(
                order(),
                vec!["Bravo", "Alpha", "Charlie"],
                "the row was moved"
            );
            assert_eq!(
                state_of(hwnd).map(|state| state.selected),
                Some(Some(1)),
                "the moved preset stays the one the editor is on"
            );

            SendMessageW(list, WM_MOUSEMOVE, WPARAM(1), at(2));
            pump();
            assert_eq!(order(), vec!["Bravo", "Charlie", "Alpha"]);
            assert_eq!(
                SendMessageW(list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0,
                2,
                "the list follows the pointer"
            );

            SendMessageW(list, WM_LBUTTONUP, WPARAM(0), at(2));
            pump();
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

        // High-contrast Save & Delete button tokens in both themes.
        assert_eq!(dark.accent_text, rgb(0x1E, 0x22, 0x2A));
        assert_eq!(dark.danger, rgb(0xE0, 0x6C, 0x75));
        assert_eq!(dark.danger_on_fill, rgb(0x1E, 0x22, 0x2A));
        assert_eq!(light.accent_text, rgb(0xFF, 0xFF, 0xFF));
        assert_eq!(light.danger, rgb(0xE4, 0x56, 0x49));
        assert_eq!(light.danger_text, rgb(0xC9, 0x3B, 0x32));
        assert_eq!(light.danger_on_fill, rgb(0xFF, 0xFF, 0xFF));
    }

    #[test]
    fn slider_math_maps_endpoints_and_midpoint_accurately() {
        assert_eq!(slider_value_from_x(0, 10, 210), 0);
        assert_eq!(slider_value_from_x(10, 10, 210), 0);
        assert_eq!(slider_value_from_x(110, 10, 210), 50);
        assert_eq!(slider_value_from_x(210, 10, 210), 100);
        assert_eq!(slider_value_from_x(300, 10, 210), 100);

        assert_eq!(slider_thumb_x(0, 10, 210), 10);
        assert_eq!(slider_thumb_x(50, 10, 210), 110);
        assert_eq!(slider_thumb_x(100, 10, 210), 210);
    }

    #[test]
    fn centered_edit_rect_centers_text_vertically_with_horizontal_padding() {
        let container = RECT {
            left: 100,
            top: 40,
            right: 300,
            bottom: 70,
        };
        let edit = centered_edit_rect(container, 18, 8);
        assert_eq!(edit.left, 108);
        assert_eq!(edit.right, 292);
        assert_eq!(edit.top, 46);
        assert_eq!(edit.bottom, 64);
        assert_eq!(edit.top - container.top, container.bottom - edit.bottom);
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

    fn relative_luminance(c: COLORREF) -> f32 {
        let lin = |ch: u8| -> f32 {
            let s = f32::from(ch) / 255.0;
            if s <= 0.04045 {
                s / 12.92
            } else {
                ((s + 0.055) / 1.055).powf(2.4)
            }
        };
        let (r, g, b) = colorref_rgb(c);
        0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
    }

    fn contrast_ratio(a: COLORREF, b: COLORREF) -> f32 {
        let la = relative_luminance(a);
        let lb = relative_luminance(b);
        let (hi, lo) = if la >= lb { (la, lb) } else { (lb, la) };
        (hi + 0.05) / (lo + 0.05)
    }

    #[test]
    fn field_label_contrast_is_high_in_both_themes() {
        let dark = Palette::of(TextTheme::Dark);
        let dark_ratio = contrast_ratio(dark.label, dark.card);
        assert!(
            dark_ratio >= 7.0,
            "One Dark Pro label/card contrast {dark_ratio:.2} < 7.0"
        );

        let light = Palette::of(TextTheme::Light);
        let light_ratio = contrast_ratio(light.label, light.card);
        assert!(
            light_ratio >= 7.0,
            "Atom One Light label/card contrast {light_ratio:.2} < 7.0"
        );
    }

    #[test]
    fn sdf_aa_coverage_is_smooth_across_circle_and_pill_boundaries() {
        // Circle boundary: inside = 1.0, exact edge = 0.5, outside = 0.0, monotone across [-0.5, 0.5].
        let c_in = aa_coverage(sd_circle(10.0, 10.0, 10.0, 10.0, 6.0));
        let c_mid = aa_coverage(sd_circle(16.0, 10.0, 10.0, 10.0, 6.0));
        let c_near_in = aa_coverage(sd_circle(15.75, 10.0, 10.0, 10.0, 6.0));
        let c_near_out = aa_coverage(sd_circle(16.25, 10.0, 10.0, 10.0, 6.0));
        let c_out = aa_coverage(sd_circle(17.0, 10.0, 10.0, 10.0, 6.0));
        assert_eq!(c_in, 1.0);
        assert!((c_near_in - 0.75).abs() < 1e-4);
        assert!((c_mid - 0.5).abs() < 1e-4);
        assert!((c_near_out - 0.25).abs() < 1e-4);
        assert_eq!(c_out, 0.0);

        // Rounded pill boundary: smooth transition across curved cap.
        let p_in = aa_coverage(sd_rounded_box(20.0, 10.0, 20.0, 10.0, 19.0, 10.0, 10.0));
        let p_edge = aa_coverage(sd_rounded_box(39.0, 10.0, 20.0, 10.0, 19.0, 10.0, 10.0));
        let p_out = aa_coverage(sd_rounded_box(40.0, 10.0, 20.0, 10.0, 19.0, 10.0, 10.0));
        assert_eq!(p_in, 1.0);
        assert!((p_edge - 0.5).abs() < 1e-4);
        assert_eq!(p_out, 0.0);
    }

    #[test]
    fn preset_listbox_has_nointegralheight_style() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };
        unsafe {
            let list = GetDlgItem(hwnd, ID_LIST).unwrap_or_default();
            let style =
                GetWindowLongPtrW(list, windows::Win32::UI::WindowsAndMessaging::GWL_STYLE) as u32;
            assert_ne!(
                style & (LBS_NOINTEGRALHEIGHT as u32),
                0,
                "ID_LIST must have LBS_NOINTEGRALHEIGHT to prevent vertical resize bounce"
            );
            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn capture_technique_combo_selects_and_updates_preset_technique() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };
        unsafe {
            let Some(state) = state_of(hwnd) else {
                panic!("the window kept no state");
            };
            state.presets.clear();
            state.presets.push(CustomPreset::named("CaptureTest"));
            state.selected = Some(0);
            rebuild_list(state);
            fill_fields(state);

            let combo = GetDlgItem(hwnd, ID_CAPTURE_TECHNIQUE).unwrap_or_default();
            assert!(IsWindowEnabled(combo).as_bool());
            assert_eq!(
                SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0,
                0,
                "defaults to InProcessHook (index 0)"
            );
            assert_eq!(
                state.presets[0].capture_technique,
                CustomCaptureTechnique::InProcessHook
            );

            assert_eq!(CAPTURE_TECHNIQUES.len(), 6);
            assert_eq!(CAPTURE_TECHNIQUES[0], CustomCaptureTechnique::InProcessHook);
            assert_eq!(
                capture_technique_label(CustomCaptureTechnique::InProcessHook),
                "In-Process Hook (Default)"
            );
            assert_eq!(
                CAPTURE_TECHNIQUES[1],
                CustomCaptureTechnique::InvisibleLayered
            );
            assert_eq!(
                capture_technique_label(CustomCaptureTechnique::InvisibleLayered),
                "Invisible Layered"
            );
            assert_eq!(
                CAPTURE_TECHNIQUES[5],
                CustomCaptureTechnique::OffscreenStaged
            );
            assert_eq!(
                capture_technique_label(CustomCaptureTechnique::OffscreenStaged),
                "Off-Screen Staged (-10000 -> Reveal)"
            );
            for (idx, expected) in CAPTURE_TECHNIQUES.iter().copied().enumerate() {
                SendMessageW(combo, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
                command(state, ID_CAPTURE_TECHNIQUE, CBN_SELCHANGE);
                assert_eq!(state.presets[0].capture_technique, expected);
                assert!(state.dirtied);
                state.dirtied = false;
            }

            // Capture Technique sits directly above Promotion Technique, which sits directly above
            // Window (both controls & labels) and follows that exact Tab order.
            let promo_combo = GetDlgItem(hwnd, ID_PROMOTION_TECHNIQUE).unwrap_or_default();
            let win_combo = GetDlgItem(hwnd, ID_WINDOW_MODE).unwrap_or_default();
            let browse_btn = GetDlgItem(hwnd, ID_BROWSE).unwrap_or_default();
            let (mut cap_rc, mut promo_rc, mut win_rc) =
                (RECT::default(), RECT::default(), RECT::default());
            let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(combo, &mut cap_rc);
            let _ =
                windows::Win32::UI::WindowsAndMessaging::GetWindowRect(promo_combo, &mut promo_rc);
            let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(win_combo, &mut win_rc);
            assert!(
                cap_rc.top < promo_rc.top && promo_rc.top < win_rc.top,
                "Row order must be Capture Technique ({}) < Promotion Technique ({}) < Window ({})",
                cap_rc.top,
                promo_rc.top,
                win_rc.top
            );
            assert_eq!(
                GetNextDlgTabItem(hwnd, browse_btn, false).ok(),
                Some(combo),
                "Tab from Browse must move to Capture Technique"
            );
            assert_eq!(
                GetNextDlgTabItem(hwnd, combo, false).ok(),
                Some(promo_combo),
                "Tab from Capture Technique must move to Promotion Technique"
            );
            assert_eq!(
                GetNextDlgTabItem(hwnd, promo_combo, false).ok(),
                Some(win_combo),
                "Tab from Promotion Technique must move to Window"
            );

            // Label 4 ("Capture Technique") < Label 5 ("Promotion Technique") < Label 6 ("Window").
            let cap_label = GetDlgItem(hwnd, ID_LABEL_BASE + 4).unwrap_or_default();
            let promo_label = GetDlgItem(hwnd, ID_LABEL_BASE + 5).unwrap_or_default();
            let win_label = GetDlgItem(hwnd, ID_LABEL_BASE + 6).unwrap_or_default();
            let (mut lrc, mut plrc, mut wlrc) = (RECT::default(), RECT::default(), RECT::default());
            let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(cap_label, &mut lrc);
            let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(promo_label, &mut plrc);
            let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(win_label, &mut wlrc);
            assert!(lrc.top < plrc.top && plrc.top < wlrc.top);
            assert!(
                plrc.right - plrc.left >= scale(140, state.dpi),
                "Promotion Technique label width {}px too narrow",
                plrc.right - plrc.left
            );

            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn promotion_technique_combo_selects_and_updates_preset_technique() {
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };
        unsafe {
            let Some(state) = state_of(hwnd) else {
                panic!("the window kept no state");
            };
            state.presets.clear();
            state.presets.push(CustomPreset::named("PromoTest"));
            state.selected = Some(0);
            rebuild_list(state);
            fill_fields(state);

            let combo = GetDlgItem(hwnd, ID_PROMOTION_TECHNIQUE).unwrap_or_default();
            assert!(IsWindowEnabled(combo).as_bool());
            assert_eq!(
                SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0,
                0,
                "defaults to DirectPromotion (index 0)"
            );
            assert_eq!(
                state.presets[0].promotion_technique,
                CustomPromotionTechnique::DirectPromotion
            );

            assert_eq!(PROMOTION_TECHNIQUES.len(), 2);
            assert_eq!(
                PROMOTION_TECHNIQUES[0],
                CustomPromotionTechnique::DirectPromotion
            );
            assert_eq!(
                promotion_technique_label(CustomPromotionTechnique::DirectPromotion),
                "Direct Promotion"
            );
            assert_eq!(
                PROMOTION_TECHNIQUES[1],
                CustomPromotionTechnique::OpenAsNewProcess
            );
            assert_eq!(
                promotion_technique_label(CustomPromotionTechnique::OpenAsNewProcess),
                "Open as New Process"
            );

            for (idx, expected) in PROMOTION_TECHNIQUES.iter().copied().enumerate() {
                SendMessageW(combo, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
                command(state, ID_PROMOTION_TECHNIQUE, CBN_SELCHANGE);
                assert_eq!(state.presets[0].promotion_technique, expected);
                assert!(state.dirtied);
                state.dirtied = false;
            }

            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn backdrop_gradient_and_primary_button_contrast_hold_in_both_themes() {
        for theme in [TextTheme::Dark, TextTheme::Light] {
            let p = Palette::of(theme);

            // Primary button vertical gradient maintains >= 7:1 text contrast on both stops
            // across normal, hover, and pressed states.
            for (hov, prs) in [(false, false), (true, false), (false, true)] {
                let (top, bot, _, fg) = primary_button_colors(p, hov, prs);
                assert!(
                    contrast_ratio(fg, top) >= 7.0,
                    "{theme:?} primary top contrast {:.2} < 7.0 (hov={hov}, prs={prs})",
                    contrast_ratio(fg, top)
                );
                assert!(
                    contrast_ratio(fg, bot) >= 7.0,
                    "{theme:?} primary bottom contrast {:.2} < 7.0 (hov={hov}, prs={prs})",
                    contrast_ratio(fg, bot)
                );
            }

            // Atmospheric backdrop has NO top neon bar or accent dot, keeps smooth gradient & cards.
            let rects = compute_layout_rects(920, 620, 96);
            let surface = render_backdrop(
                920,
                620,
                p,
                96,
                rects.titlebar.bottom,
                rects.sidebar_card,
                rects.editor_card,
            );
            let px =
                |x: i32, y: i32| surface.pixels[(y as usize) * (surface.w as usize) + x as usize];
            let (ar, ag, ab) = colorref_rgb(p.accent);
            assert_ne!(
                px(0, 0),
                pack_bgra(ar, ag, ab),
                "top neon hairline must be removed"
            );
            assert_ne!(
                px(20, rects.titlebar.bottom / 2),
                pack_bgra(ar, ag, ab),
                "titlebar neon accent dot must be removed"
            );
            assert_ne!(px(4, rects.titlebar.bottom + 4), px(915, 615));
            let (cr, cg, cb) = colorref_rgb(p.card);
            assert_eq!(
                px(rects.editor_card.left + 40, rects.editor_card.top + 40),
                pack_bgra(cr, cg, cb)
            );

            // Compact 128x96 gradient cache & 9-slice card template build cleanly in both themes.
            let grad = render_gradient_backdrop(GRADIENT_W, GRADIENT_H, p);
            assert_eq!((grad.w, grad.h), (GRADIENT_W, GRADIENT_H));
            assert_ne!(grad.pixels[0], *grad.pixels.last().unwrap());
            let (slice_px, tw, th, ..) = render_card_9slice_pixels(p, 96);
            assert_eq!(slice_px.len(), (tw * th) as usize);
            assert_eq!(
                slice_px[((th / 2) * tw + (tw / 2)) as usize],
                pack_bgra(cr, cg, cb)
            );

            // SDF anti-aliased status badge pill has background corners and solid fill center.
            let pill = render_status_badge_pill(32, 17, p.field, p.button_bg, p.border);
            let (fr, fg, fb) = colorref_rgb(p.field);
            let (br, bg, bb) = colorref_rgb(p.button_bg);
            assert_eq!(pill.pixels[0], pack_bgra(fr, fg, fb));
            assert_eq!(pill.pixels[(8 * 32 + 16) as usize], pack_bgra(br, bg, bb));
        }
    }

    #[test]
    fn list_row_pill_rect_fits_and_aligns_right_across_dpis() {
        for dpi in [96u32, 120, 144, 192] {
            let row_h = scale(26, dpi as i32);
            let row = RECT {
                left: 0,
                top: row_h * 2,
                right: scale(220, dpi as i32),
                bottom: row_h * 3,
            };
            let pill = list_row_pill_rect(row, dpi);
            assert!(pill.left >= row.left + scale(20, dpi as i32));
            assert_eq!(pill.right, row.right - scale(8, dpi as i32));
            assert_eq!(pill.right - pill.left, scale(32, dpi as i32));
            assert_eq!(pill.bottom - pill.top, scale(17, dpi as i32));
            assert!(pill.top >= row.top && pill.bottom <= row.bottom);
            let top_gap = pill.top - row.top;
            let bot_gap = row.bottom - pill.bottom;
            assert!((top_gap - bot_gap).abs() <= 1);
        }
    }

    #[test]
    fn status_pill_contrast_is_legible_in_both_themes() {
        for theme in [TextTheme::Dark, TextTheme::Light] {
            let p = Palette::of(theme);
            for hovered in [false, true] {
                let (on_fill, _, on_fg) = status_pill_colors(p, true, hovered);
                let on_ratio = contrast_ratio(on_fg, on_fill);
                assert!(
                    on_ratio >= 4.5,
                    "{theme:?} On pill contrast (hovered={hovered}) {on_ratio:.2} < 4.5"
                );

                let (off_fill, _, off_fg) = status_pill_colors(p, false, hovered);
                let off_ratio = contrast_ratio(off_fg, off_fill);
                assert!(
                    off_ratio >= 4.5,
                    "{theme:?} Off pill contrast (hovered={hovered}) {off_ratio:.2} < 4.5"
                );
            }
        }
    }

    #[test]
    fn volume_override_toggle_defaults_off_and_gates_volume_controls() {
        let _folder = crate::config::config::ConfigFolderOfItsOwn::take();
        let Some(hwnd) = create_window() else {
            panic!("the window was never made");
        };
        unsafe {
            let idx = {
                let Some(state) = state_of(hwnd) else {
                    panic!("the window kept no state");
                };
                add_preset(state);
                let idx = state.selected.expect("added preset selected");
                assert!(!state.presets[idx].volume_enabled);
                idx
            };

            let vol_toggle = GetDlgItem(hwnd, ID_VOLUME_ENABLED).unwrap_or_default();
            assert!(IsWindowEnabled(vol_toggle).as_bool());
            assert_eq!(
                SendMessageW(vol_toggle, WM_TOGGLE_GET_VALUE, WPARAM(0), LPARAM(0)).0,
                0
            );
            for id in [ID_VOLUME, ID_VOLUME_VALUE, ID_VOLUME_UP, ID_VOLUME_DOWN] {
                let ctrl = GetDlgItem(hwnd, id).unwrap_or_default();
                assert!(
                    !IsWindowEnabled(ctrl).as_bool(),
                    "volume control {id} should be disabled while volume_enabled is false"
                );
            }

            SendMessageW(
                hwnd,
                WM_TOGGLE_CHANGED,
                WPARAM(1),
                LPARAM(ID_VOLUME_ENABLED as isize),
            );
            let Some(state) = state_of(hwnd) else {
                panic!("the window kept no state");
            };
            assert!(state.presets[idx].volume_enabled);
            for id in [ID_VOLUME, ID_VOLUME_VALUE, ID_VOLUME_UP, ID_VOLUME_DOWN] {
                let ctrl = GetDlgItem(hwnd, id).unwrap_or_default();
                assert!(
                    IsWindowEnabled(ctrl).as_bool(),
                    "volume control {id} should be enabled once volume_enabled is true"
                );
            }

            let _ = DestroyWindow(hwnd);
        }
    }

    #[test]
    fn lucide_search_icon_renders_ring_and_handle_with_clean_corners() {
        for theme in [TextTheme::Dark, TextTheme::Light] {
            let p = Palette::of(theme);
            let bg = p.field;
            let fg = mix(p.muted, p.field, 18);
            let icon = render_search_icon(24, bg, fg);
            let (br, bgc, bb) = colorref_rgb(bg);
            let bg_px = pack_bgra(br, bgc, bb);
            // Outer top-left corner and interior of circle ring stay background.
            assert_eq!(icon.pixels[0], bg_px);
            assert_eq!(icon.pixels[11 * 24 + 11], bg_px);
            // Ring perimeter at (11, 3) and diagonal handle at (19, 19) are stroked.
            assert_ne!(icon.pixels[3 * 24 + 11], bg_px);
            assert_ne!(icon.pixels[19 * 24 + 19], bg_px);
        }
    }
}

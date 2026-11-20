//! The `Custom Preview` window: the one place a custom preview preset is made and edited.
//!
//! **What it is for.** A preset is the user naming a program for an extension (see
//! `config::config::custom_presets`), and this window is the list of them: a column of tabs on
//! the left, one tab a preset, a search box and an `Add` above it, and the preset's own fields on
//! the right. Every field is one of the five things the section holds — the program, the names it
//! claims, the scale its window is drawn at, the backdrop behind it, and the level it is heard at
//! — so the window and the file say the same five things and nothing else.
//!
//! **Why the controls are the platform's.** A text field with a caret, a selection, an IME and a
//! keyboard is a window's own work, and this app's preview window draws every pixel of itself
//! because a preview must composite over Explorer. A settings window has no such constraint: it
//! is an ordinary top-level window of its own, so the fields are `EDIT`, `BUTTON`, `COMBOBOX` and
//! a trackbar, and what this file owns is the theme they are painted in — the same `TextTheme`
//! the text preview and the audio card use (`Tray → System → Theme`), so a preset reads as the
//! rest of the app does. The two lists — the tabs and the combo boxes — and the buttons are
//! owner-drawn, because those are the ones a system-drawn control would paint light over a dark
//! page.
//!
//! **Why it runs on a thread of its own.** The window is opened from a tray menu, and the tray's
//! own thread is the one pumping the menu that opened it: a modal loop there would hold the tray's
//! own `WM_COMMAND` open. So the window is made, pumped and destroyed on a thread of its own, and
//! the tray is free the moment the thread is begun. One window at a time (`OPEN`): a second click
//! brings the first one forward rather than making another.
//!
//! **What a save writes.** The whole list, in the order the tabs are in, into the running
//! configuration, which is then written to `config.ini` — the file is what the watcher and the
//! next hover read, and a preset that was not written is a preset no preview uses. The window
//! also saves as it closes, so a field typed into and left is not lost.

use std::sync::atomic::{AtomicBool, AtomicIsize, Ordering};

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, CreateFontW, CreateSolidBrush, DeleteObject, DrawTextW, EndPaint, FillRect,
    GetStockObject, InvalidateRect, SelectObject, SetBkColor, SetBkMode, SetTextColor,
    UpdateWindow, CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET, DEFAULT_PITCH, DT_CENTER, DT_LEFT,
    DT_NOPREFIX, DT_SINGLELINE, DT_VCENTER, FW_NORMAL, HBRUSH, HDC, HFONT, HGDIOBJ, NULL_BRUSH,
    OUT_DEFAULT_PRECIS, PAINTSTRUCT, PROOF_QUALITY, TRANSPARENT,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{DRAWITEMSTRUCT, MEASUREITEMSTRUCT, ODS_SELECTED};
use windows::Win32::UI::Controls::{EM_SETSEL, TBM_SETPOS, TBM_SETRANGE};
use windows::Win32::UI::HiDpi::GetDpiForWindow;
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, GetClientRect, GetDlgItem,
    GetMessageW, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, LoadCursorW,
    PostMessageW, PostQuitMessage, RegisterClassExW, SendMessageW, SetForegroundWindow,
    SetWindowLongPtrW, SetWindowPos, SetWindowTextW, ShowWindow, TranslateMessage, BS_OWNERDRAW,
    CBS_DROPDOWNLIST, CBS_HASSTRINGS, CBS_OWNERDRAWFIXED, CB_ADDSTRING, CB_GETCURSEL, CB_GETLBTEXT,
    CB_GETLBTEXTLEN, CB_SETCURSEL, CS_HREDRAW, CS_VREDRAW, CW_USEDEFAULT, EN_CHANGE, GWLP_USERDATA,
    HMENU, IDC_ARROW, LBN_SELCHANGE, LBS_HASSTRINGS, LBS_NOTIFY, LBS_OWNERDRAWFIXED, LB_ADDSTRING,
    LB_GETCURSEL, LB_GETTEXT, LB_GETTEXTLEN, LB_RESETCONTENT, LB_SETCURSEL, MINMAXINFO, MSG,
    SWP_NOZORDER, SW_SHOW, SW_SHOWNORMAL, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_CLOSE,
    WM_COMMAND, WM_CREATE, WM_CTLCOLORBTN, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX, WM_CTLCOLORSTATIC,
    WM_DESTROY, WM_DRAWITEM, WM_ERASEBKGND, WM_GETMINMAXINFO, WM_HSCROLL, WM_MEASUREITEM,
    WM_NCDESTROY, WM_PAINT, WM_SETFONT, WM_SIZE, WNDCLASSEXW, WS_BORDER, WS_CAPTION, WS_CHILD,
    WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_OVERLAPPED, WS_SYSMENU, WS_TABSTOP, WS_THICKFRAME,
    WS_VISIBLE, WS_VSCROLL,
};

use crate::config::config::MAX_PRESETS;
use crate::config::config::{CustomPreset, PreviewScale, TextTheme, TransparentBackground};
use crate::text::text_theme;
use crate::CONFIG;

/// The window class this window is made from.
const CLASS_NAME: windows::core::PCWSTR = w!("RustHoverPreviewCustomPreview");

/// How many of these windows may be open at once, which is one.
static OPEN: AtomicBool = AtomicBool::new(false);
/// The window of the one that is open, so that a second click brings it forward.
static WINDOW: AtomicIsize = AtomicIsize::new(0);

/// `TBM_GETPOS`: the trackbar's own position, which this binding does not name — the one message
/// of the four the window sends it that is missing from `Win32::UI::Controls`.
const TBM_GETPOS: u32 = 0x0408;

/// A message of this window's own: bring it forward (a second click on the tray row).
const WM_BRING_FORWARD: u32 = WM_APP + 17;

// The controls' ids. The two lists and the fields are told apart by these, and `WM_COMMAND` and
// `WM_DRAWITEM` both read them.
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
const ID_SAVE: i32 = 120;
const ID_DELETE: i32 = 121;
const ID_CLOSE_BUTTON: i32 = 122;
const ID_STATUS: i32 = 130;
const ID_LABEL_BASE: i32 = 140;

/// The scale a preset's combo lists, in the order it lists them: the fit first, then the three
/// shares of the room this app's own scaling menu offers. The fourth entry is `Custom…`, chosen by
/// the window itself for a percentage none of the three is.
const SCALE_CHOICES: [(&str, Option<u32>); 4] = [
    ("Fit to screen", None),
    ("Screen 100%", Some(100)),
    ("Screen 75%", Some(75)),
    ("Screen 50%", Some(50)),
];

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

/// The window's colours, read from the configured theme.
#[derive(Clone, Copy)]
struct Palette {
    page: COLORREF,
    panel: COLORREF,
    field: COLORREF,
    ink: COLORREF,
    accent: COLORREF,
}

fn rgb(r: u8, g: u8, b: u8) -> COLORREF {
    COLORREF(u32::from(r) | (u32::from(g) << 8) | (u32::from(b) << 16))
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
    /// The palette of `theme`: its page and its ink, and the shades the window is made of — a
    /// panel a shade off the page, a field a shade off that, and an accent the theme's own.
    fn of(theme: TextTheme) -> Self {
        let (page, ink, accent) = match text_theme::loaded(theme) {
            Some(loaded) => {
                let page = loaded.background();
                let ink = loaded.foreground();
                let accent = loaded
                    .style_for_scopes(&["keyword.control", "keyword"])
                    .foreground;
                (
                    rgb(page.r, page.g, page.b),
                    rgb(ink.r, ink.g, ink.b),
                    rgb(accent.r, accent.g, accent.b),
                )
            }
            None => (rgb(250, 250, 250), rgb(30, 30, 30), rgb(0, 90, 180)),
        };

        Self {
            page,
            panel: mix(page, ink, 7),
            field: mix(page, ink, 12),
            ink,
            accent,
        }
    }
}

/// Everything the window owns while it is up: the presets being edited, which one is showing,
/// and the controls.
struct State {
    hwnd: HWND,
    presets: Vec<CustomPreset>,
    /// The preset the fields are showing, as an index into `presets`.
    selected: Option<usize>,
    /// The presets the list is showing, as indexes into `presets` — the search's answer.
    shown: Vec<usize>,
    /// Whether the fields are being filled by the window itself, which is what keeps a
    /// programmatic `SetWindowTextW` from being read back as a keystroke.
    loading: bool,
    palette: Palette,
    font: HFONT,
    brush_page: HBRUSH,
    brush_field: HBRUSH,
    brush_panel: HBRUSH,
    dpi: i32,
}

impl Drop for State {
    fn drop(&mut self) {
        unsafe {
            let _ = DeleteObject(HGDIOBJ(self.font.0));
            let _ = DeleteObject(HGDIOBJ(self.brush_page.0));
            let _ = DeleteObject(HGDIOBJ(self.brush_field.0));
            let _ = DeleteObject(HGDIOBJ(self.brush_panel.0));
        }
    }
}

/// Open the window, or bring the one that is open forward.
///
/// The tray calls this from its own `WM_COMMAND`, so the window is made on a thread of its own
/// and this returns at once (see the module doc).
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

/// The presets the running configuration holds, in the order the window lists them.
fn presets_now() -> Vec<CustomPreset> {
    CONFIG
        .lock()
        .map(|config| config.custom_presets.clone())
        .unwrap_or_default()
}

/// The theme the window is painted in: the one `Tray → System → Theme` names.
fn theme_now() -> TextTheme {
    CONFIG
        .lock()
        .map(|config| config.theme)
        .unwrap_or(TextTheme::Light)
}

/// Write the window's list into the configuration and save it: what `Save`, `Delete` and the
/// window closing all do.
///
/// A name that cannot be a section, or one that repeats another, is refused rather than written
/// — the status line says which — and nothing of the file changes.
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
    unsafe {
        let hinstance = match GetModuleHandleW(None) {
            Ok(instance) => instance,
            Err(_) => return,
        };

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
            | WS_MAXIMIZEBOX;
        let Ok(hwnd) = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            CLASS_NAME,
            PCWSTR(title.as_ptr()),
            style,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            900,
            620,
            None,
            None,
            HINSTANCE(hinstance.0),
            None,
        ) else {
            return;
        };

        WINDOW.store(hwnd.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = UpdateWindow(hwnd);

        let mut message = MSG::default();
        while GetMessageW(&mut message, HWND::default(), 0, 0).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
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
            let raw = Box::into_raw(Box::new(state));
            unsafe { SetWindowLongPtrW(hwnd, GWLP_USERDATA, raw as isize) };
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
            info.ptMinTrackSize.x = scale(640, dpi);
            info.ptMinTrackSize.y = scale(420, dpi);
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
        WM_HSCROLL => {
            if let Some(state) = unsafe { state_of(hwnd) } {
                unsafe { volume_moved(state) };
            }
            LRESULT(0)
        }
        WM_MEASUREITEM => {
            let item = unsafe { &mut *(lparam.0 as *mut MEASUREITEMSTRUCT) };
            let dpi = unsafe { state_of(hwnd) }.map_or(96, |state| state.dpi);
            item.itemHeight = scale(22, dpi) as u32;
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
            unsafe {
                SetTextColor(hdc, state.palette.ink);
                SetBkColor(hdc, state.palette.page);
            }
            LRESULT(state.brush_page.0 as isize)
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
        WM_ERASEBKGND => {
            let Some(state) = (unsafe { state_of(hwnd) }) else {
                return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            };
            let hdc = HDC(wparam.0 as *mut _);
            unsafe {
                let mut client = RECT::default();
                let _ = GetClientRect(hwnd, &mut client);
                FillRect(hdc, &client, state.brush_page);
            }
            LRESULT(1)
        }
        WM_PAINT => {
            let mut paint = PAINTSTRUCT::default();
            let hdc = unsafe { BeginPaint(hwnd, &mut paint) };
            let _ = hdc;
            let _ = unsafe { EndPaint(hwnd, &paint) };
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

unsafe fn state_of(hwnd: HWND) -> Option<&'static mut State> {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut State;
    unsafe { raw.as_mut() }
}

fn scale(value: i32, dpi: i32) -> i32 {
    (value * dpi + 48) / 96
}

/// A control's style: the window bits every child wears, plus what the caller asked for.
fn child_style(extra: u32) -> WINDOW_STYLE {
    WINDOW_STYLE(WS_CHILD.0 | WS_VISIBLE.0 | extra)
}

unsafe fn build(hwnd: HWND) -> State {
    let dpi = (unsafe { GetDpiForWindow(hwnd) } as i32).max(96);
    let palette = Palette::of(theme_now());

    let font = unsafe {
        let face = wide("Segoe UI");
        CreateFontW(
            -scale(15, dpi),
            0,
            0,
            0,
            FW_NORMAL.0 as i32,
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
    };

    let mut state = State {
        hwnd,
        presets: presets_now(),
        selected: None,
        shown: Vec::new(),
        loading: false,
        palette,
        font,
        brush_page: unsafe { CreateSolidBrush(palette.page) },
        brush_field: unsafe { CreateSolidBrush(palette.field) },
        brush_panel: unsafe { CreateSolidBrush(palette.panel) },
        dpi,
    };

    let mut controls: Vec<HWND> = Vec::new();

    // The left column: the search, the `Add` beside it, and the tabs below them.
    controls.push(unsafe {
        child(
            hwnd,
            w!("EDIT"),
            "",
            child_style(WS_BORDER.0 | WS_TABSTOP.0),
            ID_SEARCH,
        )
    });
    controls.push(unsafe { button(hwnd, "Add", ID_ADD) });
    controls.push(unsafe {
        child(
            hwnd,
            w!("LISTBOX"),
            "",
            child_style(
                WS_BORDER.0
                    | WS_TABSTOP.0
                    | WS_VSCROLL.0
                    | LBS_NOTIFY as u32
                    | LBS_OWNERDRAWFIXED as u32
                    | LBS_HASSTRINGS as u32,
            ),
            ID_LIST,
        )
    });

    // The right pane: one label and one field per thing a preset holds.
    for index in 0..6 {
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
    controls.push(unsafe {
        child(
            hwnd,
            w!("EDIT"),
            "",
            child_style(WS_BORDER.0 | WS_TABSTOP.0),
            ID_NAME,
        )
    });
    controls.push(unsafe {
        child(
            hwnd,
            w!("EDIT"),
            "",
            child_style(WS_BORDER.0 | WS_TABSTOP.0),
            ID_EXTENSIONS,
        )
    });
    controls.push(unsafe {
        child(
            hwnd,
            w!("EDIT"),
            "",
            child_style(WS_BORDER.0 | WS_TABSTOP.0),
            ID_EXE,
        )
    });
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
            ID_SCALING,
        )
    });
    controls.push(unsafe {
        child(
            hwnd,
            w!("EDIT"),
            "",
            child_style(WS_BORDER.0 | WS_TABSTOP.0),
            ID_SCALING_PERCENT,
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
    controls.push(unsafe {
        child(
            hwnd,
            w!("EDIT"),
            "",
            child_style(WS_BORDER.0 | WS_TABSTOP.0),
            ID_VOLUME_VALUE,
        )
    });

    // The bottom row: the status line, and the three buttons.
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
                )
            };
        }
    }

    // The combo boxes' items and the trackbar's range, filled once.
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
    unsafe { lay_out(&state) };
    unsafe { rebuild_list(&mut state) };
    unsafe { fill_fields(&mut state) };

    state
}

/// One child control of the window, in the theme's own font later.
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

/// One owner-drawn button: the window paints it in the theme, so the id is its only text.
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

/// Place every control for the window's current client area: the tabs on the left, the fields on
/// the right, and the buttons along the bottom.
unsafe fn lay_out(state: &State) {
    let dpi = state.dpi;
    let mut client = RECT::default();
    let _ = unsafe { GetClientRect(state.hwnd, &mut client) };
    let width = client.right;
    let height = client.bottom;

    let pad = scale(14, dpi);
    let row = scale(30, dpi);
    let gap = scale(8, dpi);
    let label_width = scale(96, dpi);
    let column = scale(250, dpi);
    let footer = scale(46, dpi);

    let place = |id: i32, x: i32, y: i32, w: i32, h: i32| {
        if let Ok(control) = unsafe { GetDlgItem(state.hwnd, id) } {
            let _ = unsafe { SetWindowPos(control, HWND::default(), x, y, w, h, SWP_NOZORDER) };
        }
    };

    // The left column.
    place(ID_SEARCH, pad, pad, column - scale(52, dpi), row);
    place(
        ID_ADD,
        pad + column - scale(44, dpi),
        pad,
        scale(44, dpi),
        row,
    );
    place(
        ID_LIST,
        pad,
        pad + row + gap,
        column,
        (height - pad * 2 - row - gap - footer).max(row),
    );

    // The right pane, one row per field.
    let left = pad * 2 + column;
    let field_left = left + label_width;
    let field_width = (width - field_left - pad).max(scale(120, dpi));

    let labels = [
        "Name",
        "Extensions",
        "Program",
        "Scaling",
        "Background",
        "Volume",
    ];
    let mut y = pad;
    for (index, _) in labels.iter().enumerate() {
        place(
            ID_LABEL_BASE + index as i32,
            left,
            y + scale(6, dpi),
            label_width,
            row,
        );
        y += row + gap;
    }

    let mut y = pad;
    place(ID_NAME, field_left, y, field_width, row);
    y += row + gap;
    place(ID_EXTENSIONS, field_left, y, field_width, row);
    y += row + gap;
    place(
        ID_EXE,
        field_left,
        y,
        (field_width - scale(88, dpi)).max(scale(80, dpi)),
        row,
    );
    place(
        ID_BROWSE,
        field_left + field_width - scale(84, dpi),
        y,
        scale(84, dpi),
        row,
    );
    y += row + gap;
    place(
        ID_SCALING,
        field_left,
        y,
        (field_width - scale(88, dpi)).max(scale(80, dpi)),
        row,
    );
    place(
        ID_SCALING_PERCENT,
        field_left + field_width - scale(84, dpi),
        y,
        scale(84, dpi),
        row,
    );
    y += row + gap;
    place(ID_BACKGROUND, field_left, y, field_width, row);
    y += row + gap;
    place(
        ID_VOLUME,
        field_left,
        y,
        (field_width - scale(70, dpi)).max(scale(80, dpi)),
        row,
    );
    place(
        ID_VOLUME_VALUE,
        field_left + field_width - scale(66, dpi),
        y,
        scale(66, dpi),
        row,
    );

    // The footer.
    let buttons = scale(84, dpi);
    place(
        ID_CLOSE_BUTTON,
        width - pad - buttons,
        height - pad - row,
        buttons,
        row,
    );
    place(
        ID_DELETE,
        width - pad - buttons * 2 - gap,
        height - pad - row,
        buttons,
        row,
    );
    place(
        ID_SAVE,
        width - pad - buttons * 3 - gap * 2,
        height - pad - row,
        buttons,
        row,
    );
    place(
        ID_STATUS,
        pad,
        height - pad - row + scale(8, dpi),
        (width - pad * 2 - buttons * 3 - gap * 2 - scale(10, dpi)).max(scale(80, dpi)),
        row,
    );

    unsafe {
        let _ = InvalidateRect(state.hwnd, None, true);
    }
}

/// The window's own handling of a click or a keystroke, by the control it came from.
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

/// Fill the fields from the selected preset.
unsafe fn fill_fields(state: &mut State) {
    let Some(preset) = state.selected.and_then(|index| state.presets.get(index)) else {
        return;
    };
    let name = preset.name.clone();
    let extensions = preset.extensions_text();
    let exe = preset.exe.clone();
    let scaling = preset.scaling;
    let background = preset.background;
    let volume = preset.volume;

    state.loading = true;
    unsafe {
        set_text(state.hwnd, ID_NAME, &name);
        set_text(state.hwnd, ID_EXTENSIONS, &extensions);
        set_text(state.hwnd, ID_EXE, &exe);

        let combo = GetDlgItem(state.hwnd, ID_SCALING).unwrap_or_default();
        let choice = match scaling {
            PreviewScale::FitToScreen => 0,
            PreviewScale::FitToScreenReduced(100) => 1,
            PreviewScale::FitToScreenReduced(75) => 2,
            PreviewScale::FitToScreenReduced(50) => 3,
            _ => SCALE_CHOICES.len() as i32,
        };
        SendMessageW(combo, CB_SETCURSEL, WPARAM(choice as usize), LPARAM(0));

        let percent = match scaling {
            PreviewScale::FitToScreen => 100,
            PreviewScale::FitToScreenReduced(percent) | PreviewScale::Percent(percent) => percent,
        };
        set_text(state.hwnd, ID_SCALING_PERCENT, &percent.to_string());
        if let Ok(edit) = GetDlgItem(state.hwnd, ID_SCALING_PERCENT) {
            let _ = EnableWindow(edit, !matches!(scaling, PreviewScale::FitToScreen));
        }

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
}

/// The tabs the search box asks for, as indexes into the preset list.
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

/// A tab was clicked: the fields become that preset's.
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

/// A new preset, named so that the list can hold it and selected so that it can be typed into.
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

/// What a keystroke in one of the text fields changes: the preset the fields are showing.
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

    let Some((_, percent)) = SCALE_CHOICES.get(choice as usize).copied() else {
        // `Custom…`: the percentage field is the one that answers, and it already does.
        if let Ok(edit) = unsafe { GetDlgItem(state.hwnd, ID_SCALING_PERCENT) } {
            let _ = unsafe { EnableWindow(edit, true) };
            let _ = unsafe { SetFocus(edit) };
        }
        return;
    };

    if let Some(preset) = state.presets.get_mut(index) {
        preset.scaling = match percent {
            None => PreviewScale::FitToScreen,
            Some(percent) => PreviewScale::FitToScreenReduced(percent),
        };
    }

    state.loading = true;
    unsafe {
        set_text(
            state.hwnd,
            ID_SCALING_PERCENT,
            &percent.unwrap_or(100).to_string(),
        );
    }
    state.loading = false;
    if let Ok(edit) = unsafe { GetDlgItem(state.hwnd, ID_SCALING_PERCENT) } {
        let _ = unsafe { EnableWindow(edit, percent.is_some()) };
    }
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
        preset.scaling = PreviewScale::Percent(percent.clamp(1, 1000));
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
    let position =
        unsafe { SendMessageW(trackbar, TBM_GETPOS, WPARAM(0), LPARAM(0)).0 }.clamp(0, 100) as u32;

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
    let volume = volume.min(100);
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

/// The `Browse…` button: the platform's own open dialog, which is the only way to name a program
/// on a machine without typing its path.
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

fn is_button(id: u32) -> bool {
    id == ID_ADD as u32
        || id == ID_SAVE as u32
        || id == ID_DELETE as u32
        || id == ID_BROWSE as u32
        || id == ID_CLOSE_BUTTON as u32
}

/// One owner-drawn control: a tab, a combo item, or a button, each in the theme's own colours.
unsafe fn draw_item(state: &State, item: &DRAWITEMSTRUCT) -> usize {
    let hdc = item.hDC;
    let palette = state.palette;
    let selected = item.itemState.0 & ODS_SELECTED.0 != 0;

    let fill = if is_button(item.CtlID) {
        if selected {
            palette.accent
        } else {
            palette.panel
        }
    } else if item.CtlID == ID_LIST as u32 && selected {
        palette.accent
    } else {
        palette.field
    };

    unsafe {
        let brush = CreateSolidBrush(fill);
        FillRect(hdc, &item.rcItem, brush);
        let _ = DeleteObject(HGDIOBJ(brush.0));

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(
            hdc,
            if fill == palette.accent {
                mix(palette.page, palette.ink, 10)
            } else {
                palette.ink
            },
        );
        let _ = SelectObject(hdc, HGDIOBJ(state.font.0));

        // The text is the control's own: a listbox or combo item by its index, a button by the
        // text the control was made with.
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
            id if id == ID_SCALING as u32 || id == ID_BACKGROUND as u32 => {
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

        let mut text: Vec<u16> = text.encode_utf16().collect();
        let mut rect = item.rcItem;
        let flags = if is_button(item.CtlID) {
            DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
        } else {
            rect.left += scale(8, state.dpi);
            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
        };
        DrawTextW(hdc, &mut text, &mut rect, flags);
    }

    1
}

fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

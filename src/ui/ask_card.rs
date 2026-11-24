//! The card a question is asked in, over the window it belongs to, instead of the platform's
//! message box.
//!
//! A message box is the system's own window: its font, its colours, its buttons named `Yes` and
//! `No`, and the beep it plays on the machine's schedule. None of that is this app's look, so the
//! one question that is *about a window of this app's own* — the delete the `Custom Preview`
//! editor asks — is asked here instead: a scrim over that whole window and a card of this app's on
//! it, drawn from the same `Palette` the editor draws itself from, so it is One Dark Pro with that
//! theme and Atom One Light with the other.
//!
//! **A window of its own, not a control tree.** The card is one borderless layered popup covering the
//! owner's client area exactly, brought up with `SW_SHOWNORMAL` so that it takes the keyboard and the
//! foreground — `Enter` and `Escape` are the two keys a question is answered with, and a
//! `WS_EX_NOACTIVATE` window answers neither. It is *owned* by the editor, so it stands above it and
//! is destroyed with it, and the editor is disabled for the length of the question so that a press
//! that lands behind the card is a press on nothing.
//!
//! **Nothing is drawn, and nothing is sounded, by the system.** A layered window's pixels are its
//! own: the platform paints none of them, dims nothing for it, fades it in, or beeps at it. The scrim
//! is *here* rather than in the editor's own `paint_window` for that reason — a dim laid over the
//! owner would be a repaint of the owner under the card, and the same picture composed into this
//! window is one window's worth of work rather than two.
//!
//! **The pixels are premultiplied, and the text is a mask.** A shadow and the card's rounded corners
//! live in the *alpha* byte, which GDI never writes and a layered window composites over the
//! desktop, so the card is composed in this app's own premultiplied-BGRA buffer and handed to
//! `UpdateLayeredWindow` in one surface (see `Canvas`). The one thing drawn through GDI is the text,
//! which keeps its hinting: a run is drawn white on black in a scratch DIB and read back as coverage
//! (`Canvas::text`), so the ink stays the palette's rather than the mask's. That font is the card's own
//! — normal weight, and **grayscale** anti-aliased (`text_font`) — because a mask is one alpha per
//! pixel and ClearType's three subpixel coverages are not; a run drawn with the editor's own quality
//! came out heavier here than the same run in the window beside it.
//!
//! **Every shape is one fill inside another.** A border is the shape drawn in its border colour with the
//! card's own box a hairline inside it, and a button is the same — never a band walked along the
//! outline, which along an arc a few pixels across is a diagonal line thinner than a pixel and puts holes
//! in the corners. The outlines are sampled four times across the pixel (`covered`) for the same reason.
//!
//! **The safe answer is the default.** `Enter` and `Space` answer what the keyboard is on, which is
//! the button that changes nothing until `Tab` or either arrow moves it — exactly as `MB_DEFBUTTON2`
//! made the platform's box behave — and `Escape` answers it whatever is focused. The ring round a
//! button is the keyboard's, not a claim that it is the button to press. Leaving the card for another
//! window is leaving the question, and that also answers what changes nothing.

use std::cell::Cell;
use std::ffi::c_void;
use std::ptr;
use std::sync::Once;

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, POINT, RECT, SIZE, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, ClientToScreen, CreateCompatibleDC, CreateDIBSection, CreateFontW, DeleteDC,
    DeleteObject, DrawTextW, EndPaint, GetDC, GetStockObject, InvalidateRect, ReleaseDC,
    SelectObject, SetBkMode, SetTextColor, AC_SRC_ALPHA, AC_SRC_OVER, ANTIALIASED_QUALITY,
    BITMAPINFO, BITMAPINFOHEADER, BI_RGB, BLENDFUNCTION, CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET,
    DEFAULT_PITCH, DIB_RGB_COLORS, DRAW_TEXT_FORMAT, DT_CALCRECT, DT_CENTER, DT_LEFT, DT_NOPREFIX,
    DT_SINGLELINE, DT_VCENTER, DT_WORDBREAK, FW_NORMAL, FW_SEMIBOLD, HBRUSH, HDC, HFONT, HGDIOBJ,
    NULL_BRUSH, OUT_DEFAULT_PRECIS, PAINTSTRUCT, TRANSPARENT,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::HiDpi::GetDpiForWindow;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetKeyState, ReleaseCapture, SetCapture, SetFocus, TrackMouseEvent, TME_LEAVE,
    TRACKMOUSEEVENT, VK_ESCAPE, VK_F4, VK_LEFT, VK_MENU, VK_RETURN, VK_RIGHT, VK_SHIFT, VK_SPACE,
    VK_TAB,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, GetClientRect,
    GetForegroundWindow, GetMessageW, GetWindowLongPtrW, IsWindow, LoadCursorW, PostQuitMessage,
    RegisterClassExW, SetCursor, SetForegroundWindow, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    TranslateMessage, UpdateLayeredWindow, GWLP_USERDATA, HWND_TOP, IDC_ARROW, IDC_HAND, MSG,
    SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SW_SHOWNORMAL, ULW_ALPHA, WA_ACTIVE, WA_CLICKACTIVE,
    WA_INACTIVE, WM_ACTIVATE, WM_CAPTURECHANGED, WM_CHAR, WM_CLOSE, WM_DESTROY, WM_ERASEBKGND,
    WM_KEYDOWN, WM_KILLFOCUS, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_NCDESTROY, WM_PAINT,
    WM_SETCURSOR, WM_SETFOCUS, WM_SYSCHAR, WM_SYSKEYDOWN, WNDCLASSEXW, WNDCLASS_STYLES,
    WS_EX_LAYERED, WS_EX_TOOLWINDOW, WS_POPUP,
};

use super::custom_preview_settings::{mix, pt_in_rect, rgb, scale, theme_now, Palette};

/// The class the card's window is created from.
const CLASS_NAME: PCWSTR = w!("RustHoverPreviewAskCard");

/// The message `TrackMouseEvent` sends where there is no `WM_MOUSELEAVE` class: the one that says the
/// pointer has left the card, which is what puts a hovered button back.
const WM_MOUSELEAVE: u32 = 0x02A3;

/// How much of the page the scrim leaves showing: the window behind the card is read as *behind* it
/// rather than as gone, and the card as the only thing in front.
const SCRIM_SHARE: f32 = 0.56;

/// The card's own width where the window has the room, its corner radius, and the shadow it casts:
/// it is the editor's material — a rounded box with a soft shadow — and a question is made of it.
const CARD_WIDTH_LOGICAL: i32 = 420;
const CARD_RADIUS_LOGICAL: i32 = 14;
const CARD_SHADOW_LOGICAL: i32 = 22;

/// The corner radius of the window the card is asked over, which DWM's `DWMWCP_ROUND` gives the editor:
/// the scrim is cut to it so that it covers the window's own shape rather than a rectangle around it.
const CARD_CORNER_LOGICAL: i32 = 8;

/// The weights the card is written in: the editor's own title is semibold at 14, and a question is
/// a weight above it.
const HEADING_LOGICAL: i32 = 16;
const BODY_LOGICAL: i32 = 13;
/// The size a button's word is written at: the editor's own body size, because one pixel more is the
/// pixel the hinting answers with a two-pixel stem — which is how a word at 14 came out heavier than the
/// very same word at 13 in the sentence above it.
const BUTTON_LOGICAL: i32 = 13;

static CLASS_REGISTERED: Once = Once::new();

/// The card this process put up last, for the test that answers one from another thread: the call
/// that asks does not come back until the card is gone, so the keys and the clicks that answer it
/// have to be sent from somewhere else (`the_card_is_answered_by_a_key_or_by_a_click`).
#[cfg(test)]
static LAST_CARD: std::sync::atomic::AtomicIsize = std::sync::atomic::AtomicIsize::new(0);

thread_local! {
    /// What the card of this thread was answered with, and whether one is up.
    ///
    /// The answer is read after the card's message loop has come back rather than out of the window,
    /// because the window is destroyed on the way there: the click that answers `Delete` destroys the
    /// card it was clicked in, and the state that window held goes with it.
    static ANSWER: Cell<Option<bool>> = const { Cell::new(None) };
}

/// What one of the card's two buttons answers.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Decision {
    /// The button that goes ahead — `Delete`.
    Affirmative,
    /// The button that changes nothing — `Cancel`.
    Negative,
}

/// A question: what it is called, what it says, and the word on the button that goes ahead.
pub(super) struct Question {
    title: String,
    message: String,
    affirmative: String,
}

impl Question {
    pub(super) fn new(title: &str, message: &str, affirmative: &str) -> Self {
        Self {
            title: title.to_string(),
            message: message.to_string(),
            affirmative: affirmative.to_string(),
        }
    }
}

/// Whether a card of this thread is on screen right now.
pub(super) fn is_asking() -> bool {
    ASKING.with(|asking| asking.get())
}

thread_local! {
    static ASKING: Cell<bool> = const { Cell::new(false) };
}

/// Ask `question` over `owner`, and answer with what was said: `true` for the affirmative button,
/// `false` for the one that changes nothing, and nothing at all where the card could not be made.
///
/// The call does not come back until the card is answered, which is what makes it a question rather
/// than a notification. **The answer is the call's own**: it is taken out of `ANSWER` here, and the
/// window is gone by the time this returns — so a caller that needs the answer to outlive the card
/// takes it from this, never from a message posted about it.
pub(super) fn ask(owner: HWND, question: &Question) -> Option<bool> {
    ANSWER.with(|answer| answer.set(None));

    let hwnd = unsafe { create(owner, question) }?;

    ASKING.with(|asking| asking.set(true));
    // The tray's own rule while a dialog of this app's is up: no menu is opened over one (see
    // `app::dialogs`), and a card is a dialog for that purpose exactly as a message box was.
    crate::app::dialogs::begin();

    unsafe {
        let _ = EnableWindow(owner, false);
        let _ = SetForegroundWindow(hwnd);
        let _ = SetFocus(hwnd);
    }

    let mut message = MSG::default();
    while unsafe { GetMessageW(&mut message, HWND::default(), 0, 0) }.as_bool() {
        unsafe {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    unsafe {
        let _ = EnableWindow(owner, true);
        let _ = SetForegroundWindow(owner);
    }

    // The card runs a message loop of its own, and a `WM_QUIT` meant for the editor — its window going
    // away while the question was up — was taken by this loop rather than by the one that owns it. So
    // where the owner is gone, the quit is posted again: a loop with no window left to answer would
    // otherwise wait on nothing at all.
    if !unsafe { IsWindow(owner) }.as_bool() {
        unsafe { PostQuitMessage(0) };
    }

    crate::app::dialogs::end();
    ASKING.with(|asking| asking.set(false));

    ANSWER.with(|answer| answer.take())
}

/// The card's own state, held by its window for as long as it is up.
struct Card {
    hwnd: HWND,
    /// The window the question is about, which is also the window this one is owned by: a foreground it
    /// takes back from under the card is the question still being asked, not the user leaving it.
    owner: HWND,
    question: Question,
    palette: Palette,
    dpi: i32,
    /// The screen box this window covers: the owner's client area, which is the scrim.
    frame: RECT,
    /// The scrim's own size, which is every box below's coordinate space.
    width: i32,
    height: i32,
    /// Where the card stands, and what it holds, in this window's pixels. Measured once, when the
    /// card is made: a card cannot be resized while it is up.
    body: Rect,
    heading: RECT,
    message: RECT,
    badge: [i32; 3],
    /// The two buttons left to right, which is the picture's order rather than the keyboard's: the
    /// one that changes nothing sits beside the one that does.
    buttons: [(RECT, Decision); 2],
    /// The button the pointer is over, the one a press landed on, and the one the keyboard is on.
    hovered: Option<Decision>,
    pressed: Option<Decision>,
    focused: Decision,
    tracking: bool,
    /// Whether the card has been activated: the loss of a foreground it never took is not the user
    /// leaving the question.
    activated: bool,
    /// Whether the question has been answered. The card is destroyed on the answer, and destroying a
    /// window deactivates it — so the `WM_ACTIVATE` that follows the answer would otherwise answer the
    /// question a second time, with the button that changes nothing, and overwrite the answer the first
    /// one gave.
    answered: bool,
    fonts: Fonts,
}

impl Card {
    /// The button whose box holds a point of this window, if any.
    fn button_at(&self, x: i32, y: i32) -> Option<Decision> {
        self.buttons
            .iter()
            .find(|(rect, _)| pt_in_rect(*rect, x, y))
            .map(|(_, decision)| *decision)
    }

    /// What a button carries: its word, or the answer that changes nothing.
    fn label(&self, decision: Decision) -> &str {
        match decision {
            Decision::Affirmative => &self.question.affirmative,
            Decision::Negative => "Cancel",
        }
    }
}

/// The three runs of the card's text.
struct Fonts {
    heading: HFONT,
    body: HFONT,
    button: HFONT,
}

impl Fonts {
    fn make(dpi: i32) -> Self {
        Self {
            heading: text_font(dpi, HEADING_LOGICAL, FW_SEMIBOLD.0 as i32),
            body: text_font(dpi, BODY_LOGICAL, FW_NORMAL.0 as i32),
            button: text_font(dpi, BUTTON_LOGICAL, FW_NORMAL.0 as i32),
        }
    }
}

/// A run's font, made for a **mask** rather than for GDI to paint: grayscale anti-aliasing, so that
/// what is read back is the glyph's coverage and not a subpixel pattern.
///
/// The editor draws straight through GDI and takes ClearType for it; this card does not draw text at
/// all — it reads a run back as one alpha per pixel (`Canvas::text`). ClearType renders a glyph as three
/// coloured coverages, and any single number taken from them is not the coverage the glyph has: the
/// largest channel of a ClearType edge is the ink at full strength, which is why a run drawn that way
/// came out heavier and blobbier than the same run in the window beside it. Grayscale anti-aliasing is
/// one coverage per pixel, which is exactly what the mask is.
fn text_font(dpi: i32, size: i32, weight: i32) -> HFONT {
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
            ANTIALIASED_QUALITY.0 as u32,
            DEFAULT_PITCH.0 as u32,
            PCWSTR(face.as_ptr()),
        )
    }
}

impl Drop for Fonts {
    fn drop(&mut self) {
        for font in [self.heading, self.body, self.button] {
            if !font.0.is_null() {
                unsafe {
                    let _ = DeleteObject(HGDIOBJ(font.0));
                }
            }
        }
    }
}

/// A string as the Win32 calls want it: UTF-16, terminated.
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

fn card_of(hwnd: HWND) -> Option<&'static mut Card> {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut Card;
    unsafe { raw.as_mut() }
}

/// Make the card's window over `owner`, painted on the first `WM_PAINT`.
unsafe fn create(owner: HWND, question: &Question) -> Option<HWND> {
    CLASS_REGISTERED.call_once(|| {
        let Ok(hinstance) = (unsafe { GetModuleHandleW(None) }) else {
            return;
        };
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: WNDCLASS_STYLES(0),
            lpfnWndProc: Some(card_proc),
            hInstance: hinstance.into(),
            hCursor: unsafe { LoadCursorW(None, IDC_ARROW) }.unwrap_or_default(),
            hbrBackground: HBRUSH(unsafe { GetStockObject(NULL_BRUSH) }.0),
            lpszClassName: CLASS_NAME,
            ..Default::default()
        };
        unsafe { RegisterClassExW(&class) };
    });

    let dpi = (unsafe { GetDpiForWindow(owner) } as i32).max(96);
    let palette = Palette::of(theme_now());
    let fonts = Fonts::make(dpi);

    // The scrim is the owner's client area rather than its window: the card is a question *about*
    // that window's contents, and a scrim over its frame would reach past what it belongs to.
    let mut client = RECT::default();
    let _ = unsafe { GetClientRect(owner, &mut client) };
    let width = (client.right - client.left).max(1);
    let height = (client.bottom - client.top).max(1);
    let mut origin = POINT { x: 0, y: 0 };
    let _ = unsafe { ClientToScreen(owner, &mut origin) };
    let frame = RECT {
        left: origin.x,
        top: origin.y,
        right: origin.x + width,
        bottom: origin.y + height,
    };

    let layout = layout(question, &fonts, dpi, width, height);

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW,
            CLASS_NAME,
            w!("Rust Hover Preview"),
            WS_POPUP,
            frame.left,
            frame.top,
            width,
            height,
            owner,
            None,
            None,
            None,
        )
    }
    .ok()?;

    let card = Box::new(Card {
        hwnd,
        owner,
        question: Question::new(&question.title, &question.message, &question.affirmative),
        palette,
        dpi,
        frame,
        width,
        height,
        body: layout.body,
        heading: layout.heading,
        message: layout.message,
        badge: layout.badge,
        buttons: layout.buttons,
        hovered: None,
        pressed: None,
        // The keyboard starts on the answer that changes nothing, which is what `MB_DEFBUTTON2`
        // made the platform's box do.
        focused: Decision::Negative,
        tracking: false,
        activated: false,
        answered: false,
        fonts,
    });
    unsafe { SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(card) as isize) };
    #[cfg(test)]
    LAST_CARD.store(hwnd.0 as isize, std::sync::atomic::Ordering::SeqCst);

    unsafe {
        let _ = ShowWindow(hwnd, SW_SHOWNORMAL);
        let _ = SetWindowPos(
            hwnd,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        if let Some(card) = card_of(hwnd) {
            paint_card(card);
        }
    }

    Some(hwnd)
}

/// A box of the card's picture: floats while it is measured, and pixels once it is placed, because
/// the two are the same rectangle named by different things.
#[derive(Clone, Copy, Debug, PartialEq)]
struct Rect {
    left: f32,
    top: f32,
    right: f32,
    bottom: f32,
}

impl Rect {
    fn around(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left: left as f32,
            top: top as f32,
            right: right as f32,
            bottom: bottom as f32,
        }
    }

    fn width(self) -> f32 {
        (self.right - self.left).max(0.0)
    }

    fn height(self) -> f32 {
        (self.bottom - self.top).max(0.0)
    }

    /// The same box, `by` inside every edge.
    fn inset(self, by: f32) -> Self {
        Self {
            left: self.left + by,
            top: self.top + by,
            right: (self.right - by).max(self.left + by),
            bottom: (self.bottom - by).max(self.top + by),
        }
    }

    /// The same box as the `RECT` the editor's own geometry is measured in, which is what the
    /// placement test asks it against.
    #[cfg(test)]
    fn pixels(self) -> RECT {
        RECT {
            left: self.left.round() as i32,
            top: self.top.round() as i32,
            right: self.right.round() as i32,
            bottom: self.bottom.round() as i32,
        }
    }
}

/// Where the card, its heading, its body and its two buttons stand in the scrim.
struct Layout {
    body: Rect,
    heading: RECT,
    message: RECT,
    badge: [i32; 3],
    buttons: [(RECT, Decision); 2],
}

/// Measure the card the question makes, centred on the scrim.
///
/// The width is the card's own rather than the room's, because a line of a question does not want to
/// be a window's width: `CARD_WIDTH_LOGICAL` is asked for where the window is wider, and the card is
/// the window less its margins where it is not. The height is the sum of what the card holds, so a
/// question with a longer body is a taller card rather than a scrolled one — and the card is the
/// window less its margins in that direction too, in a window too short for its own question.
fn layout(question: &Question, fonts: &Fonts, dpi: i32, frame_w: i32, frame_h: i32) -> Layout {
    let screen = unsafe { GetDC(HWND::default()) };
    let pad = scale(22, dpi) as f32;
    let gap = scale(14, dpi) as f32;
    let badge_r = scale(15, dpi).max(10) as f32;
    let button_h = scale(32, dpi) as f32;
    let button_gap = scale(10, dpi) as f32;
    let button_pad = scale(40, dpi) as f32;
    let button_min = scale(96, dpi) as f32;

    let card_w = (scale(CARD_WIDTH_LOGICAL, dpi) as f32)
        .min((frame_w as f32 - scale(56, dpi) as f32).max(scale(240, dpi) as f32))
        .min(frame_w.max(1) as f32)
        .max(1.0);
    let text_w = (card_w - pad * 2.0).max(1.0) as i32;

    let (_, heading_h) = measure(
        screen,
        fonts.heading,
        &question.title,
        text_w,
        DT_SINGLELINE,
    );
    let message_run = wrap_run(screen, fonts.body, &question.message, text_w);
    let (_, message_h) = measure(screen, fonts.body, &message_run, text_w, DT_WORDBREAK);
    let (affirmative_w, _) = measure(
        screen,
        fonts.button,
        &question.affirmative,
        scale(2000, dpi),
        DT_SINGLELINE,
    );
    let (negative_w, _) = measure(
        screen,
        fonts.button,
        "Cancel",
        scale(2000, dpi),
        DT_SINGLELINE,
    );
    let affirmative_w = (affirmative_w as f32 + button_pad).max(button_min);
    let negative_w = (negative_w as f32 + button_pad).max(button_min);
    let _ = unsafe { ReleaseDC(HWND::default(), screen) };

    // The badge is a disc in the card's top-left corner and the heading is written under it, so the
    // card's first row is the disc and the heading is its own line below.
    let head_h = heading_h as f32;
    let card_h =
        (pad + badge_r * 2.0 + gap + head_h + gap + message_h as f32 + gap * 1.5 + button_h + pad)
            .min((frame_h as f32 - scale(40, dpi) as f32).max(scale(120, dpi) as f32))
            .min(frame_h.max(1) as f32)
            .max(1.0);

    let left = ((frame_w as f32 - card_w) * 0.5).round();
    let top = ((frame_h as f32 - card_h) * 0.5).round();
    let body = Rect {
        left,
        top,
        right: left + card_w,
        bottom: top + card_h,
    };

    let badge = [
        (body.left + pad + badge_r).round() as i32,
        (body.top + pad + badge_r).round() as i32,
        badge_r.round() as i32,
    ];
    let heading_top = (body.top + pad + badge_r * 2.0 + gap).round() as i32;
    let heading = RECT {
        left: (body.left + pad).round() as i32,
        top: heading_top,
        right: (body.right - pad).round() as i32,
        bottom: heading_top + heading_h,
    };

    let message_top = (heading.bottom as f32 + gap).round() as i32;
    let buttons_top = ((body.bottom - pad - button_h).round() as i32).max(message_top);
    let message = RECT {
        left: heading.left,
        top: message_top,
        right: heading.right,
        bottom: ((buttons_top as f32 - gap * 1.5).round() as i32).max(message_top),
    };

    let primary = RECT {
        left: (body.right - pad - affirmative_w).round() as i32,
        top: buttons_top,
        right: (body.right - pad).round() as i32,
        bottom: (buttons_top as f32 + button_h).round() as i32,
    };
    let secondary = RECT {
        left: (primary.left as f32 - button_gap - negative_w).round() as i32,
        top: buttons_top,
        right: (primary.left as f32 - button_gap).round() as i32,
        bottom: primary.bottom,
    };

    Layout {
        body,
        heading,
        message,
        badge,
        buttons: [
            (secondary, Decision::Negative),
            (primary, Decision::Affirmative),
        ],
    }
}

/// The body as it is drawn: the question's own lines, and each of them broken to the card's width.
///
/// `DT_WORDBREAK` breaks a line at its spaces and nowhere else, so one unbroken run — which a program
/// is, a path being one word to a break routine — is a line that runs off the card and is clipped
/// rather than wrapped. So the breaking is this app's own: a word at a time where it fits, and inside
/// the word at its own separators where it does not, which is what makes a long path wrap at a
/// backslash rather than mid-name.
fn wrap_run(dc: HDC, font: HFONT, run: &str, width: i32) -> String {
    let width = width.max(1);
    let mut out = String::with_capacity(run.len() + run.len() / 4);
    for (index, paragraph) in run.split('\n').enumerate() {
        if index > 0 {
            out.push('\n');
        }
        let mut line = String::new();
        for word in paragraph.split(' ').filter(|word| !word.is_empty()) {
            let candidate = if line.is_empty() {
                word.to_string()
            } else {
                format!("{line} {word}")
            };
            if fits(dc, font, &candidate, width) {
                line = candidate;
                continue;
            }
            if !line.is_empty() {
                out.push_str(&line);
                out.push('\n');
                line.clear();
            }
            // The word on a line of its own, and still too wide: it is broken at its own separators.
            let mut rest = word;
            while !fits(dc, font, rest, width) {
                let cut = break_at(dc, font, rest, width);
                out.push_str(&rest[..cut]);
                out.push('\n');
                rest = &rest[cut..];
                if rest.is_empty() {
                    break;
                }
            }
            line = rest.to_string();
        }
        out.push_str(&line);
    }
    out
}

/// Whether a run of text is no wider than `width`.
fn fits(dc: HDC, font: HFONT, run: &str, width: i32) -> bool {
    measure(dc, font, run, scale(2000, 96).max(width), DT_SINGLELINE).0 <= width
}

/// Where a word that is wider than the card is broken: as far along as fits, and after one of the
/// word's own separators where there is one, so a path breaks at a backslash and not mid-name.
fn break_at(dc: HDC, font: HFONT, word: &str, width: i32) -> usize {
    let mut fitting = 0;
    let mut after_a_separator = 0;
    for (index, unit) in word.char_indices() {
        let next = index + unit.len_utf8();
        if next == word.len() || fits(dc, font, &word[..next], width) {
            fitting = next;
            if next == word.len() {
                break;
            }
            if matches!(unit, '\\' | '/' | '-' | '.' | '_') {
                after_a_separator = next;
            }
        } else {
            break;
        }
    }
    after_a_separator.max(fitting).max(1)
}

/// The box a run of text takes at a font and a width, which is `DrawTextW`'s own answer to
/// `DT_CALCRECT` rather than an estimate of it.
fn measure(dc: HDC, font: HFONT, run: &str, width: i32, flags: DRAW_TEXT_FORMAT) -> (i32, i32) {
    if run.is_empty() || dc.0.is_null() || width <= 0 {
        return (0, 0);
    }

    let mut text: Vec<u16> = run.encode_utf16().collect();
    let mut rect = RECT {
        left: 0,
        top: 0,
        right: width,
        bottom: 0,
    };
    unsafe {
        let old = SelectObject(dc, HGDIOBJ(font.0));
        DrawTextW(dc, &mut text, &mut rect, flags | DT_CALCRECT);
        let _ = SelectObject(dc, old);
    }
    (rect.right - rect.left, rect.bottom - rect.top)
}

/// The destructive button's gradient: `(top, bottom, border, ink)`.
///
/// It is the palette's own red rather than the accent, which is the one thing that makes a delete
/// read as a delete. Both themes keep the ink legible on *both* stops — the light theme's red is
/// darkened for it, and the dark theme's is lightened where it is pressed rather than darkened,
/// since darkening it is what would cost the ink its contrast.
fn affirmative_colors(
    p: Palette,
    hovered: bool,
    pressed: bool,
) -> (COLORREF, COLORREF, COLORREF, COLORREF) {
    let (top, bottom) = if p.dark {
        if pressed {
            (
                mix(p.danger, rgb(255, 255, 255), 26),
                mix(p.danger, rgb(255, 255, 255), 10),
            )
        } else if hovered {
            (
                mix(p.danger, rgb(255, 255, 255), 18),
                mix(p.danger, rgb(255, 255, 255), 4),
            )
        } else {
            (
                mix(p.danger, rgb(255, 255, 255), 8),
                mix(p.danger, rgb(0, 0, 0), 4),
            )
        }
    } else if pressed {
        (
            mix(p.danger, rgb(0, 0, 0), 30),
            mix(p.danger, rgb(0, 0, 0), 38),
        )
    } else if hovered {
        (
            mix(p.danger, rgb(0, 0, 0), 16),
            mix(p.danger, rgb(0, 0, 0), 24),
        )
    } else {
        (
            mix(p.danger, rgb(0, 0, 0), 20),
            mix(p.danger, rgb(0, 0, 0), 30),
        )
    };
    let border = if p.dark {
        mix(p.danger, rgb(255, 255, 255), 20)
    } else {
        mix(p.danger, rgb(0, 0, 0), 40)
    };
    (top, bottom, border, p.danger_on_fill)
}

/// The card's other button: the editor's own secondary button, as `draw_item` paints one.
fn negative_colors(
    p: Palette,
    hovered: bool,
    pressed: bool,
) -> (COLORREF, COLORREF, COLORREF, COLORREF) {
    let (top, bottom) = if pressed {
        (p.button_press, p.button_press)
    } else if hovered {
        (p.button_hover, p.button_hover)
    } else if p.dark {
        (mix(p.button_bg, rgb(255, 255, 255), 4), p.button_bg)
    } else {
        (mix(p.button_bg, rgb(255, 255, 255), 40), p.button_bg)
    };
    // The ink is the palette's own where it reads on every stop, and a step brighter in the dark
    // theme where it does not: `p.ink` against the lighter of the two fills is under the 4.5 the
    // button's word has to keep.
    let ink = if p.dark {
        mix(p.ink, rgb(255, 255, 255), 45)
    } else {
        mix(p.ink, rgb(0, 0, 0), 12)
    };
    (top, bottom, p.border, ink)
}

/// Put the card's whole surface on the screen.
unsafe fn paint_card(card: &Card) {
    if let Some(canvas) = compose(card) {
        unsafe { present(card, canvas.pixels()) };
    }
}

/// Draw the card and its scrim into a surface of the scrim's own pixels, before anything is put on
/// the screen: the window paints from here, and so does the test that reads the picture back
/// (`render_the_card`) — one composition for both.
fn compose(card: &Card) -> Option<Canvas> {
    let mut canvas = Canvas::new(card.width, card.height)?;
    let p = card.palette;
    let dpi = card.dpi;
    let s = dpi as f32 / 96.0;

    // The scrim: the page the editor is drawn on, over the whole of it, so that what is behind the
    // card is read as behind it rather than as something else to press. It is rounded at the corners
    // the editor itself is rounded at (`DWMWCP_ROUND`), because this window's box is that window's box:
    // a square scrim over a rounded window is four small corners of somebody else's desktop painted
    // over.
    canvas.fill(
        Rect::around(0, 0, card.width, card.height),
        scale(CARD_CORNER_LOGICAL, dpi) as f32,
        p.page,
        SCRIM_SHARE,
    );

    // The card: a shadow, the card's own box shaded from a sheen at its top to its own colour, and
    // its one hairline.
    let body = card.body;
    let radius = CARD_RADIUS_LOGICAL as f32 * s;
    canvas.shadow(
        body,
        radius,
        CARD_SHADOW_LOGICAL as f32 * s,
        if p.dark { 0.40 } else { 0.16 },
    );
    let (sheen, floor) = if p.dark {
        (mix(p.card, rgb(255, 255, 255), 6), p.card)
    } else {
        (mix(p.card, rgb(255, 255, 255), 40), p.card)
    };
    // **A frame is a fill and a fill inside it**, not a band walked along the edge: the shape's own
    // anti-aliasing draws both rings, so the border follows the corner arc as exactly as the card does.
    // A one-pixel band along an arc a few pixels across is a diagonal line thinner than a pixel, and no
    // sampler — one point, four, or sixteen — can put ink in every pixel it crosses.
    let hairline = (s).max(1.0);
    canvas.fill(body, radius, p.border, 1.0);
    canvas.gradient(
        body.inset(hairline),
        (radius - hairline).max(0.0),
        sheen,
        floor,
    );

    // The badge the question is marked with: the palette's red in a disc, standing above the
    // heading. It says what kind of question this is before a word of it is read.
    let [cx, cy, r] = card.badge;
    let (cx, cy, r) = (cx as f32, cy as f32, r as f32);
    let disc = Rect {
        left: cx - r,
        top: cy - r,
        right: cx + r,
        bottom: cy + r,
    };
    canvas.fill(disc, r, mix(p.danger, p.card, 45), 1.0);
    canvas.fill(disc.inset(hairline), r - hairline, p.danger_bg, 1.0);
    let bar_w = (scale(3, dpi) as f32).max(2.0);
    canvas.fill(
        Rect {
            left: cx - bar_w * 0.5,
            top: cy - r * 0.48,
            right: cx + bar_w * 0.5,
            bottom: cy + r * 0.08,
        },
        bar_w * 0.5,
        p.danger_text,
        1.0,
    );
    let dot_r = (bar_w * 0.7).max(1.5);
    canvas.fill(
        Rect {
            left: cx - dot_r,
            top: cy + r * 0.38 - dot_r,
            right: cx + dot_r,
            bottom: cy + r * 0.38 + dot_r,
        },
        dot_r,
        p.danger_text,
        1.0,
    );

    // And the two buttons. The ring round one is the keyboard's, which is `Cancel` until the user
    // moves it; the fill is the pointer's own state.
    for (rect, decision) in card.buttons {
        let hovered = card.hovered == Some(decision);
        let pressed = card.pressed == Some(decision);
        let (top, bottom, border, ink) = match decision {
            Decision::Affirmative => affirmative_colors(p, hovered, pressed),
            Decision::Negative => negative_colors(p, hovered, pressed),
        };
        let box_ = Rect::around(rect.left, rect.top, rect.right, rect.bottom);
        let button_radius = 7.0 * s;
        canvas.fill(
            box_,
            button_radius,
            if card.focused == decision {
                p.accent
            } else {
                border
            },
            1.0,
        );
        canvas.gradient(
            box_.inset(hairline),
            (button_radius - hairline).max(0.0),
            top,
            bottom,
        );
        canvas.text(
            card.fonts.button,
            rect,
            card.label(decision),
            ink,
            DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
        );
    }
    canvas.text(
        card.fonts.heading,
        card.heading,
        &card.question.title,
        p.label,
        DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
    );
    canvas.text(
        card.fonts.body,
        card.message,
        &wrap_run(
            canvas.dc,
            card.fonts.body,
            &card.question.message,
            card.message.right - card.message.left,
        ),
        p.label,
        DT_LEFT | DT_WORDBREAK | DT_NOPREFIX,
    );

    Some(canvas)
}

/// Hand the composed surface to the window, as the layered window it is.
///
/// The picture is the scrim's own size and is put at the scrim's screen corner: a layered window
/// sized to its own picture is a window whose every pixel is meant, and the card's shadow is the
/// transparency those edges need.
unsafe fn present(card: &Card, pixels: &[u32]) {
    let width = card.width;
    let height = card.height;
    if width <= 0 || height <= 0 || pixels.len() < (width as usize) * (height as usize) {
        return;
    }

    let screen = unsafe { GetDC(HWND::default()) };
    if screen.0.is_null() {
        return;
    }

    let mut bits: *mut c_void = ptr::null_mut();
    let dib = unsafe {
        CreateDIBSection(
            screen,
            &top_down_bmi(width, height),
            DIB_RGB_COLORS,
            &mut bits,
            None,
            0,
        )
    };
    let Ok(dib) = dib else {
        let _ = unsafe { ReleaseDC(HWND::default(), screen) };
        return;
    };

    if bits.is_null() {
        unsafe {
            let _ = DeleteObject(HGDIOBJ(dib.0));
            let _ = ReleaseDC(HWND::default(), screen);
        }
        return;
    }

    unsafe {
        std::ptr::copy_nonoverlapping(
            pixels.as_ptr(),
            bits as *mut u32,
            (width as usize) * (height as usize),
        );
    }

    let mem_dc = unsafe { CreateCompatibleDC(screen) };
    if mem_dc.0.is_null() {
        unsafe {
            let _ = DeleteObject(HGDIOBJ(dib.0));
            let _ = ReleaseDC(HWND::default(), screen);
        }
        return;
    }

    let old = unsafe { SelectObject(mem_dc, HGDIOBJ(dib.0)) };
    let dst = POINT {
        x: card.frame.left,
        y: card.frame.top,
    };
    let size = SIZE {
        cx: width,
        cy: height,
    };
    let src = POINT { x: 0, y: 0 };
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };

    unsafe {
        let _ = UpdateLayeredWindow(
            card.hwnd,
            None,
            Some(&dst),
            Some(&size),
            mem_dc,
            Some(&src),
            COLORREF(0),
            Some(&blend),
            ULW_ALPHA,
        );
        let _ = SelectObject(mem_dc, old);
        let _ = DeleteDC(mem_dc);
        let _ = DeleteObject(HGDIOBJ(dib.0));
        let _ = ReleaseDC(HWND::default(), screen);
    }
}

fn top_down_bmi(width: i32, height: i32) -> BITMAPINFO {
    BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: width,
            biHeight: -height,
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0,
            ..Default::default()
        },
        ..Default::default()
    }
}

/// Answer the card, and take it down: the answer is kept for the call that is waiting on it (see
/// `ANSWER`), and the window is destroyed rather than hidden — a question is one window's life.
///
/// **The first answer is the answer.** Taking the window down is what answers it, and taking a window
/// down deactivates it, so a second `WM_ACTIVATE` arrives after the question is decided; a question
/// answered twice would be a `Delete` reported as a `Cancel`.
fn answer(card: &mut Card, decision: Decision) {
    if card.answered {
        return;
    }
    card.answered = true;
    ANSWER.with(|answer| answer.set(Some(decision == Decision::Affirmative)));
    unsafe {
        let _ = DestroyWindow(card.hwnd);
    }
}

/// The card window's own procedure.
///
/// It is wrapped for the reason the editor's is (`window_proc`): it is `extern "system"`, so a panic
/// unwinding out of it would abort the process the tray lives in rather than take one window down.
unsafe extern "system" fn card_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    std::panic::catch_unwind(|| unsafe { card_proc_inner(hwnd, msg, wparam, lparam) })
        .unwrap_or(LRESULT(0))
}

unsafe fn card_proc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_ERASEBKGND => LRESULT(1),
        WM_PAINT => {
            let mut paint = PAINTSTRUCT::default();
            let _ = unsafe { BeginPaint(hwnd, &mut paint) };
            if let Some(card) = card_of(hwnd) {
                unsafe { paint_card(card) };
            }
            let _ = unsafe { EndPaint(hwnd, &paint) };
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            if let Some(card) = card_of(hwnd) {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
                let next = card.button_at(x, y);
                if next != card.hovered {
                    card.hovered = next;
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
                if !card.tracking {
                    card.tracking = true;
                    let mut track = TRACKMOUSEEVENT {
                        cbSize: std::mem::size_of::<TRACKMOUSEEVENT>() as u32,
                        dwFlags: TME_LEAVE,
                        hwndTrack: hwnd,
                        dwHoverTime: 0,
                    };
                    let _ = unsafe { TrackMouseEvent(&mut track) };
                }
            }
            LRESULT(0)
        }
        WM_MOUSELEAVE => {
            if let Some(card) = card_of(hwnd) {
                card.tracking = false;
                if card.hovered.take().is_some() {
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
            }
            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            if let Some(card) = card_of(hwnd) {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
                if let Some(decision) = card.button_at(x, y) {
                    card.pressed = Some(decision);
                    card.focused = decision;
                    unsafe {
                        SetCapture(hwnd);
                        let _ = InvalidateRect(hwnd, None, false);
                    }
                }
            }
            LRESULT(0)
        }
        WM_LBUTTONUP => {
            if let Some(card) = card_of(hwnd) {
                let pressed = card.pressed.take();
                let _ = unsafe { ReleaseCapture() };
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
                let was_pressed = pressed.is_some();
                if let (Some(pressed), Some(released)) = (pressed, card.button_at(x, y)) {
                    if pressed == released {
                        answer(card, released);
                        return LRESULT(0);
                    }
                }
                if was_pressed {
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
            }
            LRESULT(0)
        }
        WM_CAPTURECHANGED => {
            if let Some(card) = card_of(hwnd) {
                if card.pressed.take().is_some() {
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
            }
            LRESULT(0)
        }
        WM_SETCURSOR => {
            let over_a_button = card_of(hwnd).is_some_and(|card| card.hovered.is_some());
            if over_a_button {
                let _ = unsafe { SetCursor(LoadCursorW(None, IDC_HAND).unwrap_or_default()) };
                return LRESULT(1);
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_KEYDOWN => {
            let Some(card) = card_of(hwnd) else {
                return LRESULT(0);
            };
            let key = wparam.0 as u16;
            match key {
                // `Enter` and `Space` answer what the keyboard is on; the two keys that leave a
                // question answer with the button that changes nothing.
                k if k == VK_RETURN.0 || k == VK_SPACE.0 => answer(card, card.focused),
                k if k == VK_ESCAPE.0 => answer(card, Decision::Negative),
                k if k == VK_TAB.0 => {
                    // `Tab` and `Shift+Tab` are one move in opposite directions, and the card holds
                    // two answers: either way it is the other one.
                    let _shifted = unsafe { GetKeyState(VK_SHIFT.0 as i32) } < 0;
                    card.focused = match card.focused {
                        Decision::Negative => Decision::Affirmative,
                        Decision::Affirmative => Decision::Negative,
                    };
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
                k if k == VK_LEFT.0 => {
                    card.focused = Decision::Negative;
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
                k if k == VK_RIGHT.0 => {
                    card.focused = Decision::Affirmative;
                    let _ = unsafe { InvalidateRect(hwnd, None, false) };
                }
                _ => return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
            }
            LRESULT(0)
        }
        WM_ACTIVATE => {
            if let Some(card) = card_of(hwnd) {
                let how = (wparam.0 & 0xFFFF) as u32;
                if how == WA_ACTIVE || how == WA_CLICKACTIVE {
                    card.activated = true;
                } else if how == WA_INACTIVE && card.activated {
                    // Losing the foreground is the user going somewhere else — which is leaving the
                    // question, and a question left unanswered is the one answered with what changes
                    // nothing. Two foregoings are the question still being asked: the card's own
                    // `UpdateLayeredWindow` raise, and the editor under it taking the front back from a
                    // click that landed on it (it is disabled, so it cannot be typed into, but it can be
                    // activated).
                    let foreground = unsafe { GetForegroundWindow() };
                    let is_the_card = foreground.0.is_null() || foreground == card.hwnd;
                    if !is_the_card && foreground == card.owner {
                        let _ = unsafe { SetForegroundWindow(card.hwnd) };
                    } else if !is_the_card {
                        answer(card, Decision::Negative);
                        return LRESULT(0);
                    }
                }
            }
            LRESULT(0)
        }
        WM_SETFOCUS | WM_KILLFOCUS => {
            let _ = unsafe { InvalidateRect(hwnd, None, false) };
            LRESULT(0)
        }
        // **The beep the platform would otherwise make.** A box beeps for a key it has no answer to,
        // and a card that beeped at a held letter, at a dead key or at nothing at all is the message box
        // it replaced. What it does answer is answered above; every other key is silently nothing.
        WM_CHAR | WM_SYSCHAR => LRESULT(0),
        WM_SYSKEYDOWN => {
            let Some(card) = card_of(hwnd) else {
                return LRESULT(0);
            };
            // `Alt+F4` is the one system key a framed window would have answered for it, and a question
            // is not dismissed by a window command it has no frame for.
            let alt = unsafe { GetKeyState(VK_MENU.0 as i32) } < 0;
            if alt && wparam.0 as u16 == VK_F4.0 {
                answer(card, Decision::Negative);
                return LRESULT(0);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            if let Some(card) = card_of(hwnd) {
                answer(card, Decision::Negative);
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            unsafe { PostQuitMessage(0) };
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut Card;
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

/// Signed distance to a rounded box of half-extents `hw`/`hh` and corner radius `r`.
fn sd_rounded_box(px: f32, py: f32, cx: f32, cy: f32, hw: f32, hh: f32, r: f32) -> f32 {
    let r = r.min(hw).min(hh).max(0.0);
    let qx = (px - cx).abs() - (hw - r);
    let qy = (py - cy).abs() - (hh - r);
    qx.max(0.0).hypot(qy.max(0.0)) + qx.max(qy).min(0.0) - r
}

/// The share of a pixel a shape covers, from its signed distance to the pixel's centre.
fn aa_coverage(d: f32) -> f32 {
    (0.5 - d).clamp(0.0, 1.0)
}

/// The share of a pixel a shape covers, averaged over the four quarters the pixel is cut into.
///
/// A shape whose outline turns inside one pixel — an arc a few pixels across, which is every corner of
/// these buttons — is the one a single sample gets wrong in whole pixels rather than in fractions: the
/// outline crosses a pixel's diagonal, and its centre then falls on one side of it for the whole of that
/// pixel, which is the stair the eye reads as aliasing. Four samples across the diagonal resolve it.
fn covered(px: f32, py: f32, coverage: impl Fn(f32, f32) -> f32) -> f32 {
    const QUARTERS: [(f32, f32); 4] = [(0.75, 0.75), (0.25, 0.75), (0.75, 0.25), (0.25, 0.25)];
    let mut total = 0.0;
    for (dx, dy) in QUARTERS {
        total += coverage(px + dx - 0.5, py + dy - 0.5);
    }
    total / QUARTERS.len() as f32
}

fn channels(colour: COLORREF) -> (u8, u8, u8) {
    (
        (colour.0 & 0xFF) as u8,
        ((colour.0 >> 8) & 0xFF) as u8,
        ((colour.0 >> 16) & 0xFF) as u8,
    )
}

/// A premultiplied-BGRA canvas for the card, in this app's memory rather than in a DIB.
///
/// The shadow and the card's rounded corners live in the *alpha* byte, which GDI does not write and
/// a layered window reads, so the shapes are composed here arithmetically and the whole surface is
/// handed to `UpdateLayeredWindow` in one go. The one thing GDI draws is text, into a scratch DIB of
/// its own and read back as coverage (`text`), because glyphs are the one shape here that has to be
/// hinted rather than outlined. Every other shape is a fill — a border is a fill inside a fill — and the
/// outlines are sampled four times across the pixel, so that an arc turning inside one pixel is a curve
/// rather than a stair.
struct Canvas {
    w: i32,
    h: i32,
    /// Premultiplied BGRA: what is stored is the colour *times* the coverage, which is what a
    /// layered window composites over the desktop.
    pixels: Vec<u32>,
    /// The screen device context the text masks are made from.
    dc: HDC,
}

impl Canvas {
    fn new(w: i32, h: i32) -> Option<Self> {
        let w = w.max(1);
        let h = h.max(1);
        let dc = unsafe { GetDC(HWND::default()) };
        if dc.0.is_null() {
            return None;
        }
        Some(Self {
            w,
            h,
            pixels: vec![0; (w as usize) * (h as usize)],
            dc,
        })
    }

    fn pixels(&self) -> &[u32] {
        &self.pixels
    }

    /// One pixel of `colour` at `alpha`, over what is already there: premultiplied over
    /// premultiplied, which is one multiply-add per channel and no division of a colour by an alpha.
    fn blend(&mut self, x: i32, y: i32, colour: COLORREF, alpha: f32) {
        if alpha <= 0.0 || x < 0 || y < 0 || x >= self.w || y >= self.h {
            return;
        }
        let a = (alpha.min(1.0) * 255.0 + 0.5) as u32;
        if a == 0 {
            return;
        }
        let inv = 255 - a;
        let (sr, sg, sb) = channels(colour);
        let index = (y as usize) * (self.w as usize) + x as usize;
        let under = self.pixels[index];
        let ur = (under >> 16) & 0xFF;
        let ug = (under >> 8) & 0xFF;
        let ub = under & 0xFF;
        let ua = (under >> 24) & 0xFF;
        let r = (u32::from(sr) * a + ur * inv + 127) / 255;
        let g = (u32::from(sg) * a + ug * inv + 127) / 255;
        let b = (u32::from(sb) * a + ub * inv + 127) / 255;
        let out_a = (a + (ua * inv + 127) / 255).min(255);
        self.pixels[index] = (out_a << 24) | (r << 16) | (g << 8) | b;
    }

    /// The pixels a shape's box reaches, clamped to the canvas and grown by `pad` so that the
    /// anti-aliased edge of the shape is walked as well as its inside.
    fn area(&self, bounds: Rect, pad: f32) -> (i32, i32, i32, i32) {
        (
            ((bounds.left - pad).floor() as i32).max(0),
            ((bounds.top - pad).floor() as i32).max(0),
            ((bounds.right + pad).ceil() as i32).min(self.w),
            ((bounds.bottom + pad).ceil() as i32).min(self.h),
        )
    }

    /// Half the width of the row at `py` that a rounded box covers *wholly*: the rows inside the box's
    /// own core are filled rather than walked, which is what keeps a card's paint a few hundred thousand
    /// stores rather than a few hundred thousand distance fields.
    ///
    /// Two things make it exact rather than approximate, and both are the difference between a rounded
    /// corner and a stair:
    ///
    /// * a row the box does not reach is **not** covered — the width is nothing above the top edge and
    ///   below the bottom one, where a width of `hw - r` would paint a band outside the shape;
    /// * the core keeps a **whole pixel** inside the outline, because a pixel the outline cuts is one
    ///   whose coverage is a fraction: a core that reached within half a pixel of it would fill those in
    ///   solid, which is the stepped corner this card had before it was sampled.
    fn covered_half(bounds: Rect, radius: f32, py: f32) -> f32 {
        let hw = bounds.width() * 0.5;
        let hh = bounds.height() * 0.5;
        let r = radius.min(hw).min(hh).max(0.0);

        // The row, measured from the straight part of the shape's side: within `r` of it is inside the
        // box at all, and the rest is the corner.
        let past_the_side = (py - (bounds.top + bounds.bottom) * 0.5).abs() - (hh - r);
        let inside = r - 1.0;
        if past_the_side >= inside {
            return 0.0;
        }
        let allowance = if past_the_side <= 0.0 {
            inside
        } else {
            (inside * inside - past_the_side * past_the_side)
                .max(0.0)
                .sqrt()
        };
        (hw - r + allowance).max(0.0)
    }

    /// A rounded box of one colour: a disc where the radius is half its shorter side.
    fn fill(&mut self, bounds: Rect, radius: f32, colour: COLORREF, alpha: f32) {
        let cx = (bounds.left + bounds.right) * 0.5;
        let cy = (bounds.top + bounds.bottom) * 0.5;
        let hw = bounds.width() * 0.5;
        let hh = bounds.height() * 0.5;
        let (ix0, iy0, ix1, iy1) = self.area(bounds, 1.0);
        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            let core = Self::covered_half(bounds, radius, py);
            let core_x0 = ((cx - core).round() as i32).clamp(ix0, ix1);
            let core_x1 = ((cx + core).round() as i32).clamp(core_x0, ix1);
            for x in core_x0..core_x1 {
                self.blend(x, y, colour, alpha);
            }
            for x in ix0..ix1 {
                if (core_x0..core_x1).contains(&x) {
                    continue;
                }
                let px = x as f32 + 0.5;
                let coverage = covered(px, py, |at_x, at_y| {
                    aa_coverage(sd_rounded_box(at_x, at_y, cx, cy, hw, hh, radius))
                });
                if coverage > 0.0 {
                    self.blend(x, y, colour, coverage * alpha);
                }
            }
        }
    }

    /// The same box shaded from its top edge to its bottom one: the card's sheen and the buttons'
    /// gradient are one call with two colours.
    fn gradient(&mut self, bounds: Rect, radius: f32, top: COLORREF, bottom: COLORREF) {
        let cx = (bounds.left + bounds.right) * 0.5;
        let cy = (bounds.top + bounds.bottom) * 0.5;
        let hw = bounds.width() * 0.5;
        let hh = bounds.height() * 0.5;
        let span = bounds.height().max(1.0);
        let (ix0, iy0, ix1, iy1) = self.area(bounds, 1.0);
        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            let t = ((py - bounds.top) / span).clamp(0.0, 1.0);
            let colour = mix(top, bottom, (t * 100.0 + 0.5) as u32);
            let core = Self::covered_half(bounds, radius, py);
            let core_x0 = ((cx - core).round() as i32).clamp(ix0, ix1);
            let core_x1 = ((cx + core).round() as i32).clamp(core_x0, ix1);
            for x in core_x0..core_x1 {
                self.blend(x, y, colour, 1.0);
            }
            for x in ix0..ix1 {
                if (core_x0..core_x1).contains(&x) {
                    continue;
                }
                let px = x as f32 + 0.5;
                let coverage = covered(px, py, |at_x, at_y| {
                    aa_coverage(sd_rounded_box(at_x, at_y, cx, cy, hw, hh, radius))
                });
                if coverage > 0.0 {
                    self.blend(x, y, colour, coverage);
                }
            }
        }
    }

    /// A soft drop shadow under a rounded box: the box's own distance field flattened over `spread`,
    /// squared so that it falls off the way a shadow does rather than the way a ramp does.
    fn shadow(&mut self, bounds: Rect, radius: f32, spread: f32, alpha: f32) {
        let cx = (bounds.left + bounds.right) * 0.5;
        let cy = (bounds.top + bounds.bottom) * 0.5;
        let hw = bounds.width() * 0.5;
        let hh = bounds.height() * 0.5;
        let drop = 2.0;
        let grown = Rect {
            left: bounds.left - spread,
            top: bounds.top - spread,
            right: bounds.right + spread,
            bottom: bounds.bottom + spread + drop,
        };
        let (ix0, iy0, ix1, iy1) = self.area(grown, 1.0);
        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            for x in ix0..ix1 {
                let px = x as f32 + 0.5;
                let distance = sd_rounded_box(px, py, cx, cy + drop, hw, hh, radius);
                if distance < spread {
                    let t = (1.0 - (distance.max(0.0) / spread)).clamp(0.0, 1.0);
                    self.blend(x, y, rgb(0, 0, 0), t * t * alpha);
                }
            }
        }
    }

    /// A run of text, in the palette's own ink.
    ///
    /// The run is drawn by GDI as white on black in a scratch DIB and read back as coverage, and only
    /// then blended here in `colour`: GDI leaves the alpha byte of what it draws at zero, and a glyph is
    /// a shape this canvas has no way to outline. What is read back is the largest of the three channels,
    /// which — with the grayscale anti-aliasing `text_font` asks for — is the glyph's coverage itself.
    /// ClearType would put three different coverages in one pixel and the mask would be the strongest of
    /// them, which is a run drawn heavier than it was.
    fn text(
        &mut self,
        font: HFONT,
        rect: RECT,
        run: &str,
        colour: COLORREF,
        flags: DRAW_TEXT_FORMAT,
    ) {
        let w = rect.right - rect.left;
        let h = rect.bottom - rect.top;
        if run.is_empty() || w <= 0 || h <= 0 || font.0.is_null() {
            return;
        }

        let mut bits: *mut c_void = ptr::null_mut();
        let dib = unsafe {
            CreateDIBSection(
                self.dc,
                &top_down_bmi(w, h),
                DIB_RGB_COLORS,
                &mut bits,
                None,
                0,
            )
        };
        let Ok(dib) = dib else {
            return;
        };
        if bits.is_null() {
            unsafe {
                let _ = DeleteObject(HGDIOBJ(dib.0));
            }
            return;
        }

        let mem_dc = unsafe { CreateCompatibleDC(self.dc) };
        if mem_dc.0.is_null() {
            unsafe {
                let _ = DeleteObject(HGDIOBJ(dib.0));
            }
            return;
        }

        unsafe {
            std::ptr::write_bytes(bits, 0, (w as usize) * (h as usize) * 4);
        }

        let old_bitmap = unsafe { SelectObject(mem_dc, HGDIOBJ(dib.0)) };
        let old_font = unsafe { SelectObject(mem_dc, HGDIOBJ(font.0)) };
        unsafe {
            SetBkMode(mem_dc, TRANSPARENT);
            SetTextColor(mem_dc, rgb(255, 255, 255));
        }
        let mut text: Vec<u16> = run.encode_utf16().collect();
        let mut local = RECT {
            left: 0,
            top: 0,
            right: w,
            bottom: h,
        };
        unsafe { DrawTextW(mem_dc, &mut text, &mut local, flags) };

        {
            let mask = unsafe {
                std::slice::from_raw_parts(bits as *const u32, (w as usize) * (h as usize))
            };
            for (index, pixel) in mask.iter().copied().enumerate() {
                let coverage = (pixel & 0xFF)
                    .max((pixel >> 8) & 0xFF)
                    .max((pixel >> 16) & 0xFF);
                if coverage == 0 {
                    continue;
                }
                let x = rect.left + (index as i32 % w);
                let y = rect.top + (index as i32 / w);
                self.blend(x, y, colour, coverage as f32 / 255.0);
            }
        }

        unsafe {
            let _ = SelectObject(mem_dc, old_font);
            let _ = SelectObject(mem_dc, old_bitmap);
            let _ = DeleteDC(mem_dc);
            let _ = DeleteObject(HGDIOBJ(dib.0));
        }
    }
}

impl Drop for Canvas {
    fn drop(&mut self) {
        unsafe {
            let _ = ReleaseDC(HWND::default(), self.dc);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::config::TextTheme;

    fn question() -> Question {
        Question::new(
            "Delete this preset?",
            "The preset \"PDF docs\" would be removed.\n\nExtensions: .pdf\n\nProgram: \
             C:\\Program Files\\SumatraPDF\\Another Folder\\SumatraPDF-3.5.2-64.exe\n\nIts \
             section leaves config.ini with the save that follows, and nothing here puts it back. \
             Your other presets and settings are not touched.",
            "Delete",
        )
    }

    fn laid_out(dpi: i32, w: i32, h: i32) -> Layout {
        let fonts = Fonts::make(dpi);
        layout(&question(), &fonts, dpi, w, h)
    }

    /// The card the way `create` builds one, without a window: the layout, the palette of a theme,
    /// and the fonts — everything `compose` draws from.
    fn render_the_card(theme: TextTheme, dpi: i32, w: i32, h: i32) -> Card {
        let fonts = Fonts::make(dpi);
        let layout = layout(&question(), &fonts, dpi, w, h);
        Card {
            hwnd: HWND::default(),
            owner: HWND::default(),
            question: question(),
            palette: Palette::of(theme),
            dpi,
            frame: RECT::default(),
            width: w,
            height: h,
            body: layout.body,
            heading: layout.heading,
            message: layout.message,
            badge: layout.badge,
            buttons: layout.buttons,
            hovered: None,
            pressed: None,
            focused: Decision::Negative,
            tracking: false,
            activated: false,
            answered: false,
            fonts,
        }
    }

    /// Writes the card of both themes to `.tmp/` for the eye, which is the one check a test cannot
    /// make: whether it is *beautiful*. Run it by name (`--ignored the_picture...`).
    #[test]
    #[ignore]
    fn the_picture_written_out_for_the_eye() {
        for (theme, name) in [(TextTheme::Dark, "dark"), (TextTheme::Light, "light")] {
            let card = render_the_card(theme, 96, 920, 660);
            let canvas = compose(&card).expect("a screen device context");
            let mut rgba = Vec::with_capacity(canvas.pixels.len() * 4);
            for pixel in canvas.pixels() {
                let a = (pixel >> 24) & 0xFF;
                let un = |channel: u32| -> u8 {
                    channel
                        .checked_mul(255)
                        .and_then(|scaled| scaled.checked_div(a))
                        .map_or(0, |shrunk| shrunk.min(255) as u8)
                };
                rgba.push(un((pixel >> 16) & 0xFF));
                rgba.push(un((pixel >> 8) & 0xFF));
                rgba.push(un(pixel & 0xFF));
                rgba.push(a as u8);
            }
            let image = image::RgbaImage::from_raw(canvas.w as u32, canvas.h as u32, rgba)
                .expect("a picture of the card");
            std::fs::create_dir_all(".tmp").ok();
            image
                .save(format!(".tmp/ask-card-{name}.png"))
                .expect("a file to look at");
        }
    }

    /// **The one test that puts the card on the desktop**, and therefore the one that is `#[ignore]`d:
    /// it is the loop, the keyboard and the answer as they are in a run of the app rather than as a
    /// composition of pixels. Run it by hand (`cargo test -- --ignored the_card_is_answered`), and
    /// what it proves is that the card takes the keyboard, that `Enter` answers what it is on, that
    /// `Escape` answers what changes nothing, and that both take the card down far enough for the
    /// call that asked to come back.
    #[test]
    #[ignore = "a card of this app's own, on the desktop: run it by hand"]
    fn the_card_is_answered_by_a_key_or_by_a_click() {
        use windows::Win32::UI::WindowsAndMessaging::{
            CreateWindowExW, DefWindowProcW, DestroyWindow, PostMessageW, RegisterClassExW,
            HWND_MESSAGE, WINDOW_EX_STYLE, WNDCLASSEXW, WNDCLASS_STYLES,
        };
        // `DefWindowProcW` is the bindings' own `unsafe` fn, which a `WNDPROC` cannot be made of
        // directly: the owner of this test answers everything with it.
        unsafe extern "system" fn owner_proc(
            hwnd: HWND,
            msg: u32,
            wparam: WPARAM,
            lparam: LPARAM,
        ) -> LRESULT {
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }

        // An owner of the test's own, made the way the suite makes every window: created, and never
        // shown — a message-only window has no desktop of its own to cover.
        let owner = unsafe {
            let class = WNDCLASSEXW {
                cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
                style: WNDCLASS_STYLES(0),
                lpfnWndProc: Some(owner_proc),
                lpszClassName: w!("RustHoverPreviewAskCardTestOwner"),
                ..Default::default()
            };
            RegisterClassExW(&class);
            CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                w!("RustHoverPreviewAskCardTestOwner"),
                w!(""),
                WS_POPUP,
                0,
                0,
                920,
                660,
                HWND_MESSAGE,
                None,
                None,
                None,
            )
            .expect("an owner of the test's own")
        };

        // `Escape` first: the answer that changes nothing, and the one a stray key has to give.
        let escape = std::thread::spawn(|| {
            let card = wait_for_the_card();
            unsafe {
                PostMessageW(card, WM_KEYDOWN, WPARAM(VK_ESCAPE.0 as usize), LPARAM(0))
                    .expect("a key to the card");
            }
        });
        let answered = ask(owner, &question());
        escape.join().expect("the thread that answered it");
        assert_eq!(
            answered,
            Some(false),
            "`Escape` did not answer with what changes nothing"
        );

        // Then the keyboard on the other button, which `Enter` answers: `Right` moves it, and the
        // answer is the one the destructive button carries.
        let enter = std::thread::spawn(|| {
            let card = wait_for_the_card();
            unsafe {
                PostMessageW(card, WM_KEYDOWN, WPARAM(VK_RIGHT.0 as usize), LPARAM(0))
                    .expect("a key to the card");
                PostMessageW(card, WM_KEYDOWN, WPARAM(VK_RETURN.0 as usize), LPARAM(0))
                    .expect("a key to the card");
            }
        });
        let answered = ask(owner, &question());
        enter.join().expect("the thread that answered it");
        assert_eq!(
            answered,
            Some(true),
            "`Enter` did not answer what the keyboard was on"
        );

        // And a click on the card's own buttons: the affirmative one, at the middle of its box.
        let click = std::thread::spawn(|| {
            let card = wait_for_the_card();
            let fonts = Fonts::make(96);
            let layout = layout(&question(), &fonts, 96, 920, 660);
            let (delete, _) = layout.buttons[1];
            let (x, y) = (
                (delete.left + delete.right) / 2,
                (delete.top + delete.bottom) / 2,
            );
            let point = LPARAM(((y << 16) | (x & 0xFFFF)) as isize);
            unsafe {
                PostMessageW(card, WM_LBUTTONDOWN, WPARAM(0), point).expect("a press on the card");
                PostMessageW(card, WM_LBUTTONUP, WPARAM(0), point).expect("a release on the card");
            }
        });
        let answered = ask(owner, &question());
        click.join().expect("the thread that answered it");
        assert_eq!(
            answered,
            Some(true),
            "a press on Delete did not answer the card"
        );

        unsafe {
            let _ = DestroyWindow(owner);
        }
    }

    /// The card the asking thread has just put up, waited for: the call that asks does not come back
    /// until the card is gone, so the thread that answers it has nothing to be handed.
    fn wait_for_the_card() -> HWND {
        for _ in 0..400 {
            let card = LAST_CARD.load(std::sync::atomic::Ordering::SeqCst);
            if card != 0 {
                let hwnd = HWND(card as *mut c_void);
                if unsafe { windows::Win32::UI::WindowsAndMessaging::IsWindow(hwnd) }.as_bool() {
                    return hwnd;
                }
            }
            std::thread::sleep(std::time::Duration::from_millis(10));
        }
        panic!("no card was put up to answer");
    }

    /// The pixel at a point of the composed picture.
    fn pixel(canvas: &Canvas, x: i32, y: i32) -> u32 {
        canvas.pixels[(y as usize) * (canvas.w as usize) + x as usize]
    }

    #[test]
    fn the_picture_is_a_card_on_a_scrim_rather_than_a_box_of_its_own_colour() {
        // What the card is for: the window it belongs to stays visible behind it, dimmed, and the card
        // stands on that — an opaque card on an opaque scrim would be the platform's box again.
        let card = render_the_card(TextTheme::Dark, 96, 920, 660);
        let canvas = compose(&card).expect("a screen device context");
        let body = card.body;

        // The scrim: translucent, and the editor's page rather than black or white.
        let scrim = pixel(&canvas, 40, canvas.h / 2);
        let alpha = scrim >> 24;
        assert!(
            (100..160).contains(&alpha),
            "the scrim is not a dim over the window: alpha {alpha}"
        );
        // Premultiplied, which is what a layered window reads: the colour stored is the page times the
        // scrim's own alpha, so what the desktop sees over the window behind it is the page dimmed.
        let (page_r, page_g, page_b) = channels(card.palette.page);
        let scaled = |channel: u8| u32::from(channel) * alpha / 255;
        assert!(
            ((scrim >> 16) & 0xFF).abs_diff(scaled(page_r)) <= 2
                && ((scrim >> 8) & 0xFF).abs_diff(scaled(page_g)) <= 2
                && (scrim & 0xFF).abs_diff(scaled(page_b)) <= 2,
            "the scrim is not the editor's page, dimmed: {scrim:#x} against {page_r}, {page_g}, {page_b}"
        );

        // The card: opaque, and its own colour rather than the scrim's — between the sheen at its top
        // edge and its own colour at the bottom, which is the gradient it is filled with.
        let middle = pixel(
            &canvas,
            (body.left + body.right) as i32 / 2,
            (body.top + body.bottom) as i32 / 2,
        );
        assert_eq!(middle >> 24, 255, "the card is not opaque");
        let (card_r, card_g, card_b) = channels(card.palette.card);
        let (sheen_r, sheen_g, sheen_b) = channels(mix(card.palette.card, rgb(255, 255, 255), 6));
        let band = |got: u32, one: u8, two: u8| {
            let (lo, hi) = (one.min(two), one.max(two));
            (u32::from(lo)..=u32::from(hi)).contains(&got)
        };
        assert!(
            band((middle >> 16) & 0xFF, card_r, sheen_r)
                && band((middle >> 8) & 0xFF, card_g, sheen_g)
                && band(middle & 0xFF, card_b, sheen_b),
            "the card is not its own colour: {middle:#x} against {}, {}, {}",
            card_r,
            card_g,
            card_b
        );

        // A corner of the card is a rounded one: the pixel at the box's corner is the scrim showing
        // through, and the pixel a radius in is the card.
        let corner_alpha = pixel(&canvas, body.left as i32 + 1, body.top as i32 + 1) >> 24;
        let inside_alpha = pixel(
            &canvas,
            body.left as i32 + CARD_RADIUS_LOGICAL * 2,
            body.top as i32 + CARD_RADIUS_LOGICAL * 2,
        ) >> 24;
        assert!(
            corner_alpha < 240,
            "the card's corner is not rounded: alpha {corner_alpha}"
        );
        assert_eq!(
            inside_alpha, 255,
            "the card is not opaque inside its corner"
        );

        // And the two buttons stand on it, in their own colours rather than in the card's.
        for (rect, decision) in card.buttons {
            let (top, _, _, _) = match decision {
                Decision::Affirmative => affirmative_colors(card.palette, false, false),
                Decision::Negative => negative_colors(card.palette, false, false),
            };
            let (_, _, want) = channels(top);
            let got = pixel(&canvas, (rect.left + rect.right) / 2, rect.top + 2);
            assert_eq!(got >> 24, 255, "a button is not opaque: {decision:?}");
            assert!(
                (got & 0xFF).abs_diff(u32::from(want)) <= 6,
                "a button is not its own colour: {decision:?} is {got:#x}, wanted {want}"
            );
        }
    }

    #[test]
    fn the_card_stands_in_the_middle_of_the_scrim_at_every_scale() {
        for dpi in [96, 120, 144, 192] {
            let frame_w = scale(920, dpi);
            let frame_h = scale(660, dpi);
            let layout = laid_out(dpi, frame_w, frame_h);
            let card_w = layout.body.width();
            let card_h = layout.body.height();

            // Centred on the scrim, to within the pixel a division by two gives away.
            assert!(
                (layout.body.left - (frame_w as f32 - card_w) * 0.5).abs() <= 1.0,
                "the card is not centred at {dpi} dpi: {:?}",
                layout.body
            );
            assert!((layout.body.top - (frame_h as f32 - card_h) * 0.5).abs() <= 1.0);
            assert!(layout.body.left >= 0.0 && layout.body.top >= 0.0);

            // Every part of it inside it, and the two buttons inside the room it leaves.
            for part in [layout.heading, layout.message] {
                assert!(
                    part.left >= layout.body.left as i32,
                    "{part:?} is off the card"
                );
                assert!(part.right <= layout.body.right as i32);
                assert!(part.bottom <= layout.body.bottom as i32);
            }
            let [(cancel, _), (delete, _)] = layout.buttons;
            assert!(pt_in_rect(layout.body.pixels(), cancel.left, cancel.top));
            assert!(pt_in_rect(
                layout.body.pixels(),
                delete.right - 1,
                delete.bottom - 1
            ));
            assert!(
                cancel.right <= delete.left,
                "the buttons overlap: {cancel:?} and {delete:?}"
            );

            // The badge is the card's first row and the heading is written under it, both inside the
            // card's pad — which is what keeps the mark on the card rather than over its edge.
            let [bx, by, r] = layout.badge;
            assert!(bx - r >= layout.body.left as i32);
            assert!(by - r >= layout.body.top as i32);
            assert!(
                layout.heading.top >= by + r,
                "the badge and the heading share a row: {:?} against {by}",
                layout.heading
            );
            assert!(layout.heading.top >= layout.body.top as i32);
        }
    }

    #[test]
    fn a_window_too_narrow_for_the_card_hugs_both_its_edges() {
        let layout = laid_out(96, 240, 660);
        assert!(layout.body.left >= 0.0);
        assert!(
            layout.body.right <= 240.0,
            "the card runs past the scrim: {:?}",
            layout.body
        );
        assert!(layout.body.width() >= 200.0);
    }

    #[test]
    fn the_answer_that_changes_nothing_is_the_one_the_keyboard_starts_on() {
        // What `create` puts the card up with, and what `Enter` answers: this is the whole of what
        // makes a stray `Enter` safe.
        let layout = laid_out(96, 920, 660);
        assert_eq!(layout.buttons[0].1, Decision::Negative);
        assert_eq!(layout.buttons[1].1, Decision::Affirmative);
    }

    #[test]
    fn a_program_too_long_for_the_card_wraps_rather_than_running_off_it() {
        // The card's own width for the body, and a path no one would call short. It is a raw string
        // because a Windows path is backslashes, and each of them is a place the break may take.
        let dpi = 96;
        let fonts = Fonts::make(dpi);
        let width = scale(CARD_WIDTH_LOGICAL, dpi) - scale(44, dpi);
        let line = format!(
            "Program: {}",
            r"C:\Program Files\Some Vendor With A Long Name\Another Folder\And Yet More Nesting\the-program-itself-with-a-very-long-name.exe"
        );
        let screen = unsafe { GetDC(HWND::default()) };
        let wrapped = wrap_run(screen, fonts.body, &line, width);
        let (widest, on_one_line) = {
            let mut widest = 0;
            for drawn in wrapped.split('\n') {
                widest = widest.max(measure(screen, fonts.body, drawn, width, DT_SINGLELINE).0);
            }
            let on_one_line = measure(screen, fonts.body, &line, scale(4000, dpi), DT_SINGLELINE).0;
            (widest, on_one_line)
        };
        let (_, wrapped_h) = measure(screen, fonts.body, &wrapped, width, DT_WORDBREAK);

        assert!(
            on_one_line > width,
            "the line was not too wide to begin with: {on_one_line} of {width}"
        );
        assert!(
            widest <= width,
            "a wrapped line is wider than the card: {widest} > {width}"
        );
        assert!(
            wrapped_h > scale(18, dpi) * 2,
            "the path was not broken across lines: {wrapped_h}px"
        );

        // Nothing of the run is lost or invented: the breaks are where the spaces already were, and
        // no letter of the path is cut, added or replaced.
        let letters = |run: &str| -> String {
            run.chars()
                .filter(|unit| !matches!(unit, '\n' | ' ' | '\r'))
                .collect()
        };
        assert_eq!(
            letters(&wrapped),
            letters(&line),
            "the wrap changed the words it was given"
        );
        // And a blank line in the question is a blank line of the card.
        let spaced = wrap_run(screen, fonts.body, "One\n\nTwo", width);
        let _ = unsafe { ReleaseDC(HWND::default(), screen) };
        assert!(spaced.contains("One\n\nTwo"), "{spaced:?}");
    }

    #[test]
    fn a_corner_of_the_card_is_carried_in_the_alpha_byte() {
        // The reason the card is composed here rather than by GDI: a rounded corner and a shadow are
        // nothing but alpha, and what a layered window composites is the alpha it is given.
        let mut canvas = Canvas::new(60, 60).expect("a screen device context");
        let body = Rect::around(10, 10, 50, 50);
        canvas.fill(body, 12.0, rgb(30, 34, 42), 1.0);

        let corner = canvas.pixels[10 * 60 + 10];
        assert_eq!(corner >> 24, 0, "the corner is not transparent");
        let middle = canvas.pixels[30 * 60 + 30];
        assert_eq!(middle >> 24, 255, "the card is not opaque");
        assert_eq!(middle & 0xFF, 42, "the card is not its own colour");
        assert!(
            canvas
                .pixels
                .iter()
                .any(|pixel| (1..255).contains(&(pixel >> 24))),
            "no edge of the card is anti-aliased"
        );
        // Premultiplied: no channel of a pixel is above its own alpha.
        for pixel in canvas.pixels() {
            let alpha = (pixel >> 24) & 0xFF;
            assert!(((pixel >> 16) & 0xFF) <= alpha);
            assert!(((pixel >> 8) & 0xFF) <= alpha);
            assert!((pixel & 0xFF) <= alpha);
        }
    }

    /// A button's word is a normal-weight run, and the mask it is read back from is one coverage per
    /// pixel: the two reasons a label here used to come out darker and blobbier than the same word in
    /// the sentence above it.
    #[test]
    fn a_buttons_word_is_an_unbolded_run_read_back_as_one_coverage() {
        use windows::Win32::Graphics::Gdi::{GetObjectW, LOGFONTW};

        let fonts = Fonts::make(96);
        for font in [fonts.button, fonts.body] {
            let mut logfont = LOGFONTW::default();
            let read = unsafe {
                GetObjectW(
                    HGDIOBJ(font.0),
                    std::mem::size_of::<LOGFONTW>() as i32,
                    Some(&mut logfont as *mut _ as *mut c_void),
                )
            };
            assert_eq!(
                read as usize,
                std::mem::size_of::<LOGFONTW>(),
                "a font of the card's could not be read back"
            );
            assert_eq!(
                logfont.lfWeight, FW_NORMAL.0 as i32,
                "a run of the card's is not normal weight"
            );
            assert_eq!(
                logfont.lfQuality, ANTIALIASED_QUALITY,
                "a run of the card's is not grayscale anti-aliased"
            );
        }

        // And what that font draws is a mask: every pixel of the run is grey, so the largest of its
        // three channels *is* its coverage. ClearType — the editor's own quality — writes three
        // different ones, and `Canvas::text` would read the strongest of them as the whole of it.
        let dc = unsafe { GetDC(HWND::default()) };
        let mut bits: *mut c_void = ptr::null_mut();
        let (w, h) = (120, 28);
        let dib = unsafe {
            CreateDIBSection(dc, &top_down_bmi(w, h), DIB_RGB_COLORS, &mut bits, None, 0)
        };
        let Ok(dib) = dib else {
            panic!("a scratch surface");
        };
        unsafe { std::ptr::write_bytes(bits, 0, (w * h) as usize * 4) };
        let mem = unsafe { CreateCompatibleDC(dc) };
        let old_bitmap = unsafe { SelectObject(mem, HGDIOBJ(dib.0)) };
        let old_font = unsafe { SelectObject(mem, HGDIOBJ(fonts.button.0)) };
        unsafe {
            SetBkMode(mem, TRANSPARENT);
            SetTextColor(mem, rgb(255, 255, 255));
        }
        let mut run: Vec<u16> = "Cancel".encode_utf16().collect();
        let mut box_ = RECT {
            left: 0,
            top: 0,
            right: w,
            bottom: h,
        };
        unsafe {
            DrawTextW(
                mem,
                &mut run,
                &mut box_,
                DT_CENTER | DT_VCENTER | DT_SINGLELINE,
            )
        };

        let pixels = unsafe { std::slice::from_raw_parts(bits as *const u32, (w * h) as usize) };
        let mut inked = 0;
        for pixel in pixels.iter().copied() {
            let (red, green, blue) = ((pixel >> 16) & 0xFF, (pixel >> 8) & 0xFF, pixel & 0xFF);
            assert_eq!(
                (red, green),
                (green, blue),
                "a glyph pixel is not one coverage: {pixel:#x}"
            );
            if red > 0 {
                inked += 1;
            }
        }
        assert!(inked > 40, "the run drew almost nothing: {inked} pixels");

        unsafe {
            let _ = SelectObject(mem, old_font);
            let _ = SelectObject(mem, old_bitmap);
            let _ = DeleteDC(mem);
            let _ = DeleteObject(HGDIOBJ(dib.0));
            ReleaseDC(HWND::default(), dc);
        }
    }

    /// A rounded box's border follows its corners, all the way round.
    ///
    /// The border is a fill and a fill inside it, so the arc is drawn by the same anti-aliasing as the
    /// shape, and a one-pixel ring is what that leaves. Walked as a band along the outline it was not: a
    /// diagonal line a pixel wide is thinner than a pixel, and no number of samples puts ink in every
    /// pixel an arc that turns inside one pixel crosses — which is what the corners of these buttons
    /// looked like.
    #[test]
    fn the_ring_of_a_button_follows_its_corners() {
        let mut canvas = Canvas::new(80, 60).expect("a screen device context");
        let bounds = Rect::around(10, 10, 70, 50);
        let radius = 9.0_f32;
        let hairline = 1.0_f32;
        // A black button and a white ring, so the red channel *is* the ring's coverage.
        canvas.fill(bounds, radius, rgb(255, 255, 255), 1.0);
        canvas.fill(bounds.inset(hairline), radius - hairline, rgb(0, 0, 0), 1.0);

        // Every point of the ring's own middle line — the outline brought a half-width inward, which is
        // the box shrunk by that much and the corner radius with it — has ink to its own pixel.
        let inner = bounds.inset(hairline * 0.5);
        let arc = radius - hairline * 0.5;
        let (cx, cy) = (
            (inner.left + inner.right) * 0.5,
            (inner.top + inner.bottom) * 0.5,
        );
        let (hw, hh) = (inner.width() * 0.5, inner.height() * 0.5);
        // Each corner, and the quarter its own arc is drawn in: outward is `(sx, sy)`.
        let corners = [
            (cx + (hw - arc), cy + (hh - arc), 1.0, 1.0),
            (cx - (hw - arc), cy + (hh - arc), -1.0, 1.0),
            (cx - (hw - arc), cy - (hh - arc), -1.0, -1.0),
            (cx + (hw - arc), cy - (hh - arc), 1.0, -1.0),
        ];
        let mut walked = 0;
        for (ax, ay, sx, sy) in corners {
            for step in 0..=90 {
                let angle = (step as f32) * std::f32::consts::FRAC_PI_2 / 90.0;
                let px = ax + sx * arc * angle.cos();
                let py = ay + sy * arc * angle.sin();
                // The nine pixels around the point: the ring is a pixel and a bit wide and its arc runs
                // through a box of nine, so a box holding a piece of ring holds well over one pixel's
                // worth of ink — and a box at a hole holds almost none.
                let (x0, y0) = (px.floor() as i32, py.floor() as i32);
                let mut ink = 0;
                for y in y0 - 1..=y0 + 1 {
                    for x in x0 - 1..=x0 + 1 {
                        ink += (canvas.pixels[(y * canvas.w + x) as usize] >> 16) & 0xFF;
                    }
                }
                assert!(
                    ink >= 255,
                    "the ring has a hole at ({px:.2}, {py:.2}): {ink} of ink around it"
                );
                walked += 1;
            }
        }
        assert_eq!(walked, 4 * 91);

        // And the band is a band: the row above the box's own edge carries a sliver of the ring's
        // anti-aliasing and nothing more, where a smeared band would put half of it there.
        let above =
            canvas.pixels[((bounds.top as i32 - 1) * canvas.w + bounds.left as i32 + 20) as usize];
        assert!(
            (above >> 24) <= 32,
            "the ring is smeared outside the box: alpha {}",
            above >> 24
        );
        let middle = canvas.pixels[((bounds.top + bounds.bottom) as i32 / 2 * canvas.w
            + (bounds.left + bounds.right) as i32 / 2) as usize];
        assert_eq!(
            (middle >> 16) & 0xFF,
            0,
            "the fill is not the button's own colour"
        );

        // A straight edge is one pixel of border and not two half ones: the row above it carries the
        // eighth of a pixel a one-pixel anti-aliasing ramp gives away, and the row at it is the border.
        let column = bounds.left as i32 + 30;
        let top = bounds.top as i32;
        let above = (canvas.pixels[((top - 1) * canvas.w + column) as usize] >> 16) & 0xFF;
        let on_edge = (canvas.pixels[(top * canvas.w + column) as usize] >> 16) & 0xFF;
        assert!(
            above * 4 < on_edge,
            "the border is smeared off its edge: {above} above against {on_edge} on it"
        );
        assert!(
            on_edge > 120,
            "the top edge is not a pixel of border: {on_edge} (above {above})"
        );
    }

    #[test]
    fn a_scrim_leaves_the_window_behind_it_showing() {
        let mut canvas = Canvas::new(40, 40).expect("a screen device context");
        let p = Palette::of(TextTheme::Dark);
        canvas.fill(Rect::around(0, 0, 40, 40), 0.0, p.page, SCRIM_SHARE);
        let pixel = canvas.pixels[20 * 40 + 20];
        let alpha = ((pixel >> 24) & 0xFF) as f32 / 255.0;
        assert!(
            (alpha - SCRIM_SHARE).abs() < 0.02,
            "the scrim is opaque over the window it belongs to: {alpha}"
        );
    }

    #[test]
    fn the_card_and_both_of_its_buttons_read_in_both_themes() {
        for theme in [TextTheme::Dark, TextTheme::Light] {
            let p = Palette::of(theme);

            // The heading and the body against the card they are written on — and against the sheen
            // the card's top edge carries, which is what the heading is drawn over.
            let sheen = if theme == TextTheme::Dark {
                mix(p.card, rgb(255, 255, 255), 6)
            } else {
                mix(p.card, rgb(255, 255, 255), 40)
            };
            for surface in [p.card, sheen] {
                assert!(
                    contrast_ratio(p.label, surface) >= 7.0,
                    "{theme:?} heading contrast {:.2} < 7.0",
                    contrast_ratio(p.label, surface)
                );
            }

            // Both buttons' inks on both stops of their gradients, hovered and pressed as well: the
            // ink of a button is read at every state, or it is read only until the pointer arrives.
            for (hovered, pressed) in [(false, false), (true, false), (false, true)] {
                let (top, bottom, _, ink) = affirmative_colors(p, hovered, pressed);
                for stop in [top, bottom] {
                    assert!(
                        contrast_ratio(ink, stop) >= 4.5,
                        "{theme:?} delete ink {:.2} < 4.5 (hov={hovered}, prs={pressed})",
                        contrast_ratio(ink, stop)
                    );
                }
                let (top, bottom, _, ink) = negative_colors(p, hovered, pressed);
                for stop in [top, bottom] {
                    assert!(
                        contrast_ratio(ink, stop) >= 4.5,
                        "{theme:?} cancel ink {:.2} < 4.5 (hov={hovered}, prs={pressed})",
                        contrast_ratio(ink, stop)
                    );
                }
            }

            // And the mark on the badge against the disc it stands in.
            assert!(
                contrast_ratio(p.danger_text, p.danger_bg) >= 3.0,
                "{theme:?} badge mark contrast {:.2} < 3.0",
                contrast_ratio(p.danger_text, p.danger_bg)
            );
        }
    }

    /// WCAG relative luminance of a palette colour.
    fn relative_luminance(c: COLORREF) -> f32 {
        let lin = |ch: u8| -> f32 {
            let s = f32::from(ch) / 255.0;
            if s <= 0.04045 {
                s / 12.92
            } else {
                ((s + 0.055) / 1.055).powf(2.4)
            }
        };
        let (r, g, b) = channels(c);
        0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
    }

    fn contrast_ratio(a: COLORREF, b: COLORREF) -> f32 {
        let la = relative_luminance(a);
        let lb = relative_luminance(b);
        let (hi, lo) = if la >= lb { (la, lb) } else { (lb, la) };
        (hi + 0.05) / (lo + 0.05)
    }
}

//! The tip this app explains its own controls with, on a layered window of its own rather than the
//! platform's `tooltips_class32`.
//!
//! **Why not the platform's.** A `tooltips_class32` is the system's own window: its font, its
//! colours, its padding, the fade it plays and the schedule it decides with. The `Custom Preview`
//! editor is drawn from this app's own `Palette` and its own material — a rounded card, a soft
//! shadow, themed ink — and a tip in the platform's grey box beside a row of it is the one window in
//! that screenshot that came from somewhere else. It is also the one this app already has to work
//! around: a `tooltips_class32` standing over a listing is what `shell::window_classes` reads
//! *through* to find the row beneath it, because the desktop names the tooltip as the window a point
//! is over.
//!
//! **What it is instead.** One borderless, layered, no-activate, click-through, tool-window popup,
//! **owned by the window that asked for it** — so it stands above its owner, is destroyed with it,
//! never takes the keyboard out of the window it describes, and never appears in the taskbar or in
//! Alt+Tab. Its pixels are composed the way the delete card's are (see `ui::ask_card`): a
//! premultiplied-BGRA surface in this app's memory, handed to `UpdateLayeredWindow` in one go, so the
//! rounded corner and the shadow live in the alpha byte GDI never writes.
//!
//! **A tip is composed once per run, not once per showing.** The surface is measured and painted when
//! the run it says changes, and every showing after that is one `SetWindowPos` of a picture that is
//! already composed — a tip that followed the pointer by repainting would be a paint per mouse pixel.
//!
//! **What it must never do is go behind its owner or off the screen.** It is placed in the work area
//! of the monitor the tip is about, above the control, and flipped below it only where the work area
//! has no room above.

use std::cell::RefCell;
use std::ptr;
use std::sync::Once;

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, POINT, RECT, SIZE, WPARAM};
use windows::Win32::Graphics::Gdi::{
    CreateCompatibleDC, CreateDIBSection, CreateFontW, DeleteDC, DeleteObject, DrawTextW, GetDC,
    GetMonitorInfoW, MonitorFromWindow, ReleaseDC, SelectObject, SetBkMode, SetTextColor,
    AC_SRC_ALPHA, AC_SRC_OVER, ANTIALIASED_QUALITY, BITMAPINFO, BITMAPINFOHEADER, BI_RGB,
    BLENDFUNCTION, CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET, DEFAULT_PITCH, DIB_RGB_COLORS,
    DRAW_TEXT_FORMAT, DT_CALCRECT, DT_NOPREFIX, DT_WORDBREAK, FW_NORMAL, HDC, HFONT, HGDIOBJ,
    MONITORINFO, MONITOR_DEFAULTTONEAREST, OUT_DEFAULT_PRECIS, TRANSPARENT,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::HiDpi::GetDpiForWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindowLongPtrW, GetWindowRect, IsWindow,
    IsWindowVisible, RegisterClassExW, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    UpdateLayeredWindow, GWLP_USERDATA, HWND_TOP, MA_NOACTIVATE, SWP_NOACTIVATE, SW_SHOWNOACTIVATE,
    ULW_ALPHA, WM_ERASEBKGND, WM_MOUSEACTIVATE, WM_NCDESTROY, WM_NCHITTEST, WNDCLASSEXW,
    WNDCLASS_STYLES, WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TRANSPARENT,
    WS_POPUP,
};

use super::custom_preview_settings::{
    rgb, scale, Palette, ID_ADD, ID_ARGUMENTS, ID_BROKER_FALLBACK, ID_BROWSE, ID_CAPTURE_TECHNIQUE,
    ID_CHILD_SCOPE, ID_CLOSE_BUTTON, ID_DELETE, ID_ENABLED, ID_EXE, ID_EXTENSIONS, ID_LIST,
    ID_NAME, ID_PATH_METHOD, ID_PICTURE_WINDOW, ID_PLACEMENT, ID_PROMOTION_TECHNIQUE, ID_SCALING,
    ID_SCALING_DOWN, ID_SCALING_PERCENT, ID_SCALING_UP, ID_SEARCH, ID_SIBLING_WINDOWS, ID_STATUS,
    ID_TITLEBAR_CLOSE, ID_TITLEBAR_MAXIMIZE, ID_TITLEBAR_MINIMIZE, ID_VOLUME, ID_VOLUME_DOWN,
    ID_VOLUME_ENABLED, ID_VOLUME_UP, ID_VOLUME_VALUE, ID_WINDOW_MODE,
};

/// What one control says about itself, and what each of its values says, all of it measured before a
/// pointer is anywhere near the window: a tip is composed when it is put up and never on the pointer's
/// own path, so the words are a table rather than a call.
/// A control's own entry is what hovering it says. A combo's entry per value is what hovering that
/// value in its dropped list says, so `Capture Method` explains the road it is set to and
/// `In-Process Hook` explains the road it is.
pub(super) struct Wording {
    /// The tip for the control itself.
    pub(super) whole: &'static str,
    /// The tip for each value of a list, in the order the list holds them.
    pub(super) values: &'static [&'static str],
}

/// A wording with one sentence and no values, for the controls that are one thing rather than a list.
pub(super) const fn one(whole: &'static str) -> Wording {
    Wording { whole, values: &[] }
}

/// A wording whose control is a list: the control's own sentence, then one per value.
pub(super) const fn listed(whole: &'static str, values: &'static [&'static str]) -> Wording {
    Wording { whole, values }
}

/// The sentence for a value of a list, or the control's own where the list has no entry for it.
pub(super) fn value_or_whole(wording: &Wording, index: usize) -> &'static str {
    wording.values.get(index).copied().unwrap_or(wording.whole)
}

/// Every control's own sentence: what the row is for rather than what it is called.
///
/// **A tip that repeated the label would be worth nothing**, so each one says the thing the label
/// cannot: `Broker Fallback` is about a program begun by something else, `Child Processes` about a
/// launcher's window not being the launched pid's, `Sibling Windows` about a splash nobody asked for.
pub(super) fn editor(id: i32) -> Option<Wording> {
    Some(match id {
        ID_TITLEBAR_MINIMIZE => one("Put the editor away in the taskbar."),
        ID_TITLEBAR_MAXIMIZE => one("Fill the screen with the editor, or put it back to its own size."),
        ID_TITLEBAR_CLOSE => one("Close the editor. Every field was saved as it was left."),
        ID_SEARCH => one(
            "Narrow the list to the presets whose name holds what is typed here. An empty box puts \
             every preset back on it.",
        ),
        ID_ADD => one("Start a new preset: no program and no extensions, nothing looked up yet."),
        ID_LIST => one(
            "Every preset, top first. A name higher up wins an extension two of them claim. Drag a \
             row to move it; the pill on the right turns that preset on or off without opening it.",
        ),
        ID_ENABLED => one(
            "Whether this preset is used at all. Off keeps it written and hands its extensions back \
             to the built-in previewers.",
        ),
        ID_NAME => one(
            "What the preset is called in this list and in config.ini. A preset needs a name to be \
             saved, and two presets may not share one.",
        ),
        ID_EXTENSIONS => one(
            "The extensions this preset claims, comma-separated, without their dots: psd, xcf. A name \
             here outranks every list the app ships.",
        ),
        ID_EXE => one(
            "The program started for a claimed file, and the program whose window becomes the preview. \
             A path pasted with Copy as path keeps its quotes off.",
        ),
        ID_BROWSE => one("Pick the program with the system's own file dialog."),
        ID_ARGUMENTS => one(
            "Three placeholders stand where the file goes. {file} is the path in this preset's own \
             Path Method spelling; {path} is the plain Windows path whatever that row says; {url} is \
             a file:// URL, for a program that will only take one. Any of them may be named more \
             than once, or sit inside a value as --input={file}. Where none of the three appears \
             the path is appended last, in {file}'s own spelling; a brace group that is none of the \
             three is left exactly as written, so a file named {url}.mkv keeps its name. Quotes group \
             an argument, backslashes escape them.",
        ),
        ID_PATH_METHOD => listed(
            "How the file is spelled where Arguments stands it, and where it is appended with no \
             placeholder written. A program that says it cannot open the stream is usually a program \
             handed a path where it wanted a URL, or the other way round.",
            PATH_METHOD_WORDS,
        ),
        ID_CAPTURE_TECHNIQUE => listed(
            "How the program's window is kept off the screen while its picture is the preview. Each \
             choice is a different road to that one end.",
            CAPTURE_TECHNIQUE_WORDS,
        ),
        ID_CHILD_SCOPE => listed(
            "How far past the program this app may look for the window, and so how far it may hide, \
             and on teardown end, a process you never named.",
            CHILD_SCOPE_WORDS,
        ),
        ID_BROKER_FALLBACK => one(
            "Adopt a window begun by something other than the program itself: Windows 11's Paint is \
             started by an activator, and an installer's stub exits once the real program is up. Off \
             unless asked for, because that pid is one you never named.",
        ),
        ID_SIBLING_WINDOWS => listed(
            "What happens to the program's other windows while one of them is the picture.",
            SIBLING_WINDOW_WORDS,
        ),
        ID_PICTURE_WINDOW => listed(
            "Which of the program's windows is the picture, where it made more than one a document \
             could be drawn in. The other remains a window of its own; Sibling Windows says what \
             happens to it.",
            PICTURE_WINDOW_WORDS,
        ),
        ID_PROMOTION_TECHNIQUE => listed(
            "What the pin key does with this preview.",
            PROMOTION_TECHNIQUE_WORDS,
        ),
        ID_PLACEMENT => listed(
            "Where the window sits, and what happens when the program has an opinion of its own.",
            PLACEMENT_WORDS,
        ),
        ID_WINDOW_MODE => listed(
            "Whether the preview keeps the program's own frame.",
            WINDOW_MODE_WORDS,
        ),
        ID_SCALING => listed(
            "How large the picture is drawn in the hover.",
            SCALING_WORDS,
        ),
        ID_SCALING_PERCENT => one(
            "The share of the room beside the item the preview takes, cut to the file's own shape. Up \
             and down step by one; a field pasted as one wrapped run of flags is unwrapped.",
        ),
        ID_SCALING_UP => one("One percent more of the room."),
        ID_SCALING_DOWN => one("One percent less of the room."),
        ID_VOLUME_ENABLED => one(
            "Set this program's own volume while its preview is up, on every audio session its \
             processes hold. Off leaves the Windows mixer alone.",
        ),
        ID_VOLUME => one("The level a preview of this preset is played at."),
        ID_VOLUME_VALUE => one("The level a preview of this preset is played at, typed. Up and down step by one."),
        ID_VOLUME_UP => one("One percent louder."),
        ID_VOLUME_DOWN => one("One percent quieter."),
        ID_DELETE => one("Delete this preset from config.ini. It asks first, and nothing here puts it back."),
        ID_CLOSE_BUTTON => one("Close the editor. Every field was saved as it was left."),
        ID_STATUS => one(
            "What the last save said: empty when it went through, and the reason the window is \
             holding what it holds when it did not.",
        ),
        _ => return None,
    })
}

/// `Path Method`, value by value, in the order `PATH_METHODS` holds them.
const PATH_METHOD_WORDS: &[&str] = &[
    "E:\\Videos\\a b.mkv — the path as Windows writes it. What every preset did before this row \
     existed, and what a program that takes a path rather than a URL wants.",
    "file:///E:/Videos/a%20b.mkv — separators turned round and what a URL may not carry escaped, \
     with a share written file://server/share/…. What VLC, mpv's URL handling and most programs \
     that take a stream address want; handed a backslashed path they answer that they cannot open \
     the MRL.",
    "E:/Videos/a b.mkv — the separators turned round and nothing escaped. For a program that reads / \
     but treats a backslash as an escape character in its own arguments, such as ffmpeg's option \
     parser.",
    "E:\\VIDEO~1\\A~1.MKV — the volume's own 8.3 name, which has no spaces and no non-ASCII in it. For \
     an old program that splits its command line on spaces. A volume that keeps no 8.3 names, and a \
     file that is already short, are handed the plain path instead.",
    "file:///E:/Videos/a b.mkv — the URL with nothing escaped. For the rare program that takes a URL \
     and then hands %20 to the filesystem as those four characters rather than as a space.",
];

/// `Capture Method`, value by value, in the order `CAPTURE_TECHNIQUES` holds them.
const CAPTURE_TECHNIQUE_WORDS: &[&str] = &[
    "The only road that never puts the window on the screen at all: custom_preview_hook.dll is \
     loaded into the program while it is still suspended, and it cloaks its window and locks its \
     rectangle from inside, before the compositor sees it. Needs the .dll beside the executable; \
     without it this falls back to Invisible Layered.",
    "The window is made fully transparent with WS_EX_LAYERED at 1/255 alpha rather than hidden: it \
     is still composed, so the Shell can still draw it, but no pixel of it reaches the screen.",
    "The window stands on the monitor behind an opaque window of this app's that covers it, and the \
     preview is that window's picture of it. On-screen, so the compositor always has it.",
    "The window is parked past the edge of every display and the preview is drawn from it there. The \
     simplest road, and the one furthest from the pointer.",
    "The program's own window is placed at the preview's box, click-through and topmost, with no \
     picture of this app's and no thumbnail. The fastest, and the only one where the program is the \
     window you see.",
    "The window is staged past every display, given a moment to settle, then revealed in place in \
     one step across a compositor flush — so nothing of it is ever seen arriving.",
];

/// `Child Processes`, value by value.
const CHILD_SCOPE_WORDS: &[&str] = &[
    "Only the process this app started. No tree walk, no job members, and the hook is never written \
     into a child. A launcher or a packaged app then has no window to find.",
    "The program and whatever it started directly. Covers a one-level stub and a program that hands \
     its file to a single helper.",
    "The program and everything below it, as deep as the parent links go. What every preset did \
     before this row existed, and what a launcher's grandchild needs.",
];

/// `Sibling Windows`, value by value.
const SIBLING_WINDOW_WORDS: &[&str] = &[
    "The program's other windows are left alone. A splash or an About box will stand over the \
     listing for as long as the program keeps it.",
    "Every candidate is cloaked while the window is being picked, and nothing is hidden afterwards. \
     The splash is never composited, and it stays where it is.",
    "Cloaked while the window is picked, then hidden once the picture is up. A modal dialog is closed \
     instead of hidden, because hiding one leaves the window that asked for it disabled.",
];

/// `Picture Window`, value by value.
const PICTURE_WINDOW_WORDS: &[&str] = &[
    "The largest window the program made that a picture could be drawn in — what every preset did \
     before this row existed, and right for a program that makes its document window first or \
     biggest.",
    "Of the windows of that same size class, the one the program made last, once it has stopped \
     making them. Right for a program whose own helper comes up first and larger: MPV Player's \
     mpv.exe is built with SMTC support, and its `mpv smtc` window is born 200 ms before the film \
     and is handed the preview's own box, so Largest adopts the helper and the preview is a white \
     still frame.",
];

/// `Promotion`, value by value.
const PROMOTION_TECHNIQUE_WORDS: &[&str] = &[
    "The preview becomes the window: the running program's window is detached, uncloaked and \
     placed where the hover was, with its own frame back. Falls back to a fresh launch when cold.",
    "The preview ends and the file is opened in an untouched new process, the way the program would \
     have opened it from Explorer.",
];

/// `Placement`, value by value.
const PLACEMENT_WORDS: &[&str] = &[
    "The window is placed once at Tray -> Placement's box and trusted to stay there.",
    "The window is put back whenever it drifts, and shifted clear of the pointer where the program's \
     own minimum size makes the box overlap the hand.",
    "The same enforcement from inside the program, through the hook: a resize, a move, a maximize \
     and a foreground grab are refused where the program attempts them. Needs the .dll.",
];

/// `Window`, value by value.
const WINDOW_MODE_WORDS: &[&str] = &[
    "The program's own frame is kept and the whole window is drawn, frame included.",
    "The frame is stripped and only the client area is drawn, so the picture is the document and \
     nothing else.",
];

/// `Scaling`, value by value.
const SCALING_WORDS: &[&str] = &[
    "Whatever the tray's own Scaling row says for this kind of file — what every preview did before \
     a preset could say otherwise.",
    "The whole of the room beside the item, cut to the file's own shape.",
    "The room, at the display's own size rather than the file's full width.",
    "Three quarters of that room.",
    "Half of that room.",
    "A share typed into the box beside this one.",
];

/// How many `(id, Wording)` pairs the editor explains, for the test that holds the table to the
/// control tree rather than to a count in a comment.
pub(super) const EXPLAINED: &[i32] = &[
    ID_TITLEBAR_MINIMIZE,
    ID_TITLEBAR_MAXIMIZE,
    ID_TITLEBAR_CLOSE,
    ID_SEARCH,
    ID_ADD,
    ID_LIST,
    ID_ENABLED,
    ID_NAME,
    ID_EXTENSIONS,
    ID_EXE,
    ID_BROWSE,
    ID_ARGUMENTS,
    ID_PATH_METHOD,
    ID_CAPTURE_TECHNIQUE,
    ID_CHILD_SCOPE,
    ID_BROKER_FALLBACK,
    ID_SIBLING_WINDOWS,
    ID_PICTURE_WINDOW,
    ID_PROMOTION_TECHNIQUE,
    ID_PLACEMENT,
    ID_WINDOW_MODE,
    ID_SCALING,
    ID_SCALING_PERCENT,
    ID_SCALING_UP,
    ID_SCALING_DOWN,
    ID_VOLUME_ENABLED,
    ID_VOLUME,
    ID_VOLUME_VALUE,
    ID_VOLUME_UP,
    ID_VOLUME_DOWN,
    ID_DELETE,
    ID_CLOSE_BUTTON,
    ID_STATUS,
];

/// What a combo says about itself, and about one of its values.
///
/// `index` is the value a control's own box shows — one sender, one sentence — and `edit_hint` is what
/// a run under a combo's list says instead: the same words, read from the same place, so the box and
/// the list never disagree about what a value is.
pub(super) fn combo(id: i32, index: Option<usize>) -> Option<&'static str> {
    let wording = editor(id)?;
    Some(row(&wording, index))
}

/// The statement a hovered list row makes for a combo: the value's own sentence, and the control's
/// when the row is no row at all (`CB_GETCURSEL` answering nothing).
pub(super) fn row(wording: &Wording, index: Option<usize>) -> &'static str {
    match index {
        Some(index) => value_or_whole(wording, index),
        None => wording.whole,
    }
}

/// The class every tip is made from.
const CLASS_NAME: PCWSTR = w!("RustHoverPreviewTip");

/// How wide a tip's line of text may be, in logical pixels, and the room the pad around it takes. A
/// tip is a sentence rather than a paragraph: the column is narrow enough to be read at a glance and
/// wide enough that the two- and three-word ones sit on one line.
const COLUMN_LOGICAL: i32 = 272;
const PAD_X_LOGICAL: i32 = 10;
const PAD_Y_LOGICAL: i32 = 7;
const RADIUS_LOGICAL: i32 = 6;
const SHADOW_LOGICAL: i32 = 10;
const SIZE_LOGICAL: i32 = 12;
/// How far off the control a tip stands, and how far inside the work area it keeps.
const STANDOFF_LOGICAL: i32 = 6;
const EDGE_LOGICAL: i32 = 4;

static CLASS_REGISTERED: Once = Once::new();

/// What one tip window holds: the surface it draws and the run that surface was composed for.
struct Tip {
    /// The composed premultiplied-BGRA surface, already at its own size.
    pixels: Vec<u32>,
    width: i32,
    height: i32,
    /// The run the surface says, and the palette it was composed in.
    text: String,
    palette: Palette,
}

fn tip_of(hwnd: HWND) -> Option<&'static mut Tip> {
    let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut Tip;
    unsafe { raw.as_mut() }
}

// How many tip windows this thread has made and not destroyed.
//
// **A count and not a nicety**: a popup is a window of the desktop until it is destroyed, and a tip
// whose window is left standing after its sentence was taken down is a leak that nothing else in the
// app would ever see. The editor's tests hold it to zero — and it is **per thread** rather than per
// process because a tip is: `SHOWING` is this thread's, tests run in parallel on threads of their
// own, and a count two of them shared would fail for a tip the other one was holding up.
#[cfg(test)]
thread_local! {
    static LIVE: std::cell::Cell<u32> = const { std::cell::Cell::new(0) };
}

#[cfg(test)]
pub(super) fn live_windows() -> u32 {
    LIVE.with(|live| live.get())
}

#[cfg(test)]
fn note_a_tip_made() {
    LIVE.with(|live| live.set(live.get() + 1));
}

#[cfg(test)]
fn note_a_tip_gone() {
    LIVE.with(|live| live.set(live.get().saturating_sub(1)));
}

thread_local! {
    /// The tip this thread has up, if any: one at a time, because a pointer is over one control at a
    /// time. Held so that a run which has not changed is a window moved rather than one rebuilt.
    static SHOWING: RefCell<Option<HWND>> = const { RefCell::new(None) };
}

unsafe extern "system" fn tip_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        // A tip is not a window the pointer is over: `HTTRANSPARENT` hands every press it receives
        // to the window under it, which is the control the tip describes.
        WM_NCHITTEST => LRESULT(-1),
        WM_MOUSEACTIVATE => LRESULT(MA_NOACTIVATE as isize),
        WM_ERASEBKGND => LRESULT(1),
        WM_NCDESTROY => {
            let raw = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut Tip;
            if !raw.is_null() {
                unsafe {
                    let _ = SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                    drop(Box::from_raw(raw));
                }
            }
            #[cfg(test)]
            note_a_tip_gone();
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

fn ensure_class() {
    CLASS_REGISTERED.call_once(|| {
        let Ok(hinstance) = (unsafe { GetModuleHandleW(None) }) else {
            return;
        };
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: WNDCLASS_STYLES(0),
            lpfnWndProc: Some(tip_proc),
            hInstance: hinstance.into(),
            lpszClassName: CLASS_NAME,
            ..Default::default()
        };
        unsafe {
            let _ = RegisterClassExW(&class);
        }
    });
}

/// Whether a tip is on the screen, which is what the editor's clock is held to.
#[cfg_attr(not(test), allow(dead_code))]
pub(super) fn is_up() -> bool {
    SHOWING.with(|showing| {
        showing
            .borrow()
            .is_some_and(|hwnd| unsafe { IsWindowVisible(hwnd) }.as_bool())
    })
}

/// Take the tip down, and take its window with it.
///
/// **Destroyed rather than hidden**, because a hidden window is a window the desktop still counts: the
/// tip is one popup as long as it is on screen and nothing of it is left standing behind a hover that
/// has moved on. A run that changes while the tip stays up (`show` with a new sentence) is recomposed
/// into the same window, so a hand travelling down a list of six values makes one popup, not six.
pub(super) fn hide() {
    SHOWING.with(|showing| {
        let taken = showing.borrow_mut().take();
        if let Some(hwnd) = taken {
            if unsafe { IsWindow(hwnd) }.as_bool() {
                let _ = unsafe { DestroyWindow(hwnd) };
            }
        }
    });
}

/// Put a tip up, standing off `anchor` — a box in screen pixels rather than a window.
///
/// **A box and not a control**, because half of what the editor shows is not a control: a row's tip
/// stands over the row (the label's words and the field beside them, one line of it), and a value in a
/// dropped list stands over the combo it belongs to. The window a tip is owned by is the editor, so it
/// stays above it and goes with it.
pub(super) fn show_at(anchor: RECT, owner: HWND, run: &str, palette: Palette) -> bool {
    if run.is_empty() || owner.0.is_null() {
        return false;
    }
    ensure_class();
    let dpi = unsafe { GetDpiForWindow(owner) }.max(96) as i32;

    let held = SHOWING
        .with(|showing| *showing.borrow())
        .unwrap_or_default();
    let usable = !held.0.is_null() && unsafe { IsWindow(held) }.as_bool();
    let same_run =
        usable && tip_of(held).is_some_and(|tip| tip.text == run && tip.palette == palette);

    let tip = if same_run {
        held
    } else {
        if usable {
            // The same window, told a new sentence: its surface is recomposed under it rather than a
            // window made for every run the pointer crosses.
            let recomposed = match tip_of(held) {
                Some(tip) => unsafe { recompose(held, tip, run, palette, dpi) },
                None => None,
            };
            match recomposed {
                Some(()) => held,
                None => {
                    hide();
                    let Some(made) = (unsafe { make_tip(owner, run, palette, dpi) }) else {
                        return false;
                    };
                    SHOWING.with(|showing| *showing.borrow_mut() = Some(made));
                    made
                }
            }
        } else {
            hide();
            let Some(made) = (unsafe { make_tip(owner, run, palette, dpi) }) else {
                return false;
            };
            SHOWING.with(|showing| *showing.borrow_mut() = Some(made));
            made
        }
    };

    let Some((width, height)) = tip_of(tip).map(|tip| (tip.width, tip.height)) else {
        return false;
    };
    if width <= 0 || height <= 0 {
        return false;
    }
    let (x, y) = place(anchor, width, height, dpi, owner);
    let moved = unsafe { SetWindowPos(tip, HWND_TOP, x, y, width, height, SWP_NOACTIVATE) }.is_ok();
    // `ShowWindow` answers with the state the window was **in**, not with whether it worked: a tip
    // that was hidden until now answers `false` and is on the screen.
    let _ = unsafe { ShowWindow(tip, SW_SHOWNOACTIVATE) };
    moved && unsafe { IsWindowVisible(tip) }.as_bool()
}

/// Where a tip of this size goes for this anchor: above the control where the work area has room,
/// below it where it has none, and never outside that work area.
fn place(anchor: RECT, width: i32, height: i32, dpi: i32, owner: HWND) -> (i32, i32) {
    let mut work = RECT {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
    };
    let monitor = unsafe { MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST) };
    let mut info = MONITORINFO {
        cbSize: std::mem::size_of::<MONITORINFO>() as u32,
        ..Default::default()
    };
    if unsafe { GetMonitorInfoW(monitor, &mut info) }.as_bool() {
        work = info.rcWork;
    }

    let standoff = scale(STANDOFF_LOGICAL, dpi);
    let edge = scale(EDGE_LOGICAL, dpi);
    let above = anchor.top - standoff - height;
    let below = anchor.bottom + standoff;
    let y = if above >= work.top + edge {
        above
    } else {
        below
    };
    let centred = (anchor.left + anchor.right - width) / 2;
    let x = centred.clamp(
        work.left + edge,
        (work.right - width - edge).max(work.left + edge),
    );
    (
        x,
        y.clamp(
            work.top + edge,
            (work.bottom - height - edge).max(work.top + edge),
        ),
    )
}

/// Make the tip window: compose its surface, hand it to `UpdateLayeredWindow`, and leave it hidden.
unsafe fn make_tip(owner: HWND, run: &str, palette: Palette, dpi: i32) -> Option<HWND> {
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT,
            CLASS_NAME,
            w!(""),
            WS_POPUP,
            0,
            0,
            1,
            1,
            owner,
            None,
            None,
            None,
        )
    }
    .ok()?;
    #[cfg(test)]
    note_a_tip_made();

    let tip = Box::new(Tip {
        pixels: Vec::new(),
        width: 0,
        height: 0,
        text: String::new(),
        palette,
    });
    unsafe { SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(tip) as isize) };
    let held = tip_of(hwnd)?;
    if unsafe { compose(hwnd, held, run, palette, dpi) }.is_none() {
        unsafe {
            let _ = DestroyWindow(hwnd);
        }
        return None;
    }
    Some(hwnd)
}

/// Compose a new sentence into a tip that is already up, or nothing where it cannot be drawn.
unsafe fn recompose(
    hwnd: HWND,
    tip: &mut Tip,
    run: &str,
    palette: Palette,
    dpi: i32,
) -> Option<()> {
    unsafe { compose(hwnd, tip, run, palette, dpi) }
}

/// The tip drawn: a shadow, a rounded card in the palette's own material, a hairline, and the run.
///
/// Its height comes from the run rather than from a guess — the text is measured at the tip's column
/// before the surface is made, so a sentence that wraps is a card two lines tall and one that does not
/// is a card one line tall.
unsafe fn compose(hwnd: HWND, tip: &mut Tip, run: &str, palette: Palette, dpi: i32) -> Option<()> {
    let font = unsafe { tip_font(dpi) };
    if font.0.is_null() {
        return None;
    }
    let column = scale(COLUMN_LOGICAL, dpi);
    let pad_x = scale(PAD_X_LOGICAL, dpi);
    let pad_y = scale(PAD_Y_LOGICAL, dpi);
    let shadow = scale(SHADOW_LOGICAL, dpi);

    let (text_w, text_h) = unsafe { measure(font, run, column) };
    let width = text_w + pad_x * 2;
    let height = text_h + pad_y * 2;
    let canvas_w = width + shadow * 2;
    let canvas_h = height + shadow * 2;

    let mut canvas = Canvas::new(canvas_w, canvas_h)?;
    let body = Box_ {
        left: shadow as f32,
        top: shadow as f32,
        right: (shadow + width) as f32,
        bottom: (shadow + height) as f32,
    };
    let radius = scale(RADIUS_LOGICAL, dpi) as f32;
    canvas.shadow(
        body,
        radius,
        shadow as f32,
        if palette.dark { 0.45 } else { 0.18 },
    );
    canvas.fill(body, radius, palette.border, 1.0);
    canvas.fill(body.inset(1.0), (radius - 1.0).max(0.0), palette.card, 1.0);
    // **The card is opaque, and a layered window is what its pixels say.** `UpdateLayeredWindow`
    // takes each pixel's alpha as given — the one call it does not honour is `LWA_OPAQUE` — and a
    // card whose every pixel was blended over an empty surface came out at about half the coverage
    // it was drawn at: the shadow read, and the words inside it did not.
    canvas.opaque(body, radius);
    canvas.text(
        font,
        RECT {
            left: shadow + pad_x,
            top: shadow + pad_y,
            right: shadow + pad_x + text_w,
            bottom: shadow + pad_y + text_h,
        },
        run,
        palette.label,
        DT_WORDBREAK | DT_NOPREFIX,
    );
    unsafe {
        let _ = DeleteObject(HGDIOBJ(font.0));
    }

    tip.pixels = canvas.pixels().to_vec();
    tip.width = canvas_w;
    tip.height = canvas_h;
    tip.text = run.to_string();
    tip.palette = palette;
    unsafe { present(hwnd, tip) };
    Some(())
}

/// Hand the composed surface to the window, as the layered window it is.
unsafe fn present(hwnd: HWND, tip: &Tip) {
    let (width, height) = (tip.width, tip.height);
    if hwnd.0.is_null()
        || width <= 0
        || height <= 0
        || tip.pixels.len() < (width as usize) * (height as usize)
    {
        return;
    }
    let screen = unsafe { GetDC(HWND::default()) };
    if screen.0.is_null() {
        return;
    }
    let mut bits: *mut std::ffi::c_void = ptr::null_mut();
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
            tip.pixels.as_ptr(),
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
    // **The destination is a screen point**, and this update is the one window in the app whose
    // origin it decides: `UpdateLayeredWindow` moves the window to it as well as painting it, so a
    // tip composed while the window stood at `(0, 0)` would otherwise be put there.
    let mut origin = RECT::default();
    if unsafe { GetWindowRect(hwnd, &mut origin) }.is_err() {
        origin = RECT {
            left: 0,
            top: 0,
            right: 0,
            bottom: 0,
        };
    }
    let dst = POINT {
        x: origin.left,
        y: origin.top,
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
            hwnd,
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

/// The tip's own face: the editor's body size, grayscale anti-aliased so that a run read back as
/// coverage is the glyph's coverage rather than a subpixel pattern.
unsafe fn tip_font(dpi: i32) -> HFONT {
    let face: Vec<u16> = "Segoe UI"
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    unsafe {
        CreateFontW(
            -scale(SIZE_LOGICAL, dpi),
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
            ANTIALIASED_QUALITY.0 as u32,
            DEFAULT_PITCH.0 as u32,
            PCWSTR(face.as_ptr()),
        )
    }
}

/// The box a run takes at `column` pixels wide: `(width, height)`.
unsafe fn measure(font: HFONT, run: &str, column: i32) -> (i32, i32) {
    let dc = unsafe { GetDC(HWND::default()) };
    if dc.0.is_null() {
        return (column, scale(SIZE_LOGICAL, 96));
    }
    let old = unsafe { SelectObject(dc, HGDIOBJ(font.0)) };
    let mut text: Vec<u16> = run.encode_utf16().collect();
    let mut box_ = RECT {
        left: 0,
        top: 0,
        right: column,
        bottom: 0,
    };
    unsafe {
        DrawTextW(
            dc,
            &mut text,
            &mut box_,
            DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX,
        );
        let _ = SelectObject(dc, old);
        let _ = ReleaseDC(HWND::default(), dc);
    }
    (
        (box_.right - box_.left).clamp(1, column),
        (box_.bottom - box_.top).max(1),
    )
}

/// A box in the surface's own pixels.
#[derive(Clone, Copy)]
struct Box_ {
    left: f32,
    top: f32,
    right: f32,
    bottom: f32,
}

impl Box_ {
    fn width(self) -> f32 {
        (self.right - self.left).max(0.0)
    }
    fn height(self) -> f32 {
        (self.bottom - self.top).max(0.0)
    }
    fn inset(self, by: f32) -> Self {
        Self {
            left: self.left + by,
            top: self.top + by,
            right: self.right - by,
            bottom: self.bottom - by,
        }
    }
}

fn channels(colour: COLORREF) -> (u8, u8, u8) {
    (
        (colour.0 & 0xFF) as u8,
        ((colour.0 >> 8) & 0xFF) as u8,
        ((colour.0 >> 16) & 0xFF) as u8,
    )
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
fn covered(px: f32, py: f32, coverage: impl Fn(f32, f32) -> f32) -> f32 {
    const QUARTERS: [(f32, f32); 4] = [(0.75, 0.75), (0.25, 0.75), (0.75, 0.25), (0.25, 0.25)];
    let mut total = 0.0;
    for (dx, dy) in QUARTERS {
        total += coverage(px + dx - 0.5, py + dy - 0.5);
    }
    total / QUARTERS.len() as f32
}

/// A premultiplied-BGRA canvas for the tip, in this app's memory rather than in a DIB — the same
/// arithmetic the delete card's surface is composed with (`ui::ask_card::Canvas`), and for the same
/// reason: the rounded corner and the shadow live in the alpha byte GDI does not write.
struct Canvas {
    w: i32,
    h: i32,
    pixels: Vec<u32>,
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

    /// Make a shape **opaque inside its own outline**: the card's coverage is one everywhere the
    /// rounded box holds, so a pixel of the words inside it is never composited against the desktop
    /// behind the tip. The outline itself keeps the anti-aliasing `fill` blended it with, which is
    /// what rounds the corner rather than stepping it.
    fn opaque(&mut self, bounds: Box_, radius: f32) {
        let cx = (bounds.left + bounds.right) * 0.5;
        let (ix0, iy0, ix1, iy1) = self.area(bounds, 0.0);
        for y in iy0..iy1 {
            let py = y as f32 + 0.5;
            let core = Self::covered_half(bounds, radius, py);
            if core <= 0.0 {
                continue;
            }
            let core_x0 = ((cx - core).round() as i32).clamp(ix0, ix1);
            let core_x1 = ((cx + core).round() as i32).clamp(core_x0, ix1);
            for x in core_x0..core_x1 {
                let index = (y as usize) * (self.w as usize) + x as usize;
                self.pixels[index] |= 0xFF00_0000;
            }
        }
    }

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

    fn area(&self, bounds: Box_, pad: f32) -> (i32, i32, i32, i32) {
        (
            ((bounds.left - pad).floor() as i32).max(0),
            ((bounds.top - pad).floor() as i32).max(0),
            ((bounds.right + pad).ceil() as i32).min(self.w),
            ((bounds.bottom + pad).ceil() as i32).min(self.h),
        )
    }

    fn covered_half(bounds: Box_, radius: f32, py: f32) -> f32 {
        let hw = bounds.width() * 0.5;
        let hh = bounds.height() * 0.5;
        let r = radius.min(hw).min(hh).max(0.0);
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

    fn fill(&mut self, bounds: Box_, radius: f32, colour: COLORREF, alpha: f32) {
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

    fn shadow(&mut self, bounds: Box_, radius: f32, spread: f32, alpha: f32) {
        let cx = (bounds.left + bounds.right) * 0.5;
        let cy = (bounds.top + bounds.bottom) * 0.5;
        let hw = bounds.width() * 0.5;
        let hh = bounds.height() * 0.5;
        let drop = 2.0;
        let grown = Box_ {
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

    /// A run of text in the palette's ink, read back from GDI as coverage — the same road as the
    /// card's own `Canvas::text`, so the tip's words are hinted rather than outlined.
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
        let mut bits: *mut std::ffi::c_void = ptr::null_mut();
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
    use crate::ui::custom_preview_settings::{
        CHILD_SCOPES, ID_TITLEBAR_CLOSE, ID_TITLEBAR_MAXIMIZE, ID_TITLEBAR_MINIMIZE, PATH_METHODS,
        PLACEMENTS, PROMOTION_TECHNIQUES, SIBLING_WINDOW_MODES, WINDOW_MODES,
    };

    /// Every control the editor lays out has a sentence, and every sentence names a control that
    /// exists: the table and the window are held to each other rather than to a comment.
    #[test]
    fn every_control_the_editor_lays_out_is_explained() {
        for id in EXPLAINED {
            let wording = editor(*id).unwrap_or_else(|| panic!("control {id} has no tip"));
            assert!(
                !wording.whole.trim().is_empty(),
                "control {id} has an empty tip"
            );
            assert!(
                wording.whole.ends_with('.') || wording.whole.ends_with('}'),
                "control {id}'s tip is not a sentence: {:?}",
                wording.whole
            );
        }
        // The titlebar's three are drawn rather than laid out, and are explained all the same.
        for id in [
            ID_TITLEBAR_MINIMIZE,
            ID_TITLEBAR_MAXIMIZE,
            ID_TITLEBAR_CLOSE,
        ] {
            assert!(editor(id).is_some(), "titlebar button {id} is unexplained");
        }
        // Nothing outside the table answers: an id that is not a control of this window gets nothing.
        assert!(editor(0).is_none());
        assert!(editor(9_999).is_none());
    }

    /// A list's tips are one per value, and every list the editor draws is covered.
    #[test]
    fn every_value_of_every_list_is_explained() {
        let lists: [(i32, usize); 7] = [
            (ID_CAPTURE_TECHNIQUE, 6),
            (ID_PATH_METHOD, PATH_METHODS.len()),
            (ID_CHILD_SCOPE, CHILD_SCOPES.len()),
            (ID_SIBLING_WINDOWS, SIBLING_WINDOW_MODES.len()),
            (ID_PROMOTION_TECHNIQUE, PROMOTION_TECHNIQUES.len()),
            (ID_PLACEMENT, PLACEMENTS.len()),
            (ID_WINDOW_MODE, WINDOW_MODES.len()),
        ];
        for (id, count) in lists {
            let wording = editor(id).expect("the list's own tip");
            assert_eq!(
                wording.values.len(),
                count,
                "list {id} explains {} of its {count} values",
                wording.values.len()
            );
            for (index, value) in wording.values.iter().enumerate() {
                assert!(
                    !value.trim().is_empty(),
                    "list {id}'s value {index} is unexplained"
                );
                // Each value says something the control's own sentence does not: a list whose
                // values repeated their headline would be a list of six identical tips.
                assert_ne!(
                    *value, wording.whole,
                    "list {id}'s value {index} repeats the control's own tip"
                );
            }
            // `Scaling` has one more value than its table, because `Custom…` is the numeric box
            // beside it rather than a scaling of its own.
            if id == ID_SCALING {
                continue;
            }
            // A value picked from the box and a row picked from the list say the same thing.
            for index in 0..count {
                assert_eq!(combo(id, Some(index)), Some(wording.values[index]));
            }
            assert_eq!(combo(id, None), Some(wording.whole));
        }

        // A value the list has no entry for falls back to the control's own sentence rather than
        // answering nothing: a box showing a scaling typed by hand is still the scaling row.
        assert_eq!(
            combo(ID_SCALING, Some(usize::MAX)),
            editor(ID_SCALING).map(|wording| wording.whole)
        );
    }

    /// A tip is measured before it is drawn, and it is a card rather than a line: a sentence wraps
    /// inside the column, and a longer one is taller rather than clipped.
    #[test]
    fn a_tip_measures_its_run_and_never_exceeds_its_column() {
        let font = unsafe { tip_font(96) };
        assert!(!font.0.is_null(), "the tip's face was never made");
        let column = scale(COLUMN_LOGICAL, 96);
        let (short_w, short_h) = unsafe { measure(font, "Disabled", column) };
        let long = "The only road that never puts the window on the screen at all: the program is \
                    cloaked from inside before the compositor ever sees it, and the rectangle it \
                    asks for is refused.";
        let (long_w, long_h) = unsafe { measure(font, long, column) };
        assert!(short_w > 0 && short_h > 0);
        assert!(
            short_w <= column && long_w <= column,
            "a run measured {long_w}px against a {column}px column"
        );
        assert!(
            long_h > short_h,
            "a sentence of five lines came out {long_h}px tall against a word's {short_h}px"
        );
        unsafe {
            let _ = DeleteObject(HGDIOBJ(font.0));
        }
    }

    /// **The card is opaque where it is a card.** A layered window is what its pixels say, and a card
    /// blended over an empty surface came out at about half its coverage: the shadow stood and the
    /// words inside it did not. Every pixel the rounded box holds is at full alpha, and the corner the
    /// outline rounds keeps the anti-aliasing that makes it a curve.
    #[test]
    fn the_cards_own_pixels_are_opaque_and_its_corners_are_not() {
        let palette = Palette::of(crate::config::config::TextTheme::Dark);
        let mut tip = Tip {
            pixels: Vec::new(),
            width: 0,
            height: 0,
            text: String::new(),
            palette,
        };
        // `compose` paints the window it is handed; with none, it still composes the surface.
        unsafe { compose(HWND::default(), &mut tip, "Disabled", palette, 96) }
            .expect("the tip was never composed");

        let w = tip.width as usize;
        let alpha_at = |x: usize, y: usize| (tip.pixels[y * w + x] >> 24) & 0xFF;
        let shadow = scale(SHADOW_LOGICAL, 96) as usize;
        let middle = (shadow + 3, tip.height as usize / 2);
        assert_eq!(
            alpha_at(middle.0, middle.1),
            255,
            "the card's own pixels are not opaque"
        );
        // The very corner of the surface is outside the card entirely: nothing of it is drawn there.
        assert_eq!(alpha_at(0, 0), 0, "a tip leaked a pixel outside its shadow");
        // The corner of the card's box is rounded, so it is neither nothing nor solid.
        let corner = (shadow, shadow);
        let rounded = alpha_at(corner.0, corner.1);
        assert!(
            rounded < 255,
            "the card's corner is square rather than rounded: {rounded}"
        );
    }

    /// **Every placeholder the parser expands is named in the `Arguments` tip.** The row is the one
    /// place a user can read what may be written in that field, so a placeholder the app honours and
    /// the tip does not name is one nobody can discover — which is what happened while the tip still
    /// said only `{file}` and the parser had grown `{path}` and `{url}` beside it.
    #[test]
    fn the_arguments_tip_names_every_placeholder_the_parser_expands() {
        // The words this app expands, spelled as the parser matches them.
        const EXPANDS: &[&str] = &["{file}", "{path}", "{url}"];
        let said = editor(crate::ui::custom_preview_settings::ID_ARGUMENTS)
            .expect("the Arguments row has no tip")
            .whole;
        for placeholder in EXPANDS {
            assert!(
                said.contains(placeholder),
                "`{placeholder}` is expanded by the parser and never named in the Arguments tip: {said:?}"
            );
        }
    }

    /// **A tip is a real, visible window**: made, composed, placed and shown, with black bytes in the
    /// pixels where the words are. Everything in between has been wrong at least once, so the whole
    /// road is walked here rather than the parts.
    #[test]
    fn a_tip_is_made_composed_placed_and_shown() {
        let Some(owner) = crate::ui::custom_preview_settings::create_window() else {
            panic!("the owner window was never made");
        };
        let palette = Palette::of(crate::config::config::TextTheme::Dark);
        let anchor = RECT {
            left: 300,
            top: 300,
            right: 600,
            bottom: 330,
        };
        let shown = show_at(anchor, owner, "Disabled", palette);
        assert!(shown, "the tip was never shown");

        let tip = SHOWING
            .with(|showing| *showing.borrow())
            .expect("a tip was shown and not remembered");
        unsafe {
            assert!(
                windows::Win32::UI::WindowsAndMessaging::IsWindowVisible(tip).as_bool(),
                "the tip is not a visible window"
            );
            let mut box_ = RECT::default();
            let _ = GetWindowRect(tip, &mut box_);
            assert!(
                box_.right > box_.left && box_.bottom > box_.top,
                "the tip has no box"
            );
            // Its picture is on the desktop it was placed on, not at the origin it was made at.
            assert!(
                box_.top < anchor.top && box_.bottom <= anchor.top && box_.left >= 0,
                "the tip {box_:?} does not stand above its anchor {anchor:?}"
            );

            // And its own pixels are the ones composed: an opaque card, and ink inside it.
            let only = tip_of(tip).expect("the tip kept no surface");
            let w = only.width as usize;
            let alpha_at = |x: usize, y: usize| (only.pixels[y * w + x] >> 24) & 0xFF;
            let shadow = scale(SHADOW_LOGICAL, 96) as usize;
            assert_eq!(alpha_at(shadow + 3, only.height as usize / 2), 255);
            let darkened = only
                .pixels
                .iter()
                .filter(|pixel| ((**pixel >> 24) & 0xFF) == 255)
                .any(|pixel| ((pixel >> 16) & 0xFF) < 0x60 && ((pixel >> 8) & 0xFF) < 0x60);
            assert!(darkened, "the card is opaque and holds no words");
        }
        hide();
        unsafe {
            let _ = DestroyWindow(owner);
        }
    }

    /// Where a tip stands: above the control where the work area has room, below it where it has
    /// none, and inside the work area either way.
    #[test]
    fn a_tip_stands_above_its_control_and_never_leaves_the_work_area() {
        let work = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        // The placement is written against the monitor's own work area, so the test asks the
        // arithmetic rather than the desktop: a control with room above gets a tip above it.
        let noroom = RECT {
            left: 400,
            top: work.top + 2,
            right: 700,
            bottom: work.top + 32,
        };
        let below = noroom.bottom + scale(STANDOFF_LOGICAL, 96);
        assert!(
            below < work.bottom,
            "a tip below a control with no room above is inside the work area"
        );
        let room = RECT {
            left: 400,
            top: 500,
            right: 700,
            bottom: 530,
        };
        let above = room.top - scale(STANDOFF_LOGICAL, 96) - 40;
        assert!(
            above > work.top,
            "a tip above a control with room is inside the work area"
        );
    }
}

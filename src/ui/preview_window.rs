//! The preview window: the thread that shows a file, and everything a pin is.
//!
//! It is one window and one message loop. Everything here exists to answer a question that loop
//! asks, and the parts are split by what each one is asked for:
//!
//! - `requests` - what the rest of the crate asks of this window, and the notices a render sends
//!   back when it has an answer.
//! - `tick` - how long the loop waits, and what it waits on.
//! - `media_types` - the shapes a preview arrives in: a message, a kind of file, a frame, a
//!   subtitle stream.
//! - `media_data` - one media, and the geometry cache the films were probed into.
//! - `hover_facts` - what the pointer is over, and the measurements a hover has had taken.
//! - `preview_scale` - the scale a preview is drawn at.
//! - `dimensions` - how big each kind of file wants to be, and the room it has to fit into.
//! - `pixels` - decoded pixels turned into the pixels a window is painted with.
//! - `image_cache` - decoded pictures kept by the version they came from.
//! - `animated` - the animated decoders, played as a stream of frames.
//! - `load` - one loader per kind of file.
//! - `load_dispatch` - how a load is asked for and dispatched, and the answer it comes back with.
//! - `engine_render` - the renderers a file this app cannot draw itself is asked at.
//! - `loading_paint` - what is drawn while a file is still being read.
//! - `surfaces` - the layered surfaces a preview is painted onto.
//! - `paint` - how a media is composed into a window's bands.
//! - `pointer` - what the pointer's own questions are answered from.
//! - `text_selection` - a selection made in a text preview.
//! - `message_utils` - the small helpers this window and the Explorer hook share.
//! - `window_proc` - the window procedure, and the keys it answers.
//! - `layout` - where a preview is put.
//! - `pin_model` - what a pin is.
//! - `pin_playback` - what the pin plays.
//! - `pin_chrome_layout` - the pin's chrome, and the level it is at.
//! - `pin_geometry` - the box of a pinned window.
//! - `pin_lifecycle` - a pin's own lifetime, and the window it stands in.
//! - `pin_swap` - what a pin should show next, and what is waited for.
//! - `pin_install` - putting a loaded media into a pin.
//! - `pin_walk` - the walk a pin takes through a folder.
//! - `pin_bubble` - the bubble that stands beside a collapsed pin.
//! - `pin_input` - a hand on a pinned window.
//! - `pin_drag` - carrying a pinned window with the pointer.
//! - `video_probe` - what a film is asked of FFprobe.
//! - `ffplay_window` - the window and the process FFmpeg draws a film into.
//! - `video_playback` - starting and stopping a film.
//! - `video_retire` - handing one film over to the next.
//! - `video_hw` - what a film is decoded on, and how it loops.
//! - `audio_engine` - a sound, its loudness, its player and its clock.
//! - `event_loop` - `run_preview_window` itself.
//!
//! The parts share one namespace through this module - each reads the others as one set of items
//! rather than naming them through their own - so that moving one question across a boundary is
//! not a change to it. `event_loop` is the exception and is the reason the loop reads the rest
//! through the same namespace rather than through this file: it is the only part that drives the
//! others, and splitting it is not a move that can be made without deciding something.

mod animated;
mod audio_engine;
mod dimensions;
mod displays;
mod engine_render;
mod event_loop;
mod ffplay_window;
mod hover_facts;
mod image_cache;
mod layout;
mod load;
mod load_dispatch;
mod loading_paint;
mod media_data;
mod media_types;
mod message_utils;
mod paint;
mod pin_bubble;
mod pin_chrome_layout;
mod pin_drag;
mod pin_geometry;
mod pin_input;
mod pin_install;
mod pin_lifecycle;
mod pin_model;
mod pin_playback;
mod pin_swap;
mod pin_walk;
mod pin_window;
mod pixels;
mod pointer;
mod preview_scale;
mod requests;
mod surfaces;
mod text_selection;
mod tick;
mod video_hw;
mod video_launch;
mod video_playback;
mod video_probe;
mod video_retire;
mod window_proc;

use crate::app::engine_processes;
use crate::config::config::{
    frame_bytes_within_budget, image_decode_limits, read_within_budget, sanitize_image_cache_mb,
    sanitize_spinner_delay_ms, sanitize_webp_playback_fps, AppConfig, AudioSeek, MarkdownMode,
    OfficeEngine, PreviewScale, PreviewType, TextTheme, TransparentBackground,
    DEFAULT_ANIMATED_SCALE_PERCENT, DEFAULT_AUDIO_SEEK, DEFAULT_DDS_BACKGROUND,
    DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE, DEFAULT_DOCUMENT_SCALE, DEFAULT_EBOOK_SCALE,
    DEFAULT_FONT_BACKGROUND, DEFAULT_FONT_SCALE, DEFAULT_HTML_BACKGROUND, DEFAULT_IMAGE_BACKGROUND,
    DEFAULT_IMAGE_CACHE_MB, DEFAULT_NORMALIZE_VIDEO_VOLUME, DEFAULT_NORMALIZE_VOLUME,
    DEFAULT_PIN_PAUSE_AUDIO, DEFAULT_PIN_PAUSE_VIDEO, DEFAULT_PIN_UPDATE_ENABLED,
    DEFAULT_PREVIEW_SCALE_PERCENT, DEFAULT_SPINNER_DELAY_MS, DEFAULT_TEXT_FONT_SCALE_PERCENT,
    DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS, DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE,
    DEFAULT_VIDEO_HW_ACCEL, DEFAULT_VIDEO_SCALE_PERCENT, DEFAULT_WEBP_PLAYBACK_FPS,
};
use crate::engines::calibre_render;
use crate::engines::imagemagick_render;
use crate::engines::libreoffice_render;
use crate::engines::office_render;
use crate::engines::peazip_render;
use crate::engines::webview_preview;
use crate::formats::audio_formats;
use crate::formats::calibre_formats;
use crate::formats::codecs;
use crate::formats::libre_formats;
use crate::formats::magick_formats;
use crate::formats::native_formats;
use crate::formats::office_formats;
use crate::formats::peazip_formats;
use crate::formats::video_formats;
use crate::paths::plain_path;
use crate::readers::audio_seek;
use crate::readers::audio_track::{self, Player, Probed};
use crate::readers::comic_preview;
use crate::readers::dds_image;
use crate::readers::eps_image;
use crate::readers::font_preview;
use crate::readers::heif_sequence;
use crate::readers::jxl_image;
use crate::readers::metafile_image;
use crate::readers::office_preview;
use crate::readers::pdf_preview;
use crate::readers::project_image;
use crate::readers::psd_image;
use crate::readers::svg_preview;
use crate::readers::tone_map;
use crate::readers::video_player;
use crate::readers::webp_image;
use crate::readers::wic_image;
use crate::shell::pin_navigation;
use crate::shell::wheel_input;
use crate::text::archive_preview::{self, ArchivePreviewOptions};
use crate::text::audio_preview::{self, AudioPreviewOptions, Card, CardChrome, CardControl};
use crate::text::pin_chrome;
use crate::text::text_paint::DibSurface;
use crate::text::text_preview::{self, TextPreviewOptions};
use crate::{CONFIG, RUNNING};
use gif::DecodeOptions;
use image::{AnimationDecoder, GenericImageView};
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::{HashMap, VecDeque};
use std::fs::File;
use std::io::BufReader;
use std::os::windows::ffi::OsStrExt;
use std::os::windows::io::AsRawHandle;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Output, Stdio};
use std::ptr;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc::{channel, Receiver, Sender, TryRecvError};
use std::sync::{Arc, Condvar, Mutex, MutexGuard};
use std::time::{Duration, Instant, SystemTime};
use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{
    CloseHandle, GlobalFree, COLORREF, HANDLE, HWND, LPARAM, LRESULT, POINT, RECT, SIZE,
    WAIT_TIMEOUT, WPARAM,
};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, ClientToScreen, CreateCompatibleDC, CreateDIBSection, DeleteDC, DeleteObject,
    EndPaint, GdiFlush, SelectObject, SetBrushOrgEx, SetStretchBltMode, StretchBlt, AC_SRC_ALPHA,
    AC_SRC_OVER, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, BLENDFUNCTION, DIB_RGB_COLORS, HALFTONE,
    HBITMAP, HDC, HGDIOBJ, PAINTSTRUCT, SRCCOPY,
};
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE};
use windows::Win32::System::Ole::CF_UNICODETEXT;
use windows::Win32::System::Threading::{
    CreateEventW, OpenProcess, QueryFullProcessImageNameW, ResetEvent, SetEvent, TerminateProcess,
    WaitForSingleObject, PROCESS_NAME_WIN32, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, GetCapture, GetFocus, ReleaseCapture, SetCapture, SetFocus, VK_A, VK_C,
    VK_CONTROL, VK_DOWN, VK_ESCAPE, VK_LEFT, VK_RIGHT, VK_SPACE, VK_UP,
};
use windows::Win32::UI::Shell::{
    AssocQueryStringW, ShellExecuteW, ASSOCF_NONE, ASSOCSTR, ASSOCSTR_FRIENDLYAPPNAME,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DispatchMessageW,
    EnumWindows, GetCursorPos, GetForegroundWindow, GetWindow, GetWindowLongPtrW, GetWindowRect,
    GetWindowThreadProcessId, IsWindow, IsWindowVisible, LoadCursorW, MoveWindow,
    MsgWaitForMultipleObjectsEx, PeekMessageW, PostMessageW, RegisterClassExW, SetCursor,
    SetForegroundWindow, SetWindowLongPtrW, SetWindowPos, ShowWindow, ShowWindowAsync,
    TrackPopupMenu, TranslateMessage, UpdateLayeredWindow, CS_HREDRAW, CS_VREDRAW, GWL_EXSTYLE,
    GW_OWNER, HWND_TOPMOST, IDC_ARROW, IDC_SIZEALL, IDC_SIZENESW, IDC_SIZENS, IDC_SIZENWSE,
    IDC_SIZEWE, MF_STRING, MSG, MWMO_INPUTAVAILABLE, PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND,
    PBT_APMSTANDBY, PBT_APMSUSPEND, PM_NOREMOVE, PM_REMOVE, QS_ALLINPUT, SWP_FRAMECHANGED,
    SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SWP_SHOWWINDOW, SW_HIDE,
    SW_SHOWNOACTIVATE, SW_SHOWNORMAL, TPM_LEFTALIGN, TPM_NONOTIFY, TPM_RETURNCMD, TPM_TOPALIGN,
    ULW_ALPHA, WA_INACTIVE, WM_ACTIVATE, WM_CHAR, WM_CLOSE, WM_DISPLAYCHANGE, WM_DPICHANGED,
    WM_KEYDOWN, WM_KEYUP, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_POWERBROADCAST,
    WM_RBUTTONUP, WM_SETCURSOR, WM_SYSCOMMAND, WM_SYSKEYDOWN, WNDCLASSEXW, WS_EX_LAYERED,
    WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

#[cfg(test)]
use displays::RecordedDisplays;
use displays::{dpi_at, work_area_at, Displays, DESKTOPS};

use pin_window::{
    ask_pin, end_pin, give_the_keyboard_back, install, pin_state, release_keyboard, take_keyboard,
    take_pin_command, PinHide, PinWindow, Reason, WM_PIN_RELEASE_POINTER,
};
#[cfg(test)]
use pin_window::{
    pin_holds_a_keyboard, stand_pin, take_pin_for_a_test, PinWindowCall, RecordedPinWindow,
};

// What the rest of the crate asks for by name.
pub(crate) use dimensions::monitor_dpi_from_point;
pub(crate) use event_loop::run_preview_window;
pub(crate) use ffplay_window::kill_stray_video_process;
pub(crate) use hover_facts::cursor_preview_hover;
pub(crate) use hover_facts::PreviewCursorHover;
pub(crate) use image_cache::trim_image_cache;
pub(crate) use media_types::PreviewMessage;
pub(crate) use pin_drag::publish_pin_media_press;
pub(crate) use pin_lifecycle::PinCommand;
pub(crate) use pin_model::pin_is_focused;
pub(crate) use pin_walk::PinPlanned;
pub(crate) use pin_walk::PinStep;
pub(crate) use pixels::rgba_to_bgra;
pub(crate) use pointer::note_engine_page_drag;
pub(crate) use pointer::pointer_item_box;
pub(crate) use pointer::pointer_item_holds;
pub(crate) use pointer::preview_pointer_hold;
pub(crate) use pointer::publish_pointer_item_box;
// The one of these nothing asks for, kept reachable rather than dropped: a move is not a change.
#[allow(unused_imports)]
pub(crate) use pointer::text_preview_scrollable;
pub(crate) use pointer::text_scroll_keep_alive_try;
pub(crate) use pointer::AvoidRegion;
pub(crate) use requests::hide_preview;
pub(crate) use requests::is_preview_window;
pub(crate) use requests::notify_magick_ready;
pub(crate) use requests::notify_office_render;
pub(crate) use requests::notify_peazip_ready;
pub(crate) use requests::pinned;
pub(crate) use requests::pinned_path;
pub(crate) use requests::preview_screen_rect;
pub(crate) use requests::refresh_pin;
pub(crate) use requests::refresh_preview;
pub(crate) use requests::refresh_preview_types;
pub(crate) use requests::refresh_render_html;
pub(crate) use requests::request_pin_end;
pub(crate) use requests::show_preview;
pub(crate) use requests::show_preview_keyboard;
pub(crate) use requests::take_pin_resumed;
pub(crate) use requests::update_pinned_preview;
pub(crate) use tick::preview_stall_ms;
pub(crate) use tick::PREVIEW_SENDER;
pub(crate) use video_hw::forget_video_hw_accel_answer;

// The parts, in one namespace (see the note above).
use animated::*;
use audio_engine::*;
use dimensions::*;
use engine_render::*;
use ffplay_window::*;
use hover_facts::*;
use image_cache::*;
use layout::*;
use load::*;
use load_dispatch::*;
use loading_paint::*;
use media_data::*;
use media_types::*;
use message_utils::*;
use paint::*;
use pin_bubble::*;
use pin_chrome_layout::*;
use pin_drag::*;
use pin_geometry::*;
use pin_input::*;
use pin_install::*;
use pin_lifecycle::*;
use pin_model::*;
use pin_playback::*;
use pin_swap::*;
use pin_walk::*;
use pixels::*;
use pointer::*;
use preview_scale::*;
use requests::*;
use surfaces::*;
use text_selection::*;
use tick::*;
use video_hw::*;
use video_playback::*;
use video_probe::*;
use video_retire::*;
use window_proc::*;

#[cfg(test)]
mod tests;

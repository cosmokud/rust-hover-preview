use super::*;
use crate::config::config::{PinNavFileTypes, DEFAULT_FONT_SCALE_PERCENT};
// The key a pin answers with nothing, which is the one the mapping has to name explicitly
// and which nothing outside a test ever has to read.
use windows::Win32::UI::Input::KeyboardAndMouse::VK_SPACE;
// What the pointer probe speaks with: a real `SendInput`, because a drag is a sequence of
// events delivered to whatever window is under the pointer and nothing less is one.
use windows::Win32::UI::Input::KeyboardAndMouse::{
    SendInput, INPUT, INPUT_0, INPUT_TYPE, MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_LEFTDOWN,
    MOUSEEVENTF_LEFTUP, MOUSEEVENTF_MOVE, MOUSEINPUT, MOUSE_EVENT_FLAGS,
};
// What a window of the suite's own is made of and made unseen with: see `a_window_of_the_suite`.
use windows::Win32::Foundation::HINSTANCE;
use windows::Win32::UI::WindowsAndMessaging::{
    SetLayeredWindowAttributes, LWA_ALPHA, WINDOW_EX_STYLE, WINDOW_STYLE,
};

/// A display to place on: 1000 by 800 at its top-left corner.
fn bounds() -> ScreenBounds {
    ScreenBounds {
        left: 0,
        top: 0,
        right: 1000,
        bottom: 800,
    }
}

/// A window standing at a box, with the pointer somewhere on the desktop, for the tests
/// about what a press becomes.
///
/// The box is the screen's own and not the pin's, because that is what `begin_pin_drag` asks
/// for: a window the hand has carried since it was maximized is no longer standing where the
/// maximize left it, and a drag begun from the pin's remembered box begins from the screen's
/// own top border rather than from where the hand left the window.
fn a_window_at(box_: ScreenRegion) -> RecordedPinWindow {
    RecordedPinWindow::with(0x1000, Some((box_.0, box_.1)), Some(box_))
}

/// The window a restore is handed in a test, which records what is asked of it instead of asking
/// it: what a press on a stand-in owes the window it brings back — the front and the caret — is not
/// something a run of the suite may take from the user's desktop.
fn a_restored_window() -> RecordedPinWindow {
    RecordedPinWindow::new(PREVIEW_HWND.load(Ordering::SeqCst))
}

/// A lock the tests that publish the pointer's own state take, so that one of them
/// runs at a time: the item box and the hold regions are one set for the whole
/// process, so two of these tests at once is one test's box answering another test's
/// question.
static POINTER_STAND_IN: Mutex<()> = Mutex::new(());

fn layout(pos_x: i32, pos_y: i32, width: u32, height: u32) -> PreviewLayout {
    PreviewLayout {
        pos_x,
        pos_y,
        max_width: width,
        max_height: height,
        preview_w: width,
        preview_h: height,
    }
}

/// The display the placement figures are worked out for: 100%, which is what
/// distances written in logical pixels are the same as the pixels of.
const TEST_DPI: u32 = 96;

/// A placement kept off `name`, at the size the media's own scale allows — the
/// arrangement the figures are easy to read in. It is the step a hover makes, so
/// the gap is the pointer's standoff at this display's scale, and `cursor` is where
/// the pointer is for a placement that is a mouse hover's — which the ways out are
/// held clear of, where it is given (see `avoiding_text`).
fn placed(
    placement: PreviewLayout,
    media: (u32, u32),
    name: (i32, i32, i32, i32),
    cursor: Option<(i32, i32)>,
    bounds: ScreenBounds,
) -> PreviewLayout {
    placed_off(placement, media, AvoidRegion::text(name), cursor, bounds)
}

/// The same placement kept off a region of a kind of its own — the text one item
/// draws, or a column the view draws every row's in — which is what decides the
/// ways out of it; see `placed` for the rest.
fn placed_off(
    placement: PreviewLayout,
    media: (u32, u32),
    region: AvoidRegion,
    cursor: Option<(i32, i32)>,
    bounds: ScreenBounds,
) -> PreviewLayout {
    avoiding_text(
        placement,
        media,
        PreviewScale::Percent(100),
        Clearance {
            text: Some(region),
            gap: logical_px(TEST_DPI, POINTER_STANDOFF_PIXELS),
            cursor,
        },
        bounds,
        TEST_DPI,
    )
}

/// Every scale a hover is laid out by, at the shares the app starts at. A test that
/// is about one of them names that one and leaves the rest where the app has them.
fn hover_scales() -> HoverScales {
    HoverScales {
        picture: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
        video: PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT),
        animated: PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT),
        ebook: DEFAULT_EBOOK_SCALE,
        document: DEFAULT_DOCUMENT_SCALE,
        font: DEFAULT_FONT_SCALE,
        design: DEFAULT_DESIGN_SCALE,
        vector: DEFAULT_VECTOR_SCALE,
    }
}

/// A little PNG, with an `acTL` chunk ahead of its image data when `animated`: the
/// chunk is the whole of what the probe reads, and the pixel chunks after it are a
/// real still frame so that the file is a PNG either way.
fn write_test_png(path: &Path, animated: bool) {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]);

    // The header: one pixel, eight bits, colour type six.
    push_png_chunk(
        &mut bytes,
        *b"IHDR",
        &[0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0],
    );

    if animated {
        push_png_chunk(&mut bytes, *b"acTL", &[0, 0, 0, 2, 0, 0, 0, 0]);
        push_png_chunk(
            &mut bytes,
            *b"fcTL",
            &[
                0, 0, 0, 0, // the first frame's sequence number
                0, 0, 0, 1, 0, 0, 0, 1, // its width and height
                0, 0, 0, 0, 0, 0, // where it sits
                0, 0, 0, 1, 0, 0, 0, 1, // the frame's own delay
                0, 0, // how it replaces what it is drawn over
            ],
        );
    }

    push_png_chunk(
        &mut bytes,
        *b"IDAT",
        &[0x78, 0x01, 0x63, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01],
    );
    push_png_chunk(&mut bytes, *b"IEND", &[]);

    std::fs::write(path, bytes).expect("a written PNG");
}

/// One PNG chunk: its body's length, its type, its body, and the CRC of the type and
/// the body together, which is the whole of the container.
fn push_png_chunk(bytes: &mut Vec<u8>, kind: [u8; 4], body: &[u8]) {
    bytes.extend_from_slice(&(body.len() as u32).to_be_bytes());
    bytes.extend_from_slice(&kind);
    bytes.extend_from_slice(body);
    bytes.extend_from_slice(&crc32(&[&kind, body].concat()).to_be_bytes());
}

/// The CRC every PNG chunk is checked with: a table-less, bit-at-a-time take on the
/// standard polynomial, which is all a test fixture needs.
fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for byte in bytes {
        crc ^= *byte as u32;
        for _ in 0..8 {
            crc = if crc & 1 != 0 {
                (crc >> 1) ^ 0xEDB8_8320
            } else {
                crc >> 1
            };
        }
    }
    !crc
}

/// A pin standing a sound's card: the three facts its controls are gated on, read from the pin's
/// own fields rather than from the media's kind, which is why a test of them needs no media at
/// all (see `pin_shows_an_audio_card`).
fn sound_pin() -> PinnedPreview {
    let mut pin = PinnedPreview {
        overlay: false,
        hides_chrome: false,
        caption: pinned_caption_height(96, Some(MediaType::Audio)),
        content: (100, 100, 500, 300),
        ..PinnedPreview::for_test()
    };
    pin.chrome = PinChrome::always();
    pin.transport_bar = false;
    pin.frame = PinFrame::None;

    pin
}

/// The two `Pin Mode → Pause on Minimize` switches and a sound's own `Normalize`, stood where a test
/// asks and put back when the guard is dropped — the house pattern (see `CardFontSettings`).
///
/// The two switches stand *on*, which is the state a bubble is about: a test of what a press on a
/// collapsed pin's bubble does has to be able to say which of the two gates the media behind the
/// bubble answers to, and a machine's own settings are not that. `Normalize` is stood off beside
/// them because a measured gain would put the sound in FFmpeg's hands whatever the probe said, and
/// the player a test begins for one is begun off the machine's own programs when it is not.
struct PausePreviewSettings {
    video: bool,
    audio: bool,
    normalize: bool,
}

impl PausePreviewSettings {
    /// Stands the two park switches where a test wants them, holding what the machine's own
    /// settings were to put back.
    fn stood_at(video: bool, audio: bool) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = PausePreviewSettings {
            video: config.pin_pause_video,
            audio: config.pin_pause_audio,
            normalize: config.normalize_volume,
        };
        config.pin_pause_video = video;
        config.pin_pause_audio = audio;
        // The gain-scan arm is what puts a file of any probe into FFmpeg's hands, so it is stood
        // off rather than left to the machine (see `playing_player`).
        config.normalize_volume = false;

        was
    }
}

impl Drop for PausePreviewSettings {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.pin_pause_video = self.video;
            config.pin_pause_audio = self.audio;
            config.normalize_volume = self.normalize;
        }
    }
}

/// Where a pin the minimize button puts away goes, stood where a test asks and put back when the
/// guard is dropped — the house pattern (see `PausePreviewSettings`). The two picks are one group,
/// so the guard states the whole of what a menu's disc is read from (see `PinMinimizeTo`).
struct PinMinimizeToPutBack(PinMinimizeTo);

impl PinMinimizeToPutBack {
    fn put(target: PinMinimizeTo) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = config.pin_minimize_to;
        config.pin_minimize_to = target;

        Self(was)
    }
}

impl Drop for PinMinimizeToPutBack {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.pin_minimize_to = self.0;
        }
    }
}

/// The two switches a double-click on a pinned window answers to, stood where a test asks and put
/// back when the guard is dropped — the house pattern (see `PausePreviewSettings`).
///
/// The two are stood together because the second is only read where the first is on: a test of what
/// a pair does to a window's *size* is a test that has already said the pair is a centring request
/// at all.
struct DoubleClickSettings {
    center: bool,
    reset: bool,
}

impl DoubleClickSettings {
    fn stood_at(center: bool, reset: bool) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = DoubleClickSettings {
            center: config.pin_double_click_center,
            reset: config.pin_double_click_reset_size,
        };
        config.pin_double_click_center = center;
        config.pin_double_click_reset_size = reset;

        was
    }
}

impl Drop for DoubleClickSettings {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.pin_double_click_center = self.center;
            config.pin_double_click_reset_size = self.reset;
        }
    }
}

/// What a picture is drawn at, stood where a test asks and put back when the guard is dropped —
/// the way the two double-click switches are (`DoubleClickSettings`).
///
/// A share is the one thing the size a test asks for is read against rather than stated as: a test
/// that left its own share behind would be measuring every test after it.
struct PictureScale {
    was: PreviewScale,
}

impl PictureScale {
    fn stood_at(scale: PreviewScale) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = std::mem::replace(&mut config.preview_scale, scale);

        Self { was }
    }
}

impl Drop for PictureScale {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.preview_scale = self.was;
        }
    }
}

/// A pin of the kind the chrome's own questions are about: a picture, whose window is its
/// media and whose chrome is drawn over it.
fn overlay_pin(content: ScreenRegion, chrome: PinChrome) -> PinnedPreview {
    PinnedPreview {
        bound: Some((content.2 - content.0).max(content.3 - content.1).max(1)),
        content,
        chrome,
        ..PinnedPreview::for_test()
    }
}

/// The same picture with the wheel's own view of it: one drawn `scale` times its window, centred and
/// not carried — which is what a pin the wheel has taken inside its window is, and the state the two
/// halves of a press's own test differ by (see `PinZoom`).
fn with_pin_zoom(pin: PinnedPreview, scale: f32) -> PinnedPreview {
    let (width, height) = region_size(pin.content);

    PinnedPreview {
        zoom: Some(PinZoom::still(
            (width as f32 * scale, height as f32 * scale),
            (0.0, 0.0),
            centred_on(pin.content),
        )),
        ..pin
    }
}

fn edge(left: bool, top: bool, right: bool, bottom: bool) -> PinResize {
    PinResize {
        left,
        top,
        right,
        bottom,
    }
}

/// The room every resize test drags inside: a display of 1200 by 900, with a caption taken off
/// the top and a transport bar off the bottom, which is what the room of one really is.
fn drag_room() -> ScreenBounds {
    ScreenBounds {
        left: 0,
        top: 0,
        right: 1200,
        bottom: 900,
    }
}

/// The media box a drag of one edge or corner comes out with, from a media box of 400 by 300
/// standing in the middle of that room: what every test below states its expectations in.
fn dragged(edge: PinResize, dx: i32, dy: i32) -> ScreenRegion {
    dragged_from((400, 300, 800, 600), PinFrame::Shaped, edge, dx, dy)
}

fn dragged_from(
    content: ScreenRegion,
    frame: PinFrame,
    edge: PinResize,
    dx: i32,
    dy: i32,
) -> ScreenRegion {
    dragged_overlay(content, frame, edge, dx, dy, false)
}

/// The same drag for a pin whose chrome is drawn over its media: what the box that comes out of
/// it is a different question for, since the window and the media are the same box there (see
/// `PinSpace`).
fn dragged_overlay(
    content: ScreenRegion,
    frame: PinFrame,
    edge: PinResize,
    dx: i32,
    dy: i32,
    overlay: bool,
) -> ScreenRegion {
    let window = resize_pinned_content(
        PinSpace {
            content,
            room: drag_room(),
            overlay,
            transport: false,
            // A caption above the media of every kind but a sound, which is given none: this is
            // a kind that is framed or laid out, and a sound is neither (see
            // `pinned_caption_height`).
            caption: pinned_caption_height(96, None),
        },
        edge,
        dx,
        dy,
        96,
        frame,
    );

    content_box_of(window, 96, false, overlay, pinned_caption_height(96, None))
}

/// A window of the suite's own: made, shown, and invisible.
///
/// A test that asks the desktop about a window — the order it stands in, the region it is erased
/// within, whether it is on the screen at all — needs a window that *is* on the desktop: one at
/// `NOWHERE` has no visible region to be asked about, and one that was never shown covers nothing.
/// What it must not be is a window the user sees, or one that takes their keyboard and with it the
/// caret out of whatever they are typing in.
///
/// So it is layered at no alpha: `WS_EX_LAYERED` with a whole-window alpha of 0 is a window the
/// compositor draws as nothing, and nothing else about it is any different — `IsWindowVisible`
/// answers yes, it stands where it was put in the z-order, `WindowFromPoint` finds it over what is
/// under it, and its visible region is the real one it is erased and clipped within, which is what
/// the tests about the host's fill ask of it. `WS_EX_NOACTIVATE` is what keeps the caret where the
/// user left it: nothing shown here is ever activated. `WS_EX_TOOLWINDOW` is what keeps a run of the
/// suite from leaving a taskbar button and an Alt+Tab entry behind it.
///
/// **`WS_VISIBLE` at create time is the one thing that must not happen**, because a window created
/// visible is activated wherever it stands — that, and not the box it is put at, is what takes the
/// keyboard out of the user's terminal. The window is made hidden, made unseen, shown unactivated.
///
/// `extra` is the bits a caller's own window needs on top of those three (a stand-in born topmost, a
/// card's menu against), and they are added to them rather than replacing any. A class of a test's
/// own is registered with the module's instance; `a_window_of_the_suite_of` is the same window in an
/// instance a caller registered its class with.
///
/// `WS_EX_LAYERED` on a `WS_CHILD` window is legal here because the executable's manifest declares
/// the Windows this app supports, which is what turns the pre-Win8 refusal of the pair off (see
/// `build.rs`, and `pin_film_host`, which relies on the same declaration for the same pair).
pub(super) unsafe fn a_window_of_the_suite(
    extra: WINDOW_EX_STYLE,
    class: PCWSTR,
    style: WINDOW_STYLE,
    box_: ScreenRegion,
) -> HWND {
    // SAFETY: the class is one this process registered, and the window is the caller's to destroy.
    unsafe {
        a_window_of_the_suite_of(
            GetModuleHandleW(None)
                .expect("the module's own handle")
                .into(),
            extra,
            class,
            style,
            box_,
        )
    }
}

/// The same window of the suite's own, in the instance a caller registered its class with: the erase
/// probe registers a class with `lpfnWndProc` of its own, and a window of it has to be made in the
/// instance that class was registered in.
pub(super) unsafe fn a_window_of_the_suite_of(
    hinstance: HINSTANCE,
    extra: WINDOW_EX_STYLE,
    class: PCWSTR,
    style: WINDOW_STYLE,
    box_: ScreenRegion,
) -> HWND {
    let hwnd = unsafe { a_window_of_the_suite_kept_back_in(hinstance, extra, class, style, box_) };
    // SAFETY: the window is the suite's own, and showing it unactivated is what puts it on the
    // desktop without taking the caret out of whatever the user is typing in.
    unsafe { show_a_window_of_the_suite(hwnd) };

    hwnd
}

/// The same window before it is shown: what a test keeps back for a road that shows it itself — the
/// tick that finds a player drawing, the place a park's window is put back at — so that whenever it
/// does appear it is already a window nothing of it is drawn of and nothing of the user's is
/// activated.
pub(super) unsafe fn a_window_of_the_suite_kept_back(
    extra: WINDOW_EX_STYLE,
    class: PCWSTR,
    style: WINDOW_STYLE,
    box_: ScreenRegion,
) -> HWND {
    // SAFETY: the class is one this process registered, and the window is the caller's to destroy.
    unsafe {
        a_window_of_the_suite_kept_back_in(
            GetModuleHandleW(None)
                .expect("the module's own handle")
                .into(),
            extra,
            class,
            style,
            box_,
        )
    }
}

/// The three of them: made hidden, made unseen, and left for the caller to show. `WS_POPUP` is added
/// to the style because every window of the suite's own is a window of the desktop's rather than a
/// child of another — a child cannot activate, and a child of a window here is a window a test makes
/// itself, with `WS_CHILD` named where it does.
unsafe fn a_window_of_the_suite_kept_back_in(
    hinstance: HINSTANCE,
    extra: WINDOW_EX_STYLE,
    class: PCWSTR,
    style: WINDOW_STYLE,
    box_: ScreenRegion,
) -> HWND {
    // SAFETY: the class is one this process registered, in the instance handed in, and the window is
    // made at a box inside the desktop the caller named.
    let hwnd = unsafe {
        CreateWindowExW(
            extra | WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW,
            class,
            w!(""),
            style | WS_POPUP,
            box_.0,
            box_.1,
            (box_.2 - box_.0).max(1),
            (box_.3 - box_.1).max(1),
            None,
            None,
            hinstance,
            None,
        )
    }
    .expect("a window of the suite's own");
    // SAFETY: the window is this call's own, and the alpha is what makes it drawn as nothing.
    unsafe { keep_a_window_of_the_suite_unseen(hwnd) };

    hwnd
}

/// Put the bits and the alpha on a window that is on the desktop and not of the user's: the layered
/// style, no activation, no taskbar button, and a whole-window alpha of 0 — see
/// `a_window_of_the_suite` for why it is this rather than a box off every display.
///
/// Additive, because the window may be one of the app's own with bits of its own to keep (a hover's
/// host wears four): what is asked of it here is that it be composited as nothing, not that it be
/// made a window of another kind.
pub(super) unsafe fn keep_a_window_of_the_suite_unseen(hwnd: HWND) {
    let ex = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    let unseen = WS_EX_LAYERED.0 as isize | WS_EX_NOACTIVATE.0 as isize | WS_EX_TOOLWINDOW.0 as isize;
    if ex & WS_EX_LAYERED.0 as isize == 0 {
        let _ = unsafe { SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ex | unseen) };
    }
    // SAFETY: the window is layered by the line above or already was, which is the whole of what
    // this call asks of it; an alpha of 0 with `LWA_ALPHA` is the window composited as nothing.
    unsafe { SetLayeredWindowAttributes(hwnd, COLORREF(0), 0, LWA_ALPHA) }
        .expect("a window of the suite's own, at no alpha");
}

/// Show a window of the suite's own: unactivated, which is the other half of a window nothing is
/// taken from (see `a_window_of_the_suite`).
pub(super) unsafe fn show_a_window_of_the_suite(hwnd: HWND) {
    let _ = unsafe { ShowWindow(hwnd, SW_SHOWNOACTIVATE) };
}

mod audio_bands;
mod cache;
mod engine_probes;
mod hover;
mod hover_film_host_window;
mod layout;
mod media_scales;
mod resident;
mod pin_audio_facts;
mod pin_audio_pause;
mod pin_audio_seek;
mod pin_bubble_art;
mod pin_bubble_space;
mod pin_cards;
mod pin_edges;
mod pin_film_host_window;
mod pin_frames;
mod pin_gesture_no_restart;
mod pin_input;
mod pin_keys;
mod pin_menu;
mod pin_player_failure;
mod pin_resize;
mod pin_subtitle_reload;
mod pin_swap_ffplay_jigsaw;
mod pin_swap_ffplay_native;
mod pin_taskbar_window;
mod pin_volume;
mod pin_window_kind;
mod pin_zoom;
mod placement;
mod probes;
mod text_selection;
mod video_engine;
mod video_hold;
mod video_transport;
mod walk;
mod ws_g_seek_hold;
mod ws_h_supersede;
mod ws_k_chrome_nav;
mod ws_l_take_up;
mod ws_m_frame_truth;

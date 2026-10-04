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

mod cache;
mod engine_probes;
mod hover;
mod layout;
mod media_scales;
mod pin_cards;
mod pin_edges;
mod pin_frames;
mod pin_input;
mod pin_keys;
mod pin_resize;
mod pin_volume;
mod placement;
mod probes;
mod video_engine;
mod video_hold;
mod video_transport;
mod walk;
mod ws_g_seek_hold;
mod ws_h_supersede;
mod ws_j_always_kill;
mod ws_k_chrome_nav;
mod ws_l_cover_owner;

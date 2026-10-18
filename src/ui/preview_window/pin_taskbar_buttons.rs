//! The buttons in a taskbar button's hover: Previous, Play/Pause and Next.
//!
//! A pinned video or sound is a window that can be behind another program, or put away into the
//! taskbar, and the hover over its button is the one place its transport is reachable without
//! bringing the window forward. What is offered there is therefore the *pin's* own commands rather
//! than anything of this module's: a press is queued for the preview loop exactly as a press on the
//! caption or a key is, so a walk reads the same Navigation settings, the same shuffle and the same
//! end of file, and the play/pause is the same hold the card, the bar and a Space answer (see
//! `ask_pin` and `pin_command_request`).
//!
//! The Shell draws them: `ITaskbarList3`'s thumbnail toolbar is a toolbar control in the hover
//! beside the thumbnail, and the Shell keeps it for the life of the window it was added to. The
//! stand-in's window is created once and never destroyed (see `pin_taskbar`), which is what makes
//! one `ThumbBarAddButtons` the whole of the setup here: the marks go in with it, and every later
//! change of state is a `ThumbBarUpdateButtons` — a button cannot be added or removed afterwards,
//! only shown or hidden.
//!
//! What the buttons say is asked on the preview loop's own tick, of the state the chrome is already
//! drawn from, and only a change is sent to the Shell. Nothing is sent at all while the stand-in is
//! not up: a window that is hidden has no taskbar button, and a toolbar cannot be hung on one.

use super::*;

use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Gdi::HBITMAP;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::System::Registry::{
    RegCloseKey, RegOpenKeyExW, RegQueryValueExW, HKEY, HKEY_CURRENT_USER, KEY_READ,
};
use windows::Win32::UI::Controls::{
    ImageList_Add, ImageList_Create, ImageList_Destroy, HIMAGELIST, ILC_COLOR32,
};
use windows::Win32::UI::Shell::{
    ITaskbarList3, TaskbarList, THBF_ENABLED, THBF_HIDDEN, THBN_CLICKED, THB_BITMAP, THB_FLAGS,
    THB_TOOLTIP, THUMBBUTTON, THUMBBUTTONFLAGS,
};
use windows::Win32::UI::WindowsAndMessaging::{GetSystemMetrics, HICON, SM_CXSMICON};

/// The id of each button, which is what a click comes back as in the `WM_COMMAND` the Shell sends
/// the stand-in's own procedure (see `pin_taskbar_proc`).
///
/// They are this app's numbers rather than the image list's, and the two are deliberately separate:
/// a click is answered by the id, and the mark a button wears is an index into a list that could be
/// rebuilt with the marks in another order without a click meaning anything else.
const PIN_MEDIA_PREVIOUS: u32 = 1;
const PIN_MEDIA_PLAY_PAUSE: u32 = 2;
const PIN_MEDIA_NEXT: u32 = 3;

/// Where each mark sits in the image list. The middle button is the one that changes, so the two
/// halves of it are two marks rather than one redrawn.
const PIN_MEDIA_IMAGE_PREVIOUS: u32 = 0;
const PIN_MEDIA_IMAGE_PLAY: u32 = 1;
const PIN_MEDIA_IMAGE_PAUSE: u32 = 2;
const PIN_MEDIA_IMAGE_NEXT: u32 = 3;

/// The marks the list holds, in the order the buttons index them — a step back, the two halves of
/// the hold, and a step on. The constants above are positions in this array, which is what the test
/// below holds them to.
const PIN_MEDIA_MARKS: [pin_chrome::ControlGlyph; 4] = [
    pin_chrome::ControlGlyph::Previous,
    pin_chrome::ControlGlyph::Play,
    pin_chrome::ControlGlyph::Pause,
    pin_chrome::ControlGlyph::Next,
];

/// The side an icon is drawn at before any display scale is applied: the size `SM_CXSMICON` reports
/// at 100%, which is the size the marks below are proportioned against.
const PIN_MEDIA_NOMINAL_SIDE: i32 = 16;

/// What the buttons were last told, or nothing while they have never been added.
///
/// Two facts rather than a dirty flag: the tick asks once a frame, so what it needs is the last
/// answer to compare today's against, and a change of either fact is what earns the Shell a call.
///
/// shortcut: a toolbar lost to an Explorer restart is not put back — the Shell's own rule is that
/// a toolbar cannot be removed or re-added without re-creating the window it hangs on, and the
/// stand-in's window is created once for the run. Upgrade by destroying and re-creating the
/// stand-in's window when the Shell announces it has restarted (`TaskbarCreated`), which is the
/// same message the tray already re-adds its own icon on.
static PIN_MEDIA_BUTTONS: Mutex<Option<PinMediaButtons>> = Mutex::new(None);

/// The image list the buttons are drawn from, kept so that it can be given back with the window it
/// was handed to (see `forget_media_toolbar`). The Shell draws a toolbar's marks from this list for
/// as long as the toolbar lives, so it is the window's to keep and nobody else's — and the window
/// is the stand-in's, which is created once and then never taken down except by the Shell's own
/// Close (see `pin_taskbar::close_pin_taskbar`).
static PIN_MEDIA_IMAGES: AtomicIsize = AtomicIsize::new(0);

/// The two things the trio can be told.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct PinMediaButtons {
    /// Whether the pin is a video or a sound, so that the three are shown rather than hidden. A
    /// picture, a document and a card-less kind have no transport to offer, and three buttons
    /// offering a walk of the folder are not this feature (see `Pin Mode → Navigation`).
    showing: bool,
    /// Whether what is behind the middle button is going, so that it wears the pause mark.
    playing: bool,
}

/// Tell the taskbar button's hover what the pin is offering, if the stand-in is up to be told.
///
/// Called on every tick of the preview loop with the two facts the chrome is drawn from: what the
/// pin is showing, and whether that is playing. Nothing is asked of the Shell until the stand-in's
/// window is up — an unshown window has no taskbar button, and a toolbar cannot be added to one —
/// and nothing is sent once the answer stops changing. The picture the hover shows is the same two
/// facts asked of the same tick (see `sync_pin_mirror`).
pub(super) fn refresh_pin_media_buttons(kind: Option<MediaType>, playing: bool) {
    let Some(hwnd) = pin_taskbar::showing_taskbar_window() else {
        return;
    };

    let wanted = PinMediaButtons {
        showing: matches!(
            kind,
            Some(MediaType::Video) | Some(MediaType::NativeVideo) | Some(MediaType::Audio)
        ),
        playing,
    };

    let Ok(mut pushed) = PIN_MEDIA_BUTTONS.lock() else {
        return;
    };

    match *pushed {
        None => {
            // Safety: the handle is the stand-in's own window, which this thread created, and it is
            // the only window a toolbar of this app's is ever added to.
            if unsafe { add_media_toolbar(hwnd, wanted) } {
                *pushed = Some(wanted);
            }
        }
        Some(current) if current != wanted => {
            // Safety: the same window, and one the toolbar was added to on an earlier tick.
            if unsafe { push_media_buttons(hwnd, wanted) } {
                *pushed = Some(wanted);
            }
        }
        Some(_) => {}
    }
}

/// What one of the buttons' clicks asks for, as the command the pin's own caption asks for.
///
/// `None` for an id this app did not add: the number arrives in a `WM_COMMAND` from the Shell, so
/// it is a number from outside this module whatever it says, and one this app does not know is a
/// press with no meaning rather than a press to guess at.
pub(super) fn pin_media_button_command(id: u32) -> Option<PinCommand> {
    match id {
        PIN_MEDIA_PREVIOUS => Some(PinCommand::Previous),
        PIN_MEDIA_PLAY_PAUSE => Some(PinCommand::TogglePlayback),
        PIN_MEDIA_NEXT => Some(PinCommand::Next),
        _ => None,
    }
}

/// Whether a `WM_COMMAND`'s parameters are one of these buttons being pressed, and which button it
/// was: the low word is the id and the high word is the notice, which is the Shell's own
/// arrangement rather than this app's (see `pin_taskbar_proc`).
pub(super) fn pin_media_button_click(wparam: WPARAM) -> Option<PinCommand> {
    let id = (wparam.0 & 0xFFFF) as u32;
    let notice = (wparam.0 >> 16) as u32;

    (notice == THBN_CLICKED)
        .then(|| pin_media_button_command(id))
        .flatten()
}

/// The three buttons as the Shell reads them, in the order they are drawn.
fn media_buttons(state: PinMediaButtons) -> [THUMBBUTTON; 3] {
    let flags = if state.showing {
        THBF_ENABLED
    } else {
        THBF_HIDDEN
    };

    [
        media_button(
            PIN_MEDIA_PREVIOUS,
            PIN_MEDIA_IMAGE_PREVIOUS,
            "Previous",
            flags,
        ),
        media_button(
            PIN_MEDIA_PLAY_PAUSE,
            if state.playing {
                PIN_MEDIA_IMAGE_PAUSE
            } else {
                PIN_MEDIA_IMAGE_PLAY
            },
            if state.playing { "Pause" } else { "Play" },
            flags,
        ),
        media_button(PIN_MEDIA_NEXT, PIN_MEDIA_IMAGE_NEXT, "Next", flags),
    ]
}

/// One button: the id a press comes back as, the mark it wears out of the image list, the name the
/// Shell says while the pointer rests on it, and whether it is offered at all.
fn media_button(id: u32, image: u32, name: &str, flags: THUMBBUTTONFLAGS) -> THUMBBUTTON {
    THUMBBUTTON {
        dwMask: THB_BITMAP | THB_TOOLTIP | THB_FLAGS,
        iId: id,
        iBitmap: image,
        // The mark comes from the image list, which is the field the Shell is told about below;
        // an icon handle is the other way of giving it a mark and this app has none to give.
        hIcon: HICON::default(),
        szTip: tooltip(name),
        dwFlags: flags,
    }
}

/// A name as the Shell's own field: the name and its terminator, the rest of the field empty. A
/// name too long for the field is cut rather than refused, which is what a tooltip is — and it is
/// cut with the terminator's own place kept for it, since a field with no terminator in it is a
/// field the Shell reads past the end of.
fn tooltip(name: &str) -> [u16; 260] {
    let mut field = [0u16; 260];
    let room = field.len() - 1;
    for (slot, unit) in field.iter_mut().zip(name.encode_utf16()).take(room) {
        *slot = unit;
    }
    field
}

/// Add the toolbar to the stand-in's window: the marks, the three buttons, and the state they start
/// in. Answers whether the Shell took them.
///
/// It is called once, on the first tick that finds the stand-in up. A second call would fail — the
/// buttons of a toolbar cannot be added twice — which is why the state above remembers that it was
/// this tick that made them.
unsafe fn add_media_toolbar(hwnd: HWND, state: PinMediaButtons) -> bool {
    let Some(list) = taskbar_list() else {
        return false;
    };

    // The list is kept for as long as the buttons are drawn from it, which is the life of the
    // stand-in's window — created once and never destroyed (see `pin_taskbar`). Four small bitmaps
    // for the life of the process is the same price the stand-in's window itself is, and paid once.
    let images = icon_list();
    if images.is_invalid() {
        return false;
    }

    if list.ThumbBarSetImageList(hwnd, images).is_err() {
        // The marks were not taken, so the list is this call's own to give back: this is the road a
        // tick takes over and over — a window the Shell has no button for, a Shell that is
        // restarting — and a list left behind on each of them is memory nobody will ever read.
        let _ = ImageList_Destroy(images);
        return false;
    }

    // Added hidden and then told their state, rather than added in it: what a button wears is said
    // by one function either way, and a state that has gone wrong has one place to be read in.
    let hidden = PinMediaButtons {
        showing: false,
        ..state
    };
    if list
        .ThumbBarAddButtons(hwnd, &media_buttons(hidden))
        .is_err()
    {
        let _ = ImageList_Destroy(images);
        return false;
    }

    // The buttons are the Shell's from here and the list is what they are drawn from, so the list
    // stays and the state is asked for once more: a state the Shell would not take leaves the three
    // hidden rather than wrong, which is a hover with no buttons in it and not a hover with the
    // wrong ones. What is not done is adding them again — a toolbar's buttons cannot be added
    // twice — so the answer is yes from here whatever that update said.
    PIN_MEDIA_IMAGES.store(images.0, Ordering::SeqCst);
    push_media_buttons(hwnd, state);
    true
}

/// Give back what a toolbar was drawn from, for a stand-in being destroyed: after this the next
/// tick adds a toolbar and a list of its own, so a window that has gone takes nothing of the marks
/// with it and a window made later has none of the old ones.
///
/// Nothing is done here for an add that already refused: a list this call is the owner of is given
/// back where it was made (see `add_media_toolbar`), and only a list the Shell took is remembered
/// at all.
pub(super) fn forget_media_toolbar() {
    if let Ok(mut pushed) = PIN_MEDIA_BUTTONS.lock() {
        *pushed = None;
    }

    let images = PIN_MEDIA_IMAGES.swap(0, Ordering::SeqCst);
    if images != 0 {
        // Safety: the list was made by this module, the toolbar drawn from it went with the window
        // it was added to, and the handle is taken out of the slot in the same move — so this is
        // the last reference there is.
        unsafe {
            let _ = ImageList_Destroy(HIMAGELIST(images));
        }
    }
}

/// Tell the Shell what the three buttons are now, answering whether it took the answer.
unsafe fn push_media_buttons(hwnd: HWND, state: PinMediaButtons) -> bool {
    let Some(list) = taskbar_list() else {
        return false;
    };

    list.ThumbBarUpdateButtons(hwnd, &media_buttons(state))
        .is_ok()
}

/// The Shell's own taskbar object, made and dropped around one ask.
///
/// It is a singleton inside the Shell, so making one is a lookup rather than a start, and it is
/// made rather than kept because keeping it would be a COM interface living in a `static`: a handle
/// that belongs to the thread that made it, which is a thing the rest of the program would have to
/// agree about for the sake of two calls per pin.
unsafe fn taskbar_list() -> Option<ITaskbarList3> {
    CoCreateInstance(&TaskbarList, None, CLSCTX_INPROC_SERVER).ok()
}

/// The image list the buttons draw their marks from.
///
/// A list of nothing where anything of it could not be had, because a list with a hole in it is a
/// list whose indices are wrong: the buttons would wear each other's marks, which is worse than
/// buttons that are not there.
unsafe fn icon_list() -> HIMAGELIST {
    let side = small_icon_side();
    let images = ImageList_Create(side, side, ILC_COLOR32, PIN_MEDIA_MARKS.len() as i32, 0);
    if images.is_invalid() {
        return images;
    }

    let ink = peek_ink(system_uses_light_theme());
    for mark in PIN_MEDIA_MARKS {
        let Some(icon) = icon_bitmap(mark, side, ink) else {
            return HIMAGELIST::default();
        };

        // The list copies the bitmap as it is added, so the surface the mark was painted into is
        // done with by the end of this turn of the loop (see `DibSurface`).
        if ImageList_Add(images, icon.bitmap(), HBITMAP::default()) < 0 {
            return HIMAGELIST::default();
        }
    }

    images
}

/// One mark painted into a bitmap of its own, with nothing behind it.
///
/// The buffer is left at nothing and the mark is drawn over it premultiplied, which is what makes
/// the alpha channel the mark's own shape rather than the box it was drawn in: a glyph on an empty
/// surface is a glyph with a hole for a background, and a hole is what an icon is.
fn icon_bitmap(mark: pin_chrome::ControlGlyph, side: i32, ink: [u8; 3]) -> Option<DibSurface> {
    let surface = DibSurface::create(side as u32, side as u32)?;

    // Safety: the surface was made by this call and nothing else has a pointer to what is behind
    // it, so clearing the whole of it is writing what this call owns.
    unsafe {
        std::ptr::write_bytes(
            pin_chrome::surface_pixels(&surface),
            0,
            side as usize * side as usize * 4,
        );
    }

    pin_chrome::paint_card_control(
        &surface,
        RECT {
            left: 0,
            top: 0,
            right: side,
            bottom: side,
        },
        mark,
        ink,
        // The scale the mark is drawn against is the box's own against the size a mark is
        // proportioned for, so a box the display has doubled draws the mark doubled with it.
        side as f32 / PIN_MEDIA_NOMINAL_SIDE as f32,
    );

    Some(surface)
}

/// The side an icon is drawn at on this machine, which is what the Shell draws a taskbar button's
/// icon in.
///
/// `SM_CXSMICON` is asked rather than a number written here: a 16-pixel mark on a display at 200%
/// is a mark at a quarter of the size the Shell gives it, and the answer is measured rather than
/// assumed — but it is the *system's* answer, so a taskbar on another display at another scale gets
/// this one's.
///
/// shortcut: the system size and not the taskbar's own, which is right everywhere the taskbar is on
/// the primary display. Upgrade to the taskbar's monitor and DPI if a taskbar on a differently
/// scaled display is ever seen with marks the wrong size.
fn small_icon_side() -> i32 {
    // Safety: asks the system for a metric, which is a number rather than a resource.
    unsafe { GetSystemMetrics(SM_CXSMICON).max(PIN_MEDIA_NOMINAL_SIDE) }
}

/// The colour the marks are drawn in: dark on a light hover and light on a dark one.
///
/// The hover is the Shell's, so its background is the taskbar's theme rather than this app's — the
/// preview's own theme is a setting of its own and has nothing to say about it — and a mark drawn
/// in one colour is a mark that disappears on half the machines there are. The value it is read
/// against is the system's own (see `system_uses_light_theme`).
fn peek_ink(light_theme: Option<u32>) -> [u8; 3] {
    match light_theme {
        // Nought is the taskbar drawn dark, which is a light mark's background. One, a value this
        // machine does not hold, and a value that could not be read are all the light taskbar
        // Windows draws by default, which is a dark mark's.
        Some(0) => PIN_MEDIA_LIGHT_INK,
        _ => PIN_MEDIA_DARK_INK,
    }
}

/// The mark on a light hover and the mark on a dark one: near-black and near-white rather than
/// either extreme, because a mark drawn in the extreme has the antialiasing taken out of it.
const PIN_MEDIA_DARK_INK: [u8; 3] = [32, 32, 32];
const PIN_MEDIA_LIGHT_INK: [u8; 3] = [238, 238, 238];

/// Whether Windows is drawing its taskbar and its hovers light, as the value the system's own
/// Personalization settings write, or nothing where the machine has no answer.
///
/// It is read once, when the marks are painted, because the colour is the whole of what it decides.
/// A theme changed while the app is running is a hover wearing the marks of the theme before it,
/// which is a colour that is merely hard to see rather than a button that is gone.
///
/// shortcut: a theme switched at run time does not repaint the marks. Upgrade by setting the image
/// list again (`ThumbBarSetImageList`) with the marks repainted when the value is seen to change.
fn system_uses_light_theme() -> Option<u32> {
    const PERSONALIZE: PCWSTR =
        w!("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize");
    const LIGHT: PCWSTR = w!("SystemUsesLightTheme");

    // Safety: opens a key of the current user's own hive for reading, asks it for one number, and
    // closes it. Nothing else is read, and the handle is closed on every road out of the block.
    unsafe {
        let mut hkey = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, PERSONALIZE, 0, KEY_READ, &mut hkey).is_err() {
            return None;
        }

        let mut value: u32 = 0;
        let mut size = std::mem::size_of::<u32>() as u32;
        let read = RegQueryValueExW(
            hkey,
            LIGHT,
            None,
            None,
            Some(std::ptr::from_mut(&mut value).cast()),
            Some(&mut size),
        )
        .is_ok();
        let _ = RegCloseKey(hkey);

        // A value shorter than the number it is read as is a value this app has no answer from,
        // rather than one it fills the rest of with whatever was in the variable.
        (read && size >= std::mem::size_of::<u32>() as u32).then_some(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The marks and the indices the buttons read them by are two lists of one order, and this is
    /// what holds them to it: a mark moved in the array without its index is a button wearing
    /// another button's mark, which is the one mistake here a click would not reveal.
    #[test]
    fn every_mark_sits_at_the_index_its_button_reads_it_by() {
        assert_eq!(
            PIN_MEDIA_MARKS[PIN_MEDIA_IMAGE_PREVIOUS as usize],
            pin_chrome::ControlGlyph::Previous
        );
        assert_eq!(
            PIN_MEDIA_MARKS[PIN_MEDIA_IMAGE_PLAY as usize],
            pin_chrome::ControlGlyph::Play
        );
        assert_eq!(
            PIN_MEDIA_MARKS[PIN_MEDIA_IMAGE_PAUSE as usize],
            pin_chrome::ControlGlyph::Pause
        );
        assert_eq!(
            PIN_MEDIA_MARKS[PIN_MEDIA_IMAGE_NEXT as usize],
            pin_chrome::ControlGlyph::Next
        );
    }

    /// What a press on each button asks for, and what a press from anywhere else asks for nothing.
    #[test]
    fn the_three_buttons_are_the_pins_own_three_commands() {
        assert_eq!(
            pin_media_button_command(PIN_MEDIA_PREVIOUS),
            Some(PinCommand::Previous)
        );
        assert_eq!(
            pin_media_button_command(PIN_MEDIA_PLAY_PAUSE),
            Some(PinCommand::TogglePlayback)
        );
        assert_eq!(
            pin_media_button_command(PIN_MEDIA_NEXT),
            Some(PinCommand::Next)
        );

        assert_eq!(pin_media_button_command(0), None);
        assert_eq!(pin_media_button_command(4), None);
        assert_eq!(pin_media_button_command(u32::MAX), None);
    }

    /// A `WM_COMMAND` is the Shell's own arrangement of two numbers, and only the two together are
    /// a press: an id without the notice is a message about a button rather than one on it.
    #[test]
    fn a_click_is_read_from_the_id_and_the_notice_together() {
        let clicked = |id: u32| WPARAM((id | (THBN_CLICKED << 16)) as usize);

        assert_eq!(
            pin_media_button_click(clicked(PIN_MEDIA_NEXT)),
            Some(PinCommand::Next)
        );
        assert_eq!(
            pin_media_button_click(clicked(PIN_MEDIA_PLAY_PAUSE)),
            Some(PinCommand::TogglePlayback)
        );
        assert_eq!(pin_media_button_click(clicked(9)), None);
        assert_eq!(
            pin_media_button_click(WPARAM(PIN_MEDIA_NEXT as usize)),
            None,
            "an id with no notice is not a click"
        );
    }

    /// The three buttons as the Shell reads them: hidden while the pin is not a medium, and the
    /// middle one saying what a press would do rather than what the pin is.
    #[test]
    fn the_middle_button_says_what_a_press_would_do_and_the_others_say_what_they_are() {
        let held = media_buttons(PinMediaButtons {
            showing: true,
            playing: false,
        });

        assert_eq!(held[0].iId, PIN_MEDIA_PREVIOUS);
        assert_eq!(held[0].iBitmap, PIN_MEDIA_IMAGE_PREVIOUS);
        assert_eq!(held[1].iId, PIN_MEDIA_PLAY_PAUSE);
        assert_eq!(
            held[1].iBitmap, PIN_MEDIA_IMAGE_PLAY,
            "a pin that is not playing wears the mark that begins it"
        );
        assert_eq!(held[2].iId, PIN_MEDIA_NEXT);
        assert_eq!(held[2].iBitmap, PIN_MEDIA_IMAGE_NEXT);

        for button in held {
            assert_eq!(button.dwFlags, THBF_ENABLED);
            assert!(
                button.szTip.iter().any(|unit| *unit != 0),
                "every button says what it is"
            );
            assert!(button.szTip.contains(&0), "and what it says is terminated");
        }

        let going = media_buttons(PinMediaButtons {
            showing: true,
            playing: true,
        });
        assert_eq!(going[1].iBitmap, PIN_MEDIA_IMAGE_PAUSE);

        // Nothing is offered by a pin that is not a medium, and the marks are the same ones behind
        // the flag: what the buttons wear and whether they are up are two questions.
        let offers_nothing = media_buttons(PinMediaButtons {
            showing: false,
            playing: true,
        });
        for button in offers_nothing {
            assert_eq!(button.dwFlags, THBF_HIDDEN);
        }
        assert_eq!(offers_nothing[1].iBitmap, PIN_MEDIA_IMAGE_PAUSE);
    }

    /// The ink is the hover's own theme, and a machine that cannot say which theme it is in is given
    /// the light one — which is what Windows is in until its user says otherwise.
    #[test]
    fn the_marks_are_drawn_against_the_theme_the_hover_is_drawn_in() {
        assert_eq!(
            peek_ink(Some(0)),
            PIN_MEDIA_LIGHT_INK,
            "a taskbar drawn dark is a light mark's background"
        );
        assert_eq!(peek_ink(Some(1)), PIN_MEDIA_DARK_INK);
        assert_eq!(
            peek_ink(None),
            PIN_MEDIA_DARK_INK,
            "a machine with no answer is given the theme Windows is in by default"
        );

        assert!(
            PIN_MEDIA_LIGHT_INK[0] > PIN_MEDIA_DARK_INK[0] + 150,
            "and the two are opposite enough for one of them to be read on the other's background"
        );
    }

    /// A name keeps its termination: the field is the Shell's and the name is the whole of what is
    /// written into it, so a name that fills it is a name that has nowhere to end.
    #[test]
    fn a_name_fills_its_field_and_stops_at_the_end_of_it() {
        let named = tooltip("Play");
        assert_eq!(named[0], 'P' as u16);
        assert_eq!(named[3], 'y' as u16);
        assert_eq!(named[4], 0, "and it ends where the name does");

        let long = tooltip(&"n".repeat(400));
        assert_eq!(long.len(), 260);
        assert_eq!(
            long[259], 0,
            "the field is terminated rather than overflowed"
        );
    }
}

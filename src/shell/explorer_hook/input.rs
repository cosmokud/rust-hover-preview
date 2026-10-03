//! What the keyboard and the mouse are doing, read once per tick and handed to whatever
//! asks: the three buttons, the keys that matter, and the navigation a tick is allowed
//! to treat as movement in the listing.
//!
//! One reader per key on purpose. A press bit is spent by the first read of it, so two
//! readers on two threads are one answer and one silence, and which of them is which is
//! a race.

use super::*;

/// Left, right and middle buttons as one pass over them finds them.
///
/// They are deliberate input even when the cursor never moves: the folder a double-click
/// opens is user navigation, not a background change the preview has to wait out.
///
/// The left button is kept apart from the other two, and the pass is the only read of them
/// in the app, because a press bit is spent by the first read of a key: two readers on two
/// threads are one answer and one silence, and which of them it is is a race. The pin's own
/// press handling needs the left button on its own — a drag is begun from a press, and a
/// right button is not one — and it is asked from the preview thread, which therefore reads
/// what this pass published rather than reading the key again (see `publish_pin_media_press`
/// and `settle_pinned_engine_press`).
#[derive(Clone, Copy, Default)]
pub(super) struct MouseButtons {
    /// Whether any of the three is down, which is what the hold a page is dragged under and
    /// the settle a press ends are both read of.
    pub(super) active: bool,
    /// Whether any of them was pressed since the previous pass: the transition, which is
    /// what a click is known by and what a folder change is told apart from a held key by.
    pub(super) pressed: bool,
    /// Whether the left button is down on its own.
    left_down: bool,
    /// Whether the left button was pressed since the previous pass on its own.
    left_pressed: bool,
}

/// Read the buttons, and keep the left one apart as above. The index rather than the key is
/// what tells the left button from the other two, which is why `mouse_press_buttons` leads
/// with it and why that order is part of this function's contract.
pub(super) fn mouse_buttons() -> MouseButtons {
    let mut state = MouseButtons::default();

    for (index, &key) in mouse_press_buttons().iter().enumerate() {
        let raw = unsafe { GetAsyncKeyState(key.0 as i32) as u16 };
        let pressed = (raw & 0x0001) != 0;

        if is_pressed_or_down_state(raw) {
            state.active = true;
        }
        if pressed {
            state.pressed = true;
        }
        if index == 0 {
            state.left_down = is_key_down_state(raw);
            state.left_pressed = pressed;
        }
    }

    state
}

/// Enter opens the focused item, so it drives folder changes without any
/// pointer input at all.
pub(super) fn activation_key_input_state() -> (bool, bool) {
    key_input_state(&[VK_RETURN])
}

pub(super) fn key_input_state(
    keys: &[windows::Win32::UI::Input::KeyboardAndMouse::VIRTUAL_KEY],
) -> (bool, bool) {
    unsafe {
        let mut active = false;
        let mut pressed = false;
        for &key in keys {
            let state = GetAsyncKeyState(key.0 as i32) as u16;
            if is_pressed_or_down_state(state) {
                active = true;
            }
            if (state & 0x0001) != 0 {
                pressed = true;
            }
        }

        (active, pressed)
    }
}

pub(super) fn folder_probe_interval_ms(preview_active: bool) -> u64 {
    if preview_active {
        FOLDER_PROBE_MS
    } else {
        IDLE_FOLDER_PROBE_MS
    }
}

pub(super) fn is_key_down_state(state: u16) -> bool {
    (state & 0x8000) != 0
}

pub(super) fn is_key_down(vk: i32) -> bool {
    unsafe { is_key_down_state(GetAsyncKeyState(vk) as u16) }
}

pub(super) fn is_explorer_navigation_shortcut_key(
    key_vk: i32,
    alt_down: bool,
    ctrl_down: bool,
) -> bool {
    key_vk == VK_BACK_CODE
        || (alt_down
            && matches!(
                key_vk,
                key if key == VK_LEFT.0 as i32 || key == VK_RIGHT.0 as i32 || key == VK_UP.0 as i32
            ))
        || (ctrl_down && key_vk == VK_T_CODE)
}

/// Everything one poll of the keyboard says about navigation.
///
/// `active` is true while a navigation key — or one of the keys a file name is typed
/// with, which is what Explorer's own type-ahead answers — is held or was pressed since
/// the previous poll, and `pressed` is the fresh press transition alone: a held key keeps
/// reporting `active` forever, so a folder change uses `pressed` to tell a new key press
/// apart from state left over from the navigation that opened the folder (see
/// `keyboard_navigation_press_seq`).
///
/// `shortcut` is Explorer's own navigation being *asked* for — a Backspace, an arrow under
/// Alt, a Ctrl+T — which is input the app acts on rather than state it waits out.
///
/// There is no arrow here that walks a pinned window. An arrow a pin answers is a key Windows
/// routed to the pin because the pin is the window the user is in, so it arrives as a message
/// and is answered in the pin's own window procedure (see `pinned_key_command`). Reading the
/// arrow from the keyboard instead would be reading a key nobody sent this app, and would
/// answer it in addition to the listing behind, which is where a folder listing moved while a
/// preview nobody was in had come to.
#[derive(Clone, Copy, Default)]
pub(super) struct NavigationInput {
    pub(super) active: bool,
    pub(super) pressed: bool,
    pub(super) shortcut: bool,
}

/// The keys a file name is typed with: the letters and the digits, the numpad's own among
/// them. Explorer answers a name being typed with its type-ahead, which moves the selection
/// to the item the name matches — an item the keyboard is on changing with no navigation key
/// touched — and these are the keys that say it is happening (see `navigation_input`).
///
/// The codes are the keys rather than the characters, so a name typed on a layout that is
/// not this one is read as the same keys: what has to be noticed is that the keyboard has
/// moved the item, and not what the name spells. They are read only where neither Ctrl nor
/// Alt is held — a letter under one of those is a command, and a command is not a name —
/// which is what keeps the press bit of `C` where it belongs: the text preview's own Ctrl+C
/// reads that key for itself, and a press bit is consumed by whoever reads it first.
pub(super) fn type_ahead_keys() -> impl Iterator<Item = i32> {
    (0x30..=0x39).chain(0x60..=0x69).chain(0x41..=0x5a)
}

/// The navigation keys, the keys a file name is typed with, and the shortcuts made of
/// them, read in one pass.
///
/// One pass because the keys overlap: the arrows are navigation keys *and* half of a
/// shortcut, and the press bit `GetAsyncKeyState` reports is consumed by whoever reads a key
/// first — so a second read of the same key in one tick is not the same answer twice, it is
/// a read that finds the bit already taken (see the caller, which reads the navigation keys
/// ahead of everything else for exactly that reason). The arrows are read here, once, and
/// the shortcut asks its question of the same reading: whether the key is *down*, which is
/// what the second read would have found.
///
/// The keys a name is typed with are read with them, and they are not navigation keys at
/// all: Explorer answers a name being typed with its own type-ahead, which moves the
/// selection — and with it the item the keyboard is on — with none of the keys above
/// touched. A preview that did not read them would stay on the item the keyboard came from,
/// because the focus is only probed for a moment after input
/// (`should_probe_keyboard_focus`) and typing would never open that moment again. They open
/// it as a navigation key does, and their fresh press is the same fresh press: what the
/// typing selected is the user's own choice and must not be swallowed as a baseline. They
/// are read only where neither Ctrl nor Alt is held, so a command is not read as a name and
/// the keys a command is made of are left to whoever reads them (see `type_ahead_keys`).
///
/// The two keys that make a shortcut and are not navigation keys are read whatever the
/// modifiers are, and that is not an oversight. A Backspace navigates with no modifier at
/// all, so gating it on a modifier would lose the shortcut outright; and a `T` typed
/// anywhere on the machine sets its own press bit, so gating *that* read on Ctrl being held
/// would leave the bit standing until Ctrl was next pressed — and a Ctrl pressed for
/// something else would then read as a Ctrl+T that opens a tab. What leaving a press bit
/// standing costs is the read that consumes it, which is why both are read every tick.
///
/// A page the user has clicked into owns the keyboard, and everything read above belongs to
/// the page: Explorer's navigation shortcut — an arrow under a modifier, a Backspace, a `T`
/// — and Explorer's type-ahead, a letter or a digit, are gestures the user made on the page
/// and not in the folder view. The reads are still made — that is the point of them, and a
/// key not read here is a press bit left standing to be read later as somebody else's
/// gesture — but nothing this function has worked out is acted on while the page is in front
/// (see `engine_owns_the_keyboard`).
pub(super) fn navigation_input() -> NavigationInput {
    let alt_down = is_key_down(VK_MENU_CODE);
    let ctrl_down = is_key_down(VK_CONTROL_CODE);
    let mut input = NavigationInput::default();

    let navigation_keys = [
        VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT, VK_HOME, VK_END, VK_PRIOR, VK_NEXT,
    ];

    for &key in &navigation_keys {
        let key_vk = key.0 as i32;
        let state = unsafe { GetAsyncKeyState(key_vk) as u16 };

        if is_pressed_or_down_state(state) {
            input.active = true;
        }
        if (state & 0x0001) != 0 {
            input.pressed = true;
        }
        if is_key_down_state(state)
            && is_explorer_navigation_shortcut_key(key_vk, alt_down, ctrl_down)
        {
            input.shortcut = true;
        }
    }

    for key_vk in [VK_BACK_CODE, VK_T_CODE] {
        let state = unsafe { GetAsyncKeyState(key_vk) as u16 };

        if is_pressed_or_down_state(state)
            && is_explorer_navigation_shortcut_key(key_vk, alt_down, ctrl_down)
        {
            input.shortcut = true;
        }
    }

    if !ctrl_down && !alt_down {
        for key_vk in type_ahead_keys() {
            let state = unsafe { GetAsyncKeyState(key_vk) as u16 };

            if is_pressed_or_down_state(state) {
                input.active = true;
            }
            if (state & 0x0001) != 0 {
                input.pressed = true;
            }
        }
    }

    // The pass is over and the press bits it found are spent, so there is nothing left to
    // hand back: the page in front gets the keys, and the folder view is not told about any
    // of them this tick.
    if engine_owns_the_keyboard() {
        return NavigationInput::default();
    }

    input
}

pub(super) fn mouse_navigation_buttons(
) -> [windows::Win32::UI::Input::KeyboardAndMouse::VIRTUAL_KEY; 2] {
    [VK_XBUTTON1, VK_XBUTTON2]
}

pub(super) fn is_mouse_navigation_button_detected() -> bool {
    mouse_navigation_buttons()
        .iter()
        .any(|&key| unsafe { is_pressed_or_down_state(GetAsyncKeyState(key.0 as i32) as u16) })
}

/// The three press buttons, left one first.
///
/// The order is part of `mouse_buttons`' contract rather than a list: the left button is the
/// one the pin's own press handling is asked about on its own, and the index is what tells
/// it from the other two.
pub(super) fn mouse_press_buttons() -> [windows::Win32::UI::Input::KeyboardAndMouse::VIRTUAL_KEY; 3]
{
    [VK_LBUTTON, VK_RBUTTON, VK_MBUTTON]
}

/// Which of the two things that move the focus did so on one tick, as the rule that tells a pin's
/// keyboard pick from a listing that changed under it reads them (see `focus_move_input` and
/// `PinUpdateWatch::focus_moved_by_key`).
///
/// Two things move the focus in Explorer, and only one of them lands it on a file the user picked.
/// A key the user presses — an arrow, Home, End, a page key, or a letter or a digit, which Explorer
/// answers with its own type-ahead — *walks* the selection, and what it comes to rest on is the
/// user's own choice. Everything else puts the focus on an item that nobody chose: an Enter, which
/// opens the folder or the document the focus was on; a shortcut of Explorer's own, which is a
/// Backspace, an arrow under Alt, or a Ctrl+T; a Tab, which moves between the view's own elements
/// and switches Explorer's own tabs under Ctrl; a click, which is the pointer acting on the view —
/// a tab, a breadcrumb, a folder in the tree, or a file; the mouse's own navigation buttons; a
/// Delete, which takes what the focus was on out of the listing and lands the focus on whatever
/// takes its place; and any key held with Ctrl, Alt or Windows down, which is a command rather than
/// a move (a Ctrl+Tab, a Ctrl+1, a Ctrl+L and an Alt+Tab all change which listing is on screen or
/// what it is showing, and none of them walks a selection).
#[derive(Clone, Copy, Default)]
pub(super) struct FocusMoveInput {
    /// A key that walks a listing was pressed or is held.
    pub(super) walked_by_key: bool,
    /// Something that moves the focus without a key having walked it did so.
    pub(super) moved_otherwise: bool,
    /// The pointer acted on the view, which is a click. It is read here because the press bit a
    /// click is known by can only be read once a tick.
    pub(super) clicked: bool,
}

/// What the keyboard and the pointer did on one tick, for the rule above.
///
/// Read on every tick a pin is up — where the setting asks the pin to follow and where it does
/// not, so a click nobody asked about is not left standing as the answer the next read gets — and
/// a pinned tick is the one place that reads these keys at all: the loop returns at the pin before
/// its own input reads, so the press bits this spends are ones nothing else in the tick was going
/// to have (see `navigation_input`, whose one read per key per tick has to be the first).
pub(super) fn focus_move_input() -> FocusMoveInput {
    let navigation = navigation_input();
    let (_, activation_pressed) = activation_key_input_state();
    let (_, deletion_pressed) = key_input_state(&[VK_DELETE]);
    // A Tab is not a key that walks a listing: it moves between the view's own elements, and a Tab
    // under Ctrl is a tab of Explorer's switched — the one move onto another listing that neither
    // the shortcut set nor a button of the mouse's answers (see `is_explorer_navigation_shortcut_key`).
    let (_, tab_pressed) = key_input_state(&[VK_TAB]);
    // Read once per tick, and the left button kept apart within it: this pass is the only
    // read of the buttons in the app, and a press bit is spent by the first reader of a key.
    let mouse = mouse_buttons();

    // What the pin's own press handling needs, published rather than read again there — it
    // runs on the preview thread, and a read of the key from that thread would spend the
    // press this pass has just found, which is the very click the listing behind the pin is
    // answered by (see `MouseButtons` and `settle_pinned_engine_press`).
    publish_pin_media_press(mouse.left_down, mouse.left_pressed);

    FocusMoveInput {
        walked_by_key: navigation.active,
        moved_otherwise: navigation.shortcut
            || activation_pressed
            || deletion_pressed
            || tab_pressed
            || is_mouse_navigation_button_detected()
            || is_key_down(VK_CONTROL_CODE)
            || is_key_down(VK_MENU_CODE)
            || is_key_down(VK_LWIN_CODE)
            || is_key_down(VK_RWIN_CODE),
        clicked: mouse.pressed,
    }
}

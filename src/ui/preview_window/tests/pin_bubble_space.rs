use super::*;
use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, RECT, WPARAM};
use windows::Win32::UI::Input::KeyboardAndMouse::GetFocus;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, GetWindow, GetWindowRect, IsIconic, IsWindow, IsWindowVisible,
    SendMessageW, ShowWindow, GW_HWNDPREV, SC_CLOSE, SC_MINIMIZE, SW_SHOWNOACTIVATE,
    WINDOW_EX_STYLE, WM_CLOSE, WM_SYSCOMMAND, WS_POPUP,
};

/// A SPACE with a focused bubble holds a film or a sound, and nothing else.
#[test]
fn space_on_a_focused_bubble_holds_only_a_film_or_a_sound() {
    for kind in [MediaType::Video, MediaType::NativeVideo, MediaType::Audio] {
        assert!(
            bubble_space_toggles(Some(kind)),
            "a {kind:?} bubble has playback for SPACE to hold"
        );
    }

    for kind in [
        MediaType::StaticImage,
        MediaType::Text,
        MediaType::Pdf,
        MediaType::Office,
        MediaType::Unplayable,
    ] {
        assert!(
            !bubble_space_toggles(Some(kind)),
            "a {kind:?} bubble has no playback, so SPACE is not a hold"
        );
    }

    assert!(
        !bubble_space_toggles(None),
        "with nothing on screen SPACE is not a hold either"
    );
}

/// Only a click that moved nothing on a bubble the hand was already in puts the window back: the
/// focusing click stays a bubble, so SPACE has a window to hold.
#[test]
fn only_a_click_on_an_already_focused_bubble_puts_the_window_back() {
    assert!(
        bubble_click_restores(true, false),
        "a click on a focused bubble restores it"
    );
    assert!(
        !bubble_click_restores(false, false),
        "the first click only focuses, so the bubble stays up for SPACE"
    );
    assert!(
        !bubble_click_restores(true, true),
        "a drag carries the bubble rather than restoring it"
    );
    assert!(
        !bubble_click_restores(false, true),
        "a drag that also focused carries the bubble too"
    );
}

/// The bubble's menu holds the hold for a film or a sound, the two walks of the folder it
/// stands in, the switch that decides what a minimize holds back, the two places a minimize can
/// put the pin, and the two ways out of any bubble.
#[test]
fn the_bubble_menu_holds_only_what_has_playback() {
    for kind in [MediaType::Video, MediaType::NativeVideo, MediaType::Audio] {
        assert_eq!(
            bubble_menu_rows(Some(kind), false, true),
            vec![
                BubbleMenuRow::Toggle,
                BubbleMenuRow::NextFile,
                BubbleMenuRow::PreviousFile,
                BubbleMenuRow::Separator,
                BubbleMenuRow::PauseOnMinimize,
                BubbleMenuRow::Separator,
                BubbleMenuRow::MinimizeToBubble,
                BubbleMenuRow::MinimizeToTaskbar,
                BubbleMenuRow::Separator,
                BubbleMenuRow::ToTaskbar,
                BubbleMenuRow::Close,
            ],
            "a {kind:?} bubble offers the hold first"
        );
    }

    for kind in [Some(MediaType::StaticImage), Some(MediaType::Text), None] {
        assert_eq!(
            bubble_menu_rows(kind, false, true),
            vec![
                BubbleMenuRow::NextFile,
                BubbleMenuRow::PreviousFile,
                BubbleMenuRow::Separator,
                BubbleMenuRow::MinimizeToBubble,
                BubbleMenuRow::MinimizeToTaskbar,
                BubbleMenuRow::Separator,
                BubbleMenuRow::ToTaskbar,
                BubbleMenuRow::Close,
            ],
            "a {kind:?} bubble offers no hold, only the walk, the two places and the two ways out"
        );
    }
}

/// The menu names the hold for what it will do, the walk for the file it steps to, and the way
/// out for where it goes: Pause while playing and Play while held, and the taskbar spelled the
/// way the tray spells it.
#[test]
fn the_bubble_menu_names_what_each_row_will_do() {
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Maximize, false, false),
        "Maximize",
        "the caption's maximize is named for what it will do"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Maximize, false, true),
        "Restore",
        "and for the other way it will do it once the window is maximized"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Toggle, true, false),
        "Pause",
        "a hold on playing media pauses it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Toggle, false, false),
        "Play",
        "a hold on held media lets it go"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::NextFile, true, false),
        "Next File",
        "the walk forward is the file after this one"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::PreviousFile, false, false),
        "Previous File",
        "and the walk back the one before it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Separator, true, false),
        "",
        "a rule carries nothing: there is nothing on one to read"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::PauseOnMinimize, true, false),
        "Pause on Minimize",
        "the switch is spelled the way the tray spells it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::MinimizeToBubble, true, false),
        "Minimize to Bubble",
        "the place a minimize can put the pin is named for it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::MinimizeToTaskbar, false, false),
        "Minimize to Taskbar",
        "and so is the other"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::ToBubble, true, false),
        "To Bubble",
        "the bubble a pin still standing is put away into is named for it, the way the taskbar is"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::ToTaskbar, true, false),
        "To Taskbar",
        "the stand-in is the taskbar's, not the tray's"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Close, false, false),
        "Close",
        "the way out is a close"
    );
}

/// The painted menu carries the same rows in the same order, labelled for the playback behind
/// them — the panel the chrome paints cannot disagree with the rows the press answers.
#[test]
fn the_painted_menu_carries_the_rows_the_press_answers() {
    let painted: Vec<String> =
        bubble_menu_chrome_rows(Some(MediaType::Audio), true, false, true, false)
            .iter()
            .map(|row| row.label.clone())
            .collect();
    assert_eq!(
        painted,
        vec![
            "Pause",
            "Next File",
            "Previous File",
            "",
            "Pause on Minimize",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close"
        ],
        "a playing sound's menu pauses first"
    );

    let painted: Vec<String> =
        bubble_menu_chrome_rows(Some(MediaType::Video), false, false, true, false)
            .iter()
            .map(|row| row.label.clone())
            .collect();
    assert_eq!(
        painted,
        vec![
            "Play",
            "Next File",
            "Previous File",
            "",
            "Pause on Minimize",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close"
        ],
        "a held film's menu lets it go"
    );

    let painted: Vec<String> =
        bubble_menu_chrome_rows(Some(MediaType::Text), true, false, true, false)
            .iter()
            .map(|row| row.label.clone())
            .collect();
    assert_eq!(
        painted,
        vec![
            "Next File",
            "Previous File",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close"
        ],
        "a page's menu offers only the walk, the two places and the two ways out"
    );
}

/// The bubble's menu carries no flyout arrow: the shared placement points one wherever a Seek
/// row could open, and there is no Seek row here — so Close, the last row, stands unmarked.
#[test]
fn the_bubble_menu_points_no_arrow_at_any_row() {
    let rows = bubble_menu_chrome_rows(Some(MediaType::Audio), true, false, true, false);
    let popup = bubble_menu_popup((0, 44), 500, 500, 96, &rows);
    assert_eq!(
        popup.arrow,
        pin_chrome::MenuArrow::None,
        "no row of the bubble's menu opens anything"
    );
}

/// The two places a minimize can put a pin are one group rather than two switches: exactly one of
/// them carries the disc, and which one it is is the setting the tray's own `Pin Mode → Minimize`
/// writes — so a menu opened over a pick that has just been made shows it (see `PinMinimizeTo`).
#[test]
fn exactly_one_of_the_two_minimize_picks_carries_the_disc() {
    let _target = PinMinimizeToPutBack::put(PinMinimizeTo::Taskbar);

    let rows = bubble_menu_chrome_rows(Some(MediaType::Text), false, false, true, false);
    let marked: Vec<&str> = rows
        .iter()
        .filter(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet))
        .map(|row| row.label.as_str())
        .collect();
    assert_eq!(
        marked,
        ["Minimize to Taskbar"],
        "the place a minimize puts the pin is the one the disc is on, and it is alone"
    );

    let _target = PinMinimizeToPutBack::put(PinMinimizeTo::Bubble);
    let rows = bubble_menu_chrome_rows(Some(MediaType::Text), false, false, true, false);
    let marked: Vec<&str> = rows
        .iter()
        .filter(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet))
        .map(|row| row.label.as_str())
        .collect();
    assert_eq!(
        marked,
        ["Minimize to Bubble"],
        "and the other pick has it where the setting names the other place"
    );
}

/// The check the `Pause on Minimize` row carries is the switch of the kind the bubble stands
/// for: a film reads the video's and a sound reads the audio's, and a kind with no playback
/// carries no row at all.
#[test]
fn the_pause_switch_carries_the_check_of_the_kind_it_stands_for() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _switches = PausePreviewSettings::stood_at(false, true);

    let film = bubble_menu_chrome_rows(Some(MediaType::Video), false, false, true, false);
    let row = film
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a film's bubble carries the switch");
    assert!(
        matches!(film[row].mark, pin_chrome::MenuMark::Check(false)),
        "the video's switch is off, so the film's row is unchecked"
    );

    let sound = bubble_menu_chrome_rows(Some(MediaType::Audio), false, false, true, false);
    let row = sound
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a sound's bubble carries the switch");
    assert!(
        matches!(sound[row].mark, pin_chrome::MenuMark::Check(true)),
        "the audio's switch is on, so the sound's row is checked"
    );

    let page = bubble_menu_chrome_rows(Some(MediaType::Text), false, false, true, false);
    assert!(
        page.iter().all(|row| row.label != "Pause on Minimize"),
        "a page's bubble has no playback for a minimize to hold"
    );
}

/// A pin put back up out of a bubble the hand never carried comes back at its own box: the same
/// place and the same size it went down at.
///
/// The fault this is a regression test for: every restore placed the window beside the bubble, and
/// the bubble stands on the minimize button of the window the restore has just brought back — so a
/// pin cycled through the bubble with nothing moved walked sideways by most of its own width on
/// every click, shrank as it ran out of room, and ended up against the edge of the screen.
#[test]
fn a_pin_restored_from_a_bubble_that_was_never_carried_comes_back_at_its_own_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_carried = PIN_BUBBLE_CARRIED.load(Ordering::Acquire);

    // A bubble standing where the pin's own box is not: a restore that placed the window beside it
    // would have to leave a different box behind.
    // Safety: a window of this test's own, created and destroyed under the one-at-a-time lock.
    let bubble = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            700,
            500,
            120,
            120,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the bubble")
    };
    PIN_BUBBLE_HWND.store(bubble.0 as isize, Ordering::SeqCst);

    let left_at = (100, 80, 420, 320);
    let mut pin = overlay_pin(left_at, PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    PIN_BUBBLE_CARRIED.store(false, Ordering::Release);

    restore_pin(&a_restored_window(), PinRestore::Pressed);

    let restored = pin_state().and_then(|pinned| {
        pinned
            .pin()
            .map(|pin| (pin.collapsed, pin.content, pin.window_size()))
    });
    assert_eq!(
        restored,
        Some((false, left_at, (320, 240))),
        "a window placed beside the bubble would be a different box, and a pin put back up is the \
         window that went down"
    );

    // Safety: the window is this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(bubble);
    }
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(previous_carried, Ordering::Release);
    stand_pin(previous_pin);
}

/// Let the window server finish the placement just asked for before the z-order is read: a window
/// put in a band and the band's own order are the compositor's to settle, so a test that reads them
/// the instant the call returns is a test racing it — the same window reads as not yet topmost and
/// as topmost a call later.
unsafe fn settle_z_order() {
    let _ = DwmFlush();
}

/// Whether `above` stands over `below` in the z-order right now: an ask of the desktop rather than a
/// reading of what this app meant to ask for.
///
/// Safety: both handles are windows of this test's or of this app's own, and the question is a
/// measurement of them rather than a change.
unsafe fn stands_over(above: HWND, below: HWND) -> bool {
    let mut walk = GetWindow(below, GW_HWNDPREV);
    while let Ok(window) = walk {
        if window == above {
            return true;
        }
        if window.is_invalid() {
            break;
        }
        walk = GetWindow(window, GW_HWNDPREV);
    }

    false
}

/// The offset this file's windows stand at: no display is there, so a run of the suite is no window
/// of the user's and no place the hand can land in.
const NOWHERE: i32 = -32000;

/// A window of the test's own, opened where no display is.
///
/// Safety: the window is the caller's to destroy.
unsafe fn a_window_of_the_tests_own() -> HWND {
    CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("STATIC"),
        w!(""),
        WS_POPUP,
        NOWHERE,
        NOWHERE,
        320,
        240,
        None,
        None,
        None,
        None,
    )
    .expect("a window of the test's own")
}

/// The pin put away into its bubble, with its own window on the desktop and a window of the test's
/// opened over it — the arrangement a restore is asked in, with `Pin Mode → Always On Top` off.
///
/// The second window is made after the pin's, which is what puts it in front, and it is still there
/// after the restore, so that a window which came back where a hidden one sits is a window that came
/// back underneath it.
///
/// The caller holds the pin tests' lock and puts back every global written here: the handle the pin
/// is drawn by, the bubble's, the one saying a bubble was carried, the setting, and the pin itself.
///
/// Safety: both windows are the caller's to destroy.
unsafe fn a_pin_put_away_into_its_bubble() -> (HWND, HWND) {
    let pin_hwnd = a_window_of_the_tests_own();
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
    let in_front = a_window_of_the_tests_own();
    let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
    assert!(
        stands_over(in_front, pin_hwnd),
        "the window opened after the pin's is the one in front of it"
    );

    // The setting the user's own report is under: a pin that is not Always On Top.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    // A pin put away into its bubble: collapsed, its own window gone from the screen, and no stand-in
    // for it but the bubble this test holds no window for.
    PIN_BUBBLE_HWND.store(0, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(false, Ordering::Release);
    let mut pin = overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    );
    pin.collapsed = true;
    stand_pin(Some(pin));
    // What the collapse into a bubble leaves: the window off the screen, and nothing standing in for
    // it but the bubble.
    let _ = ShowWindow(pin_hwnd, SW_HIDE);
    assert!(
        !IsWindowVisible(pin_hwnd).as_bool(),
        "a pin put away into its bubble has no window on the screen"
    );

    (pin_hwnd, in_front)
}

/// A pin brought back out of its bubble comes back in front of the windows that were opened over it,
/// and takes the keyboard with it, with `Pin Mode → Always On Top` off.
///
/// The fault this is a regression test for: a click on the bubble brought the window back *behind*
/// whatever the user had been working in. What a restore asked of the window was a place — `HWND_TOP`
/// for one that was off the screen, and `HWND_NOTOPMOST`, which asks for the place a window already
/// has, for one that was not — and a place is all it got: an ordinary window that is raised is a
/// window the next activation is free to stand over, and the window the hide left in front is the one
/// Windows activated. A press on the bubble is the one thing that says which window the user means, so
/// the window the restore answers it with is the one that comes to the front and the one that holds
/// the caret, the way a press on the window itself leaves it (see `restore_pin`, `pin_take_focus`).
#[test]
fn a_pin_brought_back_by_a_press_on_its_bubble_comes_back_in_front_and_takes_the_keyboard() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_carried = PIN_BUBBLE_CARRIED.load(Ordering::Acquire);
    let before = pin_always_on_top();

    // Safety: the windows are this test's own, made and destroyed under the one-at-a-time lock, and
    // the pin is stood up under it too.
    let (pin_hwnd, in_front) = unsafe { a_pin_put_away_into_its_bubble() };

    // What is asked of the pin's window is recorded rather than done: a window put in front of the
    // user is not something a test may do to their desktop. The window the desktop answers for is the
    // one this test made, because the restore asks about the handle the app holds.
    let window = RecordedPinWindow::new(pin_hwnd.0 as isize);
    restore_pin(&window, PinRestore::Pressed);

    assert!(
        !pin_is_collapsed(),
        "the restore is what puts the pin back up"
    );
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a pin the bubble brought back is put above the windows the user had in front of it: what a \
         hide leaves behind is a place at the bottom of the z-order, and a showing that asks for the \
         place a window already has is a pin that comes back behind everything"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::SetFocusable { focusable: true },
            PinWindowCall::SetForeground,
            PinWindowCall::SetFocus,
        ],
        "the window the press brought back is made focusable again — the collapse made it \
         non-activating — and is then brought to the front and given the caret"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
        close_pin_taskbar();
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(previous_carried, Ordering::Release);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    stand_pin(previous_pin);
}

/// A pin brought back by its key while the listing holds the keyboard leaves the keyboard where the
/// user is working: only a press on this app's own stand-in asks for the front and the caret.
///
/// The key is pressed over the listing the file is picked in, and the file is what it asks for — a
/// window that took the caret with it would take the listing the user was stepping through with it
/// (see `PinRestore::Key`).
#[test]
fn a_pin_brought_back_by_its_key_leaves_the_keyboard_in_the_listing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_carried = PIN_BUBBLE_CARRIED.load(Ordering::Acquire);
    let before = pin_always_on_top();

    // Safety: the windows are this test's own, made and destroyed under the one-at-a-time lock, and
    // the pin is stood up under it too.
    let (pin_hwnd, in_front) = unsafe { a_pin_put_away_into_its_bubble() };

    let window = RecordedPinWindow::new(pin_hwnd.0 as isize);
    restore_pin(&window, PinRestore::Key);

    assert!(!pin_is_collapsed(), "the key is what puts the pin back up");
    assert_eq!(
        window.calls(),
        Vec::new(),
        "a pin brought back by its key asks nothing of the window: the listing keeps the keyboard"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
        close_pin_taskbar();
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(previous_carried, Ordering::Release);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    stand_pin(previous_pin);
}

/// An update brings a pin that is put away into its bubble back up on the file the update names, and
/// the bubble goes with it — with `Pin Mode → Update Preview → Move To Top` on and `Always On Top` off.
///
/// **The fault this is a regression test for** is the one the taskbar's own test is about, reached from
/// the bubble: a put-away pin has no window in any band, so the row that says an update brings the pin
/// in front of the listing it follows answered nothing at all for it, and the file was swapped into the
/// pin the bubble was standing in for — the user's "even the minimized pin should show on top now" (see
/// `bring_a_put_away_pin_back`).
///
/// The file is what the pin comes back up on rather than something held for the key: a bubble holds a
/// file picked behind it until the user asks for the pin (see `pin_bubble_pick`), and this road is the
/// update asking for it — so the window comes back showing the file it was brought back for, in front of
/// the windows opened over it, and the keyboard stays in the listing the pick was made in.
///
/// **The pin here stands in the preview's window**, which is the kind a picture is not: a picture's
/// pin is a window of its own, and that is the kind the taskbar's test puts away. The two between
/// them are both kinds of pin behind both ways of being away, and the split is deliberate — the
/// raise this road makes after the restore is the part the show cannot do for the kind standing in
/// the preview's window, which is put up with `HWND_TOP` (see `bring_a_put_away_pin_back`).
#[test]
fn an_update_brings_a_pin_put_away_into_its_bubble_back_on_top() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_carried = PIN_BUBBLE_CARRIED.load(Ordering::Acquire);
    let previous_own = pin_is_its_own_window();
    let before = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);
    let updates = CONFIG
        .lock()
        .map(|config| config.pin_update_enabled)
        .unwrap_or(true);

    // Safety: the windows are this test's own, made and destroyed under the one-at-a-time lock, and
    // the pin is stood up under it too. The stand-in goes first, so that a pin put away into a bubble is
    // one the bubble is showing for rather than one a button of a test before it stands for.
    let (pin_hwnd, in_front) = unsafe { a_pin_put_away_into_its_bubble() };
    unsafe { close_pin_taskbar() };
    pin_stands_in_its_own_window(false);

    // The bubble the collapse leaves on the desktop: a window of this test's own standing in for it,
    // since the bubble is painted by the loop's own thread.
    let bubble = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            NOWHERE,
            NOWHERE,
            120,
            120,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the bubble")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen — a bubble that is up
    // is the whole of what says the pin was put away into it rather than into the taskbar.
    unsafe {
        let _ = ShowWindow(bubble, SW_SHOWNOACTIVATE);
    }
    PIN_BUBBLE_HWND.store(bubble.0 as isize, Ordering::SeqCst);
    assert!(
        showing_bubble_window().is_some(),
        "the pin is put away behind a bubble standing on the desktop, which is the state an update is          answered from"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_enabled = true;
        config.pin_update_move_to_top = true;
    }

    assert!(
        an_update_brings_a_put_away_pin_back(),
        "with `Move To Top` on, an update is the pin coming forward — and a pin that is away has          nothing to bring forward but itself"
    );
    let on_the_pin = pinned_path().expect("the file the pin is showing");
    assert!(
        bring_a_put_away_pin_back(&on_the_pin),
        "so the pin the update names a file for is the pin that comes back up"
    );

    assert!(!pin_is_collapsed(), "the pin is up again");
    assert!(
        unsafe { IsWindowVisible(pin_hwnd) }.as_bool(),
        "and the window it stands in is back on the screen"
    );
    assert!(
        showing_bubble_window().is_none(),
        "the bubble goes with it: a bubble and the window it stands for are never both on the screen"
    );
    // Safety: both windows are this test's own, and the z-order is the compositor's to settle.
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "the pin is put in front of the windows opened over it, which is the one thing an update owes          a pin it brings forward"
    );
    assert!(
        !unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "and no band is kept for it: `Pin Mode → Always On Top` is the only thing that keeps a pin over          other programs"
    );
    // Safety: asks this thread's own keyboard, which a window of this test's is the only holder of.
    assert_ne!(
        unsafe { GetFocus() },
        pin_hwnd,
        "and the keyboard stays in the listing the file was picked in: a pick is a step of a walk with          the arrows as often as it is a click, and a restore that took the focus would end that walk at          the first file"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(bubble);
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
        close_pin_taskbar();
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(previous_carried, Ordering::Release);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
        config.pin_update_enabled = updates;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A preview update raises the pin into the front it owes, without taking the keyboard; with
/// `Pin Mode → Update Preview → Move To Top` off the pin is left where the user's own clicks put it.
///
/// The raise is asked about rather than read off the desktop: the band it holds is the desktop's to
/// apply, and what the update owes is the decision to hold it (`pin_holds_the_front`). The placing
/// is read off the desktop's own list, where the press that gives the band back leaves the pin.
#[test]
fn a_preview_update_brings_a_pin_in_front_and_leaves_it_when_the_row_is_off() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // Safety: both windows are this test's own, opened where no display is. The second is made
    // after the pin's, which is what puts it in front — the listing the file was picked in.
    let (pin_hwnd, in_front) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let in_front = a_window_of_the_tests_own();
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
        (pin_hwnd, in_front)
    };

    // A pin that is up and on the screen, with `Always On Top` off: an ordinary window, which is
    // the one that can be standing behind the listing the user picks a file in.
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // An update is a pick's raise: the pin is put at the front of the ordinary band, and that is
    // all it is — with `Always On Top` off the band itself is not this app's to hold.
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pin_hwnd, false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "an update raises the pin over the window the file was picked in"
    );
    assert!(
        !pin_holds_the_front(),
        "and it holds no band for it: the topmost one is the setting's, and this pin's is off"
    );
    assert_ne!(
        unsafe { GetForegroundWindow() },
        pin_hwnd,
        "the pin is raised without taking the keyboard: the pick it answers may be one made with \
         the arrows in the listing, and a raise that took the focus would end that walk"
    );

    // The press that names no file gives the band back, and the placing it makes is what can be
    // read: the pin ends behind the window the press was made in.
    unsafe { lower_pinned_preview(pin_hwnd, in_front) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the front given back is the pin behind the listing again"
    );
    assert!(
        !pin_holds_the_front(),
        "and the pin no longer owes it: the next pick decides it again"
    );

    // The row switched off: the pin is left where the user's own clicks put it.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_move_to_top = false;
    }
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pin_hwnd, false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "with the row off a preview update is a file swapped into the pin where it stands, and the \
         window stays where the user's own clicks put it"
    );
    assert!(
        !pin_holds_the_front(),
        "with the row off there is no front to owe"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A press that names no file gives the front back: the pin a pick raised is put behind the window
/// the press was made in, which is what a click on the tree, the toolbar or the gap between the
/// rows is answered with (see `lower_pinned_preview`), and a pin that is Always On Top keeps the
/// band it is in.
///
/// The front is asked about rather than read off the desktop, for the reason the update test gives:
/// holding the band is a decision this app makes, and where the pin ends is the desktop's to place.
#[test]
fn a_press_that_picks_no_file_puts_the_pin_back_behind_the_listing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // Safety: both windows are this test's own, opened where no display is. The second is made
    // after the pin's, which is what puts it in front — the listing, in the arrangement the press
    // is read in.
    let (pin_hwnd, in_front) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let in_front = a_window_of_the_tests_own();
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
        (pin_hwnd, in_front)
    };

    // A pin that is up and on the screen, with `Always On Top` off.
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The pick arms the front, as a take-up or an update does, and the press that follows names no
    // file.
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pin_hwnd, false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a pick puts the pin at the front of the ordinary band, over the listing it is answering"
    );
    assert!(
        !pin_holds_the_front(),
        "and it holds no band for it: the topmost one is the setting's, and this pin's is off"
    );

    unsafe { lower_pinned_preview(pin_hwnd, in_front) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the pin goes back behind the window the press was made in, which is what a click that \
         picked no file is answered with"
    );
    // Safety: the lower put the pin back in the ordinary band, which is the half of it that is
    // about the pin's place and not about when the band is asked for.
    let lowered_style = unsafe { GetWindowLongPtrW(pin_hwnd, GWL_EXSTYLE) } as u32;
    assert_eq!(
        lowered_style & WS_EX_TOPMOST.0,
        0,
        "a pin put back is in the ordinary band, so the listing's next activation answers for it"
    );
    assert!(
        !pin_holds_the_front(),
        "and it no longer owes the front: the next press decides it again"
    );

    // The next pick brings it forward again — otherwise a press that named no file would be a pin
    // left behind for good.
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pin_hwnd, false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a later pick raises the pin again"
    );

    // A press that names no file gives the band back once more, which is the whole of the cycle a
    // hand makes: pick a file, work in the listing, pick a file again.
    unsafe { lower_pinned_preview(pin_hwnd, in_front) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the band given back is the pin back behind the listing"
    );
    assert!(!pin_holds_the_front(), "and the cycle can be made again");

    // `Always On Top` keeps the band it is in: a press that picks no file gives back nothing there.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    assert!(
        pin_holds_the_front(),
        "a pin that is Always On Top owes the front whether or not a press raised it"
    );
    PIN_FRONT_OWED.store(true, Ordering::SeqCst);
    unsafe { lower_pinned_preview(pin_hwnd, in_front) };
    assert!(
        PIN_FRONT_OWED.load(Ordering::SeqCst),
        "a press that names no file is not the listing's answer while Always On Top is on"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A film goes under its pin, and holds the band only where there is no pin to go under it.
///
/// **The failure this pins down is a film that keeps the band its pin just gave back.** With
/// `Always On Top` off, a press that named no file puts the pin behind the listing — and the film
/// is a window of somebody else's, placed by this app again on the next tick whether or not
/// anything asked. Asked whether the pin *holds the front*, that placement put the film back in the
/// topmost band, so the pin went behind the listing while its picture stayed over every program on
/// the desktop: the pin hid, the film did not.
///
/// What asks for the band is the pin's existence and its setting: a hover's film has nothing of
/// this app's standing over it, and a pin that is Always On Top holds the band it always held. A pin
/// in the ordinary band owes its film the place directly below itself — holding the front or having
/// given it back, which is the pin's own business and moves no window of somebody else's.
#[test]
fn a_film_goes_under_its_pin_and_holds_the_band_only_with_none() {
    use windows::Win32::UI::WindowsAndMessaging::{HWND_NOTOPMOST, HWND_TOPMOST};

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    let pin_hwnd = unsafe { a_window_of_the_tests_own() };
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // No pin at all: a hover's film is the topmost band and nothing of this app's stands over it.
    stand_pin(None);
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    assert_eq!(
        player_insert_after(),
        HWND_TOPMOST,
        "a hover's film holds the band: there is no pin standing over it, and an ordinary window \
         activated over the film is a picture with the desktop showing through the band"
    );

    // The pin up with the front given back, which is what a press that named no file leaves.
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    assert!(
        !pin_holds_the_front(),
        "the front is the press's to give back, and this is the state it leaves"
    );
    assert_eq!(
        player_insert_after(),
        pin_hwnd,
        "a pin that has given the front back still stands over its film: the place below it is \
         what takes the film out of the topmost band, and without it the picture of a pin behind \
         the listing covers every program on the desktop"
    );

    // A pin that owes no band still stands over its film: the place below the pin is what takes
    // the film out of the topmost band, whether the band above it is the pin's or nobody's.
    assert_eq!(
        player_insert_after(),
        pin_hwnd,
        "the film goes under the pin in either band: the caption and the transport bar are drawn \
         into the band over it, and the press on the picture is the pin's"
    );

    // A pin with no window of its own yet has nothing for the film to go under.
    PREVIEW_HWND.store(0, Ordering::SeqCst);
    assert_eq!(
        player_insert_after(),
        HWND_NOTOPMOST,
        "a film with no pin to go under is put in the ordinary band rather than over the desktop"
    );
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);

    // `Always On Top` keeps the band it kept before any of this.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    assert_eq!(
        player_insert_after(),
        HWND_TOPMOST,
        "a pin that is Always On Top holds its band whatever any press said, and its film holds \
         the band under it"
    );

    // Safety: the window is this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A window of the player is a window of this preview, though its process is not this app's.
///
/// **The failure this pins down is the pin going behind its own film.** The foreground watch asks
/// whether the window in front is one of ours to decide whether the user has gone to work
/// somewhere else, and it asked it of the process id. A film is drawn by ffplay or mpv, so the very
/// window the pin's media band is made of answered "somebody else's" — and the pin was put behind
/// it the moment that window came up, which is the preview going out of the topmost band with a
/// picture left over it: a pin taken up on a video appearing behind Explorer, and a press that
/// picked a file blinking under the listing while the new film's window was being raised.
#[test]
fn a_window_of_the_player_is_a_window_of_the_preview() {
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);

    // No window's process is "nobody's": a pid of nothing is what a failed ask leaves, and it must
    // never read as the preview's own.
    assert!(!pid_is_the_previews(0), "a pid of nothing is no answer");

    assert!(
        pid_is_the_previews(std::process::id()),
        "this app's own windows are the preview being used, which is what the watch is about"
    );

    let player = 4242;
    VIDEO_PID.store(player, Ordering::SeqCst);
    assert!(
        pid_is_the_previews(player),
        "and so is the player's window: the film is the pin's media band, not a program the user \
         went to work in"
    );
    assert!(
        !pid_is_the_previews(player + 1),
        "while another program in front still is somebody else's, which is the whole of what the \
         watch demotes the pin for"
    );

    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// A pick's raise is a pin owing the front, and the grace that follows it answers while the front
/// is still owed: the press that names no file gives the band back, and the grace is dropped rather
/// than putting the pin back over the listing that press just took it behind.
#[test]
fn a_pin_raised_for_a_pick_answers_for_the_front_while_the_grace_lasts() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // Safety: both windows are this test's own, opened where no display is. The second is made
    // after the pin's, which is what puts it in front — the listing, in the arrangement the press
    // is read in.
    let (pin_hwnd, in_front) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let in_front = a_window_of_the_tests_own();
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
        (pin_hwnd, in_front)
    };

    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // A pick raises the pin, and with `Always On Top` off that raise is the whole of it: there is
    // no band of this app's for the listing's activation to be kept out of, so there is nothing for
    // a grace to keep.
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pin_hwnd, false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a pick's raise puts the pin at the front of the ordinary band"
    );
    assert!(
        !pin_holds_the_front(),
        "and it holds no band for it: the topmost one is the setting's, and this pin's is off"
    );

    // **The click that made the pick activates the listing after the raise.** That is the listing
    // being put back over the pin, and the grace is the pin's answer for as long as that can still
    // happen - with `Always On Top` off as much as on, because the debt a pick arms is what the grace
    // asks and the band is the setting's (see `keep_pin_at_the_front`). Asked of the band, the grace
    // ended on its first tick and every pick made after the user had clicked anything was a pin left
    // behind the listing.
    //
    // The window the user is in at this moment is still the pin: the press is read off the hook before
    // the click reaches the listing, so the activation is still to land and there is nothing yet that
    // says the pick's answer has arrived.
    let mut owed = Some(PinFrontOwed {
        until: Instant::now() + PIN_RAISE_GRACE,
        band_until: Instant::now() + PIN_RAISE_GRACE,
        listing: in_front,
    });
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, Instant::now(), pin_hwnd) };
    assert!(
        owed.is_some(),
        "a pick's grace keeps answering while the debt stands, with `Always On Top` off"
    );

    // What the click does: the listing it was made in is activated, which is the pin being put behind
    // it. The grace answers that with the raise, so the pin the user just picked a file for is the one
    // they are looking at.
    unsafe {
        let _ = SetWindowPos(
            pin_hwnd,
            in_front,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the activation the click ends in is the listing over the pin"
    );
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, Instant::now(), in_front) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "the grace raises the pin back over the listing the click activated"
    );

    // The press that names no file is the pin giving the band back. The grace it began goes with
    // it: asked again, it is dropped rather than raising the pin over the listing.
    unsafe { lower_pinned_preview(pin_hwnd, in_front) };
    unsafe { settle_z_order() };
    assert!(
        !pin_holds_the_front(),
        "a press that names no file gives the front back"
    );
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "and leaves the pin behind the listing"
    );
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, Instant::now(), in_front) };
    assert!(
        owed.is_none(),
        "a grace a press has ended is dropped rather than asked about on every tick after it"
    );
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the grace does not put the pin back over the listing the press took it behind"
    );

    // A grace that has run out is dropped the same way, so a pick is not answered for forever.
    arm_pin_to_the_front();
    let mut owed = Some(PinFrontOwed {
        until: Instant::now(),
        band_until: Instant::now(),
        listing: in_front,
    });
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, Instant::now(), in_front) };
    assert!(owed.is_none(), "a grace that has run out is dropped");

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A pick's raise is answered again when the listing's own activation lands after it: the answer the
/// stand-in kinds already had, and the one a pin that is a window of its own was left without.
///
/// **The fault this is a regression test for.** The click that picks a file reaches this app before it
/// reaches the listing: the press is read off the mouse hook and posted, the pick's raise takes the
/// band and gives it straight back, and only then is the window the click landed in activated — which
/// is that window being put over the pin again, a moment after the pin was put over it. The raise is
/// one answer made once, and the grace a pick begins is the answer to the activation that follows it;
/// it was asked of the stand-in kinds only (see
/// `a_pin_raised_for_a_pick_answers_for_the_front_while_the_grace_lasts`), so a picture pin stayed
/// behind the listing it had just followed, and every file picked after that was picked out of sight
/// of the pin.
///
/// What the grace asks is where the two windows stand, and its raise is the band taken and given back
/// rather than held, so a pin with `Always On Top` off is still left behind by the window the user's
/// next click puts in front (see `keep_pin_at_the_front` and `take_the_band_and_give_it_back`).
#[test]
fn a_pick_chases_the_listing_when_the_activation_lands_after_its_raise() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin's own window and the listing it follows: two windows of the test's own, opened where no
    // display is, and the pin standing in the first rather than in the loop's.
    let (pin_hwnd, listing) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let listing = a_window_of_the_tests_own();
        PREVIEW_HWND.store(listing.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(listing, SW_SHOWNOACTIVATE);
        (pin_hwnd, listing)
    };
    pin_real_window::stand_real_pin_window_for_a_test(Some(pin_hwnd));
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    pin_stands_in_its_own_window(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The pick's answer, made once: the pin comes over the listing it answers.
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pinned_window(), false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, listing) },
        "a pick's raise puts the pin over the listing it answers"
    );

    // And then the click that made the pick reaches the listing, which is activated: the window the
    // user is working in is put over the pin again. `Always On Top` off means no band is held that
    // could keep it out, which is the setting meaning what it says.
    unsafe {
        let _ = SetWindowPos(
            pin_hwnd,
            listing,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(listing, pin_hwnd) },
        "the activation the click ends in is the listing over the pin"
    );

    // The grace the pick began is the answer to it, and the pin the user has just picked a file for is
    // the window they are looking at, whichever window the pin stands in.
    let mut owed = Some(PinFrontOwed {
        until: Instant::now() + PIN_RAISE_GRACE,
        band_until: Instant::now() + PIN_RAISE_GRACE,
        listing,
    });
    unsafe { keep_pin_at_the_front(pinned_window(), &mut owed, Instant::now(), listing) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, listing) },
        "the grace brings the pin back over the listing the click activated"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(listing);
        let _ = DestroyWindow(pin_hwnd);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A pick's press holds the pin over the listing until the click that made it has activated it: the
/// activation is never a moment the listing stands over the pin, which is the flash the user reported
/// between a pick and the answer for it.
///
/// **The fault this is a regression test for.** The click reaches this app before it reaches the
/// listing, so the raise a press makes lands while the listing is still where it was — and the
/// activation that follows is the listing put in front of the ordinary band, which is over the pin.
/// The grace answers that a tick later, and a tick is long enough for the user to see the listing
/// through the pin: `STATIC_PIN_WAIT_MS` of it, plus the paint the pick's own file costs before the
/// tick that answers gets to ask. A window in the topmost band is not one an activation can be put
/// over, so the band is held for the length of that one click and given back the moment the listing is
/// the window the user is in (see `raise_pinned_preview` and `keep_pin_at_the_front`).
#[test]
fn a_pick_holds_the_pin_over_the_listing_until_the_click_activates_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let previous_band = PIN_BAND_HELD.load(Ordering::SeqCst);
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin in a window of its own, and the listing it follows: two windows of this test's own,
    // opened where no display is.
    let (pin_hwnd, listing) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let listing = a_window_of_the_tests_own();
        PREVIEW_HWND.store(listing.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(listing, SW_SHOWNOACTIVATE);
        (pin_hwnd, listing)
    };
    pin_real_window::stand_real_pin_window_for_a_test(Some(pin_hwnd));
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    pin_stands_in_its_own_window(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The press, read off the hook before the click reaches the listing: the pin is raised for the pick
    // it names, and the band is kept rather than given straight back.
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    PIN_BAND_HELD.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pinned_window(), true) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "the press holds the band until the click it answers has landed"
    );

    // And the click lands: the listing is activated, which is that window being put in front of the
    // ordinary band. The pin is not in it, so nothing the activation can do puts the listing over the
    // pin — there is no moment the user sees it through.
    unsafe {
        let _ = SetWindowPos(
            listing,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, listing) },
        "the pin the user picked a file for is never a window the listing's activation covers"
    );

    // **The click landing is not the band's signal to go.** Windows sets the foreground before it raises
    // the window it names, so a tick that gave the band back because the listing is now the window the user
    // is in would hand the pin to a raise still on its way: the listing over the pin, for a tick, which is
    // the flash. What the band's own deadline is measured from is the press.
    let pressed = Instant::now();
    let mut owed = Some(PinFrontOwed {
        until: pressed + PIN_RAISE_GRACE,
        band_until: pressed + PIN_RAISE_GRACE,
        listing,
    });
    unsafe {
        keep_pin_at_the_front(
            pin_hwnd,
            &mut owed,
            pressed + Duration::from_millis(1),
            listing,
        )
    };
    unsafe { settle_z_order() };
    assert!(
        owed.is_some(),
        "the pick is still owed while the press's grace has time on it"
    );
    assert!(
        unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "and the band is not given back on the activation's word: the foreground is set before the window it names is raised"
    );

    // And on the press's clock the band goes back, with the pin over the listing it answered for: the front
    // of the ordinary band is the pin landing on the file, not the pin put behind the window the user went on
    // working in.
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, pressed + PIN_RAISE_GRACE, listing) };
    unsafe { settle_z_order() };
    assert!(
        owed.is_none(),
        "the front the band was held for is answered, and the debt goes with it"
    );
    assert!(
        !unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "the band is the setting's again: with `Always On Top` off the next click puts the pin behind it"
    );
    assert!(
        unsafe { stands_over(pin_hwnd, listing) },
        "and giving the band back leaves the pin the window over the listing it followed"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(listing);
        let _ = DestroyWindow(pin_hwnd);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    PIN_BAND_HELD.store(previous_band, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A band held for a pick goes back behind the window the user moved to. The grace answers one click,
/// and a pick whose activation never lands where it was expected — the user gone to another program, a
/// listing closed under the click, a press that named no file — is not a licence to leave the pin over
/// the program they went to, which with `Always On Top` off is the one thing that setting forbids.
///
/// **The band is given back where the user is, not where the pin was raised.** `HWND_NOTOPMOST` alone
/// would put the pin at the front of the ordinary band — over the window the user just moved to, which
/// is the same fault the hold was a fix for, one program along. So the release asks where the user is
/// and puts the pin below that window, which is what every other press of theirs would have done.
#[test]
fn a_band_held_for_a_pick_goes_back_behind_the_window_the_user_moved_to() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let previous_band = PIN_BAND_HELD.load(Ordering::SeqCst);
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin in a window of its own, the listing its click was made in, and the program the user went
    // to instead: three windows of this test's own, opened where no display is.
    let (pin_hwnd, listing, moved_to) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let listing = a_window_of_the_tests_own();
        let moved_to = a_window_of_the_tests_own();
        PREVIEW_HWND.store(listing.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(listing, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(moved_to, SW_SHOWNOACTIVATE);
        (pin_hwnd, listing, moved_to)
    };
    pin_real_window::stand_real_pin_window_for_a_test(Some(pin_hwnd));
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    pin_stands_in_its_own_window(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The press holds the band, and the user is then in a program of their own — the one the pin's
    // hold put it over, and the one it must not be left over.
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    PIN_BAND_HELD.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pinned_window(), true) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "the press holds the band until the click it answers has landed"
    );

    // And the grace runs out with the user in another program: the band goes, and the pin goes back
    // behind the window they are in rather than over it.
    let mut owed = Some(PinFrontOwed {
        until: Instant::now(),
        band_until: Instant::now(),
        listing,
    });
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, Instant::now(), moved_to) };
    unsafe { settle_z_order() };
    assert!(
        owed.is_none(),
        "a grace that has run out is dropped, whatever held the band"
    );
    assert!(
        !unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "a band held for a pick is never the band the pin keeps"
    );
    assert!(
        unsafe { stands_over(moved_to, pin_hwnd) },
        "and the pin goes behind the program the user moved to, not over it"
    );

    // Safety: all three windows are this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(moved_to);
        let _ = DestroyWindow(listing);
        let _ = DestroyWindow(pin_hwnd);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    PIN_BAND_HELD.store(previous_band, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A pick's own placement keeps the band the press held.
///
/// **The fault this is a regression test for.** Between the press and the activation that click ends in,
/// the pick is installed: the file the press named is read, decoded and shown in the window, and the show
/// places that window the way every placement asks — `band_place()`. With the band held for the press and
/// the placement's answer written for a pin that owes nothing, every pick put its own window back in the
/// ordinary band a few milliseconds *before* the listing was activated over it, so the band was given back
/// before the click that took it had landed and the user saw Explorer over the pin for a tick — on a
/// click, and not on the arrows, which walk the same files without a press to hold anything.
#[test]
fn a_picks_own_placement_keeps_the_band_the_press_held() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let previous_band = PIN_BAND_HELD.load(Ordering::SeqCst);
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin in a window of its own, and the listing the click that picked the file was made in.
    let (pin_hwnd, listing) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let listing = a_window_of_the_tests_own();
        PREVIEW_HWND.store(listing.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(listing, SW_SHOWNOACTIVATE);
        (pin_hwnd, listing)
    };
    pin_real_window::stand_real_pin_window_for_a_test(Some(pin_hwnd));
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    pin_stands_in_its_own_window(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The press, read before the click reaches the listing: the band is held for it.
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    PIN_BAND_HELD.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pinned_window(), true) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "the press holds the band until the click it answers has landed"
    );

    // The pick is installed while the click is still on its way: the window is placed the one way every
    // placement is asked, which is the question a show puts to the band.
    unsafe {
        let _ = SetWindowPos(
            pin_hwnd,
            band_place(),
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "a pick's own placement is no band given back: what holds the band outlives the show"
    );

    // And the click lands after it: the listing is activated, and nothing of it can be put over the pin.
    unsafe {
        let _ = SetWindowPos(
            listing,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, listing) },
        "the pin the user picked a file for is never a window the listing's activation covers"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(listing);
        let _ = DestroyWindow(pin_hwnd);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    PIN_BAND_HELD.store(previous_band, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A band the desktop did not give is no band to give back, and the pick's front is still owed.
///
/// **The road the take can fail on.** The desktop answers a band assignment as made and can still leave the
/// window where it was (see `raise_pinned_preview`, which asks again from outside the band for that reason),
/// and a hold the desktop has not taken is no hold: the pick's own placement is a placement in the ordinary
/// band, the activation lands over the pin, and the release on the press's clock finds nothing to give back.
/// Dropping the answer there would be a pin left behind the listing it followed for the rest of the walk —
/// the fault the whole grace exists for, on the one road the take is not trusted. So the answer stays with
/// the pick's own clock, and it is answered the way a stand-in's is: the pin raised over the listing, a take
/// and a give back, no band kept.
#[test]
fn a_pick_whose_band_the_desktop_did_not_give_is_still_owed_the_front() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let previous_band = PIN_BAND_HELD.load(Ordering::SeqCst);
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin in a window of its own and the listing its click was made in, both in the ordinary band: the
    // hold is owed but never taken, which is the take that was answered and not made.
    let (pin_hwnd, listing) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let listing = a_window_of_the_tests_own();
        PREVIEW_HWND.store(listing.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(listing, SW_SHOWNOACTIVATE);
        let _ = SetWindowPos(
            pin_hwnd,
            HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        (pin_hwnd, listing)
    };
    pin_real_window::stand_real_pin_window_for_a_test(Some(pin_hwnd));
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    pin_stands_in_its_own_window(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }
    assert!(
        unsafe { !window_holds_the_topmost_band(pin_hwnd) },
        "the window the band was asked of and not given is no window in the topmost band"
    );

    // The click lands while the band is not the desktop's: the listing is activated, which is that window
    // put over the pin — the flash, on a take that did not take. The press's clock then runs out.
    unsafe {
        let _ = SetWindowPos(
            listing,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        settle_z_order();
    }
    assert!(
        unsafe { stands_over(listing, pin_hwnd) },
        "the activation the band was meant to answer covers the pin it was not given for"
    );

    // A file that reached the window late, which is the only road the press's clock runs out first: the
    // answer has been moved on (see `extend_pin_front_owed`) and the band's hold has not.
    let pressed = Instant::now();
    let mut owed = Some(PinFrontOwed {
        until: pressed + PIN_RAISE_GRACE + PIN_RAISE_GRACE,
        band_until: pressed + PIN_RAISE_GRACE,
        listing,
    });
    arm_pin_to_the_front();
    PIN_BAND_HELD.store(true, Ordering::SeqCst);
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, pressed + PIN_RAISE_GRACE, listing) };
    unsafe { settle_z_order() };
    assert!(
        owed.is_some(),
        "a band that was never the desktop's to give back leaves the pick's front still owed"
    );
    assert!(
        unsafe { stands_over(pin_hwnd, listing) },
        "and the answer is the pin over the file the user picked, not the pin left behind the listing"
    );
    assert!(
        unsafe { !window_holds_the_topmost_band(pin_hwnd) },
        "while the answer is owed the chase keeps no band: this is the stand-in kinds' road, taken and given back"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(listing);
        let _ = DestroyWindow(pin_hwnd);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    PIN_BAND_HELD.store(previous_band, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A band held for a pick is given back on the press's own clock, however long the file it named takes
/// to arrive: the hold is the answer to one click, and a pick whose click never brings the listing to the
/// front — the folder closed under it, the frame the press was read in never activated — is not a licence
/// to keep the pin over every program the user works in for as long as files keep being stepped through.
///
/// **The grace is extended and the hold is not.** An update that reaches the window late is a file the
/// pin must still come back on, so the debt lasts another `PIN_RAISE_GRACE` from it (see
/// `extend_pin_front_owed`); the band the press took is not the debt's, though, and extending it would
/// make the longest the pin stands over the user's programs as long as their hand keeps moving.
#[test]
fn a_band_held_for_a_pick_is_given_back_on_the_presss_own_clock() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let previous_band = PIN_BAND_HELD.load(Ordering::SeqCst);
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin in a window of its own, the listing its click was made in, and the program the user is in
    // while the file it named is still being read.
    let (pin_hwnd, listing, elsewhere) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let listing = a_window_of_the_tests_own();
        let elsewhere = a_window_of_the_tests_own();
        PREVIEW_HWND.store(listing.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(listing, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(elsewhere, SW_SHOWNOACTIVATE);
        (pin_hwnd, listing, elsewhere)
    };
    pin_real_window::stand_real_pin_window_for_a_test(Some(pin_hwnd));
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    pin_stands_in_its_own_window(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    let pressed = Instant::now();
    PIN_FRONT_OWED.store(false, Ordering::SeqCst);
    PIN_BAND_HELD.store(false, Ordering::SeqCst);
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pinned_window(), true) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "the press holds the band until the click it answers has landed"
    );

    // The file the pick named takes longer to read than the grace the press began, so the debt is
    // extended — and the hold is not: the tick the press's own grace runs out on is a tick that gives the
    // band back to the program the user is in, whatever the file still costs.
    let mut owed = Some(PinFrontOwed {
        until: pressed + PIN_RAISE_GRACE,
        band_until: pressed + PIN_RAISE_GRACE,
        listing,
    });
    extend_pin_front_owed(&mut owed);
    let owed = owed.expect("a pick's grace is still owed for the file on its way");
    let still_owed = owed.until;
    unsafe {
        let _ = SetWindowPos(
            elsewhere,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    unsafe {
        keep_pin_at_the_front(
            pin_hwnd,
            &mut Some(owed),
            pressed + PIN_RAISE_GRACE + Duration::from_millis(1),
            elsewhere,
        )
    };
    unsafe { settle_z_order() };
    assert!(
        !unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "a hold is measured from the press, not from the last update a file cost"
    );
    assert!(
        still_owed > pressed + PIN_RAISE_GRACE,
        "while the answer itself was owed for longer than the press's own grace"
    );
    assert!(
        unsafe { stands_over(elsewhere, pin_hwnd) },
        "and the band going back is the pin behind the program the user went to, not over it"
    );

    // Safety: all three windows are this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(elsewhere);
        let _ = DestroyWindow(listing);
        let _ = DestroyWindow(pin_hwnd);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    PIN_BAND_HELD.store(previous_band, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A pick's grace answers the listing its click was made in, and no other window: a second listing
/// put over the pin is the user having gone to another folder, not the answer this grace owes, and a
/// pin raised over the window they chose is a pin with `Always On Top` off standing over it — the one
/// thing that setting must never do (see `PinFrontOwed`).
#[test]
fn a_pick_answers_only_the_listing_its_click_was_made_in() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();

    // The pin, the listing the pick was made in, and the second listing the user moved to inside the
    // grace: three windows of this test's own, opened where no display is.
    let (pin_hwnd, picked_in, moved_to) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let picked_in = a_window_of_the_tests_own();
        let moved_to = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(picked_in, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(moved_to, SW_SHOWNOACTIVATE);
        (pin_hwnd, picked_in, moved_to)
    };
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The pin stands over the listing the pick was made in, and the second listing comes over the pin:
    // a click into it is that window being activated, and it is not a window the pin follows. Placed
    // rather than activated, because `SetForegroundWindow` is another program's to do to this one.
    unsafe {
        let _ = SetWindowPos(
            picked_in,
            pin_hwnd,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        let _ = SetWindowPos(
            pin_hwnd,
            moved_to,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(moved_to, pin_hwnd) },
        "another listing is over the pin, which is where a click into it puts it"
    );
    assert!(
        unsafe { stands_over(pin_hwnd, picked_in) },
        "with the pin still standing over the listing the pick was made in"
    );

    // The grace asks about that listing, which the pin stands over: nothing is raised, so the window
    // the user went to keeps the front its own click gave it — and the front the pick owed is still
    // owed, for the listing the pick was made in.
    arm_pin_to_the_front();
    let mut owed = Some(PinFrontOwed {
        until: Instant::now() + PIN_RAISE_GRACE,
        band_until: Instant::now() + PIN_RAISE_GRACE,
        listing: picked_in,
    });
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, Instant::now(), moved_to) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(moved_to, pin_hwnd) },
        "the grace does not raise the pin over a window the pick's click was never made in"
    );
    assert!(
        owed.is_some(),
        "and the front the pick owes is still owed for the listing it answers"
    );

    // Safety: every window is this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(moved_to);
        let _ = DestroyWindow(picked_in);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
    }
    stand_pin(previous_pin);
}

/// A pick whose listing is closed inside its grace is answered for nothing: the window the click was
/// made in is gone, so there is nothing left for the raise to answer — and asking the z-order of a
/// handle that is no longer a window would have the grace raise the pin over whatever the user had
/// moved to instead, which is the one thing `Always On Top` off must never do (see `PinFrontOwed`).
#[test]
fn a_pick_whose_listing_is_gone_answers_for_nothing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin, the listing the press was made in, and the window the user moved to after closing it.
    let (pin_hwnd, closed_listing, moved_to) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        let closed_listing = a_window_of_the_tests_own();
        let moved_to = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(closed_listing, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(moved_to, SW_SHOWNOACTIVATE);
        let _ = DestroyWindow(closed_listing);
        (pin_hwnd, closed_listing, moved_to)
    };
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The user is in the window they moved to, which stands over the pin.
    unsafe {
        let _ = SetWindowPos(
            pin_hwnd,
            moved_to,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(moved_to, pin_hwnd) },
        "the window the user moved to stands over the pin"
    );

    // The grace the press began still carries the listing, which is no longer a window: nothing is
    // raised, and the debt goes with the window rather than standing until the grace runs out.
    arm_pin_to_the_front();
    let mut owed = Some(PinFrontOwed {
        until: Instant::now() + PIN_RAISE_GRACE,
        band_until: Instant::now() + PIN_RAISE_GRACE,
        listing: closed_listing,
    });
    unsafe { keep_pin_at_the_front(pin_hwnd, &mut owed, Instant::now(), moved_to) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(moved_to, pin_hwnd) },
        "the grace raises nothing for a listing that is no longer a window"
    );
    assert!(
        owed.is_none(),
        "and the front a closed listing was owed goes with it"
    );

    // Safety: every window is this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(moved_to);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A press that names no file puts a pin that is a window of its own behind the window it was made in,
/// and the front an update owes that pin is still asked for the one way Windows allows.
///
/// **The lowering is the app's, because for this pin Windows cannot do it.** The Shell puts an ordinary
/// window behind the one a click activates, but a click in the window that is *already* active reorders
/// nothing: Explorer is the foreground window while the user works in it, so a press on its toolbar, its
/// tree or the gap between its rows left the pin standing over the listing it follows — the fault the
/// user reported, beside every other program covering the pin. A press that names no file is the listing
/// being worked in rather than a file being picked, and the pin goes behind the window that press was
/// made in (see `PreviewMessage::PinNoPick`). The demotion is recoverable rather than final: the pin's
/// own window is one a click may activate, so a click on the pin brings it back (see `pin_set_focusable`).
/// A lowering this window cannot answer is the desktop's, which the tick hands in as no window: there is
/// no place to go behind, and the bottom of the desktop's list is below the wallpaper.
///
/// **The raise is the rest of what is left of the stand-in kinds**, because `HWND_TOP` is refused to a
/// process that is not in the foreground: the band is taken for the moment the pin comes forward over
/// the listing, and given straight back. With `Always On Top` on the band is the setting's rather than
/// either road's, and with `Update Preview → Move To Top` off the pin is left where the user's own
/// clicks put it.
#[test]
fn a_pin_that_is_a_window_of_its_own_goes_behind_the_press_and_comes_forward_on_an_update() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let before = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // The pin's own window and the window this loop holds are two different windows, which is the
    // arrangement every forced raise got wrong (see `pinned_window`).
    let (in_front, pin_hwnd) = unsafe {
        let in_front = a_window_of_the_tests_own();
        let pin_hwnd = a_window_of_the_tests_own();
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        (in_front, pin_hwnd)
    };
    PREVIEW_HWND.store(in_front.0 as isize, Ordering::SeqCst);
    pin_real_window::stand_real_pin_window_for_a_test(Some(pin_hwnd));

    stand_pin(Some(overlay_pin((100, 100, 420, 340), PinChrome::always())));
    pin_stands_in_its_own_window(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "the window made after the other stands over it"
    );

    // What a pick leaves: the pin at the front of the ordinary band, over the listing it follows. The
    // raise is the road below; this is the arrangement the press that answers it is read in.
    arm_pin_to_the_front();
    unsafe { raise_pinned_preview(pinned_window(), false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a pick leaves the pin over the listing it answers"
    );

    // With no window to go behind — the desktop, which the tick hands in as no window — nothing is
    // asked: the bottom of the desktop's own list is below the wallpaper, and a pin sent there is one
    // nobody can click back. The front a pick owed ends with it, as it does for the stand-in kinds: a
    // press that names no file is the listing's answer whether or not it can say where the listing is.
    assert!(
        PIN_FRONT_OWED.load(Ordering::SeqCst),
        "the pick armed the front this press answers"
    );
    unsafe { lower_pinned_preview(pinned_window(), HWND(ptr::null_mut())) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a press with no window to go behind leaves the pin where it is"
    );
    assert!(
        !PIN_FRONT_OWED.load(Ordering::SeqCst),
        "and the front it owed is given up: nothing is left ticking to raise it back"
    );

    // The press that names no file — the toolbar, the tree, the gap between the rows: the pin goes
    // behind the window the press was made in, which is the listing the user is working in. The handle
    // handed in is the pin's own, which for a picture is not the loop's: a lowering given the loop's
    // window moves a window nobody can see (see `pinned_window`).
    arm_pin_to_the_front();
    unsafe { lower_pinned_preview(pinned_window(), in_front) };
    unsafe { settle_z_order() };
    assert_eq!(
        unsafe { GetWindow(pin_hwnd, GW_HWNDPREV) }.ok(),
        Some(in_front),
        "a press that names no file puts the pin directly behind the window it was made in"
    );
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "which is the listing over it, not the bottom of anybody's list"
    );
    assert!(
        !PIN_FRONT_OWED.load(Ordering::SeqCst),
        "and the front a pick owed is given back with it"
    );
    // Safety: the style is read off this test's own window rather than changed. The bit is the one the
    // take-down of `8a537f9` must never leave on a pin's own window, whatever road runs (`WS_EX_NOACTIVATE`
    // is the hover's, whose press is not the pin's), and this is the road that now runs on every press
    // that names no file.
    assert_eq!(
        unsafe { GetWindowLongPtrW(pin_hwnd, GWL_EXSTYLE) } as u32 & WS_EX_NOACTIVATE.0,
        0,
        "and it stays one a click may activate: the press that put it behind is one a click on the \
         pin takes it back out of"
    );

    // And the raise: with `Always On Top` on, the raise a fake pin is owed is the topmost band — a band
    // a real pin is given by the setting when it is shown, and not taken for it here.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    assert!(!unsafe { window_holds_the_topmost_band(pin_hwnd) });
    unsafe { raise_pinned_preview(pinned_window(), false) };
    unsafe { settle_z_order() };
    assert!(
        !unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "the band is the setting's, put on when the pin is shown, and not taken by a raise"
    );

    // An update, with `Always On Top` off and `Move To Top` on: the pin is brought over the window it
    // follows, which is the one thing Windows will not do for a window of a process that is not in the
    // foreground (see `raise_pinned_preview`). The listing is put over the pin first, since that is
    // where a pick leaves it.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }
    unsafe {
        let _ = SetWindowPos(
            pin_hwnd,
            in_front,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the listing is over the pin before the update answers it"
    );
    unsafe { raise_pinned_preview(pinned_window(), false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "an update brings a pin that is a window of its own over the listing it was behind"
    );
    assert!(
        !unsafe { window_holds_the_topmost_band(pin_hwnd) },
        "and the band is given straight back: with `Always On Top` off the next click on another window \
         puts the pin behind it, exactly as any window behaves"
    );

    // And with the row off the pin is left where the user's own clicks put it.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_move_to_top = false;
    }
    unsafe {
        let _ = SetWindowPos(
            pin_hwnd,
            in_front,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    unsafe { settle_z_order() };
    unsafe { raise_pinned_preview(pinned_window(), false) };
    unsafe { settle_z_order() };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "with `Move To Top` off an update is a file swapped into the pin where it stands"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
        let _ = DestroyWindow(in_front);
    }
    pin_real_window::stand_real_pin_window_for_a_test(None);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    pin_stands_in_its_own_window(false);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// The Shell's Show Desktop — `Win`+`D`, or the taskbar's button at the end of the bar — minimizes
/// every window that offers a minimize box, and the pin's own window is one of them. The pin goes
/// down with the desktop and comes back up with it, and Pin Mode is still on.
///
/// **The fault this is a regression test for.** Show Desktop ended Pin Mode. The window procedure
/// answered every `WM_SYSCOMMAND` as a close, and the Shell's minimize of a window arrives as one of
/// those (`SC_MINIMIZE`): the gesture meant to clear the screen took the pin away with it, and the
/// pin the user had put up had to be put up again by hand. A `Win`+`D` never reaches an app as a key
/// and no message of it names the desktop, so what is answered here is the only road the gesture has
/// to this pin.
///
/// Only `SC_CLOSE` is a close: the window's own button and the button's menu row, both of them the
/// Shell's ask rather than a key's — `Alt+F4` never gets as far as a message, because a system
/// keystroke on this window is swallowed where the pin's own keys are read. Everything else a system
/// command can be is the desktop's business with the window — `SC_MINIMIZE` and `SC_RESTORE` above
/// all — and it belongs to `DefWindowProcW`, the road every other window in the system takes for it.
/// A pin of another kind stands in a window of `pin_taskbar`'s, whose own procedure answers those
/// two commands because that window is not this one: there is no pin there to minimize, and the pin
/// is put away into its bubble instead (see `handle_taskbar_restore`).
#[test]
fn the_shells_show_desktop_minimizes_a_pin_that_is_a_window_of_its_own() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_own = pin_is_its_own_window();
    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_end = PIN_END_REQUESTED.swap(false, Ordering::SeqCst);

    // The box the pin says it stands in, where no display is, and the window the pin's own window is
    // made from for it: what answers the Shell's command here is the class and procedure the app's
    // own window answers with rather than a stand-in's.
    let box_ = (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240);
    stand_pin(Some(overlay_pin(box_, PinChrome::always())));
    pin_stands_in_its_own_window(true);
    let pin_hwnd = unsafe {
        pin_real_window::stand_real_pin_window_for_a_test(None);
        pin_real_window::real_pin_window_at(box_)
    }
    .expect("the pin's own window");

    // What Show Desktop sends every window with a minimize box, and what it therefore sends this
    // one: the Shell is the sender, and the command is the window's own.
    unsafe {
        let _ = SendMessageW(
            pin_hwnd,
            WM_SYSCOMMAND,
            WPARAM(SC_MINIMIZE as usize),
            LPARAM(0),
        );
    }

    assert!(
        !PIN_END_REQUESTED.load(Ordering::SeqCst),
        "Show Desktop minimizes a window rather than closing it, so the pin is not asked to end"
    );
    assert!(
        pinned(),
        "the pin stands with its window down with the rest of the desktop"
    );
    assert!(
        unsafe { IsIconic(pin_hwnd).as_bool() },
        "and it is the window that went down: the command reaches the default procedure, which is \
         where a window minimizes"
    );

    // And the road that puts it back up: the tick's own show, which runs on a tick where the pin's
    // window is down and the desktop is not the window in front — and a window the Shell minimized
    // is a window down, whatever the visibility of it says (see `the_pin_window_is_down`). A pin the
    // Shell minimized that nothing restored would be a pin gone for good, since it is not a window
    // the Shell keeps in its list.
    unsafe { show_pinned_window(pin_hwnd) };
    assert!(
        !unsafe { IsIconic(pin_hwnd).as_bool() },
        "the pin comes back up with the desktop it went down with"
    );
    assert_eq!(
        unsafe {
            let mut rect = RECT::default();
            let _ = GetWindowRect(pin_hwnd, &mut rect);
            (rect.left, rect.top, rect.right, rect.bottom)
        },
        box_,
        "and it comes back where it stood, which is not where a minimized window sits"
    );

    // A close is still a close: the button's own menu row.
    unsafe {
        let _ = SendMessageW(
            pin_hwnd,
            WM_SYSCOMMAND,
            WPARAM(SC_CLOSE as usize),
            LPARAM(0),
        );
    }
    assert!(
        PIN_END_REQUESTED.load(Ordering::SeqCst),
        "the Shell's Close is the pin's end, and the pin is ended rather than its window destroyed"
    );
    assert!(
        unsafe { IsWindow(pin_hwnd).as_bool() },
        "the pin is ended rather than its window destroyed, because the window is this app's for the \
         life of the process"
    );

    // And the window's own close, which is what `Alt+F4` would be where the system keys were not
    // swallowed (see `WM_SYSKEYDOWN`), is the same end.
    PIN_END_REQUESTED.store(false, Ordering::SeqCst);
    unsafe {
        let _ = SendMessageW(pin_hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
    }
    assert!(
        PIN_END_REQUESTED.load(Ordering::SeqCst),
        "and a `WM_CLOSE`, which is the window's own close, is one too"
    );

    // Safety: the window is this test's own, and every global written here is put back.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
    }
    PIN_END_REQUESTED.store(previous_end, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// A picture pin owns the topmost band for the moment it appears, and gives it back — whatever the
/// setting says, and never again after that moment.
///
/// What the user asked for: "the moment we click SPACE the window is exactly on top of everything",
/// with `Always On Top` off meaning the pin is then an ordinary window that the next click puts behind.
#[test]
fn a_pin_that_is_a_window_of_its_own_takes_the_band_to_appear_and_gives_it_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let before = pin_always_on_top();

    stand_pin(Some(overlay_pin((100, 100, 420, 340), PinChrome::always())));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    pin_stands_in_its_own_window(true);
    assert!(
        showing_takes_the_band_to_appear(false),
        "a pin that is a window of its own, off the screen and not Always On Top appears over          everything: `HWND_TOP` is refused for a process that does not hold the foreground"
    );
    assert!(
        !showing_takes_the_band_to_appear(true),
        "a pin already on the screen is left where the user's own clicks put it: a repaint is not an          appearance, and a band on every repaint is a pin that could not be put behind anything"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    assert!(
        !showing_takes_the_band_to_appear(false),
        "a pin that is Always On Top is going to the band above everything anyway, so it is not          taken and given back first"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }
    pin_stands_in_its_own_window(false);
    assert!(
        !showing_takes_the_band_to_appear(false),
        "the kinds still standing in the preview's window keep the plain raise they have always had"
    );

    pin_stands_in_its_own_window(false);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    stand_pin(previous_pin);
}

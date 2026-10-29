//! The stand-in window the taskbar button belongs to: the frame a hover is drawn of, what a picture
//! of a put-away pin is taken from, and what a run does when the window the button belonged to has
//! gone.
//!
//! The window itself is the one part of the taskbar button a test can hold in its hand — the button
//! is the Shell's — so what is asserted here is the window: that nothing of Windows is drawn around
//! the picture, that a window the Shell has taken is made again rather than kept as a handle, and
//! that the picture of a pin put away into the bubble is the bubble.

use super::*;
use std::time::Duration;
use windows::core::w;
use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, RECT, WPARAM};
use windows::Win32::Graphics::Dwm::DwmFlush;
use windows::Win32::Graphics::Gdi::{CreateSolidBrush, FillRect, GetDC, ReleaseDC};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, GetClientRect, GetWindow, GetWindowLongPtrW, GetWindowRect,
    IsIconic, IsWindow, IsWindowVisible, RegisterClassExW, SendMessageW, SetWindowPos, ShowWindow,
    GWL_EXSTYLE, GW_HWNDPREV, HWND_TOPMOST, SC_RESTORE, SWP_NOACTIVATE, SW_MINIMIZE, SW_RESTORE,
    SW_SHOWNOACTIVATE, WA_ACTIVE, WINDOW_EX_STYLE, WM_ACTIVATE, WM_SIZE, WM_SYSCOMMAND,
    WNDCLASSEXW,
    WS_EX_APPWINDOW, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

/// Register the class the stand-in is made from, once for the run.
///
/// The real registration is the preview loop's own, at its start (see `run_preview_window`), and a
/// test that asked for a window of an unregistered class would be told only that the class is not
/// there. The class is registered under the name the module creates windows with, so what is made
/// here is the app's own stand-in and not a stand-in of the test's.
fn register_the_stand_in_class() {
    static REGISTERED: std::sync::Once = std::sync::Once::new();
    REGISTERED.call_once(|| {
        // Safety: a class of this process's own, registered once, with the window procedure the app
        // creates its stand-in with.
        unsafe {
            let hinstance = GetModuleHandleW(None).expect("the module's own handle");
            let class = WNDCLASSEXW {
                cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
                lpfnWndProc: Some(pin_taskbar_proc),
                hInstance: hinstance.into(),
                lpszClassName: PIN_TASKBAR_CLASS,
                ..Default::default()
            };
            RegisterClassExW(&class);
        }
    });
}

/// The stand-in as the app holds it, making one where there is none.
fn the_stand_in() -> HWND {
    register_the_stand_in_class();
    // Safety: the class above is registered by this point, and what comes back is the app's own
    // stand-in or nothing at all — a window of this app's, created on this thread.
    unsafe { pin_taskbar_window().expect("the stand-in the app holds") }
}

/// Let go of whatever stand-in the app holds, so that a test beside this one — or the same test a
/// second time — makes one of its own.
fn forget_the_stand_in() {
    // Safety: asks the app to let go of its own stand-in the way its Shell's own Close does.
    unsafe { close_pin_taskbar() };
}

/// The frame the Shell draws around a window's client area, measured rather than read off the style:
/// it is what a taskbar hover shows of the window around the picture the Shell mirrors into it.
unsafe fn the_frame_around(hwnd: HWND) -> (i32, i32) {
    let mut window = RECT::default();
    let mut client = RECT::default();
    let _ = GetWindowRect(hwnd, &mut window);
    let _ = GetClientRect(hwnd, &mut client);
    (
        (window.right - window.left) - (client.right - client.left),
        (window.bottom - window.top) - (client.bottom - client.top),
    )
}

/// The hover of the button is the pin's own picture and nothing of Windows around it: the window the
/// Shell draws the hover from has no frame at all, so no native title bar is drawn over the picture.
///
/// The fault this is a regression test for: the stand-in was an ordinary overlapped window with a
/// caption, and the Shell draws a window's hover *frame and all* — so the pin's picture came with
/// the app's own title bar over it, which is what the user reported seeing in the hover.
#[test]
fn the_stand_in_wears_no_frame_over_the_picture() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // A window made for this test rather than one left by the test before it: what is measured is
    // the style the window is created with, and a window somebody else made says nothing about it.
    forget_the_stand_in();
    let hwnd = the_stand_in();

    // Safety: the handle is this app's own stand-in window, and every question below is a
    // measurement of it rather than a change to it.
    let frame = unsafe { the_frame_around(hwnd) };
    assert_eq!(
        frame,
        (0, 0),
        "the stand-in is a borderless popup: a hover drawn as the window would be a picture with a \
         title bar and a border around it"
    );

    // And it is still the app's button in the taskbar: the style that asks for one is what a
    // borderless popup could quietly lose.
    let ex_style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    assert_ne!(
        ex_style & WS_EX_APPWINDOW.0 as isize,
        0,
        "and the window is still one the Shell gives a taskbar button to"
    );

    // And a borderless popup is still a window the Shell minimizes: the put-away form of the
    // stand-in is the minimized one, and a window that refused to go down would leave the app's
    // button standing over a pin the user put away (see `show_pin_taskbar`).
    unsafe {
        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(hwnd, SW_MINIMIZE);
    }
    assert!(
        unsafe { IsIconic(hwnd).as_bool() },
        "the stand-in goes down to the taskbar the way the Shell's own button expects"
    );

    forget_the_stand_in();
}

/// A button the Shell has taken away comes back on the tick, in the form the pin asks for.
///
/// This is the user's own report, as a regression test: minimizing a pin to the taskbar left the app
/// with no taskbar button at all for the rest of the run — until the app was restarted — because the
/// window the button belonged to was gone and the handle the run held was only ever asked about
/// again when a *new* pin's own road showed it. The next pin does not heal a pin that is already
/// there, so the form is read back on the tick (see `keep_the_stand_in_up`).
#[test]
fn a_button_the_shell_has_taken_away_comes_back_on_the_tick() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let before = pin_always_on_top();

    // A pin on the screen with the setting that wears a button, whose window the Shell has taken.
    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }
    PIN_BUBBLE_HWND.store(0, Ordering::SeqCst);

    forget_the_stand_in();
    let gone = the_stand_in();
    // Safety: the window is this app's own stand-in, destroyed behind the app's back the way a Shell
    // that has restarted takes one.
    unsafe {
        DestroyWindow(gone).expect("the stand-in let go of");
    }

    keep_the_stand_in_up();

    let back = showing_taskbar_window().expect("the button the pin asks for");
    assert!(
        unsafe { IsWindow(back).as_bool() },
        "the tick makes a stand-in of its own rather than holding a handle to a window that is gone"
    );
    assert!(
        !unsafe { IsIconic(back).as_bool() },
        "and it is the form a pin on the screen wants: up and mirroring, because a minimized window          is one the Shell is not composing a hover from"
    );

    // And the same pin put away into the taskbar asks for the other form: the minimized one, with
    // the frame the Shell cached on the way down.
    if let Some(mut pin) = take_pin_for_a_test() {
        pin.collapsed = true;
        stand_pin(Some(pin));
    }
    forget_the_stand_in();
    let gone = the_stand_in();
    // Safety: the same window, destroyed behind the app's back again.
    unsafe {
        DestroyWindow(gone).expect("the stand-in let go of");
    }

    keep_the_stand_in_up();

    let back = showing_taskbar_window().expect("the button a put-away pin asks for");
    assert!(
        unsafe { IsIconic(back).as_bool() },
        "a pin put away into the taskbar wears the minimized form, which is the frozen frame"
    );

    // A pin put away into the *bubble* asks for no taskbar button at all, and the tick does not make
    // one behind the round bubble's back (see `a_button_is_owed`).
    let bubble = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            0,
            0,
            120,
            120,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the bubble")
    };
    // Safety: the window is this test's own, and it is only asked to be on the screen.
    unsafe {
        let _ = ShowWindow(bubble, SW_SHOWNOACTIVATE);
    }
    PIN_BUBBLE_HWND.store(bubble.0 as isize, Ordering::SeqCst);
    forget_the_stand_in();
    keep_the_stand_in_up();
    assert!(
        showing_taskbar_window().is_none(),
        "a pin put away into its bubble wears the bubble and no taskbar button"
    );

    // Safety: the bubble's window is this test's and the stand-in is the app's own.
    unsafe {
        let _ = DestroyWindow(bubble);
    }
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    forget_the_stand_in();
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    stand_pin(previous_pin);
}

/// A stand-in the Shell has taken is made again rather than kept as a handle.
///
/// The fault this is a regression test for: the window was held in a slot that was only ever
/// cleared by the app's own Close, so a run whose window was destroyed — by the Shell, by a Shell
/// restart, by whatever else — went on holding a handle to a window that was not there. Every road
/// that shows the stand-in then showed nothing at all, which is a run with no taskbar button and no
/// hover for the rest of its life: the user's report was that the button was gone until the app was
/// restarted.
#[test]
fn a_stand_in_the_shell_has_taken_is_made_again() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let first = the_stand_in();

    // Safety: the window is this app's own stand-in, and it is destroyed behind the app's back the
    // way a window the Shell has taken is — the slot is left holding it, which is the state the
    // next ask has to survive.
    unsafe {
        DestroyWindow(first).expect("the stand-in let go of");
    }
    assert!(
        !unsafe { IsWindow(first).as_bool() },
        "the window the app holds has gone"
    );

    let second = the_stand_in();
    assert!(
        unsafe { IsWindow(second).as_bool() },
        "and the next ask makes a window of its own rather than handing back the one that has gone"
    );
    assert_eq!(
        unsafe { the_frame_around(second) },
        (0, 0),
        "which is a stand-in like any other, down to what a hover shows of it"
    );

    forget_the_stand_in();
}

/// The picture of a pin put away into the bubble is the bubble: the pin's own windows are off the
/// screen behind it, and the bubble is the only window of the pin's there is to be a picture of.
///
/// The fault this is a regression test for: a bubble's own `To Taskbar` showed the pin's hover with
/// nothing in it at all. The picture is taken as the stand-in is put up, the pin's windows are
/// already hidden by then, and the mirror fell through to the preview window — which is off the
/// screen and composed of nothing.
#[test]
fn the_picture_of_a_pin_put_away_into_a_bubble_is_the_bubble() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    // A bubble standing on the screen, stood in by a real window of the machine's own: what the
    // mirror reads is whether the window is visible, and a window that was never shown is not.
    let bubble = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            0,
            0,
            120,
            120,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the bubble")
    };
    // Safety: the window is this test's own, and it is only asked to be on the screen for as long as
    // the question is asked.
    unsafe {
        let _ = ShowWindow(bubble, SW_SHOWNOACTIVATE);
    }
    PIN_BUBBLE_HWND.store(bubble.0 as isize, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    assert_eq!(
        mirror_source(Some(MediaType::Video)),
        Some(bubble),
        "a put-away pin's picture is the bubble it left"
    );

    // And a pin that is *not* put away is pictured by the window it is standing in, not by the
    // bubble: the bubble is not on the screen while the pin is up (see `collapse_pin_to`).
    if let Some(mut pin) = take_pin_for_a_test() {
        pin.collapsed = false;
        stand_pin(Some(pin));
    }
    assert_ne!(
        mirror_source(Some(MediaType::Video)),
        Some(bubble),
        "a pin that is up is not pictured by a bubble"
    );

    // Safety: the window is this test's own and the stand-in is the app's own; both are let go of,
    // and the handles the app holds are put back as they were.
    unsafe {
        let _ = DestroyWindow(bubble);
    }
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A pin brought back by the taskbar button goes back exactly where it stood.
///
/// The fault this is a regression test for: the button's own click took the stand-in down *before*
/// the pin was asked to come back, and the pin reads which place to come back to from whether that
/// stand-in is up — so a pin restored from the taskbar was placed beside the pointer on the taskbar
/// instead of at the box it was put away at, which is the one thing a restore from a taskbar button
/// must not do. It only showed with `Pin Mode → Always On Top` ON, because that is the setting the
/// early take-down was gated on.
#[test]
fn a_pin_restored_from_its_button_goes_back_where_it_stood() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let before = pin_always_on_top();
    let stand_in = the_stand_in();

    // The pin as a put-away pin is: collapsed, standing at the box the user left it at, and with the
    // stand-in up in the form a hover holds the frozen frame in.
    let left_at = (100, 80, 420, 320);
    let mut pin = overlay_pin(left_at, PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    // Safety: the window is the app's own stand-in and it is only asked to be up, which is the state
    // the button's own click finds it in.
    unsafe {
        let _ = ShowWindow(stand_in, SW_SHOWNOACTIVATE);
    }

    // The click itself: the Shell sends the window its restore, and the window asks the pin for it.
    // Safety: a message to a window of this app's own, with the numbers the Shell sends with it.
    unsafe {
        SendMessageW(
            stand_in,
            WM_SYSCOMMAND,
            WPARAM(SC_RESTORE as usize),
            LPARAM(0),
        );
    }
    assert!(
        pin_taskbar_is_showing(),
        "the click leaves the stand-in standing: the pin's own restore is what takes it down, and          what it reads the place to come back to from is whether there is one"
    );

    restore_pin(&a_restored_window(), PinRestore::Pressed);

    let restored =
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.collapsed, pin.content)));
    assert_eq!(
        restored,
        Some((false, left_at)),
        "a pin restored from the taskbar comes back where it stood, and a place taken from the \
         pointer on the taskbar would be a different one"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    forget_the_stand_in();
    stand_pin(previous_pin);
}

/// A pin minimized a second time, after its button has brought it back once, goes away again.
///
/// The fault this is a regression test for: the user's own run had the first minimize work — the
/// button went down to the taskbar in its put-away form — and the minimize after the button brought
/// the pin back did nothing at all, with the pin's window flashing where it stood rather than going
/// away. Windows asks an iconic window whether it may be opened for the mirroring step's own
/// showing of the stand-in as much as for the user's click on the button, and answered as the click
/// it queued the pin's own restore behind the collapse that had just asked for it (see
/// `ASKING_THE_STAND_IN_UP`).
#[test]
fn a_pin_minimized_again_after_a_button_restore_goes_away_again() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let before = pin_always_on_top();

    let left_at = (100, 80, 420, 320);
    let mut pin = overlay_pin(left_at, PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }

    forget_the_stand_in();
    register_the_stand_in_class();
    let _ = take_pin_command();

    // The first minimize: the pin goes away and its button stands in the taskbar, minimized.
    collapse_pin_to(PinMinimizeTo::Taskbar);
    let stand_in = showing_taskbar_window().expect("the button the first minimize leaves");
    assert!(
        unsafe { IsIconic(stand_in).as_bool() },
        "the first minimize puts the button down in the form a hover holds the frozen frame in"
    );

    // The click on that button: the Shell asks the window to restore, and the window asks the pin.
    // Safety: a message to a window of this app's own, with the numbers the Shell sends with it.
    unsafe {
        SendMessageW(
            stand_in,
            WM_SYSCOMMAND,
            WPARAM(SC_RESTORE as usize),
            LPARAM(0),
        );
    }
    assert!(
        matches!(take_pin_command(), Some(PinCommand::Restore)),
        "the click on the button is the pin's own restore"
    );
    restore_pin(&a_restored_window(), PinRestore::Pressed);
    assert!(!pin_is_collapsed(), "the button brings the pin back");
    assert_eq!(
        take_pin_command(),
        None,
        "and the restore the click asked for is the last thing left in hand"
    );

    // And the minimize the user makes next, which used to do nothing: the pin stays put away, and
    // the stand-in's own showing for the picture leaves no restore behind it.
    collapse_pin_to(PinMinimizeTo::Taskbar);
    assert_eq!(
        take_pin_command(),
        None,
        "the picture taken for the hover is not a reason to bring the pin back"
    );
    assert!(
        pin_is_collapsed(),
        "the second minimize puts the pin away rather than leaving it on the screen"
    );
    let stand_in = showing_taskbar_window().expect("the button the second minimize owes");
    assert!(
        unsafe { IsIconic(stand_in).as_bool() },
        "and the button goes down with it, in the form a hover holds the frozen frame in"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    forget_the_stand_in();
    stand_pin(previous_pin);
}

/// A stand-in whose pin is put away is not brought back by an activation.
///
/// The fault this is a regression test for: the user's own run had the first minimize after the
/// button had brought the pin back flash the window up over the desktop and leave it standing — the
/// same press a second time went through. Windows hands the keyboard on to the next window when this
/// app takes the pin's window off the screen, and the stand-in the taskbar click a moment before had
/// activated is the window it picks; read as a click on the button it asked for the pin's own
/// restore, behind the collapse that had just put the pin away.
#[test]
fn a_put_away_stand_in_is_not_brought_back_by_an_activation() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let before = pin_always_on_top();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    forget_the_stand_in();
    let _ = take_pin_command();

    // The minimize: the pin goes away into the taskbar, and the button goes down in its put-away
    // form with it.
    collapse_pin_to(PinMinimizeTo::Taskbar);
    let stand_in = showing_taskbar_window().expect("the button the minimize leaves");
    let _ = take_pin_command();

    // What Windows sends that window as the pin's own window goes off the screen: the message it
    // activates a window with, and no click on the button.
    // Safety: a message to a window of this app's own, with the numbers the Shell sends with it.
    unsafe {
        SendMessageW(stand_in, WM_ACTIVATE, WPARAM(WA_ACTIVE as usize), LPARAM(0));
    }

    assert_eq!(
        take_pin_command(),
        None,
        "an activation of a stand-in whose pin is put away is not a click on its button, and the \
         pin it would bring back is the one the collapse has just put away"
    );
    assert!(
        pin_is_collapsed(),
        "and the pin is left where the minimize put it"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    forget_the_stand_in();
    stand_pin(previous_pin);
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

/// A pin the taskbar's own button brings back comes back in front of the windows that were in front
/// of it.
///
/// The fault this is a regression test for: the user's own run had a pin restored from the taskbar
/// come back behind everything — the window "should show up over/on top of all the other windows,
/// this used the last state instead". A window that has been hidden has no place of its own left in
/// the z-order, and the showing of it asked for `HWND_NOTOPMOST`, which does nothing at all for a
/// window that is already ordinary: what the pin came back to was the slot it went down in, which is
/// behind everything opened over it while it was away (see `show_pinned_window`).
#[test]
fn a_pin_the_button_brings_back_comes_back_in_front_of_the_windows_it_left_behind() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let before = pin_always_on_top();

    register_the_stand_in_class();
    forget_the_stand_in();
    let _ = take_pin_command();

    // The window the pin is drawn by, and the window standing in front of it — both standing where
    // no display is, so that a run of the suite is no window of the user's and no place the hand can
    // land in. The second is made after the first, which is what puts it in front.
    let pin_hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the pin")
    };
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
    let in_front = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a window in front of the pin")
    };
    // Safety: both windows are this test's own, and both are asked to be on the screen — an offscreen
    // window is still one the Shell lists, which is the whole of what the z-order below is read from.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
    }
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the window this test opened after the pin's is the one in front of it"
    );

    // The setting the user's own report is under: a pin that is not Always On Top.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    let mut pin = overlay_pin(
        (
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN + 320,
            TASKBAR_OFFSCREEN + 240,
        ),
        PinChrome::always(),
    );
    pin.collapsed = false;
    stand_pin(Some(pin));

    // The minimize: the pin goes away into the taskbar, and its own window goes off the screen with
    // it.
    collapse_pin_to(PinMinimizeTo::Taskbar);
    let stand_in = showing_taskbar_window().expect("the button the minimize leaves");
    assert!(
        !unsafe { IsWindowVisible(pin_hwnd).as_bool() },
        "the pin's own window comes off the screen with the pin, and a window that is not on the \
         screen is one the Shell has no place in the z-order for"
    );

    // The click on that button: the Shell asks the window it holds for a restore, and the window
    // asks the pin for it (see `handle_taskbar_restore`).
    // Safety: a message to a window of this app's own, with the numbers the Shell sends with it.
    unsafe {
        SendMessageW(
            stand_in,
            WM_SYSCOMMAND,
            WPARAM(SC_RESTORE as usize),
            LPARAM(0),
        );
    }
    assert!(
        matches!(take_pin_command(), Some(PinCommand::Restore)),
        "the click on the button is the pin's own restore"
    );
    restore_pin(&a_restored_window(), PinRestore::Pressed);

    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a pin the button brought back is put above the windows the user had in front of it: what \
         `HWND_NOTOPMOST` does for a window that is already ordinary is nothing at all, so the pin \
         would be shown back into the place it went down in rather than over the other windows"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    forget_the_stand_in();
    stand_pin(previous_pin);
}

/// A stand-in the Shell left iconic under a pin that is on the screen comes back up on the tick.
///
/// The setting is `Pin Mode → Always On Top` OFF, where a pin on the screen wears a button of its
/// own: the Shell's own minimize of that button leaves it iconic, and putting it back up is the
/// tick's repair. The repair is a showing of an iconic window, so Windows asks the window whether
/// it may open — and answering that as the button's gesture left the window iconic for the rest of
/// the run, with the pin's window raised for it on every tick (see `keep_the_stand_in_up`).
#[test]
fn a_stand_in_left_iconic_under_a_pin_on_the_screen_comes_back_up_on_the_tick() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let before = pin_always_on_top();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    forget_the_stand_in();
    let stand_in = the_stand_in();
    // Safety: this test's own stand-in is asked up and then minimized, which is the form the Shell
    // leaves a button in when it minimizes one.
    unsafe {
        let _ = ShowWindow(stand_in, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(stand_in, SW_MINIMIZE);
    }
    assert!(unsafe { IsIconic(stand_in).as_bool() });
    let _ = take_pin_command();

    keep_the_stand_in_up();

    assert!(
        !unsafe { IsIconic(stand_in).as_bool() },
        "the tick puts a button a pin on the screen wears back up, which it cannot do while the \
         window's own question is answered as the user's click"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "and a pin that is already on the screen holds nothing for the button to ask of it"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    forget_the_stand_in();
    stand_pin(previous_pin);
}

/// **A stand-in a pin leaves standing with nothing behind it goes down on the tick**, and what it
/// leaves is the nothing a pin that never wore one leaves. A pin's own close hides the stand-in with
/// the windows it takes down, but the close is one road out of a pin and the tick is what runs after
/// all of them: the watchdog's end, a Shell gesture on the button, and a stand-in shown for a pin
/// that is already gone all leave a window of this app's in the Shell's list with nothing behind it —
/// the button a hover then finds over a file the app is not showing (see `keep_the_stand_in_up` and
/// `a_button_is_owed`).
#[test]
fn a_stand_in_left_for_a_pin_that_is_gone_leaves_what_a_close_without_one_leaves() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // What a close leaves, with a stand-in left behind it or without one: whether a pin is up at
    // all, and whether the app still holds a window the Shell lists a button for.
    fn what_a_close_leaves(a_stand_in_is_up: bool) -> (bool, bool) {
        let previous_pin = take_pin_for_a_test();
        forget_the_stand_in();
        stand_pin(Some(overlay_pin((100, 80, 420, 320), PinChrome::always())));

        if a_stand_in_is_up {
            let stand_in = the_stand_in();
            // Safety: the handle is this app's own stand-in, and it is only asked up.
            unsafe {
                let _ = ShowWindow(stand_in, SW_SHOWNOACTIVATE);
            }
            assert!(
                showing_taskbar_window().is_some(),
                "the stand-in is up before the pin goes"
            );
        }

        // The end that is not the loop's own: nothing of the pin's windows is hidden with it, so
        // whatever the Shell has standing is still standing after it (see `Reason::Hung`).
        let _ = end_pin(Reason::Hung, &Win32PinWindow);

        keep_the_stand_in_up();

        let left = (pinned(), showing_taskbar_window().is_some());
        forget_the_stand_in();
        stand_pin(previous_pin);
        left
    }

    let with_no_stand_in = what_a_close_leaves(false);
    let with_a_stand_in = what_a_close_leaves(true);

    assert_eq!(
        with_a_stand_in, with_no_stand_in,
        "a stand-in left by a pin that is gone leaves what a pin that wore none leaves"
    );
    assert_eq!(
        with_a_stand_in,
        (false, false),
        "which is no pin up and no window of ours for the Shell to list"
    );
}

/// The client box of a window, which is the box the Shell draws a taskbar hover's picture in.
///
/// Safety: the handle is a window of this app's own or of the test's, and the question is a
/// measurement of it rather than a change to it.
unsafe fn the_client_box(hwnd: HWND) -> (i32, i32) {
    let mut client = RECT::default();
    let _ = GetClientRect(hwnd, &mut client);
    (client.right - client.left, client.bottom - client.top)
}

/// A pin put away into the taskbar keeps the picture its collapse took, and is never pictured by a
/// window that is not on the screen.
///
/// The fault this is a regression test for: the road that puts a pin away into the taskbar takes the
/// pin's picture *while the bubble is still up* and then puts the pin's own windows away behind it —
/// but the tick went on settling the mirror afterwards, and what it found for a pin the player draws
/// was that player's own *hidden* window. The Shell composes nothing for a window that is not on the
/// screen, so the stand-in was re-registered and re-sized for a picture that was not there, the
/// registration the collapse had taken was dropped with it — and the hover it belonged to was left
/// showing the stand-in's own rows, which this app has never painted.
#[test]
fn a_pin_put_away_keeps_the_picture_its_collapse_took() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_video = VIDEO_HWND.load(Ordering::SeqCst);

    register_the_stand_in_class();

    // The bubble a collapse leaves, standing on the screen at the circle's own size: what the mirror
    // is taken of while the pin is being put away (see `showing_bubble_window`).
    let bubble = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            0,
            0,
            44,
            44,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the bubble")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen: a bubble the
    // mirror is taken of is one the Shell is composing (see `showing_bubble_window`).
    unsafe {
        let _ = ShowWindow(bubble, SW_SHOWNOACTIVATE);
    }
    PIN_BUBBLE_HWND.store(bubble.0 as isize, Ordering::SeqCst);
    let bubble_box = unsafe { the_client_box(bubble) };

    // The window the pin is drawn by — mpv's, for the user's setting — already put away behind the
    // bubble by the collapse, which is what the pin's own windows are while it is a bubble. A window
    // of this test's own stands in for it: what matters is that it is not on the screen.
    let player = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            0,
            0,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a player's window")
    };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    // A pin put away into the taskbar, with the stand-in the button belongs to up.
    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    forget_the_stand_in();
    let stand_in = the_stand_in();
    // Safety: the window is this app's own stand-in, asked only to be up — a hidden window has no
    // taskbar button and no picture to hold (see `showing_taskbar_window`).
    unsafe {
        let _ = ShowWindow(stand_in, SW_SHOWNOACTIVATE);
    }

    // The collapse's own taking: the picture is the bubble's, and the stand-in is given its box.
    sync_pin_mirror(Some(MediaType::Video));
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        bubble_box,
        "the picture a collapse takes is the bubble and no other window, so the stand-in is the \
         bubble's own box"
    );

    // And the bubble goes with the collapse's next breath (see `bubble_to_taskbar`).
    hide_pin_bubble();

    // What the tick asks a moment later: nothing of the pin is on the screen any more, so there is
    // no window left to take a picture of and the picture already taken is the one to hold.
    assert_eq!(
        mirror_source(Some(MediaType::Video)),
        None,
        "a pin put away has no window on the screen left to be pictured by"
    );
    sync_pin_mirror(Some(MediaType::Video));
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        bubble_box,
        "and the stand-in is left the box the frame the hover holds was taken in, rather than one \
         re-made for a window the Shell composes nothing for"
    );

    // Safety: the two windows are this test's own and the stand-in is the app's own; the handles the
    // app holds are put back as they were.
    unsafe {
        let _ = DestroyWindow(player);
        let _ = DestroyWindow(bubble);
    }
    VIDEO_HWND.store(previous_video, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    forget_the_stand_in();
    stand_pin(previous_pin);
}
/// The picture a minimize into the taskbar takes is taken of a window that is still up.
///
/// The other half of the rule above, and the road that has no bubble in it: a pin put away by its
/// minimize button is pictured by the window it is drawn by, and the collapse takes that picture
/// *before* it puts that window away — the pin is marked away first and its windows are hidden after
/// (see `collapse_pin_to`). A rule that looked for a window on the screen only once the collapse had
/// finished would find none and never size the stand-in for a picture at all, so the hover would fall
/// back on the stand-in's own rows.
#[test]
fn a_pin_minimized_into_the_taskbar_is_pictured_before_its_windows_go() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_video = VIDEO_HWND.load(Ordering::SeqCst);
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let _ = take_pin_command();

    register_the_stand_in_class();

    // The window the pin is drawn by — mpv's, for the user's setting — standing in the pin's band and
    // on the screen, which is what the collapse finds it as (see `keep_the_pin_for_the_taskbar`).
    let player = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            0,
            0,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a player's window")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen: a player's window is
    // a picture of the pin while it is showing, and the collapse takes the picture before it hides it.
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
    }
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    let player_box = unsafe { the_client_box(player) };

    // The film the pin is showing, which is the media the collapse asks its picture of.
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    forget_the_stand_in();
    let stand_in = the_stand_in();

    // The moment the collapse takes its picture in: the pin is already marked away — every road that
    // takes one sets that first — and its own windows are still up, because the collapse puts them
    // away after (see `collapse_pin_to`, `hide_pinned_windows`).
    if let Some(mut pin) = take_pin_for_a_test() {
        pin.collapsed = true;
        stand_pin(Some(pin));
    }
    unsafe {
        keep_the_pin_for_the_taskbar();
    }
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        player_box,
        "a pin put away is pictured while its own window is still up, so the stand-in is left that \
         window's box"
    );

    // And once the window has gone down with the collapse there is nothing left to take a picture of,
    // so the picture already taken is the one to hold (see `a_pin_put_away_keeps_the_picture_its_collapse_took`).
    unsafe {
        hide_pinned_windows();
    }
    assert!(
        !unsafe { IsWindowVisible(player).as_bool() },
        "the collapse puts the window the picture was taken of away"
    );
    assert_eq!(
        mirror_source(Some(MediaType::Video)),
        None,
        "and there is no window of the pin's left on the screen to picture"
    );
    sync_pin_mirror(Some(MediaType::Video));
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        player_box,
        "so the stand-in is left the box the frame the hover holds was taken in"
    );

    // Safety: the window is this test's own and the stand-in is the app's own; the handles the app
    // holds are put back as they were.
    unsafe {
        let _ = DestroyWindow(player);
    }
    VIDEO_HWND.store(previous_video, Ordering::SeqCst);
    forget_the_stand_in();
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    stand_pin(previous_pin);
}

/// A pin drawn by a window of its own is pictured by that window, and by nothing else.
///
/// The fault this is a regression test for: a film's player window is off the screen on roads of its
/// own — one that has ended, one being relaunched inside a walk, one a display change has taken down
/// — and the pin's own window was asked next for the picture. That window is a *hole* where the film
/// was: what a mirror of it holds is the caption and the bar around an empty band. And so is a player
/// window this app has not cleared yet, which is a handle that is no window at all — both are
/// registrations the Shell draws nothing into, and taking either is what takes down the registration
/// the frame already standing in the stand-in was taken with (see `mirror_source`).
#[test]
fn a_pin_drawn_by_a_window_of_its_own_is_pictured_only_by_that_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_video = VIDEO_HWND.load(Ordering::SeqCst);
    let previous_own = pin_is_its_own_window();

    register_the_stand_in_class();
    forget_the_stand_in();
    // The film is FFmpeg's player's, which is the kind of film this test is about: a window of its
    // own standing in the pin's band, so the player is the picture and the pin's window is the hole
    // (see `mirror_source`).
    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

    // The pin's own window, up: where a picture, a page, a card and a folder's listing are drawn.
    let pin_hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the pin")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen — which is what
    // makes it a picture rather than one more window the Shell composes nothing for.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    }
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);

    // And a player's window that is *not* on the screen, which is a player ended and begun again.
    let player = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            640,
            480,
            None,
            None,
            None,
            None,
        )
        .expect("a player's window")
    };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    assert_eq!(
        mirror_source(Some(MediaType::Video)),
        Some(pin_hwnd),
        "a hosted film is drawn by the pin's own window, so the pin is its picture: a stale player \
         window that is not on the screen is no picture of a film"
    );
    assert_eq!(
        mirror_source(Some(MediaType::EngineSvg)),
        None,
        "and a document the engine draws is the same case"
    );
    assert_eq!(
        mirror_source(Some(MediaType::StaticImage)),
        Some(pin_hwnd),
        "while a picture is drawn by the pin's own window, which is up: that *is* the picture"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(player);
        let _ = DestroyWindow(pin_hwnd);
    }
    VIDEO_HWND.store(previous_video, Ordering::SeqCst);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    forget_the_stand_in();
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// A sound's peek is the card the pin draws for it.
///
/// The fault this is a regression test for: a sound was shown as nothing, deliberately — the three
/// buttons beside the hover were taken to be the whole of what a sound put away had to offer — so the
/// stand-in was left at the square it is created at, with no registration in it, whether the pin was
/// on the screen or put away. The card is drawn in the pin's own window exactly as a picture is, and
/// it is the picture the button's hover was missing: a peek of a sound is its card the way a peek of
/// a film is the film (see `mirror_source`).
#[test]
fn a_sound_is_pictured_by_the_card_the_pin_draws_for_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    register_the_stand_in_class();
    forget_the_stand_in();

    // The window the card is drawn in, up: the picture a pin's own window holds for every kind that
    // has no window of its own.
    let card = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the pin")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen — a picture's source
    // is a window the Shell is composing (see `on_the_screen`).
    unsafe {
        let _ = ShowWindow(card, SW_SHOWNOACTIVATE);
    }
    PREVIEW_HWND.store(card.0 as isize, Ordering::SeqCst);
    let card_box = unsafe { the_client_box(card) };

    // The sound the pin is showing, and the pin itself put away into the taskbar — the road the hover
    // of a sound's button is reached by.
    let mut sound = create_loading_media(320, 240);
    sound.media_type = MediaType::Audio;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(sound);
    }

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    let stand_in = the_stand_in();

    assert_eq!(
        mirror_source(Some(MediaType::Audio)),
        Some(card),
        "a sound's card is drawn in the pin's own window, so that window is the picture a hover is \
         given rather than nothing at all"
    );

    // The collapse's own taking: a put-away pin's picture is taken while the window it is drawn in is
    // still up (see `keep_the_pin_for_the_taskbar`).
    unsafe {
        keep_the_pin_for_the_taskbar();
    }
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        card_box,
        "and the stand-in is left the card's own box, which is the box the hover's picture is drawn \
         in"
    );

    // Safety: the window is this test's own and the stand-in is the app's own; the handles the app
    // holds are put back as they were.
    unsafe {
        let _ = DestroyWindow(card);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    forget_the_stand_in();
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    stand_pin(previous_pin);
}

/// The colour a test paints a window so that what the button holds can be told from a window with
/// nothing in it.
const A_COLOUR_OF_ITS_OWN: COLORREF = COLORREF(0x0020_C040);

/// Where a test stands the window a pin's picture is taken of: on the desk and away from its corner,
/// so that a picture read off the wrong place is read as the wrong colour rather than as a
/// coincidence.
const A_BOX_ON_THE_DESK: (i32, i32) = (30, 40);

/// Fill a window's client area with a colour of its own: what a picture read off the desktop is read
/// for, since a window that paints nothing is not a picture of anything.
///
/// Safety: the window is the caller's own, and painting it is a change to it that the caller is the
/// one asking for.
unsafe fn fill_the_window(hwnd: HWND, colour: COLORREF) {
    let mut client = RECT::default();
    let _ = GetClientRect(hwnd, &mut client);
    let dc = GetDC(hwnd);
    let brush = CreateSolidBrush(colour);
    FillRect(dc, &client, brush);
    let _ = windows::Win32::Graphics::Gdi::DeleteObject(brush);
    let _ = ReleaseDC(hwnd, dc);
}

/// The colour a picture is holding at a point, as the app's own `COLORREF` colour: the three bytes a
/// picture read off the desktop is made of, in the order a colour is written in.
fn the_colour_in(pixels: &[u8], width: u32, x: u32, y: u32) -> u32 {
    let at = ((y * width + x) * 4) as usize;
    pixels[at + 2] as u32 | (pixels[at + 1] as u32) << 8 | (pixels[at] as u32) << 16
}

/// A pin put away into the taskbar holds a picture of its own, taken while its windows are still up.
///
/// The fault this is a regression test for: the user's own report of a put-away pin whose hover held
/// "a windows PE icon stretched like crazy depending on the pin mode aspect ratio" for every kind of
/// pin, while the same pin's hover was its live self as long as the pin was up. What the Shell holds of
/// a window that has gone down is a still of the window's own rows, and a stand-in whose only content
/// is a mirror of another window has none of its own: the mirror is composed *over* the stand-in while
/// it is up, and it goes with the window's going down. So the button's still has to be a picture of the
/// pin's own — one taken while the pin's windows are still on the screen to read it off (see
/// `take_the_pin_picture`), and handed over when the Shell asks for it (see
/// `answer_the_thumbnail_request`).
#[test]
fn a_pin_put_away_holds_the_picture_taken_while_its_windows_are_up() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_video = VIDEO_HWND.load(Ordering::SeqCst);
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    register_the_stand_in_class();
    forget_the_stand_in();

    // The window the pin is drawn by — mpv's, for the user's setting — on the screen at a size and in
    // a colour of its own, and above whatever else is standing where it is: a picture of a pin is a
    // picture of what the desktop is showing where that window is, so a test that read one back
    // through anything else would be reading the desk rather than the pin.
    let player = unsafe {
        CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            A_BOX_ON_THE_DESK.0,
            A_BOX_ON_THE_DESK.1,
            64,
            48,
            None,
            None,
            None,
            None,
        )
        .expect("a player's window")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen — a window the Shell
    // composes nothing for is no picture at all (see `mirror_source`) — and to be painted, since a
    // window that paints nothing is not what a picture of the pin is read for.
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
        fill_the_window(player, A_COLOUR_OF_ITS_OWN);
    }
    let _ = unsafe { DwmFlush() };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    let player_box = unsafe { the_client_box(player) };

    // The film the pin is showing — the kind whose picture is its player's own window — and the pin
    // put away with no bubble standing in its place (see `current_media_type`, `mirror_source`).
    let mut video = create_loading_media(player_box.0 as u32, player_box.1 as u32);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    PIN_BUBBLE_HWND.store(0, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 164, 128), PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    let stand_in = the_stand_in();

    // The collapse's own taking, while the window the pin is drawn by is still up (see
    // `keep_the_pin_for_the_taskbar`).
    unsafe {
        keep_the_pin_for_the_taskbar();
    }
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        player_box,
        "the picture is the player's window, so the stand-in is left that window's own box"
    );

    let (pixels, width, height) =
        the_pin_picture().expect("a pin put away into the taskbar holds a picture of its own");
    assert_eq!(
        (width, height),
        (player_box.0 as u32, player_box.1 as u32),
        "and it is taken of the window the mirror is of, at that window's own size"
    );
    assert_eq!(
        the_colour_in(&pixels, width, 0, 0),
        A_COLOUR_OF_ITS_OWN.0,
        "and it is that window's own colour: what a put-away button's hover holds is the pin, rather \
         than the icon the Shell falls back on for a window with no still of its own"
    );

    // And the window it was taken of goes with the collapse, which is the road every put-away pin
    // leaves it by (see `collapse_pin_to`, `hide_pinned_windows`).
    unsafe {
        hide_pinned_windows();
        keep_the_pin_for_the_taskbar();
    }
    let (held, ..) = the_pin_picture().expect("the picture the button already holds");
    assert_eq!(
        held, pixels,
        "a picture once taken is the picture the button keeps: nothing of the pin is on the screen \
         any more to take another of"
    );

    // Safety: the window is this test's own and the stand-in is the app's own; the handles the app
    // holds are put back as they were.
    unsafe {
        let _ = DestroyWindow(player);
    }
    forget_the_stand_in();
    VIDEO_HWND.store(previous_video, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    stand_pin(previous_pin);
}

/// A button made again holds no picture of the pin the button before it belonged to.
///
/// The picture of a hosted film is the pin's own window's, because the player is a child of the
/// pin's film host rather than a window of the desktop's (see `mirror_source`).
///
/// The road this is a regression test for: the Shell's own `Close` takes the stand-in's window down
/// with the pin (see `close_pin_taskbar`, `pin_taskbar_window`), and the next pin makes another one —
/// while the picture the button holds is this app's own and not the window's, so a picture kept past
/// the window it was handed for is a new pin's button showing the file the pin before it was of.
#[test]
fn a_button_made_again_holds_no_picture_of_the_pin_the_button_before_belonged_to() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_video = VIDEO_HWND.load(Ordering::SeqCst);
    let previous_own = pin_is_its_own_window();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    register_the_stand_in_class();
    forget_the_stand_in();

    // The window the pin is drawn by, on the screen and painted: what a picture of it is taken of.
    // A hosted film — mpv's or FFmpeg's, adopted into the pin's own host — is drawn by the pin's
    // own window, so that is the picture the button holds; the stale player window beside it is a
    // handle that is not on the screen and never the picture (see `mirror_source`).
    let pin_window = unsafe {
        CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            A_BOX_ON_THE_DESK.0,
            A_BOX_ON_THE_DESK.1,
            64,
            48,
            None,
            None,
            None,
            None,
        )
        .expect("the pin's own window")
    };
    let previous_preview = PREVIEW_HWND.swap(pin_window.0 as isize, Ordering::SeqCst);
    let player = unsafe {
        CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            A_BOX_ON_THE_DESK.0,
            A_BOX_ON_THE_DESK.1,
            64,
            48,
            None,
            None,
            None,
            None,
        )
        .expect("a stale player's window")
    };
    // Safety: the window is this test's own, asked to be on the screen and painted — the two things
    // that make it the picture the button of a film pin is taken of.
    unsafe {
        let _ = ShowWindow(pin_window, SW_SHOWNOACTIVATE);
        fill_the_window(pin_window, A_COLOUR_OF_ITS_OWN);
    }
    let _ = unsafe { DwmFlush() };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    let mut video = create_loading_media(64, 48);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    PIN_BUBBLE_HWND.store(0, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 164, 128), PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    // **The pin stands in a window of its own**, which is where a film's host is made: the picture
    // the button holds is the pin's window rather than the player's (see `mirror_source`).
    pin_stands_in_its_own_window(true);
    the_stand_in();

    // The picture the button of this pin holds, taken the way a collapse takes it.
    //
    // **Asked again while it is still nothing**, because the picture is read off another window's
    // pixels and that read is a read of what the compositor has drawn: asked once, this was a red in
    // one run out of three with nothing changed but the moment the run started (the same read stands
    // behind `DwmFlush` above). The wait is bounded by a few frames.
    let mut picture = None;
    for _ in 0..40 {
        unsafe { keep_the_pin_for_the_taskbar() };
        picture = the_pin_picture();
        if picture.is_some() {
            break;
        }
        std::thread::sleep(Duration::from_millis(5));
    }
    assert!(
        picture.is_some(),
        "the button of a pin put away holds the pin's own picture"
    );

    // And the Shell's own Close takes the button down with the pin, which is the only road that
    // clears the picture with the window it was taken for (see `close_pin_taskbar`).
    forget_the_stand_in();
    assert!(
        the_pin_picture().is_none(),
        "a button made again holds no picture of the pin the button before it belonged to"
    );

    // Safety: the window is this test's own; the handles the app holds are put back as they were.
    unsafe {
        let _ = DestroyWindow(player);
        let _ = DestroyWindow(pin_window);
    }
    VIDEO_HWND.store(previous_video, Ordering::SeqCst);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// **A film host's button holds a picture of the host itself**, and it is what the Shell is handed when
/// it asks for the still of a button that is down.
///
/// The fault this is a regression test for: a pin whose film a player of its own draws is put away by
/// minimizing the host that owns its chrome, and the Shell's still of a window that is down is taken of
/// the window's own rows — while the film in a host is a thumbnail the Shell composes *over* those rows,
/// which are the black the host is erased to. So the button of a put-away film pin had a blank hover.
/// What the button holds has to be a picture of the host as the desktop shows it, taken before the
/// collapse (see `pin_film_host::keep_the_pins_picture_for_the_taskbar`), and handed over when the Shell
/// asks (see `answer_a_thumbnail_ask`).
#[test]
fn a_film_hosts_button_holds_a_picture_of_the_host_itself() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    // Safety: the pin and its host are this test's own, and every global written here is put back when
    // the stand is dropped.
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });

    // The host on the desk and in a colour of its own: the picture is read off the desktop where the
    // window stands, so one read anywhere else is the desk rather than the pin.
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        let _ = SetWindowPos(
            host,
            HWND_TOPMOST,
            A_BOX_ON_THE_DESK.0,
            A_BOX_ON_THE_DESK.1,
            64,
            48,
            SWP_NOACTIVATE,
        );
        fill_the_window(host, A_COLOUR_OF_ITS_OWN);
    }
    let _ = unsafe { DwmFlush() };
    // The box the picture is to be of, read before the collapse: a minimized window has no client area
    // for anything to be measured in.
    let host_box = unsafe { the_client_box(host) };

    collapse_pin_to(PinMinimizeTo::Taskbar);

    let (pixels, width, height) = the_pin_picture()
        .expect("a put-away pin whose film a player draws holds a picture of its host");
    assert_eq!(
        (width as i32, height as i32),
        host_box,
        "the picture is taken of the host's own client area, which is where the film and the chrome over \
         it stand"
    );
    assert_eq!(
        the_colour_in(&pixels, width, 0, 0),
        A_COLOUR_OF_ITS_OWN.0,
        "and it is the host's own colour rather than the black its rows are erased to"
    );

    // And the Shell's own ask for the still of the button that is down is answered with it: the high
    // word of the message's `lparam` is the widest thumbnail it will take and the low word the tallest.
    let handed_over = thumbnails_handed_over();
    unsafe { answer_a_thumbnail_ask(host, LPARAM((160 << 16) | 90)) };
    assert_eq!(
        thumbnails_handed_over(),
        handed_over + 1,
        "the ask is answered with a bitmap made in the box it asked in"
    );
}

/// **A pin moved from the bubble into the taskbar keeps the picture its bubble collapse took.**
///
/// The fault this is a regression test for: the bubble's own `To Taskbar` row asks the host for the
/// picture the button is to hold when the host has been off the screen for as long as the bubble has
/// been up, so nothing was taken and the app's still was never forced — the hover of a film pin
/// brought back from the bubble into the taskbar was blank. **The last moment the host is on the
/// screen is the bubble's own collapse**, and the frame composed there is the last one the user saw,
/// which is exactly what a hover of a put-away pin is for (see taking the picture at every collapse,
/// `pin_film_host::keep_the_pins_picture_for_the_taskbar`).
#[test]
fn a_film_host_moved_from_the_bubble_into_the_taskbar_holds_the_picture_its_bubble_collapse_took() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    // Safety: the pin and its host are this test's own, and every global written here is put back
    // when the stand is dropped.
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        let _ = SetWindowPos(
            host,
            HWND_TOPMOST,
            A_BOX_ON_THE_DESK.0,
            A_BOX_ON_THE_DESK.1,
            64,
            48,
            SWP_NOACTIVATE,
        );
        fill_the_window(host, A_COLOUR_OF_ITS_OWN);
    }
    let _ = unsafe { DwmFlush() };
    let host_box = unsafe { the_client_box(host) };

    // The road the user takes: the pin goes away into the round bubble first.
    collapse_pin_to(PinMinimizeTo::Bubble);
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "the bubble's collapse is what takes the host off the screen"
    );

    let (pixels, width, height) = the_pin_picture().expect(
        "the bubble's collapse takes the picture the button of a pin moved from it into the taskbar \
         is to hold",
    );
    assert_eq!(
        (width as i32, height as i32),
        host_box,
        "and it is taken of the host's own client area, before the collapse takes it off the screen"
    );
    assert_eq!(
        the_colour_in(&pixels, width, 0, 0),
        A_COLOUR_OF_ITS_OWN.0,
        "and it is the host as the desktop showed it rather than the black its rows are erased to"
    );

    // The bubble's own menu row moves the pin into the taskbar, where the button is the host's.
    let forced = stills_forced();
    bubble_to_taskbar();
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool() && unsafe { IsIconic(host) }.as_bool(),
        "the host is the window the Shell's button belongs to, and it is put away visibly: a hidden \
         window wears no button at all"
    );
    assert_eq!(
        stills_forced(),
        forced + 1,
        "and the button is told to hold the app's own still for as long as the pin is away, whether or          not this row had a window of the pin's on the screen to picture"
    );
    let (pixels, width, _) = the_pin_picture().expect("and the picture taken at the collapse is kept");
    assert_eq!(
        the_colour_in(&pixels, width, 0, 0),
        A_COLOUR_OF_ITS_OWN.0,
        "a row that could take no picture of a window that is not there hands over no second one"
    );
}

/// **A film host moved from the bubble into the taskbar comes back with its film and its source.**
///
/// The fault this is a regression test for: the pin restored from the taskbar this way had a black
/// band with a moving seekbar — the host and the source the Shell draws through it stayed down, or
/// the registration between them was gone, and the seek that followed had no picture to take. The
/// pin's own restore re-establishes the pair in the one pass it happens in (see
/// `pin_film_host::bring_the_film_back`).
#[test]
fn a_film_host_moved_from_the_bubble_into_the_taskbar_comes_back_with_its_film() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    // Safety: the pin and its host are this test's own, and every global written here is put back
    // when the stand is dropped.
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    let source = unsafe { super::pin_film_host_window::a_source_of_the_host(host) };
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(source, SW_SHOWNOACTIVATE);
    }

    collapse_pin_to(PinMinimizeTo::Bubble);
    bubble_to_taskbar();
    super::pin_film_host_window::let_the_posted_shows_land();
    assert!(
        !unsafe { IsWindowVisible(source) }.as_bool(),
        "the source goes down with the host it is drawn in"
    );

    // What the Shell does when the host's button is clicked: it restores the window it holds, and
    // nothing of this app's was pressed for it (see `the_pin_came_back`).
    unsafe {
        let _ = ShowWindow(host, SW_RESTORE);
    }
    assert!(
        matches!(take_pin_command(), Some(PinCommand::CameBack)),
        "the pin's state comes back with the window the Shell restored"
    );

    restore_pin(&a_restored_window(), PinRestore::TheShell);
    assert!(!pin_is_collapsed(), "the pin is up again");
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool() && !unsafe { IsIconic(host) }.as_bool(),
        "and the film's host is on the screen with it"
    );
    super::pin_film_host_window::let_the_posted_shows_land();
    assert!(
        unsafe { IsWindowVisible(source) }.as_bool(),
        "and the source the Shell draws through the host is up again"
    );
    assert!(
        pin_film_thumbnail::the_standing_source_process(host).is_some(),
        "and the film's registration stands for that host again: a host with no registration in it \
         is the black it erases to behind the picture"
    );

    // Safety: the source is this test's own window.
    unsafe {
        let _ = DestroyWindow(source);
    }
}

/// **A pin restored before its film's hide has landed still brings the film back.**
///
/// The fault this is a regression test for: a hide of the source is `ShowWindowAsync` — a post to
/// another process — so at the restore the window can still be up while the hide is on its way. The
/// restore used to read the window, find it up, ask for nothing, and let the queued hide land last:
/// the host on the screen with the source down, which is the black band with a moving seekbar and a
/// seek that changes nothing, for as long as no later pass asks again. What the restore compares
/// against is the last ask rather than the window (see `pin_film_thumbnail::set_the_source_visible`).
#[test]
fn a_film_host_restored_before_its_films_hide_has_landed_brings_the_film_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    // Safety: the pin and its host are this test's own, and every global written here is put back
    // when the stand is dropped.
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    let source = unsafe { super::pin_film_host_window::a_source_of_the_host(host) };
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(source, SW_SHOWNOACTIVATE);
    }

    collapse_pin_to(PinMinimizeTo::Bubble);
    bubble_to_taskbar();
    // **The posted hides are deliberately left where they are**: that is the state this test is
    // about, and the drain below is what lets them land after the restore has said what it wants.
    assert!(
        unsafe { IsWindowVisible(source) }.as_bool(),
        "the hide the collapse asked for is still on its way to the player"
    );

    unsafe {
        let _ = ShowWindow(host, SW_RESTORE);
    }
    restore_pin(&a_restored_window(), PinRestore::TheShell);
    super::pin_film_host_window::let_the_posted_shows_land();
    assert!(
        unsafe { IsWindowVisible(source) }.as_bool(),
        "the hide that was in flight does not land last: the source is up with the pin that asked for it"
    );
    assert!(
        pin_film_thumbnail::the_standing_source_process(host).is_some(),
        "and the film's registration stands for that host"
    );

    // Safety: the source is this test's own window.
    unsafe {
        let _ = DestroyWindow(source);
    }
}

/// **The size the Shell sends is the pin coming back, even while the host is still iconic.**
///
/// The fault this is a regression test for: the Shell restores the window it holds a message at a time,
/// and the gate that reads the restore also asked the window whether it was iconic. A `SIZE_RESTORED`
/// that arrived before the iconic bit had gone was refused — the pin stood on the screen with a black
/// band, a seek that changed nothing, and the film coming back only when some later pass happened to
/// show the window, which is the "sometimes it recovers by itself" of the user's own report. What the
/// gate asks is whether the window is hidden: a window the Shell is bringing back is not (see
/// `the_pin_came_back`).
#[test]
fn a_film_host_still_iconic_when_the_size_says_it_came_back_is_the_pins_restore() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    // Safety: the pin and its host are this test's own, and every global written here is put back
    // when the stand is dropped.
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });

    collapse_pin_to(PinMinimizeTo::Bubble);
    bubble_to_taskbar();
    assert!(
        unsafe { IsIconic(host) }.as_bool() && unsafe { IsWindowVisible(host) }.as_bool(),
        "the host is the window the Shell's button put away, and it is visibly away"
    );

    // What the window is sent as it comes back, before its iconic bit has gone: the pin asked for a
    // restore is a pin whose state has to follow the window (see `the_pin_came_back`).
    unsafe {
        SendMessageW(host, WM_SIZE, WPARAM(SIZE_RESTORED as usize), LPARAM(0));
    }
    assert!(
        matches!(take_pin_command(), Some(PinCommand::CameBack)),
        "the pin comes back with the window the Shell is restoring"
    );
}

/// **A pin whose host the Shell has not finished restoring comes up with its film anyway.**
///
/// The fault this is a regression test for: the window the Shell restores is the host, and the size
/// that says it is back can arrive before its iconic bit has gone. The restore left a still-iconic
/// host alone — the picture never came back, the band was black, and only whatever later pass
/// happened to show the host recovered it, which is the "sometimes it plays perfectly" of the user's
/// own report. A pin the user asked for is up with its film, so the host is asked up rather than
/// waited for (see `pin_film_host::bring_the_film_back`).
#[test]
fn a_film_host_still_iconic_when_its_pin_is_restored_comes_back_up_with_its_film() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    // Safety: the pin and its host are this test's own, and every global written here is put back
    // when the stand is dropped.
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    let source = unsafe { super::pin_film_host_window::a_source_of_the_host(host) };
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(source, SW_SHOWNOACTIVATE);
    }

    collapse_pin_to(PinMinimizeTo::Bubble);
    bubble_to_taskbar();
    super::pin_film_host_window::let_the_posted_shows_land();
    assert!(
        unsafe { IsIconic(host) }.as_bool(),
        "the host is the window the Shell's button put away"
    );

    // The pin's own restore, with the host left in the state the Shell's restore can be read in
    // while it is still finishing.
    restore_pin(&a_restored_window(), PinRestore::TheShell);
    assert!(
        !unsafe { IsIconic(host) }.as_bool() && unsafe { IsWindowVisible(host) }.as_bool(),
        "the host is up with the pin the user asked for rather than left to a later pass"
    );
    super::pin_film_host_window::let_the_posted_shows_land();
    assert!(
        unsafe { IsWindowVisible(source) }.as_bool(),
        "and the source the Shell draws through it is up again"
    );

    // Safety: the source is this test's own window.
    unsafe {
        let _ = DestroyWindow(source);
    }
}

/// A picture already taken is not taken down by a window that is not there any more.
///
/// The fault this is a regression test for: the user's own report of a button whose peek held nothing
/// at all "in a particular order". Every road that ends the window the pin is drawn by — a film that
/// has run out, a player killed by a swap or by a display change, a relaunch — leaves this app
/// holding a handle to a window that is gone for a moment, and the tick settled the mirror against
/// that handle: the registration was unregistered and the stand-in was re-sized to the square it is
/// created at. What is left of the stand-in after that is its own surface, which this app has never
/// painted — and the frame a collapse takes is taken off *that* surface, so a minimize in the middle
/// of one of those roads left the button's peek empty for the rest of the run (see `sync_pin_mirror`,
/// `mirror_source`).
#[test]
fn a_picture_already_taken_is_not_taken_down_by_a_window_that_is_not_there() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_video = VIDEO_HWND.load(Ordering::SeqCst);
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    register_the_stand_in_class();
    forget_the_stand_in();

    // The window the pin is drawn by, standing in the pin's band and on the screen.
    let player = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a player's window")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen — which is what makes
    // it the film's picture rather than one more window the Shell composes nothing for.
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
    }
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    let player_box = unsafe { the_client_box(player) };

    // The film the pin is showing, so that every ask below is asked about the kind it is about rather
    // than about whatever was on screen before this test (see `current_media_type`).
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    let stand_in = the_stand_in();
    // Safety: the window is this app's own stand-in, asked only to be up — a hidden window has no
    // taskbar button and no picture to hold (see `showing_taskbar_window`).
    unsafe {
        let _ = ShowWindow(stand_in, SW_SHOWNOACTIVATE);
    }

    // The picture as the tick takes it while the film is playing.
    sync_pin_mirror(Some(MediaType::Video));
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        player_box,
        "a film's picture is its player's window while that window is up"
    );

    // And the player is gone — not hidden but *destroyed*, with this app still holding the number of
    // the window it was, which is the instant a film that has run out, a swap's own take-down and a
    // display change all pass through.
    // Safety: the window is this test's own and it is only asked to go away.
    unsafe {
        let _ = DestroyWindow(player);
    }

    sync_pin_mirror(Some(MediaType::Video));

    assert_eq!(
        unsafe { the_client_box(stand_in) },
        player_box,
        "a picture once taken is the picture to hold: the stand-in is left the box the frame was \
         taken in rather than re-made for a window that is not there, which is what leaves the hover \
         with nothing in it"
    );

    // And a minimize in the middle of that road is the same rule where it costs the most: what the
    // collapse takes into the button is the picture standing in the stand-in, so a stand-in stripped
    // of it is a button whose peek is empty for the rest of the run (see `keep_the_pin_for_the_taskbar`).
    if let Some(mut pin) = take_pin_for_a_test() {
        pin.collapsed = true;
        stand_pin(Some(pin));
    }
    unsafe {
        keep_the_pin_for_the_taskbar();
    }
    assert_eq!(
        unsafe { the_client_box(stand_in) },
        player_box,
        "so a pin put away after the film's window went keeps the frame that was on the screen"
    );

    // Safety: the stand-in is the app's own, and the handles it holds are put back as they were.
    forget_the_stand_in();
    VIDEO_HWND.store(previous_video, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    stand_pin(previous_pin);
}

/// A button is not shown for a pin that is put away into its round bubble.
///
/// The fault this is a regression test for: `Pin Mode → Always On Top` switched off while a pin was
/// put away into the bubble showed the stand-in — minimized, at the square it is created at, and with
/// no registration in it — because that road shows a button without asking which of the two put the
/// pin away. Two stand-ins stood over one pin, and the one on the taskbar held no picture of anything,
/// since every window of the pin's was off the screen (see `show_pin_taskbar`, `a_button_is_owed`).
#[test]
fn a_button_is_not_shown_for_a_pin_put_away_into_its_bubble() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let before = pin_always_on_top();

    register_the_stand_in_class();
    forget_the_stand_in();
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    // The round bubble a collapse into the bubble leaves, standing on the screen: a window of this
    // test's own standing in for it, since the bubble is painted by the loop's own thread.
    let bubble = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            TASKBAR_OFFSCREEN,
            TASKBAR_OFFSCREEN,
            120,
            120,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the bubble")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen — which is what says
    // the pin is reached by the bubble rather than by the taskbar (see `a_button_is_owed`).
    unsafe {
        let _ = ShowWindow(bubble, SW_SHOWNOACTIVATE);
    }
    PIN_BUBBLE_HWND.store(bubble.0 as isize, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));

    // Safety: the road the setting's own toggle takes, which is a road into a button that knows
    // nothing of where the pin was put away (see the `PinChanged` arm of the loop's tick).
    unsafe {
        show_pin_taskbar();
    }

    assert!(
        showing_taskbar_window().is_none(),
        "a pin put away into the bubble is reached by the bubble and not by the taskbar, so no \
         button is owed it — and one shown anyway is a button with nothing in it, since every window \
         of the pin's is off the screen"
    );

    // Safety: the bubble's window is this test's own and the stand-in is the app's own; the handles
    // the app holds are put back as they were.
    unsafe {
        let _ = DestroyWindow(bubble);
    }
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    forget_the_stand_in();
    stand_pin(previous_pin);
}

/// A pin that stands in a window of its own is minimized *as a window* when it is put away into the
/// taskbar: the button it wears is that window's own, and hiding the window is the one thing that takes
/// the button away with it.
///
/// **The fault this is a regression test for.** `Minimize → To Taskbar` took every pinned window off the
/// screen and then stood the app's own taskbar window in its place — and a real pin is refused that
/// stand-in, because the window it stands in brings a button of its own. So the collapse hid the only
/// window that had a button and put nothing in its place: the pin left the screen and the taskbar
/// together, with nothing anywhere left to bring it back (see `collapse_pin_to` and `a_button_is_owed`).
/// A window of the Shell's is minimized by the Shell and restored by it, so the collapse asks Windows for
/// the one thing Windows already does rather than hiding what Windows is drawing a button for.
#[test]
fn a_pin_that_is_a_window_keeps_its_button_when_it_is_minimized_to_the_taskbar() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);

    /// Off every display, so that the window this test puts up is no window of the user's — and wide
    /// enough that the pin's own chrome has room in it.
    const NO_MANS_LAND: i32 = -20000;
    let box_ = (
        NO_MANS_LAND,
        NO_MANS_LAND,
        NO_MANS_LAND + 320,
        NO_MANS_LAND + 240,
    );

    let mut pin = overlay_pin(box_, PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    // The stand-in is offered a class and no window of its own, so that a stand-in the collapse asks for
    // is one it can really make: what is asserted below is that nothing is owed it, not that it is
    // impossible to make.
    forget_the_stand_in();
    register_the_stand_in_class();
    pin_real_window::stand_real_pin_window_for_a_test(None);
    let pin_hwnd =
        unsafe { pin_real_window::real_pin_window_at(box_) }.expect("the pin's own window");
    // On the screen, off every display: the state a pin's window is in when its button is owed.
    // Safety: the window is this test's own.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    }

    collapse_pin_to(PinMinimizeTo::Taskbar);

    assert!(
        pin_is_collapsed(),
        "the pin is put away, and what is playing behind its window is held by that"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "and the minimize this app asked its own window for is no restore: the window is shown on its \
         way down to the button, and a `SIZE_RESTORED` read as the Shell's would bring the pin back in \
         the same breath it was put away in"
    );
    assert!(
        unsafe { IsWindowVisible(pin_hwnd) }.as_bool() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "the pin's own window is the one Windows minimizes: a window that is minimized keeps its \
         taskbar button, and a window that is hidden takes the button away with it"
    );
    assert!(
        showing_taskbar_window().is_none(),
        "and nothing stands in for it: one pin wears one button, and this one is the window's own"
    );

    // Safety: the window is this test's own, and every global written here is put back as it was.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
    forget_the_stand_in();
}

/// A pin put away as a window comes back as one: the Shell restores the window, and the pin's own state
/// follows it.
///
/// **The fault this is a regression test for.** A collapse into the taskbar is a state this app holds —
/// held-back previews, parked playback, and a minimize that is answered by nothing because the pin
/// believes it is away — while the window a real pin stands in is put back by a button of the Shell's
/// that asks this app nothing. Left alone, the pin is on the screen and still away as far as the app is
/// concerned: what it plays stays parked, and the next minimize does not land (see `collapse_pin_to`).
#[test]
fn a_pins_own_window_brought_back_is_the_pins_own_restore() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);

    const NO_MANS_LAND: i32 = -20000;
    let box_ = (
        NO_MANS_LAND,
        NO_MANS_LAND,
        NO_MANS_LAND + 320,
        NO_MANS_LAND + 240,
    );

    let mut pin = overlay_pin(box_, PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    forget_the_stand_in();
    register_the_stand_in_class();
    pin_real_window::stand_real_pin_window_for_a_test(None);
    let pin_hwnd =
        unsafe { pin_real_window::real_pin_window_at(box_) }.expect("the pin's own window");
    // Safety: the window is this test's own.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    }

    collapse_pin_to(PinMinimizeTo::Taskbar);
    assert!(
        pin_is_collapsed() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "the pin is put away as a window, which is the state this test is about"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "and the ask this app made of its own window is no restore of the Shell's"
    );

    // What the button Windows draws for that window sends when it is clicked: the Shell's own restore,
    // which reaches the window rather than any stand-in of this app's.
    // Safety: a message to a window of this app's own, with the numbers the Shell sends with it.
    unsafe {
        SendMessageW(
            pin_hwnd,
            WM_SYSCOMMAND,
            WPARAM(SC_RESTORE as usize),
            LPARAM(0),
        );
    }

    assert!(
        !unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "the Shell's own call puts the window back on the screen"
    );
    assert!(
        matches!(take_pin_command(), Some(PinCommand::CameBack)),
        "and the pin's state comes back with the window: a pin left believing it is away holds \
         everything a collapse holds back, and answers no minimize with anything"
    );

    // And the road the answer takes, which is the restore without a window of the app's pressed for it:
    // the pin is up again, and no window of the app's is put in front of the one the Shell chose.
    restore_pin(&a_restored_window(), PinRestore::TheShell);
    assert!(!pin_is_collapsed(), "the pin is up again");

    // Safety: the window is this test's own, and every global written here is put back as it was.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
    forget_the_stand_in();
}

/// The bubble's own `To Taskbar` row is the same door: a pin that stands in a window of its own is put
/// in the taskbar as a window, whatever door it was asked from.
///
/// The fault this is a regression test for is the one above, reached from the other side: a pin inside
/// the round bubble has its window *hidden*, and the row that moves it to the taskbar put the app's
/// stand-in up instead of the window — which a real pin is refused — so the button the row was asked for
/// was the one thing that did not appear (see `bubble_to_taskbar`).
#[test]
fn a_pin_that_is_a_window_goes_to_the_taskbar_from_its_bubble_as_a_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);

    const NO_MANS_LAND: i32 = -20000;
    let box_ = (
        NO_MANS_LAND,
        NO_MANS_LAND,
        NO_MANS_LAND + 320,
        NO_MANS_LAND + 240,
    );

    let mut pin = overlay_pin(box_, PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    forget_the_stand_in();
    register_the_stand_in_class();
    pin_real_window::stand_real_pin_window_for_a_test(None);
    let pin_hwnd =
        unsafe { pin_real_window::real_pin_window_at(box_) }.expect("the pin's own window");
    // Safety: the window is this test's own.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    }

    // The bubble first, which is the state the row is asked from: the pin's window goes down with it.
    collapse_pin_to(PinMinimizeTo::Bubble);
    assert!(
        pin_is_collapsed() && !unsafe { IsWindowVisible(pin_hwnd) }.as_bool(),
        "a pin in the bubble has its window down, which is what the row has to undo"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "a pin hidden into the bubble is not one the Shell brought back: nothing of this app's asked the \
         window down there, so nothing of it is a restore"
    );

    bubble_to_taskbar();

    assert_eq!(
        take_pin_command(),
        None,
        "and the row's own minimize is no restore either, though the window is shown to become iconic"
    );
    assert!(
        pin_is_collapsed(),
        "the pin is still put away, only somewhere else"
    );
    assert!(
        unsafe { IsWindowVisible(pin_hwnd) }.as_bool() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "and it is the pin's own window that is minimized: a window that is hidden wears no button, so \
         the taskbar the row asks for is a taskbar with nothing in it"
    );
    assert!(
        showing_taskbar_window().is_none(),
        "and nothing stands in for it: one pin wears one button, and this one is the window's own"
    );

    // Safety: the window is this test's own, and every global written here is put back as it was.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
    hide_pin_bubble();
    forget_the_stand_in();
}

/// An update brings a pin that is put away into the taskbar back on top of the listing the file was
/// picked in — and leaves it where the minimize put it where `Pin Mode → Update Preview → Move To Top`
/// is off.
///
/// **The fault this is a regression test for.** `Move To Top` is the row that says an update brings the
/// pin over the listing it follows, and a put-away pin has no window in any band to bring: the file was
/// swapped into a window nobody could see, behind the taskbar button the minimize left, so the row
/// answered nothing at all for a pin the user had minimized — "even the minimized pin should show on top
/// now" (see `bring_a_put_away_pin_back`).
///
/// What the pin is brought back **to** is the window the Shell minimized, and the front is asked for
/// after it is back: a window that is off the screen has no place of its own in the z-order, and the
/// place a hidden window is shown into is behind everything opened over it since (see
/// `show_pinned_window`).
#[test]
fn an_update_brings_a_pin_put_away_into_the_taskbar_back_on_top() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);
    let updates = CONFIG
        .lock()
        .map(|config| config.pin_update_enabled)
        .unwrap_or(true);

    /// Off every display, so that the windows this test puts up are no windows of the user's — and wide
    /// enough that the pin's own chrome has room in it.
    const NO_MANS_LAND: i32 = -20000;
    let box_ = (
        NO_MANS_LAND,
        NO_MANS_LAND,
        NO_MANS_LAND + 320,
        NO_MANS_LAND + 240,
    );

    let mut pin = overlay_pin(box_, PinChrome::always());
    pin.collapsed = false;
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    // The stand-in is offered a class and no window of its own, so that a stand-in the collapse asks for
    // is one it can really make: what is asserted below is that nothing is owed it, not that it is
    // impossible to make.
    forget_the_stand_in();
    register_the_stand_in_class();
    pin_real_window::stand_real_pin_window_for_a_test(None);
    let pin_hwnd =
        unsafe { pin_real_window::real_pin_window_at(box_) }.expect("the pin's own window");

    // The window the loop holds, which the property a restore asks about is asked of: there is always
    // one, and here it is the pin's own — the own-window pin is drawn into that window alone.
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
    // Safety: the window is this test's own.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    }

    // And the listing the file is picked in: a window opened after the pin's, which is what puts it in
    // front of the pin — the arrangement an update is answered from.
    let listing = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            NO_MANS_LAND,
            NO_MANS_LAND,
            320,
            240,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the listing")
    };
    // Safety: the window is this test's own, and it is asked to be on the screen — an offscreen window
    // is still one the Shell lists, which is the whole of what the z-order below is read from.
    unsafe {
        let _ = ShowWindow(listing, SW_SHOWNOACTIVATE);
    }
    assert!(
        unsafe { stands_over(listing, pin_hwnd) },
        "the window this test opened after the pin's is the one in front of it"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_enabled = true;
        config.pin_update_move_to_top = true;
    }

    // The minimize: the pin goes away as a window, which is the state the update below is answered from.
    collapse_pin_to(PinMinimizeTo::Taskbar);
    assert!(
        pin_is_collapsed() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "the pin is put away, and the window it stands in is the one Windows minimized"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "and the ask this app made of its own window is no restore of the Shell's"
    );

    // The update: the file a pick names is the file the pin comes back up on, and the window comes back
    // with it, in front of the listing the pick was made in.
    assert!(
        an_update_brings_a_put_away_pin_back(),
        "with `Move To Top` on, an update is the pin coming forward — and a pin that is away has nothing
         to bring forward but itself"
    );
    // The file the pin is showing, which is the file this door answers for: an update's own pick names
    // a file nothing has read yet, and the window is not put up until that file is on it (see
    // `bring_a_put_away_pin_back` and `a_put_away_pin_comes_back_on_the_file_the_update_names_rather_than_before_it`).
    let on_the_pin = pinned_path().expect("the file the pin is showing");
    assert!(
        bring_a_put_away_pin_back(&on_the_pin),
        "so the pin the update names a file for is the pin that comes back up"
    );

    assert!(
        !pin_is_collapsed(),
        "the pin is up again, and what the collapse held back is the tick's to give back"
    );
    assert!(
        unsafe { IsWindowVisible(pin_hwnd) }.as_bool() && !unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "and it is the window Windows minimized that comes back, rather than a stand-in for it: a pin
         put away as a window is the window minimized, and the window is what the update brings up"
    );
    assert!(
        unsafe { stands_over(pin_hwnd, listing) },
        "the pin is put in front of the window the file was picked in, which is the one thing an
         update owes a pin it brings forward"
    );
    assert_eq!(
        unsafe { GetWindowLongPtrW(pin_hwnd, GWL_EXSTYLE) } as u32 & WS_EX_TOPMOST.0,
        0,
        "and no band is kept for it: `Pin Mode → Always On Top` is the only thing that keeps a pin over
         other programs, and the next click on another window puts this one behind it"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "and the restore this road asked for is no restore of the Shell's to answer as well"
    );

    // And the row switched off: an update is a file swapped into the pin where it stands, and a pin that
    // stands nowhere stays away until the user brings it back.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_move_to_top = false;
    }
    collapse_pin_to(PinMinimizeTo::Taskbar);
    let _ = take_pin_command();
    assert!(
        pin_is_collapsed() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "the pin is put away again, which is the state this half is about"
    );
    assert!(
        !an_update_brings_a_put_away_pin_back(),
        "with the row off an update does not bring a put-away pin back"
    );
    assert!(
        !bring_a_put_away_pin_back(&on_the_pin),
        "and the road answers that there was nothing to bring back"
    );
    assert!(
        pin_is_collapsed() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "so the file is swapped into the pin where the minimize put it, behind the button that stands
         for it, until the user brings it back themselves"
    );

    // Safety: both windows are this test's own, and every global written here is put back as it was.
    unsafe {
        let _ = DestroyWindow(listing);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
    forget_the_stand_in();
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_enabled = updates;
        config.pin_update_move_to_top = move_to_top;
    }
}

/// A put-away pin comes back on the file an update names, and not a tick before that file is on its
/// window: the window is not put up on the file it was showing for as long as the read and the decode
/// take.
///
/// **The fault this is a regression test for.** The restore was answered where the pick was read, and
/// a pick names a file nothing has read yet — so the window came up on the file before it and the new
/// one was swapped in a tick or more later, a flash between the two images every time an update brought
/// a put-away pin back (see `bring_a_put_away_pin_back`).
///
/// What the door is asked here is the pin's own state rather than the caller's word, because that is
/// what leaves the order to the window: the install writes the file and the media in one breath (see
/// `install_pinned_media`), so a pin naming the file is a pin showing it — and until then there is
/// nothing for the window to come back on.
#[test]
fn a_put_away_pin_comes_back_on_the_file_the_update_names_rather_than_before_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_own = pin_is_its_own_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);
    let updates = CONFIG
        .lock()
        .map(|config| config.pin_update_enabled)
        .unwrap_or(true);

    /// Off every display, so that the window this test puts up is no window of the user's — and wide
    /// enough that the pin's own chrome has room in it.
    const NO_MANS_LAND: i32 = -20000;
    let box_ = (
        NO_MANS_LAND,
        NO_MANS_LAND,
        NO_MANS_LAND + 320,
        NO_MANS_LAND + 240,
    );

    let mut pin = overlay_pin(box_, PinChrome::always());
    pin.collapsed = false;
    let showing_before = pin.path.clone();
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    // The stand-in is offered a class and no window of its own, so that a stand-in the collapse asks
    // for is one it can really make: what is asserted below is that the pin stays away, not that no
    // button could have stood for it.
    forget_the_stand_in();
    register_the_stand_in_class();
    pin_real_window::stand_real_pin_window_for_a_test(None);
    let pin_hwnd =
        unsafe { pin_real_window::real_pin_window_at(box_) }.expect("the pin's own window");
    // The window the loop holds, which the property a restore asks about is asked of: there is always
    // one, and here it is the pin's own.
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
    // Safety: the window is this test's own.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    }

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_enabled = true;
        config.pin_update_move_to_top = true;
    }

    // The minimize: the pin goes away as a window, which is the state the update below is answered
    // from.
    collapse_pin_to(PinMinimizeTo::Taskbar);
    assert!(
        pin_is_collapsed() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "the pin is put away, and the window it stands in is the one Windows minimized"
    );

    // The file the update names, which nothing has read yet: the pin still has the file before it, so
    // there is nothing to come back on, and a window put up here would show that file for as long as
    // the read and the decode take.
    let update = PathBuf::from("picture-after-it.png");
    assert_ne!(
        update, showing_before,
        "the update names another file than the one the pin is showing"
    );
    assert!(
        !bring_a_put_away_pin_back(&update),
        "a restore answered before the file is on the window is no restore: the pin is not showing the file it names"
    );
    assert!(
        pin_is_collapsed() && unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "so the pin stays where the collapse left it rather than coming up on the file before it, \
         which is the flash the user sees between the two images"
    );

    // The install's own write, which is where the file an update names becomes the file the pin is
    // showing and the media behind it the window's (see `install_pinned_media`). The show that follows
    // it paints before the window appears, so the window comes up on this file and no other.
    if let Some(mut state) = pin_state() {
        if let Some(pin) = state.pin_mut() {
            pin.path = update.clone();
        }
    }

    assert!(
        bring_a_put_away_pin_back(&update),
        "and the pin comes back where the file it was brought back for is the file on it"
    );
    assert!(!pin_is_collapsed(), "the pin is up again");
    assert!(
        unsafe { IsWindowVisible(pin_hwnd) }.as_bool() && !unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "and the window Windows minimized is the one that comes back"
    );

    // Safety: the window is this test's own, and every global written here is put back as it was.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
    forget_the_stand_in();
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_enabled = updates;
        config.pin_update_move_to_top = move_to_top;
    }
}

/// **The button is the setting's, and it is asked of the window that wears it.** `Pin Mode -> Always
/// On Top` on is a pin over everything already, which owes no taskbar button; off is an ordinary window
/// with an ordinary window's button. The pair's button belongs to the host that owns the pin's chrome,
/// and the pin's own window is the one that must not carry a second one.
///
/// The fault this is a regression test for is the intermittent one: the app-window bit was added to a
/// host that was *already on the screen* at the promotion, with a style re-read and no show, and a
/// button is the Shell's answer to a window being shown - so the pair wore the button it happened to
/// be shown with. And the row that changes the setting recomputed nothing, so a pin with the setting
/// turned on kept the button it was taken up with.
#[test]
fn the_pairs_taskbar_button_is_the_settings() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    let pin = pin_real_window::real_pin_window().expect("the pin's own window");
    let previous = pin_always_on_top();
    let wears_a_button = |hwnd: HWND| unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } & WS_EX_APPWINDOW.0 as isize != 0;

    // The setting off: the host is the window the Shell lists, and the pin's own window is not.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    assert!(
        wears_a_button(host),
        "a pin that is not Always On Top wears a taskbar button, on the host that owns its chrome"
    );
    assert!(
        !wears_a_button(pin),
        "and the pin's own window is listed with the host that owns it rather than beside it"
    );

    // The row turns `Always On Top` on with the pin standing: the button goes with it.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    unsafe { pin_film_host::the_button_follows_the_setting() };
    assert!(
        !wears_a_button(host) && !wears_a_button(pin),
        "a pin that is Always On Top owes no button, and the row takes the one it had"
    );

    // And off again: the same button comes back, on the same window.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }
    unsafe { pin_film_host::the_button_follows_the_setting() };
    assert!(
        wears_a_button(host) && !wears_a_button(pin),
        "and a setting turned off gives the pair its button back, once"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = previous;
    }
}

/// **A pair promoted out of a hover asks for no button while the setting says the pin is on top**,
/// and the bit is written before the host is shown rather than on an already-visible window's back.
///
/// The promotion is the road the intermittent button was reported on: it added `WS_EX_APPWINDOW` to a
/// host that was on the screen as a hover's, re-read the style and showed nothing, and the Shell's list
/// is a window being shown (see `pin_real_window::set_the_pin_window_appwindow`). Here the bit is
/// written with the rest of the pin's bits, which is the style write the promotion was already making.
#[test]
fn a_promoted_pair_owes_no_button_with_always_on_top() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let film = super::pin_film_host_window::StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { super::pin_film_host_window::HostedHover::up(&film, &path, 3.0) };
    let previous = pin_always_on_top();
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    pin_stands_in_its_own_window(true);
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert_eq!(
        pin_film_host::film_host(),
        Some(hover.host),
        "the hover's host is the pin's host"
    );
    assert!(
        unsafe { GetWindowLongPtrW(hover.host, GWL_EXSTYLE) } & WS_EX_APPWINDOW.0 as isize == 0,
        "a pair promoted out of a hover with Always On Top on asks for no button at all"
    );

    // And the setting off puts the button on that same host, since it is the window the Shell lists.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }
    unsafe { pin_film_host::the_button_follows_the_setting() };
    assert!(
        unsafe { GetWindowLongPtrW(hover.host, GWL_EXSTYLE) } & WS_EX_APPWINDOW.0 as isize != 0,
        "and a setting turned off gives the promoted pair its button"
    );

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = previous;
    }
    drop(hover);
}

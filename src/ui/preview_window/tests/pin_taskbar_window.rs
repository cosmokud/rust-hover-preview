//! The stand-in window the taskbar button belongs to: the frame a hover is drawn of, what a picture
//! of a put-away pin is taken from, and what a run does when the window the button belonged to has
//! gone.
//!
//! The window itself is the one part of the taskbar button a test can hold in its hand — the button
//! is the Shell's — so what is asserted here is the window: that nothing of Windows is drawn around
//! the picture, that a window the Shell has taken is made again rather than kept as a handle, and
//! that the picture of a pin put away into the bubble is the bubble.

use super::*;
use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, GetClientRect, GetWindowRect, IsIconic, IsWindow,
    RegisterClassExW, SendMessageW, ShowWindow, SC_RESTORE, SW_MINIMIZE, SW_SHOWNOACTIVATE,
    WA_ACTIVE, WINDOW_EX_STYLE, WM_ACTIVATE, WM_SYSCOMMAND, WNDCLASSEXW, WS_EX_APPWINDOW, WS_POPUP,
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

    restore_pin();

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
    restore_pin();
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

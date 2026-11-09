use super::*;
use windows::core::w;
use windows::Win32::Foundation::{LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, RegisterClassW, WINDOW_EX_STYLE, WNDCLASSW,
    WS_CHILD, WS_POPUP,
};

/// **A press outside every window of ours puts it behind the window it landed in**, and the press is the
/// whole of the question: nothing here asks whether Explorer came to the front, which is the transition a
/// press inside a listing never makes (see `a_press_puts_us_behind`).
///
/// The two halves are asserted apart because they answer different questions: the rule is about where a
/// press landed, and `press_is_inside_ours` is about whether it landed on us.
#[test]
fn a_press_outside_our_windows_puts_us_behind_them() {
    let elsewhere = press_is_inside_ours(HWND(0x3000 as *mut _), POINT { x: 4, y: 4 });
    assert!(
        !elsewhere,
        "a window that is neither a viewer's, the pin's menu's nor a page's is not one of ours"
    );
    assert!(
        a_press_puts_us_behind(true, elsewhere, true, false),
        "so the press that landed in it is the press the pin goes behind — a listing's tree, its rows, \
         the gap between them, another program and the desktop are all this one answer"
    );
    assert!(
        !a_press_puts_us_behind(false, elsewhere, true, false),
        "and no press is the press that asks for anything"
    );
}

/// **A press on a window of ours is the pin being used**, not the listing behind it being worked in: the
/// pin's own chrome, its bubble, its menu, its film's host, its page, its stand-in. Nothing is asked.
#[test]
fn a_press_inside_our_own_windows_asks_for_nothing() {
    assert!(
        !a_press_puts_us_behind(true, true, true, false),
        "the press that landed on us is not the press that goes behind anything"
    );
    assert!(
        !a_press_puts_us_behind(true, true, true, true),
        "and Always On Top answers the same way, for the other of its two reasons"
    );
}

/// **And the answer to "inside one of ours" is the broad one.** A pin's film is drawn by mpv in a window
/// of its own process, so the viewer's process alone answers no for it — and a press on the pin's own
/// picture would then put the pin behind that picture. `is_our_own_window` is the listing's question and
/// answers for a viewer's windows and for the browser's page inside the engine's; the pin's player window
/// is the half beside it (see `the_pin_of_the_window`).
///
/// The `Always On Top` flag is the other fact the press road reads, and it is read from what the preview
/// loop publishes rather than from the configuration, which a hook callback may not lock (see
/// `our_window_stands_over_everything`).
///
/// One test for the three published facts, so that none of them is written while another is read.
#[test]
fn the_press_road_reads_what_the_preview_loop_publishes() {
    let film = HWND(0xAB01 as *mut _);
    let (owner_before, video_before) =
        crate::app::seam::the_windows_of_the_pin(crate::app::seam::the_facts());
    let topmost_before = crate::app::seam::PIN_KEYS_TOPMOST.load(Ordering::Acquire);
    let standing_before = crate::app::seam::PIN_UP.load(Ordering::Acquire);

    // **A window of its own publishes a key owner and no pin of the run's**, which is why the press road
    // asks the owner and not `PIN_UP`: a detached window has had `PIN_UP` cleared with the detach, and a
    // gate written against it would ask nothing for the one window that is nobody's preview and still a
    // window of ours standing over the desktop.
    crate::app::seam::publish_pin_key_owner(0xAB00, 0xAB01);
    crate::app::seam::PIN_UP.store(false, Ordering::Release);
    assert!(
        a_window_of_ours_stands(),
        "a standing window is standing whatever the run calls it"
    );

    assert!(
        !is_our_own_window(film),
        "mpv's window is not the viewer's process and is not the listing's own"
    );
    assert!(
        press_is_inside_ours(film, POINT { x: 4, y: 4 }),
        "and it is a window of ours all the same: the press on the pin's film is the pin's"
    );

    // **And the hit test under a pin is not what answers for it.** A film's host answers `HTTRANSPARENT`
    // so the press reaches the pin's chrome, so the window `WindowFromPoint` names over the pin's own
    // picture is the listing the pin was fitted over — read as that listing's, a press on the pin's
    // picture puts the pin behind its own picture. The box the pin's window stands in answers instead.
    let box_before = crate::app::seam::preview_window_rect();
    crate::app::seam::PIN_WINDOW_BOX.publish(Some((100, 100, 300, 200)));
    assert!(
        press_is_inside_ours(HWND(0x3000 as *mut _), POINT { x: 150, y: 150 }),
        "a point inside the pin's own box is the pin's, whatever the hit test names under it"
    );
    assert!(
        !press_is_inside_ours(HWND(0x3000 as *mut _), POINT { x: 350, y: 150 }),
        "and one outside it is not"
    );
    crate::app::seam::PIN_WINDOW_BOX.publish(box_before);

    crate::app::seam::publish_pin_always_on_top(true);
    assert!(
        !a_press_puts_us_behind(
            true,
            press_is_inside_ours(HWND(0x3000 as *mut _), POINT { x: 4, y: 4 }),
            a_window_of_ours_stands(),
            our_window_stands_over_everything()
        ),
        "Always On Top is never put behind by a press: the setting keeps the band it holds"
    );

    crate::app::seam::publish_pin_key_owner(0, 0);
    assert!(
        !a_window_of_ours_stands(),
        "and with no window of ours standing a press on the desktop costs one atomic and asks nothing"
    );

    crate::app::seam::publish_pin_key_owner(owner_before, video_before);
    crate::app::seam::publish_pin_always_on_top(topmost_before);
    crate::app::seam::PIN_UP.store(standing_before, Ordering::Release);
}

/// A window of this test's own, shown nowhere: the tests below are about a tree of windows and not about a
/// pixel of one.
///
/// Safety: the caller puts every window it makes away, and none is ever shown.
unsafe fn a_window(class: windows::core::PCWSTR, parent: Option<HWND>) -> HWND {
    unsafe extern "system" fn an_empty_window(
        hwnd: HWND,
        msg: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        // Safety: the default procedure answers a message this window's own class never handles.
        unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
    }

    let instance = GetModuleHandleW(None).expect("this module");
    let wc = WNDCLASSW {
        lpfnWndProc: Some(an_empty_window),
        hInstance: instance.into(),
        lpszClassName: class,
        ..Default::default()
    };
    let _ = unsafe { RegisterClassW(&wc) };

    let style = if parent.is_some() { WS_CHILD } else { WS_POPUP };
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class,
            w!(""),
            style,
            -32000,
            -32000,
            320,
            240,
            parent.unwrap_or_default(),
            None,
            None,
            None,
        )
    }
    .expect("a window of the test's own")
}

/// **A page is drawn by the browser's own windows, which are nobody's preview's process.** The engine's
/// window is the viewer's; a page inside it is the browser's, so the viewer's process alone answers no for
/// it — and a press on a page the user is reading would be a press that puts the pin behind its own page
/// (see `is_our_own_window`, which asks `IsChild` for exactly this).
#[test]
fn a_press_inside_the_page_we_draw_in_is_ours() {
    let before = crate::app::seam::ENGINE_HWND.load(Ordering::Acquire);
    let class = w!("rhp-press-road-test-window");
    // Safety: every window here is this test's own, made where no display is, and put away below.
    let (engine, page) = unsafe {
        let engine = a_window(class, None);
        let page = a_window(class, Some(engine));
        (engine, page)
    };

    crate::app::seam::ENGINE_HWND.store(engine.0 as isize, Ordering::Release);
    assert!(
        unsafe { IsChild(engine, page) }.as_bool(),
        "the page stands inside the engine's window, which is the whole of the test"
    );
    assert!(
        press_is_inside_ours(page, POINT { x: 4, y: 4 }),
        "and a press on it is a press on ours: the pin's own page is not a place to go behind"
    );
    crate::app::seam::ENGINE_HWND.store(before, Ordering::Release);

    // Safety: the windows are this test's own and are never shown.
    unsafe {
        let _ = DestroyWindow(page);
        let _ = DestroyWindow(engine);
    }
}

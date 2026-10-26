//! The window FFmpeg draws a film into, and the process standing behind it: finding the
//! window, keeping it out of the foreground, sending it a key, and killing what is left over.

use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    GetClientRect, GetParent, SetParent, GWL_STYLE, SWP_FRAMECHANGED, SWP_SHOWWINDOW, WS_CAPTION,
    WS_CHILD, WS_CLIPCHILDREN, WS_CLIPSIBLINGS, WS_EX_APPWINDOW, WS_EX_ACCEPTFILES,
    WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_SYSMENU, WS_THICKFRAME, WS_VISIBLE,
};

use super::*;

// Expected executable names of the players a video or a sound is spawned as, used to verify a
// recorded PID still belongs to that process before killing it.
pub(super) const FFPLAY_IMAGE_NAME: &str = "ffplay.exe";
pub(super) const MPV_IMAGE_NAME: &str = "mpv.exe";

// Track the player's window HWND for cursor-over-preview detection
pub(super) static VIDEO_HWND: AtomicIsize = AtomicIsize::new(0);
// Track the player's process ID to re-find the window if needed
pub(super) static VIDEO_PID: AtomicU32 = AtomicU32::new(0);

/// Which program the id in `VIDEO_PID` is the id of, written with that id and read before
/// anything is killed by it.
///
/// **A kill by id alone is a kill of whatever the system has since handed that id to**, which is
/// why every one of them verifies the image name first (see `process_image_matches`), and what is
/// verified against therefore has to be the player that was actually begun rather than the one
/// this app began before mpv existed: a machine with mpv plays videos through mpv, and an mpv id
/// checked as `ffplay.exe` is a stray this app would never sweep and a player it would never
/// retire. The audio route's player is FFmpeg's and keeps the name it has always had.
static VIDEO_PROCESS_IMAGE: Lazy<Mutex<&'static str>> = Lazy::new(|| Mutex::new(FFPLAY_IMAGE_NAME));

/// Write down the player this app has just begun: its process, and the program that process runs.
///
/// The two are written together because they are one fact (see `VIDEO_PROCESS_IMAGE`), and this is
/// the one place either is written for a player — the audio route writes its own ffplay id through
/// here too, since a sound's player is swept by the same checks as a film's.
pub(super) fn note_player_process(pid: u32, image: &'static str) {
    VIDEO_PID.store(pid, Ordering::SeqCst);

    if let Ok(mut held) = VIDEO_PROCESS_IMAGE.lock() {
        *held = image;
    }
}

/// The image name the id in `VIDEO_PID` may be verified as, which is the player this app last
/// begun (see `note_player_process`).
///
/// It answers a second question as well — whether the window in a pinned band is mpv's, which is
/// the one window that travels with the band rather than being left behind a strip of the pin's
/// chrome (see `ensure_video_window_topmost`) — and it answers it here because a second record of
/// which player is on screen would be a second thing to keep in step with this one.
pub(super) fn video_process_image() -> &'static str {
    VIDEO_PROCESS_IMAGE
        .lock()
        .map(|held| *held)
        .unwrap_or(FFPLAY_IMAGE_NAME)
}

/// The image name the player of a route is begun, recorded and verified under.
///
/// Every route that begins a player at all is answered here rather than left to a fallback arm, so
/// that a fourth engine has to be answered for rather than quietly inheriting FFmpeg's name.
pub(super) fn player_image(route: VideoRoute) -> &'static str {
    match route {
        VideoRoute::Mpv => MPV_IMAGE_NAME,
        VideoRoute::Ffplay | VideoRoute::MediaEngine | VideoRoute::NoPreview => FFPLAY_IMAGE_NAME,
    }
}
// Guard to ensure we only run a single style-monitor thread.
pub(super) static NOACTIVATE_MONITOR_STARTED: AtomicBool = AtomicBool::new(false);
// Flag set when the system resumes from sleep, so the main loop can reset state.
pub(super) static RESUME_FROM_SLEEP: AtomicBool = AtomicBool::new(false);
// Flag set when the display under the preview changed — a monitor's scale, or the
// desktop rearranged — a frame the window proc can only discard. What was on screen
// is put back by the loop, which is the side that holds the hover it came from.
pub(super) static DISPLAY_RESET: AtomicBool = AtomicBool::new(false);

/// Data passed to the EnumWindows callback to find a player's window
pub(super) struct EnumWindowsData {
    pub(super) target_pid: u32,
    pub(super) found_hwnd: HWND,
    pub(super) best_area: i64,
}

/// Callback for EnumWindows to find a window belonging to a specific process
pub(super) unsafe extern "system" fn enum_windows_callback(
    hwnd: HWND,
    lparam: LPARAM,
) -> windows::Win32::Foundation::BOOL {
    let data = &mut *(lparam.0 as *mut EnumWindowsData);
    let mut window_pid: u32 = 0;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));

    if window_pid != data.target_pid {
        return windows::Win32::Foundation::BOOL(1);
    }

    // Prefer visible, top-level windows (skip hidden and owned/popups behind owners)
    if !IsWindowVisible(hwnd).as_bool() {
        return windows::Win32::Foundation::BOOL(1);
    }

    if let Ok(owner) = GetWindow(hwnd, GW_OWNER) {
        if !owner.is_invalid() {
            return windows::Win32::Foundation::BOOL(1);
        }
    }

    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return windows::Win32::Foundation::BOOL(1);
    }

    let width = (rect.right - rect.left).max(0) as i64;
    let height = (rect.bottom - rect.top).max(0) as i64;
    let area = width * height;
    if area <= 0 {
        return windows::Win32::Foundation::BOOL(1);
    }

    // Keep the largest candidate; this is typically the player's real output window.
    if area > data.best_area {
        data.best_area = area;
        data.found_hwnd = hwnd;
    }

    windows::Win32::Foundation::BOOL(1)
}

/// The extended styles this app asserts on a player's window every time it finds it.
///
/// One function because it is one list asserted on a window of another process
/// from a thread of its own, about five times a millisecond for as long as a film
/// plays, and a list written out twice is a list whose two copies drift: the
/// second one silently drops whatever the first has learned (see
/// `ensure_video_window_topmost`, which asserts the same list from the tick's own
/// road rather than one of its own).
///
/// **`WS_EX_TRANSPARENT` is what makes a click on the band reach nothing of the
/// player's own.** A video pin's band is alpha zero by design, so that the window
/// underneath is both seen and clicked through it — which is FFmpeg's window, and
/// its mouse bindings are not this app's to change: `left double-click toggle
/// full screen` is compiled into the player and no option unbinds it (`-draw_mouse
/// 0` hides the drawn pointer and leaves the binding live, verified against
/// ffplay 9.0.2). So a hand on the band entered fullscreen and relative-mouse
/// mode — the cursor warp and the flash — with no flag this app could pass to
/// prevent it. Passing the click through the player's own window is the window
/// property that does it, and it costs the player's own pointer bindings, which
/// is the trade: the band is this app's chrome, and this app's bar already carries
/// a seek that works.
///
/// `WS_EX_LAYERED` is deliberately *not* set with it: a layered window's contents
/// come from `UpdateLayeredWindow`, and this one is SDL's, which paints and
/// swaps its own surface. A video that stopped rendering would be a far worse
/// answer than a click that falls through to the desktop.
pub(super) fn player_ex_style_for(current: isize) -> isize {
    // **A film in the pin's own host wears none of this.** The list is a window of the desktop's own
    // wearing bits this app puts on another program's window: a hosted player is mpv's child under a
    // host the pin owns and disables, so it neither takes a press nor holds a band of its own (see
    // `pin_film_host`).
    if pin_film_host::the_pinned_player_was_begun_in_the_host() {
        return current;
    }

    let mut style = current
        | WS_EX_NOACTIVATE.0 as isize
        | WS_EX_TOOLWINDOW.0 as isize
        | WS_EX_TRANSPARENT.0 as isize;
    if !pinned() || pin_holds_the_front() {
        style |= WS_EX_TOPMOST.0 as isize;
    } else {
        style &= !(WS_EX_TOPMOST.0 as isize);
    }
    style
}

/// The host this app has recorded as holding the player `pid`, where it has recorded one.
///
/// It is where a player begun for a host is to be adopted from its own top-level window: mpv is
/// given the host as `--wid` and already draws there, while FFmpeg's player only gets a window of
/// the desktop's of its own, which is moved in once it exists.
pub(super) fn host_for_a_hosted_player(pid: u32) -> Option<HWND> {
    if pid == 0 {
        return None;
    }
    if pin_film_host::hosted_player_pid() == Some(pid) {
        return pin_film_host::film_host();
    }
    if hover_film_host::hover_film_host_player() == Some(pid) {
        return hover_film_host::hover_film_host();
    }
    None
}

/// Adopt a player's own top-level window into a host this app owns: the button press that begins an
/// FFmpeg film is already on screen at `child`, and the pin's (or hover's) band is `host`.
///
/// SDL2 has no embed switch on Windows, so the window cannot be asked to draw into the host — it is
/// restyled to a disabled `WS_CHILD` and re-parented after it exists, which is the `--wid` shape mpv
/// is given at its launch. The frame the window was born with is stripped here, and it was born off
/// every display (see `start_video_playback_where`), so no border is ever seen.
///
/// Idempotent: a window already parented to the host is only fitted to it, so the monitor that runs
/// every few milliseconds while a film plays can ask again without restyling a child the compositor
/// is drawing.
pub(super) unsafe fn adopt_the_players_window(child: HWND, host: HWND) -> bool {
    if child.is_invalid() || host.is_invalid() {
        return false;
    }

    if unsafe { GetParent(child) }.is_ok_and(|parent| parent == host) {
        unsafe { fit_the_hosted_player_to_the_host(child, host) };
        return true;
    }

    let style = unsafe { GetWindowLongPtrW(child, GWL_STYLE) };
    let style = style
        & !(WS_POPUP.0 as isize
            | WS_CAPTION.0 as isize
            | WS_THICKFRAME.0 as isize
            | WS_SYSMENU.0 as isize
            | WS_MINIMIZEBOX.0 as isize
            | WS_MAXIMIZEBOX.0 as isize)
        | (WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN).0 as isize;
    unsafe { SetWindowLongPtrW(child, GWL_STYLE, style) };

    // The bits that make a window of the desktop's one of the Shell's go; ACCEPTFILES is ffplay's
    // own and is left exactly as it was found.
    let ex = unsafe { GetWindowLongPtrW(child, GWL_EXSTYLE) };
    let ex = ex
        & !(WS_EX_TOPMOST.0 as isize
            | WS_EX_APPWINDOW.0 as isize
            | WS_EX_TOOLWINDOW.0 as isize
            | WS_EX_NOACTIVATE.0 as isize)
        | WS_EX_ACCEPTFILES.0 as isize;
    unsafe { SetWindowLongPtrW(child, GWL_EXSTYLE, ex) };

    let _ = unsafe { SetParent(child, host) };
    // `SetParent` puts the topmost bit back for a window that had one; the host wears none.
    unsafe { SetWindowLongPtrW(child, GWL_EXSTYLE, ex) };
    // Disabled, so a press on the picture is a press on the pin (see `pin_film_host`).
    let _ = unsafe { EnableWindow(child, false) };
    unsafe { fit_the_hosted_player_to_the_host(child, host) };
    true
}

/// Fit a hosted player's window to the host's client box: the band's box changes and SDL will not
/// follow its parent on its own.
pub(super) unsafe fn fit_the_hosted_player_to_the_host(child: HWND, host: HWND) {
    let mut client = RECT::default();
    if unsafe { GetClientRect(host, &mut client) }.is_err() {
        return;
    }
    let _ = unsafe {
        SetWindowPos(
            child,
            HWND::default(),
            0,
            0,
            (client.right - client.left).max(1),
            (client.bottom - client.top).max(1),
            SWP_FRAMECHANGED | SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW,
        )
    };
}

/// Style and raise a player's known window.
pub(super) unsafe fn apply_noactivate_to_hwnd(hwnd: HWND) -> bool {
    // **A hosted player's window is not this app's to style or to publish.** It is mpv's child under
    // a host the pin owns, and every road that would place, style or raise it is answered by the
    // host instead (see `player_ex_style_for`, `place_player_window`).
    if pin_film_host::the_pinned_player_was_begun_in_the_host()
        || hover_film_host::the_hover_player_was_begun_in_the_host(VIDEO_PID.load(Ordering::SeqCst))
    {
        return false;
    }

    // Store the video window HWND for cursor-over-preview detection
    VIDEO_HWND.store(hwnd.0 as isize, Ordering::SeqCst);

    // The decision this function is the raise half of, asked once so that the style it writes and
    // the raise it holds off are the same answer (see `player_order`).
    let order = player_order(
        video_process_image(),
        pinned(),
        pin_chrome_up(),
        pin_volume_open(),
        pin_menu_open(),
        pin_player_is_parked(),
    );

    // Add WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST and WS_EX_TRANSPARENT to its extended
    // style (see `player_ex_style_for`).
    let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    SetWindowLongPtrW(hwnd, GWL_EXSTYLE, player_ex_style_for(current_style));

    // The style is put on whatever else is happening, and the *raise* is not. This is the one
    // place every raise of the player's window goes through — the monitor thread calls it about
    // every hundred milliseconds for as long as a film plays, whether or not the tick has asked for
    // anything — and that is what made a pinned window's volume popup open and vanish: the popup is
    // drawn over the media band, the media band of a video is this very window, and a raise puts
    // this window on top of the pin's own window every hundred milliseconds however carefully the
    // tick is holding off. The tick's hold-off was the right idea against the wrong caller. A strip
    // of the pin's chrome is painted over the same band for a kind whose chrome is drawn over its
    // media, so a raise buries a showing strip the same way (see `pin_overlay_chrome`).
    //
    // So the raise waits while the popup is up, any strip of the pin's chrome is showing, or the
    // card's menu is up, and the style does not: `WS_EX_NOACTIVATE` is what keeps the player from
    // taking the keyboard, and that has to keep being asserted throughout, because the player's
    // window is created without it and this is the only thing that ever puts it on (see `Bug 2`'s
    // note over `try_apply_noactivate_style`).
    // Nothing is asked for on the way out — the popup's closing raises the pin's own window, a
    // strip coming up is answered by the tick's chrome block with the same raise, and the card's
    // menu is raised over the film by the tick's own `render_pin_menu` (see `toggle_pin_volume`,
    // `pin_chrome_up`); the next raise through here settles the order either way.
    //
    // **And it waits while a drag is parking the window, which is the same arrangement for the same
    // reason: this is a thread of its own, asking every hundred milliseconds whether the player's
    // window is still the window it styled.** Neither the drag's own raises nor the tick's can cover
    // it — they are two callers, and this one answers for a film nobody asked about. A park this
    // function undid would last from one monitor pass to the next, which at the pace it runs is a
    // hundred milliseconds of a drag with the film back on screen and the compositor re-blitting it
    // at the pointer's rate: the stutter the park was written for, and a band painted flat under a
    // window that is standing in it (see `pin_player_is_parked`).
    //
    // **What is held off here is the `Leave` answer and nothing else.** Those are the park, a strip
    // of the pin's chrome, an open volume popup and an open card's menu (see `pin_player_is_parked`,
    // `pin_chrome_up`, `pin_volume_open`, `pin_menu_open`): a raise over any of the last three puts
    // the film over what the user is looking at — the caption and the transport bar the pin draws
    // into the band over the media, and the panel the card's menu floats in (see the note above and
    // `pin_overlay_chrome`, `pin_menu_window`). mpv's film holds its own band while they are up
    // (`--ontop`, see `mpv_args`), the pin's own raise carries the press, and a band that has moved
    // is placed without a raise (see `ensure_pinned_sibling_box`). A collapsed pin holds off every
    // film alike: its band went with the window, so a raise here is a chrome-less film back over
    // Explorer beside the bubble (see `hide_pinned_windows`).
    //
    // **And a collapsed pin's film is put *away* rather than merely left unraised, which is the
    // whole of what a player begun behind a bubble costs without it.** A player is given its file
    // and a box and nothing else — neither one takes a switch that means "come up off the screen",
    // and both show the window they made the moment they have one (see `start_video_playback`,
    // `mpv_args`) — so the film a *walk* lands on behind a bubble stands on the desktop,
    // chrome-less, at the box the pin's own band had, while the bubble beside it goes on painting
    // the same film out of frames read from the file: two pictures of one film, one of them a window
    // nobody asked for.
    // What hid the film the collapse found was `hide_pinned_windows`, and it is one call at the
    // collapse rather than a tick — nothing hides a window that arrives afterwards.
    //
    // This is the one place that finds such a window as soon as it exists, which is why the hide is
    // here rather than at the launch: the launch has a process id and no window yet (the player puts
    // its window up after it has read the file's header), and the monitor thread walks the player's
    // process on the five-to-a-hundred-millisecond cadence it already keeps — so the newly begun
    // film is off the screen within a pass of the monitor, and stays off it for as long as the pin
    // is a bubble, however often the player shows its own window again.
    if pin_is_collapsed() {
        let _ = ShowWindow(hwnd, SW_HIDE);
        return true;
    }

    if order == PlayerOrder::Leave {
        return true;
    }

    // Raise the window above everything so it doesn't hide behind Explorer — and put the pin back
    // over it where one is up, so the pin's own chrome and the press on the picture stay on top
    // (see `place_player_window`).
    place_player_window(
        hwnd,
        order,
        (0, 0, 0, 0),
        SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
    );
    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    true
}

/// Apply WS_EX_NOACTIVATE style to a window
/// Returns true if the window was found and modified
pub(super) unsafe fn try_apply_noactivate_style(pid: u32) -> bool {
    // **A player in a host of this app's has no window of the desktop's to style.** mpv's child
    // under the host is already disabled, and there is nothing to enumerate or to publish: the
    // host stands in every top-level road instead (see `pin_film_host` and `hover_film_host`).
    // FFmpeg's player is the one that is not there yet: it makes its own top-level window, which
    // is adopted into the host here — once, and then never found again by the enumeration below.
    if let Some(host) = host_for_a_hosted_player(pid) {
        // **The thumbnail road's player is a window of the desktop's and is kept one.** mpv is begun
        // with no `--wid`, so what is found under this pid is the film's own window rather than a
        // child of the host's: it is adopted as the picture the host shows by a DWM thumbnail, and
        // never moved into the host the way FFmpeg's player is (see `pin_film_thumbnail`).
        if pin_film_host::hosted_player_pid() == Some(pid)
            && pin_film_host::the_pin_shows_its_film_by_thumbnail()
            && video_process_image() == MPV_IMAGE_NAME
        {
            if !pin_film_thumbnail::the_source_is_up(pid) {
                let mut data = EnumWindowsData {
                    target_pid: pid,
                    found_hwnd: HWND::default(),
                    best_area: 0,
                };
                let _ = EnumWindows(
                    Some(enum_windows_callback),
                    LPARAM(&mut data as *mut EnumWindowsData as isize),
                );
                if !data.found_hwnd.is_invalid() {
                    unsafe { pin_film_thumbnail::adopt_the_source(data.found_hwnd, pid) };
                }
            } else {
                // **And this is the pass that keeps the source invisible**: the style is put on at
                // adoption and mpv can come back from a video output reinit without it, so a source
                // whose layered bit is gone is a film standing over the pin until the next adopt — and
                // a window that is *the same handle* styled again is never adopted again (see
                // `pin_film_thumbnail::keep_the_source_invisible`).
                unsafe { pin_film_thumbnail::keep_the_source_invisible() };
            }
            return false;
        }

        let mut data = EnumWindowsData {
            target_pid: pid,
            found_hwnd: HWND::default(),
            best_area: 0,
        };
        let _ = EnumWindows(
            Some(enum_windows_callback),
            LPARAM(&mut data as *mut EnumWindowsData as isize),
        );
        if !data.found_hwnd.is_invalid() {
            unsafe { adopt_the_players_window(data.found_hwnd, host) };
        }
        return false;
    }

    // Reuse the window already found while it still exists and still belongs to
    // the player. This runs every few milliseconds for as long as a video plays,
    // so the desktop enumeration below is needed once, or again if ffplay
    // recreates its window (the cached handle stops being a window).
    let cached = VIDEO_HWND.load(Ordering::SeqCst);
    if cached != 0 {
        let hwnd = HWND(cached as *mut std::ffi::c_void);
        if IsWindow(hwnd).as_bool() {
            let mut window_pid: u32 = 0;
            GetWindowThreadProcessId(hwnd, Some(&mut window_pid));
            if window_pid == pid {
                return apply_noactivate_to_hwnd(hwnd);
            }
        }
        VIDEO_HWND.store(0, Ordering::SeqCst);
    }

    let mut data = EnumWindowsData {
        target_pid: pid,
        found_hwnd: HWND::default(),
        best_area: 0,
    };

    let _ = EnumWindows(
        Some(enum_windows_callback),
        LPARAM(&mut data as *mut EnumWindowsData as isize),
    );

    if !data.found_hwnd.is_invalid() {
        return apply_noactivate_to_hwnd(data.found_hwnd);
    }

    VIDEO_HWND.store(0, Ordering::SeqCst);
    false
}

/// Set WS_EX_NOACTIVATE on a window belonging to the given process
/// This prevents the window from stealing focus
/// Uses a singleton monitor thread so repeated previews don't spawn extra workers.
///
/// The thread spends its time between players waiting on an event rather than polling: a
/// window belongs to the player that is playing, and while none is, there is nothing to
/// re-assert — so what the wait is for is the hand-over a new player makes (see
/// `set_noactivate_for_process`), and a thread that woke every eighty milliseconds to find
/// that nothing had changed is a thread this app does not need. A wait that is never
/// signalled is not a lost wake-up either: the timeout it is given is the cadence the
/// window is kept in step at, so a hand-over that raced the wait is caught by the next one.
pub(super) fn ensure_noactivate_monitor() {
    if NOACTIVATE_MONITOR_STARTED
        .compare_exchange(false, true, Ordering::AcqRel, Ordering::Acquire)
        .is_err()
    {
        return;
    }

    std::thread::spawn(|| {
        let mut monitored_pid: u32 = 0;
        let mut pid_started = Instant::now();

        while RUNNING.load(Ordering::Acquire) {
            let pid = VIDEO_PID.load(Ordering::Acquire);

            if pid != monitored_pid {
                monitored_pid = pid;
                pid_started = Instant::now();
                if pid == 0 {
                    VIDEO_HWND.store(0, Ordering::SeqCst);
                }
            }

            if pid != 0 {
                unsafe {
                    let _ = try_apply_noactivate_style(pid);
                }

                let elapsed = pid_started.elapsed();
                let delay_ms = if elapsed < Duration::from_millis(250) {
                    5
                } else if elapsed < Duration::from_secs(2) {
                    20
                } else {
                    100
                };
                wait_for_noactivate_wake(delay_ms);
            } else {
                // Nothing is playing: the wait is the player's own start, which is what
                // signals this thread rather than a clock.
                wait_for_noactivate_wake(NOACTIVATE_IDLE_WAIT_MS);
            }
        }

        NOACTIVATE_MONITOR_STARTED.store(false, Ordering::Release);
    });
}

/// How long the monitor thread waits while nothing is playing, which is a bound on how long
/// it takes to notice the run ending rather than a cadence anything is kept in step at.
pub(super) const NOACTIVATE_IDLE_WAIT_MS: u64 = 500;

/// The event the monitor thread waits on: signalling it is a player's window having been
/// handed over (see `set_noactivate_for_process`). A handle that could not be created is the
/// null one, and a wait on that fails at once — which leaves the thread polling at the
/// timeout it was given, the way it did before there was an event at all.
///
/// It is kept as the number a handle is rather than as the handle, for the reason
/// `VIDEO_HWND` is: a handle is a raw pointer, and what is shared across threads here is a
/// number that the calls it is handed to take as a handle.
pub(super) static NOACTIVATE_WAKE: Lazy<isize> = Lazy::new(|| {
    unsafe { CreateEventW(None, false, false, None) }
        .map(|handle| handle.0 as isize)
        .unwrap_or_default()
});

/// The wake event as the handle a Windows call takes.
pub(super) fn noactivate_wake_handle() -> HANDLE {
    HANDLE(*NOACTIVATE_WAKE as *mut core::ffi::c_void)
}

/// Wait for a player's window to be handed over, or for `timeout_ms` to pass.
pub(super) fn wait_for_noactivate_wake(timeout_ms: u64) {
    let handle = noactivate_wake_handle();
    if handle.0.is_null() {
        // An event that could not be created is not one to wait on: a failed wait returns
        // at once, and a thread that returned at once every time is a thread spinning on a
        // machine that has no event. What is left is the sleep this was before there was an
        // event at all, which costs the wakeups it always did and nothing more.
        std::thread::sleep(Duration::from_millis(timeout_ms));
        return;
    }

    // A failed wait is the timeout's, and it leaves the caller doing what it would have done
    // anyway: looking at the process it is watching.
    let _ = unsafe { WaitForSingleObject(handle, timeout_ms as u32) };
}

/// Wake the monitor thread: a player is playing, and the window it draws in is the
/// monitor's to keep in step (see `ensure_noactivate_monitor`).
pub(super) fn wake_noactivate_monitor() {
    let _ = unsafe { SetEvent(noactivate_wake_handle()) };
}

pub(super) fn set_noactivate_for_process(pid: u32, image: &'static str) {
    note_player_process(pid, image);

    // First, do a few immediate synchronous checks with very tight timing
    // This minimizes the window where focus can be stolen
    unsafe {
        for _ in 0..10 {
            if try_apply_noactivate_style(pid) {
                // Found and modified - but keep monitoring in case window is recreated
                break;
            }
            // Very short spin-wait for the first attempts
            std::thread::yield_now();
        }
    }

    ensure_noactivate_monitor();
    // The monitor is woken rather than left to notice on its own clock: between players it
    // is waiting on this, and a player whose window appears while that wait runs is one the
    // thread would otherwise look at up to half a second later (see
    // `ensure_noactivate_monitor`).
    wake_noactivate_monitor();
}

/// A key as FFmpeg's player has to be given it, and the `lParam` that has to carry.
///
/// The three numbers are the virtual-key code, the scan code the same key has on a US layout
/// (`MapVirtualKey`, which is what a real `WM_KEYDOWN` for it carries), and the message
/// parameter built from the scan code. They are written out rather than computed because they
/// were measured rather than reasoned about, and the measurement overturned the obvious
/// answer: a `WM_KEYDOWN` posted with `lParam` of zero does *not* pause the player, and one
/// posted with the scan code does — every time.
///
/// That is not this app being fussy. FFmpeg's player is an SDL program, and SDL turns a Win32
/// key message into its own by looking the virtual-key code up to a scan code and the scan code
/// up to a key symbol; a scan code of zero is not a key, so the symbol it arrives as is not
/// `p`, and a player switching on the symbol it was given does nothing at all for a message
/// whose scan code is missing. Reading the argument as "the message needs its scan code" is
/// what the zero was missing, and no flag in the code says so. mpv is posted the same message and
/// reads its key the same way, so the one shape serves both.
///
/// The one bit deliberately *not* set is bit 30, which is Windows' own "this is an auto-repeat"
/// and which SDL reads as a repeat and drops: a posted pause that repeated would toggle, toggle
/// and toggle back, so a hand resting on the bar's button would flicker rather than hold. The
/// low bit is the repeat *count*, which a real key-down of one press is 1.
pub(super) fn ffplay_key_lparam(scan: u32) -> LPARAM {
    LPARAM((scan as isize) << 16 | 1)
}

/// A key the player is toggled with, and the scan code it is named by.
///
/// One key code for both players this app can be running, because both of them read it as `p`:
/// FFmpeg's player has `p` compiled in as its pause key, and the binding file an mpv video is
/// launched with holds `p cycle pause` (see `video_launch::MPV_PAUSE_BINDING`).
pub(super) const VIDEO_PAUSE_KEY: (u32, u32) = (b'P' as u32, 0x19);

/// The window belonging to `pid`, if the player this app believes is playing has a window of its
/// own up — which is not the same question as whether a player is running: a process opens its
/// window after it has read the file's header, so a player in the middle of starting is running
/// and has no window yet.
///
/// The pid is asked about as well as the window, and that is what a relaunch made necessary.
/// A relaunch now leaves two players alive at once — the one on screen and the one beginning over
/// it (see `retire_replaced_player`) — and `VIDEO_HWND` is a single published handle that names
/// whichever window the monitor found last. So "there is a window" is not a question that can be
/// asked of the handle alone: a key posted through it during the overlap would reach the player
/// being retired, which is a player this app is about to end, and the pin would be paused against
/// a picture that is on its way out. A handle that belongs to another pid is therefore not this
/// player's window at all, and neither is a handle whose player is not the one being played.
///
/// A handle that is not a window any more is cleared on the way out, because the only other paths
/// that clear it are the ones that end a player, and this is the one place that can notice a
/// window has gone on its own — a player that puts its window up, has it torn down, and goes on
/// playing is a player this app must not post a key to for ever after.
pub(super) fn video_window_for(pid: u32) -> Option<HWND> {
    if pid == 0 {
        return None;
    }

    // **A hover's hosted player has no window of the desktop's**, only mpv's child under this
    // app's host: that child is the player's own window, already disabled and non-activating, and
    // it is what the roads asking "is the player's window up" are answered with (see
    // `the_players_window_is_up`).
    if let Some(child) = hover_film_host::the_hover_players_window(pid) {
        return Some(child);
    }

    // The pin's host holds the film the same way: a hosted FFmpeg player is its child, and a key
    // has to reach that window rather than a top-level one that no longer exists (see
    // `pin_film_host::a_child_of_the_host`).
    if pin_film_host::hosted_player_pid() == Some(pid) {
        return pin_film_host::film_host()
            .and_then(|host| pin_film_host::a_child_of_the_host(host, pid));
    }

    let hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
    if hwnd_val == 0 {
        return None;
    }

    // SAFETY: `hwnd_val` is a window handle published by `apply_noactivate_to_hwnd` from the
    // handle Windows gave back for a player's own window, and it is cleared by the same paths
    // that end a player — so it is either a live window of some player of this app's or a handle
    // to nothing, which `IsWindow` settles before the handle is used for anything. A window that
    // is on its way out is discarded by the system rather than delivered to, which is the answer
    // `PostMessageW` gives back as a failure and what every caller falls back on.
    let hwnd = HWND(hwnd_val as *mut core::ffi::c_void);
    unsafe {
        if !IsWindow(hwnd).as_bool() {
            VIDEO_HWND.store(0, Ordering::SeqCst);
            return None;
        }

        let mut owner: u32 = 0;
        GetWindowThreadProcessId(hwnd, Some(&mut owner));
        (owner == pid).then_some(hwnd)
    }
}

/// Whether the player behind a pinned file is holding a level the pin no longer claims, and so is
/// not the player a pause key may be posted to.
///
/// The level a running player is playing at cannot be changed: FFmpeg's player is told nothing
/// once it is running, so a level is another player begun at it (see `restart_pinned_player`). So
/// letting a film go of its hold with a key — which is the cheap way out, and the one that keeps the
/// player that is already there — would start it at the level it was stopped at while the bar is
/// drawn at the level the hand set. Answering that it cannot be done is what sends the caller to
/// its own fallback, `restart_pinned_player`, which begins the file again at the second the hold
/// was taken at, at the level the pin is holding (see `toggle_pinned_playback`).
///
/// Every fact it is refused over is a refusal rather than a permission, and each is a different
/// mistake. A file that is not held is a file this key is being used to *stop*: ending its player
/// would take a picture away to answer a pause, which is the fallback of last resort (see
/// `hold_pinned_player_without_a_key`). A hold that has not reached its player yet is not a player
/// holding at any level — it is a player that is playing, and the key owed it is the thing being
/// delivered (see `settle_pending_hold`), which a level turned in the meantime does not cancel.
/// And a hold that is a *gesture's* is the end of that gesture rather than a press of the pause
/// button's: a window dragged for a second and released has to carry on from the second it was
/// watching, which a player begun again is the one way not to do (see `video_drag_hold_apply`).
pub(super) fn level_is_owed_to_the_next_player(
    held: bool,
    drag_held: bool,
    hold_owed: bool,
    playing_at: u32,
    level: u32,
) -> bool {
    held && !drag_held && !hold_owed && playing_at != level
}

/// The same answer as `level_is_owed_to_the_next_player`, over the pin that is up.
pub(super) fn pinned_level_is_owed_to_the_next_player() -> bool {
    pinned_playback_state().is_some_and(|(_, _, transport, volume)| {
        level_is_owed_to_the_next_player(
            transport.paused_at.is_some(),
            transport.drag_held,
            transport.pending_hold,
            volume.playing_at,
            volume.level,
        )
    })
}

/// How many pause keys this app has posted, counted where a test can read it.
///
/// The key lands on another process's message queue and a test run has none of those: the player a
/// test stands in is a number no process holds (see `spawn_player`), `video_window_for` answers
/// nothing for it, and the post below is never reached — so a key that is posted and a key that is
/// not look exactly alike from the transport, which is the only trace a test has. That is not
/// enough for the one decision the key stands behind: whether a gesture tells the player to hold
/// at all (see `video_drag_hold_apply`). What is counted is the telling rather than the delivery,
/// and it is a test's alone — the release binary carries no part of it.
#[cfg(test)]
static PAUSE_POSTS: AtomicU64 = AtomicU64::new(0);

/// The pause keys posted so far (see `PAUSE_POSTS`).
#[cfg(test)]
pub(super) fn pause_posts() -> u64 {
    PAUSE_POSTS.load(Ordering::Acquire)
}

/// Forget the pause keys posted so far, so each test counts its own.
#[cfg(test)]
pub(super) fn clear_pause_posts() {
    PAUSE_POSTS.store(0, Ordering::SeqCst);
}

/// Put the pause key on the player's own message queue, answering whether it was delivered.
///
/// It is one code path for both players because it is one key: FFmpeg's player reads a posted `P`
/// as its own `p`, and the binding file an mpv video is launched with holds `p cycle pause` (see
/// `VIDEO_PAUSE_KEY`).
///
/// Nothing is asked of the player about what it is doing — it reports nothing, which is the
/// whole reason `PinTransport` keeps what this app has asserted rather than what a player has
/// answered (see `transport_playing`). All that is done here is put a key on the player's own
/// message queue, which is a window's own business and is read off the handle already published
/// for the rest of the window's handling (see `apply_noactivate_to_hwnd`).
///
/// One player is refused before any of that, and it is refused for the reason a key is the only
/// thing a running player can be told: a player holding a film at a level the pin has moved on
/// from cannot be let go of with a key, because the level cannot be given to it afterwards (see
/// `level_is_owed_to_the_next_player`). A hold still owed to a player is the other case, and it is
/// answered here rather than refused — the key is what it is waiting for.
pub(super) fn video_key_pause() -> bool {
    if pinned_level_is_owed_to_the_next_player() {
        return false;
    }

    // Counted where the app commits to telling the player rather than where the key lands, because
    // it lands nowhere a test can see: there is no player's window here to post one through (see
    // `PAUSE_POSTS`).
    #[cfg(test)]
    PAUSE_POSTS.fetch_add(1, Ordering::Relaxed);

    let Some(hwnd) = video_window_for(VIDEO_PID.load(Ordering::SeqCst)) else {
        return false;
    };

    // SAFETY: the handle was just confirmed to be a live window of the player this app believes
    // is playing, so both messages below reach the player the caller means. Posting is the whole
    // of what is asked of it, and the two answers are the whole of what is known in return.
    unsafe {
        let (vk, scan) = VIDEO_PAUSE_KEY;
        let lparam = ffplay_key_lparam(scan);
        // The release is as much a part of the press as the press is: a key the player has been
        // sent down and never sent up is a key it will not report again until it is sent up,
        // so the second pause press would arrive as the *first* one and the bar would stop
        // answering altogether.
        let down = PostMessageW(hwnd, WM_KEYDOWN, WPARAM(vk as usize), lparam).is_ok();
        let up = PostMessageW(hwnd, WM_KEYUP, WPARAM(vk as usize), lparam).is_ok();
        down && up
    }
}

/// Check if the current ffplay process is still running
/// Clears stored state if the process has exited
pub(super) fn is_video_process_running() -> bool {
    if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
        if let Some(ref mut media) = *media_guard {
            if let Some(ref mut process) = media.video_process {
                match process.try_wait() {
                    Ok(Some(_)) => {
                        let pid = process.id();
                        media.video_process = None;
                        VIDEO_HWND.store(0, Ordering::SeqCst);
                        VIDEO_PID.store(0, Ordering::SeqCst);
                        // The player is confirmed gone, so the record of it goes with
                        // it rather than being left for the next run to look for.
                        engine_processes::forget(pid);
                        return false;
                    }
                    Ok(None) => return true,
                    Err(_) => {
                        media.video_process = None;
                        VIDEO_HWND.store(0, Ordering::SeqCst);
                        // Keep VIDEO_PID: the process is not confirmed dead, so
                        // the leftover-process checks can still find and kill it.
                        return false;
                    }
                }
            }
        }
    }
    false
}

/// True when the handle refers to a process whose executable file name matches
/// `expected_name` (compared case-insensitively).
pub(super) unsafe fn process_image_matches(handle: HANDLE, expected_name: &str) -> bool {
    let mut buffer = [0u16; 1024];
    let mut len = buffer.len() as u32;
    if QueryFullProcessImageNameW(
        handle,
        PROCESS_NAME_WIN32,
        PWSTR(buffer.as_mut_ptr()),
        &mut len,
    )
    .is_err()
    {
        return false;
    }

    let path = String::from_utf16_lossy(&buffer[..len as usize]);
    std::path::Path::new(&path)
        .file_name()
        .is_some_and(|name| name.eq_ignore_ascii_case(std::ffi::OsStr::new(expected_name)))
}

/// True when `pid` still refers to a live player of the kind this app last began.
pub(super) fn is_player_pid_alive(pid: u32) -> bool {
    unsafe {
        let Ok(handle) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) else {
            return false;
        };
        let matches = process_image_matches(handle, video_process_image());
        let _ = CloseHandle(handle);
        matches
    }
}

/// Terminate `pid` when it is still the player we spawned. Non-blocking: it only
/// requests the termination, it never waits for the process to exit.
pub(super) fn terminate_player_pid(pid: u32) {
    unsafe {
        let Ok(handle) = OpenProcess(
            PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            pid,
        ) else {
            return;
        };
        if process_image_matches(handle, video_process_image()) {
            let _ = TerminateProcess(handle, 1);
        }
        let _ = CloseHandle(handle);
    }
}

/// Kill the pinned player the way a gesture's press does: terminate, and reap
/// where the death confirms — never a wait, never the UI thread.
///
/// A kill confirmed on the spot clears the pid with it, so `VIDEO_PID == 0`
/// asserts silence for the whole dead interval. A kill not yet taken is
/// handed to the orphan reaper that outlives every gesture, which asks again
/// on every tick until the death confirms: never dropped, never waited on
/// (see `retire_orphaned_player`).
///
/// Answers whether nothing is left: pid zero is already nothing, and a kill
/// confirmed gone leaves nothing either.
pub(super) fn kill_pinned_player_async() -> bool {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid == 0 {
        return true;
    }

    // The media's handle and background work belong to the player being
    // killed: the only thing that can end it now is the pid, so nothing else
    // may keep asking the handle. What the tick asks instead is the pid,
    // which the reaper below retries to confirmation.
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.cancel_background_work();
            media.video_process = None;
        }
    }
    retire_orphaned_player(pid);

    VIDEO_PID.load(Ordering::SeqCst) != pid
}

/// Clear the recorded video process state once `pid` is confirmed gone.
pub(super) fn clear_video_process_state(pid: u32) {
    if VIDEO_PID
        .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst)
        .is_ok()
    {
        VIDEO_HWND.store(0, Ordering::SeqCst);
        // A hover's host outlives nothing: the player it held is confirmed gone, so the window is
        // taken down rather than left as a black top-level rectangle (see `hover_film_host`).
        unsafe { hover_film_host::take_down_the_hover_film_host_if_it_held(pid) };
        // The player is confirmed gone, so the record of it goes with it rather than
        // being left for the next run to look for.
        engine_processes::forget(pid);
    }
}

/// Kill the last spawned player process when it is still alive, whichever engine it is.
///
/// A process can outlive its `Child` handle: a kill may not take effect, or the
/// handle may be dropped before the process is confirmed gone. Without this a
/// surviving player keeps its window on screen and the next hover would spawn a
/// second one next to it. The PID is only cleared once the process is confirmed
/// gone, so a later call retries instead of losing track of it — and the player is
/// verified by the image name it was begun under, so an mpv stray is ended as mpv
/// and never as ffplay (see `note_player_process`).
pub fn kill_stray_video_process() {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid == 0 {
        return;
    }

    if !is_player_pid_alive(pid) {
        clear_video_process_state(pid);
        return;
    }

    terminate_player_pid(pid);

    if !is_player_pid_alive(pid) {
        clear_video_process_state(pid);
    }
}

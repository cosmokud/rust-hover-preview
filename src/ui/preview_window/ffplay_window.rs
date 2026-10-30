//! The window FFmpeg draws a film into, and the process standing behind it: finding the
//! window, keeping it out of the foreground, sending it a key, and killing what is left over.

use super::pin_trace::note_pin_trace;
use super::*;
use windows::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, Thread32First, Thread32Next, TH32CS_SNAPTHREAD, THREADENTRY32,
};
use windows::Win32::System::Threading::{
    OpenThread, ResumeThread, SuspendThread, THREAD_SUSPEND_RESUME,
};

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
    // **A film drawn through the pin's own host wears none of this.** The list is a window of the
    // desktop's own wearing bits this app puts on another program's window: a source is adopted by
    // `pin_film_thumbnail::adopt_the_source`, which gives it the invisibility and the press-through
    // both engines' windows need, so nothing of it is styled as a player's own window here (see
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
/// It is where a player of a film is adopted as the picture, from the window the player made for
/// itself: neither engine is drawn into a window of this app's, so what is found under the pid is a
/// window of the desktop's and the host shows it through the Shell's thumbnail (see
/// `pin_film_thumbnail`).
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

/// Style and raise a player's known window.
pub(super) unsafe fn apply_noactivate_to_hwnd(hwnd: HWND) -> bool {
    // **A hosted player's window is not this app's to style or to publish.** It is the source the
    // host draws through the Shell's thumbnail, and every road that would place, style or raise it is
    // answered by the host instead (see `player_ex_style_for`, `place_player_window`).
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
    );

    // Add WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST and WS_EX_TRANSPARENT to its extended
    // style (see `player_ex_style_for`).
    let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    SetWindowLongPtrW(hwnd, GWL_EXSTYLE, player_ex_style_for(current_style));

    // **And the keyboard the window took before this style was on it is given back here.** The style
    // above stops the player's window taking the keyboard; it does nothing about the activation that
    // already happened when the player opened its window, and that activation is not cosmetic: a
    // window of ours in front is the handle `key_input::pin_owns_caret` compares against, so for as
    // long as it holds the foreground a standing pin answers every `Space` pressed anywhere (see
    // `pin_lifecycle::a_players_foreground_goes_to`).
    pin_lifecycle::give_the_foreground_back_from(hwnd);

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
    // **What is held off here is the `Leave` answer and nothing else.** Those are a strip of the
    // pin's chrome, an open volume popup and an open card's menu (see `pin_chrome_up`,
    // `pin_volume_open`, `pin_menu_open`): a raise over any of them puts the film over what the
    // user is looking at — the caption and the transport bar the pin draws
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
    // **A film the Shell draws into a host of this app's is a window of the desktop's that is kept
    // one.** Neither player is begun into a window of this app's — mpv is launched with no `--wid`
    // and FFmpeg's player has no embed switch — so the window under this pid is the film's own,
    // and the host shows it by the Shell's thumbnail rather than holding it (see
    // `pin_film_thumbnail`).
    if let Some(host) = host_for_a_hosted_player(pid) {
        // **The source is adopted once, and the host is whichever one is showing the film** — a
        // pin's, or the hover's the pin will take up, so that pinning it is a re-pointing of the same
        // picture rather than a second player (see `hover_film_host`). A window that replaced itself
        // is another source: `the_source_is_up` says this pid's picture is not the one recorded, and
        // the adoption takes the new registration and drops the old (see
        // `pin_film_thumbnail::note_the_source`).
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
                unsafe { pin_film_thumbnail::adopt_the_source(data.found_hwnd, pid, host) };
            }
        } else {
            // **And this is the pass that keeps the source invisible**: the style is put on at
            // adoption and a player can come back from a video output reinit without it, so a source
            // whose layered bit is gone is a film standing over the pin until the next adopt — and
            // a window that is *the same handle* styled again is never adopted again (see
            // `pin_film_thumbnail::keep_the_source_invisible`).
            unsafe { pin_film_thumbnail::keep_the_source_invisible() };
            // **And the pass that sizes a film adoption had no answer for**: the player opens the
            // file after its window exists, and the one placement is made here as soon as it
            // answers rather than at the box the player chose for itself (see
            // `pin_film_thumbnail::place_the_source_when_the_player_answers`).
            pin_film_thumbnail::place_the_source_when_the_player_answers(host);
            // **And the pass that says whether a source adopted to replace this one can take the
            // picture yet**: the registration that is standing goes on drawing the window before it
            // until the new one can draw what it was placed for, and the swap is one step made here
            // rather than a picture taken down before its replacement exists (see
            // `pin_film_thumbnail::hand_the_picture_over`).
            pin_film_thumbnail::hand_the_picture_over(host);
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
            // **A window that must never wear a taskbar button is taken out of the Shell's lists from
            // its first appearance**, whether or not this app is showing it as a film: the warm resident
            // is a real top-level window with no file in it, and a button titled `No file` is not one the
            // user asked for (see `unlist_the_awaited_window` and `mpv_resident`).
            let a_window_waits_to_be_unlisted = unsafe { unlist_the_awaited_window() };

            if pid != monitored_pid {
                monitored_pid = pid;
                pid_started = Instant::now();
                if pid == 0 {
                    VIDEO_HWND.store(0, Ordering::SeqCst);
                }
            }

            // **A hover's window is styled before it is shown, and this is the pass that can** (see
            // `HOVER_BEFORE_THE_SHOW`): the show is what activates a player's window, so a style put
            // on the window the moment it exists — before anybody has shown it — is the one that
            // keeps the foreground where it was.
            let eager = (pid != 0)
                .then(|| the_player_awaiting_its_show(Instant::now()))
                .flatten();

            if pid != 0 {
                if eager.is_some_and(|(awaited, _)| awaited == pid) {
                    let road = match eager {
                        Some((_, true)) => "hover",
                        _ => "pin",
                    };
                    match unsafe { style_the_window_before_the_show(pid) } {
                        BeforeTheShow::NotYet => {}
                        BeforeTheShow::Styled => {
                            note_pin_trace!("{road} no-activate pid {pid} styled before the show");
                            forget_the_player_awaiting_its_show();
                        }
                        BeforeTheShow::ShownFirst => {
                            note_pin_trace!(
                                "{road} no-activate pid {pid} lost the race (visible already)"
                            );
                            forget_the_player_awaiting_its_show();
                        }
                    }
                }

                unsafe {
                    let _ = try_apply_noactivate_style(pid);
                }

                let elapsed = pid_started.elapsed();
                // **A replacement that is being waited for is looked for at the start's cadence, not
                // the idle one.** The cover is already on screen (see
                // `pin_playback::relaunch_pinned_player`), so the only thing left of a relaunch's
                // wait is its window being found and handed the picture — and the pass that finds it is
                // the pass the cover ends. It is bounded by the same wait the handover is bounded by,
                // and it costs nothing at all while no relaunch is pending.
                let launching = elapsed < Duration::from_secs(VIDEO_START_WAIT_SECS)
                    && pinned()
                    && !pin_is_collapsed()
                    && pin_playback::a_relaunch_is_in_flight();
                // And a player waiting to be caught before its show is looked for at the start's
                // cadence too, whatever age the player it belongs to is.
                wait_for_noactivate_wake(the_monitor_wait(
                    elapsed,
                    launching || eager.is_some() || a_window_waits_to_be_unlisted,
                ));
            } else if a_window_waits_to_be_unlisted {
                // A window this app is waiting for is looked for at the start's cadence: it is a player's
                // first quarter of a second, and the bit has to be on the window before the player shows
                // it (see `A_WINDOW_COMING_MS`).
                wait_for_noactivate_wake(A_WINDOW_COMING_MS);
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

/// How long the monitor waits before its next look at the player it follows: the start's own cadence
/// while the player is young or a relaunch is waiting for its window, and the idle one once it is up and
/// nothing is pending (see `ensure_noactivate_monitor`).
pub(super) fn the_monitor_wait(elapsed: Duration, launching: bool) -> u64 {
    if elapsed < Duration::from_millis(250) || launching {
        A_WINDOW_COMING_MS
    } else if elapsed < Duration::from_secs(2) {
        20
    } else {
        100
    }
}

/// The player whose window is to be styled **before** it is shown: the id it was begun under, the
/// moment it was begun, and whether the road that began it was a hover's.
///
/// **A window is activated by the show, not by anything this app does after it.** mpv makes a window
/// of its own and shows it as soon as it has opened the file, so the ordinary passes find a window
/// that has already taken the foreground and give it back a frame too late (see
/// `apply_noactivate_to_hwnd`). A launch that must not take the foreground says here that it has
/// begun a player, and the monitor's first passes put `WS_EX_NOACTIVATE` alone on that player's
/// window the instant it exists: a window carrying it cannot be activated when it is shown, whatever
/// the player asks for (see `video_playback::start_hover_video_playback`).
///
/// **Both roads write it, and the road is carried rather than assumed**: a pin's player is the film
/// under the pin's own window, and a launch that activated itself would take the front off the listing
/// the click was made in and give it to a window the pin makes invisible (see
/// `pin_lifecycle::a_hand_back_goes_to`). The word is only what the trace names the road by.
///
/// The marker is cleared once the film's window is on screen — whether the pass caught it in time or
/// lost the race — and it runs out on its own after `PLAYER_BEFORE_THE_SHOW_WAIT` either way, so a
/// launch whose player never makes a window leaves the monitor on the cadence it already keeps (see
/// `a_player_awaits_its_show` and `BeforeTheShow`).
static PLAYER_BEFORE_THE_SHOW: Mutex<Option<(u32, Instant, bool)>> = Mutex::new(None);

/// How long a player's window is waited for, counted from the launch: a player that has made no window
/// in this long is one whose show the eager pass has missed anyway (see `PLAYER_BEFORE_THE_SHOW`).
pub(super) const PLAYER_BEFORE_THE_SHOW_WAIT: Duration = Duration::from_millis(2000);

/// Write down a player as one whose window is to be styled before it is shown, with the road that
/// began it (see `PLAYER_BEFORE_THE_SHOW`).
pub(super) fn note_a_player_about_to_show(pid: u32, for_a_hover: bool) {
    if let Ok(mut held) = PLAYER_BEFORE_THE_SHOW.lock() {
        *held = Some((pid, Instant::now(), for_a_hover));
    }
}

/// The player still waiting for its window, as the id it was begun under and the road that began it —
/// the marker's own answer, asked at `now` rather than at the moment it was written.
pub(super) fn a_player_awaits_its_show(
    held: Option<(u32, Instant, bool)>,
    now: Instant,
) -> Option<(u32, bool)> {
    held.filter(|(_, begun, _)| now.saturating_duration_since(*begun) < PLAYER_BEFORE_THE_SHOW_WAIT)
        .map(|(pid, _, for_a_hover)| (pid, for_a_hover))
}

/// The marker as it stands, or nothing where it has run out — and one that has run out is dropped
/// here, so nothing goes on being asked about it.
pub(super) fn the_player_awaiting_its_show(now: Instant) -> Option<(u32, bool)> {
    let held = PLAYER_BEFORE_THE_SHOW.lock().ok().and_then(|held| *held);
    let player = a_player_awaits_its_show(held, now);
    if player.is_none() {
        if let Ok(mut held) = PLAYER_BEFORE_THE_SHOW.lock() {
            *held = None;
        }
    }
    player
}

/// Forget the player whose window has been found, whether the pass caught it before the show or not.
pub(super) fn forget_the_player_awaiting_its_show() {
    if let Ok(mut held) = PLAYER_BEFORE_THE_SHOW.lock() {
        *held = None;
    }
}

/// The style the eager pass writes: the two bits a player's window must not be shown without.
///
/// `WS_EX_NOACTIVATE` is the bit the show itself reads, and it is the whole of what this pass was once
/// for: a window carrying it cannot take the foreground when the player shows it (see
/// `PLAYER_BEFORE_THE_SHOW`).
///
/// **`WS_EX_TOOLWINDOW` is the other half, and it is the one a window with no film in it needs.** A
/// player opened idle with a window to draw into makes a real top-level window and shows it before any
/// file is loaded — the warm resident is exactly that (see `mpv_resident`) — and every window of a
/// player's own that this app shows is out of the Shell's taskbar and out of Alt+Tab in its end state
/// too: `player_ex_style_for` and `pin_film_thumbnail::source_ex_style` both put this bit on, and
/// neither ever takes it off. Written here it is on the window from the first frame the Shell can list
/// it, which is the difference between a warm player and a button titled `No file`.
///
/// What is *not* written here is everything the ordinary passes decide: the topmost band, the
/// press-through and the invisibility of a source (see `player_ex_style_for`). Those are answers of the
/// road the window is being put on, and a second answer written here would be one the passes then have
/// to disagree with.
pub(super) fn eager_ex_style(current: isize) -> isize {
    current | WS_EX_NOACTIVATE.0 as isize | WS_EX_TOOLWINDOW.0 as isize
}

/// The window this app has been told must never wear a taskbar button, with the moment the player that
/// makes it was begun: the warm resident's, and no other player's (see `note_a_window_to_unlist`).
///
/// **Why the resident needs a road of its own.** A player begun for a hover is caught by the eager
/// pass before its show or is the film the user is looking at a moment later, so a bit written after it
/// is on the screen is not worth a hide and a show there: that is a frame of film blanked. The resident
/// is the other way round — a real top-level window at a corner of no display, one pixel square, with
/// nothing in it and nobody looking at it (see `mpv_resident::PARKED_AT`) — so the bit can be written
/// and the show that makes the Shell read it can be paid at any moment, and the earliest one is the
/// window's first appearance.
static A_WINDOW_TO_UNLIST: Mutex<Option<(u32, Instant)>> = Mutex::new(None);

/// Ask for a player's film window to be taken out of the Shell's lists from its first appearance.
///
/// The monitor is started and woken rather than left to its own clock: between players it waits on that
/// event, and the window this is for appears a quarter of a second after the player is begun. A resident
/// is the one player that can be begun with no film behind it, so the thread that would otherwise be
/// started by a launch is started here as well (see `ensure_noactivate_monitor`).
pub(super) fn note_a_window_to_unlist(pid: u32) {
    if let Ok(mut held) = A_WINDOW_TO_UNLIST.lock() {
        *held = Some((pid, Instant::now()));
    }
    ensure_noactivate_monitor();
    wake_noactivate_monitor();
}

/// The process whose window is still to be taken out of the Shell's lists (see `A_WINDOW_TO_UNLIST`).
#[cfg(test)]
pub(super) fn the_window_to_unlist() -> Option<u32> {
    A_WINDOW_TO_UNLIST
        .lock()
        .ok()
        .and_then(|held| held.map(|(pid, _)| pid))
}

/// Forget the window this run was waiting to take out of the Shell's lists (see
/// `A_WINDOW_TO_UNLIST`).
#[cfg(test)]
pub(super) fn forget_the_window_to_unlist() {
    if let Ok(mut held) = A_WINDOW_TO_UNLIST.lock() {
        *held = None;
    }
}

/// How long the monitor waits between looks while a window it has been told about has not appeared yet:
/// the start's own cadence, because the window is a player's first frame of existence and the bit has to
/// be on it before the player shows it (see `the_monitor_wait` and `unlist_the_awaited_window`).
const A_WINDOW_COMING_MS: u64 = 5;

/// Take the awaited player's film window out of the Shell's lists, and answer whether one is still being
/// waited for (see `A_WINDOW_TO_UNLIST`).
///
/// The marker runs out the way a player's own wait for its window does: a process that made no window in
/// that time is one this pass has missed anyway, and a monitor that kept looking would keep its start
/// cadence for the rest of the run (see `PLAYER_BEFORE_THE_SHOW_WAIT`).
///
/// Safety: the window this asks about is a player's own window of another process, and the style and
/// show calls are the one road this app styles a player's window on (see `try_apply_noactivate_style`
/// and `mpv_resident`).
unsafe fn unlist_the_awaited_window() -> bool {
    let Some((pid, begun)) = A_WINDOW_TO_UNLIST.lock().ok().and_then(|held| *held) else {
        return false;
    };
    if begun.elapsed() >= PLAYER_BEFORE_THE_SHOW_WAIT {
        if let Ok(mut held) = A_WINDOW_TO_UNLIST.lock() {
            *held = None;
        }
        return false;
    }

    !unsafe { give_a_players_window_no_button(pid) }
}

/// Put the bits that keep a player's window out of the user's lists on it, wherever it is in its life:
/// answering whether it was there to be styled yet (see `unlist_the_awaited_window`).
unsafe fn give_a_players_window_no_button(pid: u32) -> bool {
    let mut data = TheFilmsWindowSearch {
        target_pid: pid,
        found: HWND::default(),
    };
    let _ = unsafe {
        EnumWindows(
            Some(the_films_window_callback),
            LPARAM(&mut data as *mut TheFilmsWindowSearch as isize),
        )
    };
    if data.found.is_invalid() {
        return false;
    }

    if let Ok(mut held) = A_WINDOW_TO_UNLIST.lock() {
        *held = None;
    }

    let window = data.found;
    let current = unsafe { GetWindowLongPtrW(window, GWL_EXSTYLE) };
    let wanted = eager_ex_style(current);
    if wanted != current {
        let _ = unsafe { SetWindowLongPtrW(window, GWL_EXSTYLE, wanted) };
        // A style written behind a window's back is not one it presents with until it is asked to read
        // it (see `pin_real_window::set_the_pin_window_appwindow`, which is the same fact for a pin).
        let _ = unsafe {
            SetWindowPos(
                window,
                None,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        };
        // **And the Shell lists a window as it is shown**, so a window it has already listed wears the
        // button it was listed with however the bit is re-read. This window is the resident's: parked
        // at a corner of no display, one pixel square, with no film loaded and nobody looking at it, so
        // the show is paid for here rather than left to the first hover — which would leave a button
        // titled `No file` standing until the resident was picked up, or for ever.
        if unsafe { IsWindowVisible(window) }.as_bool() {
            let _ = unsafe { ShowWindow(window, SW_HIDE) };
            let _ = unsafe { ShowWindow(window, SW_SHOWNOACTIVATE) };
        }
        note_pin_trace!("player window {:?} taken out of the Shell's lists", window);
    }
    true
}

/// What the eager pass found of a hover's player's film window.
pub(super) enum BeforeTheShow {
    /// The film's window is not on screen yet — it was styled and the pass goes on looking, or the
    /// player has not made it yet.
    NotYet,
    /// The film's window is up and wears `WS_EX_NOACTIVATE`: an earlier pass put it there, so the show
    /// it is now in could not activate it.
    Styled,
    /// The film's window is up without the bit: it was shown before this app ever had it, and today's
    /// road is what is left.
    ShownFirst,
}

/// The class name of the window a film is drawn in.
///
/// **It is what tells the film's window from the windows a player makes for everything else, and the
/// distinction is not cosmetic: styling a helper breaks the launch it belongs to.** Measured on mpv
/// 0.41.0 here: mpv makes a hidden `mpv-smtc` window (the media keys' own) about 20ms into the
/// launch and its film's `mpv` window some 250ms later, which is then shown some 15–25ms after that.
/// Giving `WS_EX_NOACTIVATE` to the helper did not merely fail to help: the film's window came out at
/// ~500ms *and already visible*, so the pass lost the very show it exists to beat — while styling the
/// film's own window left every timing in place and the window shown still carrying the bit.
///
/// A build that names its window something else is not a wrong window styled — the pass finds
/// nothing and today's road answers it (see `style_the_window_before_the_show`).
const A_FILMS_WINDOW_CLASS: &str = "mpv";

/// What the eager enumeration carries: which process, and the film's window where one was found.
struct TheFilmsWindowSearch {
    target_pid: u32,
    found: HWND,
}

/// Callback for `EnumWindows` that **does not require the window to be visible**, which is the one
/// thing the ordinary callback refuses and the whole point of this pass: a window nobody has shown
/// yet is the window this is for (see `enum_windows_callback` and `A_FILMS_WINDOW_CLASS`).
unsafe extern "system" fn the_films_window_callback(
    hwnd: HWND,
    lparam: LPARAM,
) -> windows::Win32::Foundation::BOOL {
    let data = &mut *(lparam.0 as *mut TheFilmsWindowSearch);
    let mut window_pid: u32 = 0;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));

    if window_pid != data.target_pid {
        return windows::Win32::Foundation::BOOL(1);
    }
    if let Ok(owner) = GetWindow(hwnd, GW_OWNER) {
        if !owner.is_invalid() {
            return windows::Win32::Foundation::BOOL(1);
        }
    }

    let mut class: [u16; 64] = [0; 64];
    let named = GetClassNameW(hwnd, &mut class);
    if named <= 0
        || !String::from_utf16_lossy(&class[..named as usize])
            .eq_ignore_ascii_case(A_FILMS_WINDOW_CLASS)
    {
        return windows::Win32::Foundation::BOOL(1);
    }

    data.found = hwnd;
    windows::Win32::Foundation::BOOL(0)
}

/// Put `WS_EX_NOACTIVATE` alone on a hover's player's film window the moment it exists, and say
/// whether that window is on screen yet (see `BeforeTheShow`).
pub(super) unsafe fn style_the_window_before_the_show(pid: u32) -> BeforeTheShow {
    let mut data = TheFilmsWindowSearch {
        target_pid: pid,
        found: HWND::default(),
    };
    let _ = EnumWindows(
        Some(the_films_window_callback),
        LPARAM(&mut data as *mut TheFilmsWindowSearch as isize),
    );

    if data.found.is_invalid() {
        return BeforeTheShow::NotYet;
    }

    // **The window that is up is the one that says whether the race was won**, asked off the style it
    // is wearing rather than off a record kept across passes: a window that already carries the bit
    // was styled by an earlier pass before it was shown, and one that does not was shown before this
    // app ever had its handle.
    let current = GetWindowLongPtrW(data.found, GWL_EXSTYLE);
    if IsWindowVisible(data.found).as_bool() {
        return if current & WS_EX_NOACTIVATE.0 as isize != 0 {
            BeforeTheShow::Styled
        } else {
            BeforeTheShow::ShownFirst
        };
    }

    SetWindowLongPtrW(data.found, GWL_EXSTYLE, eager_ex_style(current));
    BeforeTheShow::NotYet
}

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

    // **A film the Shell draws into a host of this app's is asked about as its source**, which is
    // the window the player made and the host thumbnails: a key has to reach that window, and it is
    // the one handle the road publishing nothing else leaves (see `pin_film_thumbnail`).
    if pin_film_thumbnail::the_source_is_up(pid) {
        return pin_film_thumbnail::the_source_window();
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
pub(super) fn level_is_owed_to_the_next_player(
    held: bool,
    hold_owed: bool,
    playing_at: u32,
    level: u32,
) -> bool {
    held && !hold_owed && playing_at != level
}

/// The same answer as `level_is_owed_to_the_next_player`, over the pin that is up.
pub(super) fn pinned_level_is_owed_to_the_next_player() -> bool {
    pinned_playback_state().is_some_and(|(_, _, transport, volume)| {
        level_is_owed_to_the_next_player(
            transport.paused_at.is_some(),
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
/// enough for the one decision the key stands behind: whether the pin tells the player to hold at
/// all. What is counted is the telling rather than the delivery,
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

    video_key_pause_for(VIDEO_PID.load(Ordering::SeqCst))
}

/// The same key, posted to the player that is to be given it: the press's own hold everywhere.
///
/// **The key toggles, and which way it lands is the caller's fact to know.** A player already
/// holding a film is already a still frame, and a key posted to it would set it going again.
///
/// A key posted is directed at the window it is posted to, whatever that window's extended style
/// says, so the press-through and no-activate bits a source wears do not keep the player from
/// reading it.
pub(super) fn video_key_pause_for(pid: u32) -> bool {
    // Counted where the app commits to telling the player rather than where the key lands, because
    // it lands nowhere a test can see: there is no player's window here to post one through (see
    // `PAUSE_POSTS`).
    #[cfg(test)]
    PAUSE_POSTS.fetch_add(1, Ordering::Relaxed);

    let Some(hwnd) = video_window_for(pid) else {
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

/// The players this app has suspended and not let go of, each with how many of its threads were
/// suspended (see `suspend_player_threads`).
///
/// It is the whole record a freeze of this kind leaves: the player is never resumed on the road that
/// suspends it — it is ended once its replacement has the picture — so what is written down is only
/// what a kill the system refuses would otherwise leave standing frozen for ever (see
/// `terminate_player_pid`).
static SUSPENDED_PLAYERS: Lazy<Mutex<Vec<(u32, usize)>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// **Suspend every thread of the player being replaced, which is the hold for an engine with no
/// command channel.** FFmpeg's player is told nothing once it runs, and the one thing it can be told —
/// the pause key — is a toggle the pin's own transport uses. A player whose threads are suspended
/// draws nothing, clears nothing and presents nothing, so the frame it last put up is the frame that
/// stays. It is the hold a hover is given for the pin that takes it up (see
/// `pin_playback::freeze_the_outgoing_picture`).
///
/// It is a handful of `OpenThread`/`SuspendThread` calls over the threads the snapshot names, made
/// once per relaunch and never on a gesture. A player already suspended is answered with the threads
/// that were suspended for it rather than suspended again, because a second suspension takes a second
/// resume to undo (see `resume_suspended_player`).
pub(super) fn suspend_player_threads(pid: u32) -> usize {
    if pid == 0 {
        return 0;
    }

    if let Ok(held) = SUSPENDED_PLAYERS.lock() {
        if let Some((_, threads)) = held.iter().find(|(suspended, _)| *suspended == pid) {
            return *threads;
        }
    }

    let threads = suspend_the_threads(pid);
    if threads > 0 {
        if let Ok(mut held) = SUSPENDED_PLAYERS.lock() {
            held.push((pid, threads));
        }
    }
    threads
}

/// Let a suspended player go again, answering how many of its threads were suspended: **the road a
/// kill that is not taken takes** (see `terminate_player_pid`), and nothing else.
///
/// A player this app froze is ended, never resumed — but a process left suspended for ever is a film
/// standing frozen behind a window nothing owns, which is worse than the film going on, so the one
/// road that can leave one behind lets it go instead.
pub(super) fn resume_suspended_player(pid: u32) -> usize {
    let threads = {
        let Ok(mut held) = SUSPENDED_PLAYERS.lock() else {
            return 0;
        };
        let Some(at) = held.iter().position(|(suspended, _)| *suspended == pid) else {
            return 0;
        };
        held.remove(at).1
    };

    #[cfg(test)]
    if SUSPEND_STAND_IN.lock().is_ok_and(|held| held.is_some()) {
        return threads;
    }

    resume_the_threads(pid);
    threads
}

/// Forget that a player was suspended, answered where a kill is requested or refused: the record is
/// a claim about a live process, and the system hands a process id to another process (see
/// `suspend_player_threads`).
pub(super) fn forget_suspended_player(pid: u32) {
    if let Ok(mut held) = SUSPENDED_PLAYERS.lock() {
        held.retain(|(suspended, _)| *suspended != pid);
    }
}

/// The players this app has suspended and not ended, as a test reads the freeze: it is how a run says
/// that the player on screen was the one suspended and the replacement was not (see
/// `suspend_player_threads`).
#[cfg(test)]
pub(super) fn suspended_players() -> Vec<u32> {
    SUSPENDED_PLAYERS
        .lock()
        .map(|held| held.iter().map(|(pid, _)| *pid).collect())
        .unwrap_or_default()
}

/// Forget every suspension a test stood in, so each test reads its own (see `suspended_players`).
#[cfg(test)]
pub(super) fn clear_suspended_players() {
    if let Ok(mut held) = SUSPENDED_PLAYERS.lock() {
        held.clear();
    }
}

/// How many of a player's threads are suspended, or the answer a test stood in for it (see
/// `SUSPEND_STAND_IN`).
fn suspend_the_threads(pid: u32) -> usize {
    #[cfg(test)]
    if let Some(threads) = SUSPEND_STAND_IN.lock().ok().and_then(|held| *held) {
        return threads;
    }
    // Safety: the snapshot is this app's own to close, and each thread handle is opened for the one
    // right `SuspendThread` needs and closed straight after it.
    unsafe {
        let Ok(snapshot) = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0) else {
            return 0;
        };
        let mut entry = THREADENTRY32 {
            dwSize: std::mem::size_of::<THREADENTRY32>() as u32,
            ..Default::default()
        };
        let mut suspended = 0;
        if Thread32First(snapshot, &mut entry).is_ok() {
            loop {
                if entry.th32OwnerProcessID == pid {
                    if let Ok(thread) = OpenThread(THREAD_SUSPEND_RESUME, false, entry.th32ThreadID)
                    {
                        if SuspendThread(thread) != u32::MAX {
                            suspended += 1;
                        }
                        let _ = CloseHandle(thread);
                    }
                }
                if Thread32Next(snapshot, &mut entry).is_err() {
                    break;
                }
            }
        }
        let _ = CloseHandle(snapshot);
        suspended
    }
}

/// Resume every thread of `pid` that the snapshot names and is still openable: a thread that ended
/// between the snapshot and the open is one `OpenThread` answers nothing for, which is not an error.
fn resume_the_threads(pid: u32) {
    // Safety: as above — the snapshot and each thread handle are this call's own and closed here.
    unsafe {
        let Ok(snapshot) = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0) else {
            return;
        };
        let mut entry = THREADENTRY32 {
            dwSize: std::mem::size_of::<THREADENTRY32>() as u32,
            ..Default::default()
        };
        if Thread32First(snapshot, &mut entry).is_ok() {
            loop {
                if entry.th32OwnerProcessID == pid {
                    if let Ok(thread) = OpenThread(THREAD_SUSPEND_RESUME, false, entry.th32ThreadID)
                    {
                        let _ = ResumeThread(thread);
                        let _ = CloseHandle(thread);
                    }
                }
                if Thread32Next(snapshot, &mut entry).is_err() {
                    break;
                }
            }
        }
        let _ = CloseHandle(snapshot);
    }
}

/// The suspension a test is answered with in place of the machine's own.
///
/// The player a test stands in is a process id no process holds — and on the hover's road it is the
/// test process's own — so the real road must never run from a test. The machine's own answer is
/// asked of real processes by the two tests that begin one (see `suspend_the_threads`).
#[cfg(test)]
static SUSPEND_STAND_IN: Mutex<Option<usize>> = Mutex::new(Some(1));

/// Terminate `pid` when it is still the player we spawned. Non-blocking: it only
/// requests the termination, it never waits for the process to exit.
///
/// **Every end of a player in the preview comes through here**, which is why the one line that says a
/// player went while its window was the picture is written here: an end made before its replacement
/// stands is the black band the relaunch's order exists to keep out (see
/// `video_retire::note_an_end_while_standing`).
///
/// **A suspended player is ended by the same request** — `TerminateProcess` ends a process whether its
/// threads are running or suspended — so what is left here is the one road a request the system does
/// not take leaves. The record goes with the request either way, and a player still alive after a
/// refusal is let go of rather than left frozen: a frozen process nothing will ever resume is the
/// leak this record exists to rule out (see `resume_suspended_player`).
pub(super) fn terminate_player_pid(pid: u32) {
    note_an_end_while_standing(pid);
    // Safety: the handle is opened for the two rights the calls below need and closed before this
    // answers, and the image check is what keeps a kill by id alone impossible.
    let requested = unsafe {
        OpenProcess(
            PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            pid,
        )
        .is_ok_and(|handle| {
            let requested = process_image_matches(handle, video_process_image())
                && TerminateProcess(handle, 1).is_ok();
            let _ = CloseHandle(handle);
            requested
        })
    };

    if !requested && is_player_pid_alive(pid) && resume_suspended_player(pid) > 0 {
        note_pin_trace!("suspended player could not be ended pid {pid}: let go of");
    }
    forget_suspended_player(pid);
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

    // **The resident is never a stray.** It is the player this app keeps between hovers, and a sweep
    // is for a player that outlived the hover it was begun for (see `mpv_resident`).
    if mpv_resident::holds_the_pid(pid) {
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

#[cfg(test)]
mod tests {
    use super::*;

    /// **A relaunch is looked for at the start's cadence, and nothing else is.** What is left of a
    /// relaunch's wait behind a frozen picture is its window being found, so the monitor passes every
    /// few milliseconds while a replacement is pending — bounded by the player-start wait — and goes
    /// back to the idle cadence the moment nothing is pending (see `the_monitor_wait`).
    #[test]
    fn a_pending_relaunch_is_looked_for_at_the_starts_cadence() {
        assert_eq!(
            the_monitor_wait(Duration::from_millis(10), false),
            5,
            "a player that has just been begun"
        );
        assert_eq!(
            the_monitor_wait(Duration::from_secs(30), true),
            5,
            "a replacement waiting for its window, however long the pin has been up"
        );
        assert_eq!(
            the_monitor_wait(Duration::from_secs(1), true),
            5,
            "and the same inside the first two seconds"
        );
        assert_eq!(
            the_monitor_wait(Duration::from_secs(1), false),
            20,
            "a young player is watched closer than the idle cadence"
        );
        assert_eq!(
            the_monitor_wait(Duration::from_secs(30), false),
            100,
            "and nothing pending is the idle cadence the monitor always had"
        );
    }

    /// **Every launch that must not take the foreground asks for the eager style, and only while its
    /// window is still to come.** The marker is what such a launch writes — a hover's and a pin's
    /// alike — and the road travels with it, so this is the whole of the question the monitor asks
    /// before it styles anything early (see `note_a_player_about_to_show` and
    /// `a_player_awaits_its_show`).
    #[test]
    fn the_eager_style_is_asked_for_both_roads_and_only_until_the_window_arrives() {
        let begun = Instant::now();
        assert_eq!(
            a_player_awaits_its_show(Some((42, begun, true)), begun),
            Some((42, true)),
            "a hover that has just begun its player is the one the pass runs for"
        );
        assert_eq!(
            a_player_awaits_its_show(Some((42, begun, false)), begun),
            Some((42, false)),
            "and so is a pin's, which is the window a launch must not activate either"
        );
        assert_eq!(
            a_player_awaits_its_show(None, begun),
            None,
            "a launch that noted nothing has no player to style before its show"
        );
        assert_eq!(
            a_player_awaits_its_show(Some((42, begun, true)), begun + PLAYER_BEFORE_THE_SHOW_WAIT),
            None,
            "and a player whose window never came is not looked for past its wait"
        );
    }

    /// **The eager style takes the player's window out of the taskbar as well as off the focus**, and it
    /// adds nothing the passes that follow would have to disagree with.
    ///
    /// `WS_EX_TOOLWINDOW` is what keeps a window out of the Shell's taskbar and out of Alt+Tab, and a
    /// player's window wears it in every end state this app has for one: `player_ex_style_for` puts it on
    /// a hover's or a pin's window and `pin_film_thumbnail::source_ex_style` puts it on a source (see
    /// `eager_ex_style`).
    #[test]
    fn the_eager_style_takes_the_window_out_of_the_taskbar_and_off_the_focus() {
        let noactivate = WS_EX_NOACTIVATE.0 as isize;
        let toolwindow = WS_EX_TOOLWINDOW.0 as isize;
        let inert = WS_EX_TOPMOST.0 as isize | WS_EX_TRANSPARENT.0 as isize;

        assert_eq!(
            eager_ex_style(0),
            noactivate | toolwindow,
            "a bare window is given the two bits a player's window must not be shown without"
        );
        assert_eq!(
            eager_ex_style(toolwindow),
            toolwindow | noactivate,
            "and a bit the player already wears is left exactly as it was"
        );
        assert_eq!(
            eager_ex_style(inert),
            inert | noactivate | toolwindow,
            "while the band, the press-through and the invisibility of a source stay the passes' own"
        );
    }

    /// **And the end state every pass leaves a player's window in is the one the eager style names.**
    ///
    /// The eager pass is shared with hover players, so the only way its losing a button is not a change
    /// of behaviour is that both ordinary style roads already ask for the same bit: a hover's film and a
    /// pin's are styled by `player_ex_style_for`, and the source the Shell draws through a host by
    /// `pin_film_thumbnail::source_ex_style` (whose own test holds that bit and the two beside it in
    /// `the_invisible_knob_is_on_unless_it_says_zero_and_is_the_layered_bit`). Nothing either of them
    /// decides is taken away by the eager pass, and neither ever takes the toolwindow bit off.
    #[test]
    fn a_players_end_style_still_has_the_bit_the_eager_pass_writes() {
        let toolwindow = WS_EX_TOOLWINDOW.0 as isize;
        let noactivate = WS_EX_NOACTIVATE.0 as isize;

        // A player's window as a launch opens it, and one that already wears the bit: the two ends the
        // ordinary passes are asked about.
        let app_window = windows::Win32::UI::WindowsAndMessaging::WS_EX_APPWINDOW.0 as isize;
        for current in [app_window, app_window | toolwindow] {
            let styled = player_ex_style_for(current);
            assert_ne!(
                styled & toolwindow,
                0,
                "a player's window is out of the taskbar when the passes are done with it"
            );
            assert_ne!(styled & noactivate, 0, "and it never takes the keyboard");
            assert_eq!(
                eager_ex_style(styled),
                styled,
                "so the eager pass asks for nothing the end state does not already have: {styled:x}"
            );
        }
    }

    /// A pid with no window of its own is a engine that has not made one yet, which is the answer the
    /// monitor keeps looking on rather than a race it has lost (see `style_the_window_before_the_show`).
    #[test]
    fn a_hover_with_no_window_yet_is_awaited_rather_than_styled() {
        // A test that runs the desktop's enumeration takes the suite's lock, however hopeless the
        // process it asks about (see `PIN_TESTS_ONE_AT_A_TIME`).
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        assert!(
            matches!(
                unsafe { style_the_window_before_the_show(u32::MAX - 7) },
                BeforeTheShow::NotYet
            ),
            "a process that has made no window is no window to style yet"
        );
    }

    /// A window of the class a film's window has, made by this process and left where no display is:
    /// what the passes that style a player before its show are asked about, and what the caller
    /// destroys. `shown` is the other half of the question — a player shows its own window the moment
    /// it has configured it, so a pass either catches the window before that or finds it already
    /// listed.
    ///
    /// Safety: the window is the caller's to destroy.
    unsafe fn a_film_window_of_the_tests_own(shown: bool) -> HWND {
        use windows::Win32::UI::WindowsAndMessaging::{
            CreateWindowExW, DefWindowProcW, RegisterClassW, WINDOW_EX_STYLE, WNDCLASSW, WS_POPUP,
        };

        unsafe extern "system" fn an_empty_window(
            hwnd: HWND,
            msg: u32,
            wparam: windows::Win32::Foundation::WPARAM,
            lparam: windows::Win32::Foundation::LPARAM,
        ) -> windows::Win32::Foundation::LRESULT {
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }

        unsafe {
            let instance = GetModuleHandleW(None).expect("this module");
            let class = w!("mpv");
            let wc = WNDCLASSW {
                lpfnWndProc: Some(an_empty_window),
                hInstance: instance.into(),
                lpszClassName: class,
                ..Default::default()
            };
            // A class left registered by a test that ran before this one is the same class.
            let _ = RegisterClassW(&wc);

            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                class,
                w!(""),
                WS_POPUP,
                -32000,
                -32000,
                320,
                180,
                None,
                None,
                instance,
                None,
            )
            .expect("a film's window, off the desktop");
            if shown {
                let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
            }
            hwnd
        }
    }

    /// **The eager pass styles the film's own window as soon as it exists and before it is shown.** A
    /// window of this process is made with the class a film's window has and left unshown the way a
    /// player leaves it for the few milliseconds it takes to configure it — and the pass finds it, gives
    /// it the two bits a player's window must not be shown without, and then says the show it later saw
    /// was the one the bits were already on the window for (see `style_the_window_before_the_show` and
    /// `A_FILMS_WINDOW_CLASS`).
    #[test]
    fn the_eager_pass_styles_a_film_window_that_is_not_shown_yet() {
        // **This is a window of this process's own, wearing the player's class**, so it is one the
        // other pass that enumerates by process and class would answer with — and that pass's
        // window is one this one would answer with. One at a time (see `PIN_TESTS_ONE_AT_A_TIME`).
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

        unsafe {
            let hwnd = a_film_window_of_the_tests_own(false);
            let noactivate = WS_EX_NOACTIVATE.0 as isize;
            let toolwindow = WS_EX_TOOLWINDOW.0 as isize;

            assert!(
                matches!(
                    style_the_window_before_the_show(GetCurrentProcessId()),
                    BeforeTheShow::NotYet
                ),
                "a window nobody has shown is styled and the pass goes on looking"
            );
            let ex = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
            assert_ne!(ex & noactivate, 0, "and it is given `WS_EX_NOACTIVATE`");
            assert_ne!(
                ex & toolwindow,
                0,
                "with the bit that keeps it out of the Shell's taskbar from its first appearance"
            );

            let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
            assert!(
                matches!(
                    style_the_window_before_the_show(GetCurrentProcessId()),
                    BeforeTheShow::Styled
                ),
                "and the show it is now in is one the bits were already on the window for"
            );

            let _ = DestroyWindow(hwnd);
        }
    }

    /// **A player's window is taken out of the Shell's lists from its first appearance**, which is the
    /// road a warm resident's window gets: the resident is begun idle with a window to draw into and no
    /// file to draw, so its window is a real top-level window of the user's from the moment mpv makes it
    /// — a taskbar button titled `No file` (see `note_a_window_to_unlist` and `mpv_resident`).
    ///
    /// Both lives of that window are held here, because the Shell reads a taskbar bit as a window is
    /// *shown*: one caught before its show wears the bits from the start, and one the pass finds already
    /// listed is styled and put down and up again for it — which costs nothing, because the window is at
    /// a corner of no display with nothing in it, and is the whole reason this road is the resident's
    /// rather than a hover player's.
    #[test]
    fn a_players_window_is_taken_out_of_the_shells_lists_from_its_first_appearance() {
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

        let pid = unsafe { GetCurrentProcessId() };
        let the_bits = |hwnd: HWND| {
            let ex = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
            (
                ex & (WS_EX_TOOLWINDOW.0 as isize) != 0,
                ex & (WS_EX_NOACTIVATE.0 as isize) != 0,
            )
        };

        for shown_first in [false, true] {
            forget_the_window_to_unlist();
            let window = unsafe { a_film_window_of_the_tests_own(shown_first) };
            let was_up = unsafe { IsWindowVisible(window) }.as_bool();
            assert_eq!(
                the_bits(window),
                (false, false),
                "a window as a player opens it is listed and may be activated"
            );

            note_a_window_to_unlist(pid);
            // **The pass is asked directly rather than waited for.** The thread that runs it is the
            // app's, and a test holds the road rather than the clock — which is also why the answer
            // here is only that nothing is left waiting: the monitor may have spent the marker first.
            assert!(
                !unsafe { unlist_the_awaited_window() },
                "the window is there to be taken out of the lists, so nothing is left waiting for it"
            );
            assert_eq!(
                the_bits(window),
                (true, true),
                "so it wears the bits that keep it out of the taskbar and off the focus ({shown_first})"
            );
            assert_eq!(
                unsafe { IsWindowVisible(window) }.as_bool(),
                was_up,
                "and a window that was on the screen is left on the screen by the show the Shell                  read the bits with, where one caught before its show is left for the player to show"
            );
            assert_eq!(
                the_window_to_unlist(),
                None,
                "the marker is spent, so the monitor goes back to its own cadence"
            );

            let _ = unsafe { DestroyWindow(window) };
        }
        forget_the_window_to_unlist();
    }

    /// A process begun for a test to suspend: it makes no window and plays nothing, and its own thirty
    /// seconds are bounded by the kill and the wait in whatever this test leaves behind.
    #[allow(clippy::zombie_processes)]
    fn a_process_to_suspend() -> std::process::Child {
        use std::os::windows::process::CommandExt;
        std::process::Command::new("cmd")
            .args(["/c", "ping", "-n", "30", "127.0.0.1"])
            .stdin(std::process::Stdio::null())
            .stdout(std::process::Stdio::null())
            .stderr(std::process::Stdio::null())
            // The console program's own window is not this suite's to put on the desktop.
            .creation_flags(0x08000000)
            .spawn()
            .expect("a process to suspend")
    }

    /// **A process this app has suspended can still be ended**, which is the fact a relaunch's freeze
    /// rests on: the end of a player is a termination and not a request for its threads to stop, so a
    /// player frozen where it stands dies like any other and is never left behind (see
    /// `terminate_player_pid`).
    ///
    /// Asked of a real process because it is a fact about the machine rather than about this code: the
    /// suspension below is the same enumeration and the same `SuspendThread` calls a relaunch makes.
    #[test]
    fn a_suspended_process_can_still_be_ended() {
        let mut child = a_process_to_suspend();
        assert!(
            suspend_the_threads(child.id()) >= 1,
            "every thread of it is suspended"
        );

        assert!(child.kill().is_ok(), "the kill is still taken");
        assert!(
            child.wait().is_ok(),
            "and a suspended process goes: nothing is left standing frozen"
        );
    }

    /// **And a suspension is let go of where the end is refused.** The other road out of a freeze is a
    /// resume, and a resumed process is running again rather than merely alive.
    #[test]
    fn a_suspended_process_runs_again_once_it_is_resumed() {
        let mut child = a_process_to_suspend();
        assert!(suspend_the_threads(child.id()) >= 1);
        resume_the_threads(child.id());
        assert!(
            child.try_wait().is_ok_and(|ended| ended.is_none()),
            "the process the resume let go of is still running"
        );

        let _ = child.kill();
        let _ = child.wait();
    }

    /// **One suspension per player, and one resume to let it go.** A second suspension of the same
    /// player would take a second resume to undo, which is the frozen process a refused kill would
    /// leave behind (see `suspend_player_threads` and `resume_suspended_player`).
    #[test]
    fn a_suspension_is_recorded_once_and_let_go_of_once() {
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        let pid = 4242;
        clear_suspended_players();
        assert_eq!(suspend_player_threads(pid), 1);
        assert_eq!(
            suspend_player_threads(pid),
            1,
            "a player already suspended is not suspended again"
        );
        assert_eq!(suspended_players(), vec![pid]);
        assert_eq!(resume_suspended_player(pid), 1);
        assert_eq!(
            resume_suspended_player(pid),
            0,
            "and there is nothing left to let go of"
        );
        assert!(suspended_players().is_empty());
    }
}

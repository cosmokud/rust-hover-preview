//! The window FFmpeg draws a film into, and the process standing behind it: finding the
//! window, keeping it out of the foreground, sending it a key, and killing what is left over.

use super::*;

// Expected executable name of the playback process spawned below, used to
// verify a recorded PID still belongs to that process before killing it.
pub(super) const VIDEO_PROCESS_IMAGE_NAME: &str = "ffplay.exe";

// Track the ffplay video window HWND for cursor-over-preview detection
pub(super) static VIDEO_HWND: AtomicIsize = AtomicIsize::new(0);
// Track the ffplay process ID to re-find the window if needed
pub(super) static VIDEO_PID: AtomicU32 = AtomicU32::new(0);
// Guard to ensure we only run a single style-monitor thread.
pub(super) static NOACTIVATE_MONITOR_STARTED: AtomicBool = AtomicBool::new(false);
// Flag set when the system resumes from sleep, so the main loop can reset state.
pub(super) static RESUME_FROM_SLEEP: AtomicBool = AtomicBool::new(false);
// Flag set when the display under the preview changed — a monitor's scale, or the
// desktop rearranged — a frame the window proc can only discard. What was on screen
// is put back by the loop, which is the side that holds the hover it came from.
pub(super) static DISPLAY_RESET: AtomicBool = AtomicBool::new(false);

/// Data passed to the EnumWindows callback to find ffplay window
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

    // Keep the largest candidate; this is typically the real ffplay output window.
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
/// `ensure_video_window_topmost`, which re-asserts only the three it also knows
/// about and so preserves these rather than re-deciding them).
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
pub(super) fn player_ex_style(current: isize) -> isize {
    current
        | WS_EX_NOACTIVATE.0 as isize
        | WS_EX_TOOLWINDOW.0 as isize
        | WS_EX_TOPMOST.0 as isize
        | WS_EX_TRANSPARENT.0 as isize
}

/// How long a player's window is kept off the screen from the moment it is first seen.
///
/// FFmpeg's player makes its window before it opens the file — SDL builds it, and the film is
/// demuxed and decoded into it afterwards — so for the first moments of its life the window is
/// a rectangle of the video's own size with nothing in it. An SDL window has no background of
/// its own to erase, so what the compositor holds for that rectangle is composited until the
/// player paints: whatever preview was on screen there a moment earlier, at its own size and
/// with black around it, inside a window the size of the video. It lasted a few milliseconds
/// and it was the previous file's frame, whichever kind it was.
///
/// Nothing on this side can ask the player whether it has a frame — `PrintWindow` reads blank
/// for the window, because its content is a GPU surface — so this is a bound rather than an
/// observation. It is deliberately short, because the wait it sits inside is not free: a window
/// held back is a hole in the desktop shaped like a video, which is the thing the pin's own cover
/// exists for (see `compose_parked_band`) and which a hover has never had. Too short leaves the
/// window being raised empty, which is what it did before; too long is a hole nobody is covering,
/// which is the same fault a few milliseconds wider.
pub(super) const VIDEO_FIRST_FRAME_HOLD: Duration = Duration::from_millis(120);

/// The player's window, and the moment this app first laid eyes on it. One player plays at a
/// time, so one slot is the whole of it, and it is keyed on the window rather than on the pid for
/// the reason the rest of this file is careful about handles: two players are alive at once
/// during a relaunch, and a handle says nothing about which one is behind it (see
/// `video_window_for`). A pid's own wait is started over rather than inherited, which is what
/// `forget_player_window_wait` is for.
static PLAYER_WINDOW_SEEN: Lazy<Mutex<Option<(isize, Instant)>>> = Lazy::new(|| Mutex::new(None));

/// Whether a window that has been standing for `seen_for` may go on the screen.
fn player_window_is_ready(seen_for: Duration) -> bool {
    seen_for >= VIDEO_FIRST_FRAME_HOLD
}

/// How long `hwnd` has been standing, which is this call that notes down the moment it was
/// first seen — a window other than the one held starts its wait over.
fn player_window_standing_for(hwnd: HWND) -> Duration {
    let key = hwnd.0 as isize;
    let mut seen = PLAYER_WINDOW_SEEN
        .lock()
        .unwrap_or_else(|slot| slot.into_inner());

    match seen.as_mut() {
        Some((held, since)) if *held == key => since.elapsed(),
        _ => {
            *seen = Some((key, Instant::now()));
            Duration::ZERO
        }
    }
}

/// Forget the wait a player's window was holding, so the next one this app finds starts its own
/// rather than inheriting a clock that has been running since the last film.
fn forget_player_window_wait() {
    PLAYER_WINDOW_SEEN
        .lock()
        .unwrap_or_else(|slot| slot.into_inner())
        .take();
}

/// Whether the player's window is up, which is not the same question as whether it exists: a
/// window this app is still holding back is a window it has found and nothing more, and the wait
/// for a player starting reads this so that it goes on standing — with this app's own spinner at
/// the pointer — until there is a film to see rather than a hole in the desktop (see `player_wait`).
pub(super) fn player_window_is_up() -> bool {
    let hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
    hwnd_val != 0 && player_window_is_ready(player_window_standing_for(HWND(hwnd_val as *mut _)))
}

/// Style and raise a known ffplay window.
pub(super) unsafe fn apply_noactivate_to_hwnd(hwnd: HWND) -> bool {
    // Store the video window HWND for cursor-over-preview detection. This is published before the
    // window is on screen, and has to be: `try_apply_noactivate_style` reaches the window again
    // through this handle, and `enum_windows_callback` skips hidden windows, so a window that was
    // held back and not published could never be found a second time to be let up. What the rest
    // of this app asks of it is therefore asked through `player_window_is_up`, which knows the
    // difference between a window that exists and one that is showing.
    VIDEO_HWND.store(hwnd.0 as isize, Ordering::SeqCst);

    // Add WS_EX_NOACTIVATE and WS_EX_TOPMOST to its extended style. This goes on before anything
    // else here, and whether or not the window is on screen: `WS_EX_NOACTIVATE` is what keeps the
    // player from taking the keyboard, and a window that is being held back for the want of a
    // frame is still a window that can take it (see the note over `try_apply_noactivate_style`).
    let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    SetWindowLongPtrW(hwnd, GWL_EXSTYLE, player_ex_style(current_style));

    // A window the player has made and not yet painted in is put away rather than raised, which
    // is the whole of what this hold is for (see `VIDEO_FIRST_FRAME_HOLD`).
    if !player_window_is_ready(player_window_standing_for(hwnd)) {
        let _ = ShowWindow(hwnd, SW_HIDE);
        return true;
    }

    // The style is put on whatever else is happening, and the *raise* is not. This is the one
    // place every raise of the player's window goes through — the monitor thread calls it about
    // every hundred milliseconds for as long as a film plays, whether or not the tick has asked for
    // anything — and that is what made a pinned window's volume popup open and vanish: the popup is
    // drawn over the media band, the media band of a video is this very window, and a raise puts
    // this window on top of the pin's own window every hundred milliseconds however carefully the
    // tick is holding off. The tick's hold-off was the right idea against the wrong caller.
    //
    // So the raise waits while the popup is up, and the style does not: `WS_EX_NOACTIVATE` is what
    // keeps the player from taking the keyboard, and that has to keep being asserted throughout,
    // because the player's window is created without it and this is the only thing that ever puts
    // it on (see `Bug 2`'s note over `try_apply_noactivate_style`). Nothing is asked for on the way
    // out — the popup's closing raises the pin's own window, and the next raise through here
    // settles the order (see `toggle_pin_volume`).
    //
    // **And it waits while a drag is parking the window, which is the same arrangement for the same
    // reason: this is a thread of its own, asking every hundred milliseconds whether the player's
    // window is still the window it styled.** Neither the drag's own raises nor the tick's can cover
    // it — they are two callers, and this one answers for a film nobody asked about. A park this
    // function undid would last from one monitor pass to the next, which at the pace it runs is a
    // hundred milliseconds of a drag with the film back on screen and the compositor re-blitting it
    // at the pointer's rate: the stutter the park was written for, and a band painted flat under a
    // window that is standing in it (see `pin_player_is_parked`).
    if pin_volume_open() || pin_player_is_parked() {
        return true;
    }

    // Force the video preview window to topmost so it doesn't hide behind Explorer
    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
    );
    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    true
}

/// Apply WS_EX_NOACTIVATE style to a window
/// Returns true if the window was found and modified
pub(super) unsafe fn try_apply_noactivate_style(pid: u32) -> bool {
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

pub(super) fn set_noactivate_for_process(pid: u32) {
    VIDEO_PID.store(pid, Ordering::SeqCst);
    // A player this app has not seen before starts its own wait, whichever handle it turns up
    // with: a window that was held back for a film that has ended must not be let straight on
    // to the screen because the next one reuses its handle (see `player_window_is_ready`).
    forget_player_window_wait();

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
/// what the zero was missing, and no flag in the code says so.
///
/// The one bit deliberately *not* set is bit 30, which is Windows' own "this is an auto-repeat"
/// and which SDL reads as a repeat and drops: a posted pause that repeated would toggle, toggle
/// and toggle back, so a hand resting on the bar's button would flicker rather than hold. The
/// low bit is the repeat *count*, which a real key-down of one press is 1.
pub(super) fn ffplay_key_lparam(scan: u32) -> LPARAM {
    LPARAM((scan as isize) << 16 | 1)
}

/// A key FFmpeg's player is toggled with, and the scan code it is named by.
pub(super) const FFPLAY_PAUSE_KEY: (u32, u32) = (b'P' as u32, 0x19);

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

/// Whether a player's window is there to be told something, which is not the same question as
/// whether a player is running (see `video_window_for`).
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
pub(super) fn ffplay_key_pause() -> bool {
    if pinned_level_is_owed_to_the_next_player() {
        return false;
    }

    let Some(hwnd) = video_window_for(VIDEO_PID.load(Ordering::SeqCst)) else {
        return false;
    };

    // SAFETY: the handle was just confirmed to be a live window of the player this app believes
    // is playing, so both messages below reach the player the caller means. Posting is the whole
    // of what is asked of it, and the two answers are the whole of what is known in return.
    unsafe {
        let (vk, scan) = FFPLAY_PAUSE_KEY;
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

/// True when `pid` still refers to a live ffplay process.
pub(super) fn is_ffplay_pid_alive(pid: u32) -> bool {
    unsafe {
        let Ok(handle) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) else {
            return false;
        };
        let matches = process_image_matches(handle, VIDEO_PROCESS_IMAGE_NAME);
        let _ = CloseHandle(handle);
        matches
    }
}

/// Terminate `pid` when it is still the ffplay process we spawned. Non-blocking:
/// it only requests the termination, it never waits for the process to exit.
pub(super) fn terminate_ffplay_pid(pid: u32) {
    unsafe {
        let Ok(handle) = OpenProcess(
            PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            pid,
        ) else {
            return;
        };
        if process_image_matches(handle, VIDEO_PROCESS_IMAGE_NAME) {
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
        // The player is confirmed gone, so the record of it goes with it rather than
        // being left for the next run to look for.
        engine_processes::forget(pid);
    }
}

/// Kill the last spawned ffplay process when it is still alive.
///
/// A process can outlive its `Child` handle: a kill may not take effect, or the
/// handle may be dropped before the process is confirmed gone. Without this a
/// surviving ffplay keeps its window on screen and the next hover would spawn a
/// second one next to it. The PID is only cleared once the process is confirmed
/// gone, so a later call retries instead of losing track of it.
pub fn kill_stray_video_process() {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid == 0 {
        return;
    }

    if !is_ffplay_pid_alive(pid) {
        clear_video_process_state(pid);
        return;
    }

    terminate_ffplay_pid(pid);

    if !is_ffplay_pid_alive(pid) {
        clear_video_process_state(pid);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A window the player has made is not put on screen until it has stood long enough to
    /// have been painted into, and that is the whole of the rule.
    ///
    /// Both ends are the point. A window let up at the moment it is found is composited with
    /// nothing in it, and an SDL window has no background of its own to erase, so what the
    /// compositor holds for that rectangle — which is the preview that was on screen there a
    /// moment earlier, at its own size and with black around it — is what the hand sees, at the
    /// video's own resolution, for as long as the first frame takes. And a window held back for
    /// good is a hover that never shows its video, which is why the answer is a bound rather than
    /// a refusal.
    #[test]
    fn a_players_window_is_held_off_the_screen_until_it_can_have_a_frame() {
        assert!(
            !player_window_is_ready(Duration::ZERO),
            "a window found this instant has nothing in it yet: showing it now is what composited \
             the previous preview's frame inside the video's box"
        );
        assert!(
            !player_window_is_ready(VIDEO_FIRST_FRAME_HOLD / 2),
            "and it is a wait rather than a single tick: half the bound is still a window that \
             has not been painted into"
        );
        assert!(
            player_window_is_ready(VIDEO_FIRST_FRAME_HOLD),
            "the bound itself is the moment the window may go up, or the film never arrives"
        );
        assert!(
            player_window_is_ready(VIDEO_FIRST_FRAME_HOLD * 4),
            "and a window past the bound is never held back again, however long it has been up"
        );
    }

    /// The wait belongs to the window, and a window this app has not seen starts its own.
    ///
    /// The second window here is one Windows handed back the handle the first film used, which
    /// is the case that lets a film's own clock carry over: nothing about the handle says which
    /// player is behind it, so the wait has to be taken on the window being different rather
    /// than on the handle being different. The wait running rather than restarting is what the
    /// bound above already covers.
    #[test]
    fn a_window_that_is_not_the_one_being_held_starts_its_own_wait() {
        forget_player_window_wait();

        let first = HWND(0x1000 as *mut _);
        let second = HWND(0x2000 as *mut _);

        assert!(
            !player_window_is_ready(player_window_standing_for(first)),
            "a window found this instant has nothing in it, whatever the handle says"
        );
        assert!(
            !player_window_is_ready(player_window_standing_for(first)),
            "and asking again while it is still inside the bound is not a new window"
        );

        forget_player_window_wait();
        assert!(
            !player_window_is_ready(player_window_standing_for(first)),
            "forgetting the wait puts the next film back at the start of it, so a film that ends \
             inside the bound does not leave the next one already held past it"
        );
        assert!(
            !player_window_is_ready(player_window_standing_for(second)),
            "and a window this app has not seen before is not answered off an old film's clock"
        );
    }
}

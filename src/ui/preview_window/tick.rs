//! How long the loop waits between ticks, and what it waits on: the preview channel, the
//! watchdogs that notice a thread which has stopped answering, and the message pump each wait
//! ends at.

use super::*;

/// How long the preview thread waits when there is nothing on screen: no preview
/// to animate, nothing to repaint, and no pointer region left to keep in step.
/// The wait is the preview channel rather than a sleep, so a hover is answered
/// the moment it arrives and this interval is only a ceiling on how long a window
/// message — a resume, a display change — waits to be noticed.
pub(super) const IDLE_WAIT_MS: u64 = 500;
/// How long the preview thread waits when a preview is on screen but nothing in
/// it moves: a static picture or page of text. The wait is the preview channel
/// rather than a sleep, so a Hide/Show answers the moment it arrives and this
/// interval is only a ceiling on polled work (pin key counts, resume flag) and
/// window messages. ~7x/s instead of ~60x/s; pinned static ticks use their own
/// shorter ceiling so caption buttons stay snappy (see the wait below).
pub(super) const STATIC_WAIT_MS: u64 = 150;
/// Ceiling for a static tick while a pin is up: the wait wakes on window input
/// at once (see `wait_preview_channel`), so a drag follows the hand at the
/// pointer's pace rather than at this pace — this stays short so caption
/// buttons still feel instant while waking 3x less often than the frame loop.
pub(super) const STATIC_PIN_WAIT_MS: u64 = 50;
/// How long the preview thread waits between the frames of something that is moving. The wait
/// wakes on window input as well as on this interval, so this is a ceiling and not a rate: a
/// drag follows the hand because the wait ends when the pointer does, not because a tick came
/// round.
pub(super) const FRAME_WAIT_MS: u64 = 16;

/// How long the preview thread waits before it comes round again, which is the loop's whole
/// cadence in one place.
///
/// Four bands, and which one a tick is in is a decision rather than an arithmetic result: it is
/// the difference between a loop that wakes sixty times a second for the life of the process
/// doing nothing, and one that wakes twice a second and still answers every button.
///
/// *Nothing on screen* does not answer to its cadence at all. There is nothing to animate,
/// nothing to repaint and nothing left to keep in step, so the wait becomes the preview channel
/// itself and a hover is answered as it arrives rather than on the next tick. `dynamic` and
/// `pinned` are ignored in that band and are named that way rather than left to the caller to
/// work out, because a caller that has to decide which arguments matter has to re-decide the
/// question this answers.
///
/// *Something moving* is the frame rate, and it is held for anything that has a picture to
/// advance: a video, an animation, a sound's card, a transport bar, a spinner.
///
/// *Something still* is the band the two settings above divide, and the split is the whole of
/// the argument for having two: a pinned static document wakes at `STATIC_PIN_WAIT_MS` and a
/// hover's at `STATIC_WAIT_MS`, because a pin's caption carries the buttons that close it and a
/// static preview's carries nothing at all. Three times as often is not three times the CPU —
/// both bands sleep on the channel and wake on input — but it is the difference between a
/// caption that feels instant and one that does not.
///
/// Two commits tuned these numbers and disagreed with each other, which is what a decision
/// with nowhere to be tested goes: there was no place to ask whether the number was still right,
/// so each change was argued from what it seemed to cost.
pub(super) fn wait_before_the_next_tick(
    nothing_on_screen: bool,
    dynamic: bool,
    pinned: bool,
) -> u64 {
    if nothing_on_screen {
        IDLE_WAIT_MS
    } else if dynamic {
        FRAME_WAIT_MS
    } else if pinned {
        STATIC_PIN_WAIT_MS
    } else {
        STATIC_WAIT_MS
    }
}

/// Whether this tick takes a full look at the media behind the preview, or trusts the hint.
///
/// A full look is the expensive one — it takes the media lock and reads what is on screen — and
/// a static tick skips it entirely, which is what made the loop cheap enough to run on a
/// battery. Four things force one anyway, and each of them is a way the hint can be wrong:
///
/// * `dynamic` — the hint says something moved, and a hint that is wrong here costs a stale
///   frame rather than a late one.
/// * `a_wait` — a load, a walk, a video's start or a first frame is outstanding, and what is on
///   screen is a stand-in for something that has not arrived.
/// * `generation_moved` — a swap bumps the generation, and a new hover's media is a different
///   answer from the last one's.
/// * `since_the_last_one` — the backstop, and the only one of the four that catches a kind that
///   changes without a swap: streaming frames that land after the fact, which is a file whose
///   type this app only discovers by looking. Without it those frames are noticed on the next
///   swap rather than within half a second, which is a preview that stays a spinner under a
///   stream that is already playing.
///
/// The backstop is the one arm here that exists because of a symptom rather than a principle,
/// and it is also the one most likely to be argued away as a cost — so it is named rather than
/// left as a bare `elapsed()` in a four-hundred-line loop.
pub(super) fn needs_a_full_media_look(
    dynamic_hint: bool,
    a_wait: bool,
    generation_moved: bool,
    since_the_last_one: Duration,
) -> bool {
    dynamic_hint
        || a_wait
        || generation_moved
        || since_the_last_one >= Duration::from_millis(STATIC_MEDIA_REFRESH_MS)
}

/// How often the player's window is put back in front, which is split by what is competing with
/// it for the top of the z-order.
///
/// A pinned video competes with the pin's own window and so keeps the tight band; a hover's
/// video competes with nothing, and re-asserting topmost for a tooltip is five DWM reorders a
/// second spent on a window that was in front when nobody clicked anything. This is the one
/// number in the loop that was moved for that reason alone, and it is two arms of an `if` where
/// the arms are the decision.
pub(super) fn topmost_cadence_ms(pinned: bool) -> u64 {
    if pinned {
        PIN_TOPMOST_REASSERT_MS
    } else {
        HOVER_TOPMOST_REASSERT_MS
    }
}

/// How often the page an engine is drawing is looked for on disk.
///
/// The look is a content probe, a folder index and two cache keys, and it is worth doing
/// several times less often than the loop turns while a render takes seconds. Nothing is
/// gained by asking sooner: a page that has not landed has not landed however often it is
/// looked for, and the spinner turning a little longer is the whole of the cost.
pub(super) const ENGINE_PAGE_POLL: Duration = Duration::from_millis(150);
/// How often the pointer hold region is republished while a static preview is
/// up. The region only changes on show/move/resize/swap/hide, so a slow
/// heartbeat plus an immediate publish on those transitions is the whole of
/// what the hook needs, without the per-tick lock + `GetWindowRect`.
pub(super) const POINTER_HOLD_HEARTBEAT_MS: u64 = 500;
/// Topmost re-assertion cadence split by context: a pinned video needs the
/// tighter band (it competes with the pin's own window), a hover video does
/// not and gets the slow one so DWM is not reordered 5x/s for a tooltip.
pub(super) const PIN_TOPMOST_REASSERT_MS: u64 = 200;
pub(super) const HOVER_TOPMOST_REASSERT_MS: u64 = 1000;
/// How often a static tick still takes one full look at the media, so a kind
/// change that arrives without a swap (streaming frames landing late) is
/// noticed within half a second rather than never.
pub(super) const STATIC_MEDIA_REFRESH_MS: u64 = 500;
// Message passing for thread communication
/// The event the preview channel's sends are announced on, so that a thread waiting for one
/// sleeps on the channel as well as on the message queue.
///
/// The wait the loop does was a wait on the queue alone, polled in eight-millisecond slices
/// because nothing would otherwise wake it: a thread with nothing on screen spent the whole of
/// `IDLE_WAIT_MS` waking to ask an empty channel whether it had anything, several times a
/// second, for the whole life of the process. Signalling this on the send is what lets the
/// wait block on the channel too — so a hover is answered when it is sent rather than at the
/// end of a slice, and an idle loop costs one kernel wait per tick instead of one per slice.
pub(super) fn preview_channel_event() -> HANDLE {
    // Kept as the number a handle is rather than as the handle, for the reason
    // `NOACTIVATE_WAKE` is: a handle is a raw pointer, and this is shared across threads.
    static EVENT: Lazy<isize> = Lazy::new(|| {
        // Safety: an event with no name and no security descriptor, manual-reset, which is
        // what a wait that may be woken several times before it is read wants.
        unsafe { CreateEventW(None, true, false, None) }
            .map(|handle| handle.0 as isize)
            .unwrap_or_default()
    });

    HANDLE(*EVENT as *mut core::ffi::c_void)
}

/// Send a message to the preview loop and wake it, so that a wait in progress ends at once.
pub(super) fn send_preview(message: PreviewMessage) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(tx) = sender.as_ref() {
            let _ = tx.send(message);
            // Signalled after the send rather than before, so an armed wait can only be woken
            // by a message that is already in the channel to be read.
            unsafe {
                let _ = SetEvent(preview_channel_event());
            }
        }
    }
}

pub static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

// Use AtomicIsize for the HWND pointer (thread-safe)
pub(super) static PREVIEW_HWND: AtomicIsize = AtomicIsize::new(0);

/// The clock the preview loop's liveness is read on, and when that loop was last
/// seen running.
///
/// One thread owns the answer and another reads it: the Explorer hook is what
/// decides about a preview loop that has stopped answering — the engines such a loop
/// is holding warm are ended from there instead (see `PREVIEW_STALL_MS` in
/// `explorer_hook`) — and it cannot ask the loop itself, because a question put to a
/// loop that has stopped is the one question that would not be answered. So the loop
/// notes its own tick and the hook reads how long ago the last one was. One relaxed
/// store per tick and one relaxed load per look is the whole cost of it.
pub(super) static PREVIEW_CLOCK: Lazy<Instant> = Lazy::new(Instant::now);
pub(super) static PREVIEW_ALIVE_MS: AtomicU64 = AtomicU64::new(0);

/// Note that the preview loop has run a tick — whether the tick did anything or not,
/// since a loop waiting on the channel for a hover is a loop that is working (see
/// `preview_stall_ms`).
pub(super) fn note_preview_alive() {
    PREVIEW_ALIVE_MS.store(
        PREVIEW_CLOCK.elapsed().as_millis() as u64,
        Ordering::Relaxed,
    );
}

/// The clock the pin's own liveness is read on, and when the loop last turned a tick
/// with a pin up.
///
/// This is a second clock from the one above, and it is a second because the two are read
/// by different watchers for different reasons. `PREVIEW_ALIVE_MS` answers "is the preview
/// loop working at all", and the Explorer hook is what asks it — a loop stopped on an idle
/// channel is a loop that is working. This one answers "is the loop turning while a window
/// the user is looking at is up", which is a much narrower question with a much shorter
/// bound, and nobody inside the loop can ask it: a loop that has stopped is a loop that
/// cannot notice it has stopped (see `spawn_pin_watchdog`).
pub(super) static PIN_ALIVE_MS: AtomicU64 = AtomicU64::new(0);
pub(super) static PIN_CLOCK: Lazy<Instant> = Lazy::new(Instant::now);

/// How long a pin may hold the loop before it is given up on.
///
/// It is generous on purpose, and deliberately so: the loop is a busy one and a pinned
/// window legitimately does work on it — starting a player, laying out a large picture,
/// reading a folder the first time it is walked. A bound that fires on those would close
/// a window out from under a user who was reading it. What it is long enough to outlast is
/// the *repeatable* case: a window that is frozen rather than busy, where the loop is not
/// coming back at all and every button on it is dead.
pub(super) const PIN_STALL_MS: u64 = 20_000;

/// Note that the loop turned a tick with a pin up. One relaxed store per tick (see
/// `note_preview_alive`, which is the same idea for the loop as a whole).
pub(super) fn note_pin_alive() {
    PIN_ALIVE_MS.store(PIN_CLOCK.elapsed().as_millis() as u64, Ordering::Relaxed);
}

/// Watch a pinned window's loop and give up on a window that is not coming back.
///
/// This is the last-resort half of the answer to a pin whose buttons do nothing. The other
/// half is that nothing slow runs on the loop at all (see `PinPlanner`), and this exists for
/// whatever that does not anticipate: a window that cannot be closed is a bug with no way
/// out but restarting the app, and a pin the user can lose is a far better answer than that.
///
/// So it is blunt by design. When the loop has not turned for `PIN_STALL_MS` while a pin is
/// up, the pin is ended from outside the loop, and the loop is left alone — whatever it is
/// stuck in is not something another thread can safely interrupt.
///
/// The take-down is the whole of the road out of a pin and not part of it, because a pin
/// taken down in halves is worse than one not taken down at all: the player's window and the
/// browser's are somebody else's windows, and leaving either standing puts a playing video
/// or a rendered document over the live previews that resume behind it, with no transport
/// bar left to stop them. The bubble is taken down for the same reason — a collapsed pin
/// whose state has been cleared but whose bubble is still on screen is a window nothing
/// will ever take away again.
///
/// The state and the window are one call now, because they were one thing and this thread
/// had drifted a copy of its own: it cleared the pin and two flags and left the keyboard
/// the pin had claimed, the focusable window, and the queued walk all standing. So the
/// caller says why the pin is going rather than choosing a road — `Reason::Hung`, which is
/// the only reason that is not the loop's own tick — and the road is derived from it. What
/// the road changes is what this thread may do of the window directly, which is the pointer
/// (a capture belongs to the thread that took it) and the hide (a window procedure that is
/// the thing being waited on would only make this wait too). Losing this thread costs
/// nothing; losing the only one that can take a window down costs the window.
pub(super) fn spawn_pin_watchdog() {
    std::thread::spawn(|| {
        while RUNNING.load(Ordering::Acquire) {
            std::thread::sleep(Duration::from_millis(500));

            if !pinned() {
                // Nothing to watch. The clock is left where it is, and the next pin is
                // given a full bound rather than inheriting the age of the last one.
                continue;
            }

            let now = PIN_CLOCK.elapsed().as_millis() as u64;
            if now.saturating_sub(PIN_ALIVE_MS.load(Ordering::Relaxed)) < PIN_STALL_MS {
                continue;
            }

            // The whole of what a road out of a pin does, state and window, from whichever
            // end is taking it. The state list was made one list by `end_pin_state_guards`;
            // the window work beside it was still written out by hand here, and *this* copy
            // was the one that drifted — it kept the keyboard the pin had claimed, kept the
            // window focusable, and left the walk a caption button had queued to be answered
            // into a pin that no longer existed, so a pin killed for being hung left a desktop
            // nothing could be clicked on. Both halves are now one call with the reason as its
            // only argument, and the reason is what says which end is taking it (see
            // `pin_window::end_pin`).
            let window = Win32PinWindow;
            // TEMP-WEDGE: see the block below `preview_stall_ms`.
            wedge_log("WATCHDOG ending hung pin");
            let owed = end_pin(Reason::Hung, &window);

            // What is not done on this thread is the window work, because the loop is the
            // thread that owns the windows and this one is the thread that has given up on
            // it. The loop notices the cleared pin on its next turn and takes the windows down
            // itself; a loop that never comes back cannot, which is what the hide below and
            // the message the road posted are for.
            if owed == PinHide::OnItsOwnThread {
                std::thread::spawn(move || window.hide_pin_windows());
            }
        }
    });
}

/// How long the preview loop has been quiet, in milliseconds: the age of its last
/// tick, growing for as long as the loop is inside work that has not come back.
///
/// A loop that has not ticked since the app started answers with the age of the app,
/// which is the same answer a loop that is gone is owed — nothing here waits on the
/// loop or looks for it; it is only ever a number read.
pub fn preview_stall_ms() -> u64 {
    let alive = PREVIEW_ALIVE_MS.load(Ordering::Relaxed);
    (PREVIEW_CLOCK.elapsed().as_millis() as u64).saturating_sub(alive)
}

// TEMP-WEDGE instrumentation: pin down where the preview thread wedges on an
// ffplay -> native pin swap. Everything in this block goes away once diagnosed.
static WEDGE_TICK: AtomicU64 = AtomicU64::new(0);
/// Set once a native hold is outstanding; the settle entry/exit marks below
/// only log from then on, so an ordinary run stays quiet.
pub(super) static WEDGE_ARMED: AtomicBool = AtomicBool::new(false);
// TEMP-WEDGE: ring of the last pin-state acquisitions (thread, file, line,
// age), so the stackwatch can name whoever took the pin lock last. A holder
// that vanished leaves its thread id behind even where no live stack shows it.
static WEDGE_RING: Mutex<Vec<(u32, &'static str, u32, u64)>> = Mutex::new(Vec::new());
// TEMP-WEDGE: the preview thread's id, for the stack capture below.
static PREVIEW_TID: AtomicU32 = AtomicU32::new(0);

/// TEMP-WEDGE: record one acquisition of the pin lock.
pub(super) fn wedge_note_acquire(location: &'static std::panic::Location<'static>) {
    if let Ok(mut ring) = WEDGE_RING.lock() {
        ring.push((
            unsafe { windows::Win32::System::Threading::GetCurrentThreadId() },
            location.file(),
            location.line(),
            WEDGE_CLOCK.elapsed().as_millis() as u64,
        ));
        while ring.len() > 32 {
            ring.remove(0);
        }
    }
}

/// TEMP-WEDGE: dump the ring into the log.
pub(super) fn wedge_dump_ring() {
    let ring = WEDGE_RING.lock().map(|ring| ring.clone()).unwrap_or_default();
    let now = WEDGE_CLOCK.elapsed().as_millis() as u64;
    for (tid, file, line, at) in ring {
        let file = file.rsplit(['/', '\\']).next().unwrap_or(file);
        wedge_log(&format!(
            "pinlock tid={tid} {file}:{line} {}ms ago",
            now.saturating_sub(at),
        ));
    }
}
static WEDGE_STACK_DONE: AtomicBool = AtomicBool::new(false);
static WEDGE_CLOCK: Lazy<Instant> = Lazy::new(Instant::now);
static WEDGE_STALL_LOGGED: AtomicBool = AtomicBool::new(false);

/// The wedge log: a line per swap event plus a heartbeat while a pin is up,
/// so the tail shows the last thing the loop did before it stopped turning.
pub(super) fn wedge_log(line: &str) {
    let path = std::env::temp_dir().join("rhp-wedge.log");
    let stamp = WEDGE_CLOCK.elapsed().as_millis();
    let line = format!("[{stamp:>10} ms] {line}\n");
    if let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        use std::io::Write;
        let _ = file.write_all(line.as_bytes());
    }
}

/// Note one more turn of the preview loop (see `wedge_spawn_watchdog`).
pub(super) fn wedge_tick() {
    WEDGE_TICK.fetch_add(1, Ordering::Relaxed);
    // The thread being watched (see `wedge_spawn_stackwatch`).
    let tid = unsafe { windows::Win32::System::Threading::GetCurrentThreadId() };
    PREVIEW_TID.store(tid, Ordering::Release);
}

/// Run `f`, logging where it takes longer than `slow_ms` to come back —
/// a call that never comes back is the wedge, and its label is the last line.
pub(super) fn wedge_timed<T>(label: &str, slow_ms: u64, f: impl FnOnce() -> T) -> T {
    let started = Instant::now();
    let out = f();
    let ms = started.elapsed().as_millis() as u64;
    if ms >= slow_ms {
        wedge_log(&format!("SLOW {label} took {ms} ms"));
    }
    out
}

/// Watch an armed hold and, when the tick stops turning behind it, write the
/// wedged preview thread's own stack into the log: the exact call it is stuck
/// in, with symbol names from the PDBs beside the exe. Suspends the thread
/// only to copy its context out, then resumes it at once — everything slow
/// (walking, symbol lookup, logging) runs after the resume, so a thread held
/// mid-allocator cannot wedge this one on the allocator.
pub(super) fn wedge_spawn_stackwatch() {
    std::thread::spawn(|| {
        loop {
            std::thread::sleep(Duration::from_millis(500));
            if !RUNNING.load(Ordering::Acquire) {
                return;
            }
            if !WEDGE_ARMED.load(Ordering::Acquire) {
                continue;
            }
            let before = WEDGE_TICK.load(Ordering::Relaxed);
            std::thread::sleep(Duration::from_secs(6));
            if WEDGE_TICK.load(Ordering::Relaxed) != before {
                continue;
            }
            if WEDGE_STACK_DONE.swap(true, Ordering::AcqRel) {
                return;
            }
            wedge_log("stackwatch: tick stalled behind an armed hold, capturing");
            wedge_capture_stack();
            wedge_log("stackwatch: pin-lock acquisitions (newest last)");
            wedge_dump_ring();
            return;
        }
    });
}

/// Suspend the preview thread, copy its context, resume it, then walk and log.
fn wedge_capture_stack() {
    use windows::Win32::Foundation::HANDLE;
    use windows::Win32::System::Diagnostics::Debug::*;
    use windows::Win32::System::Threading::*;

    let tid = PREVIEW_TID.load(Ordering::Acquire);
    if tid == 0 {
        wedge_log("stackwatch: no preview thread recorded");
        return;
    }

    let process = unsafe { GetCurrentProcess() };
    // PDBs live beside the exe; name the folder explicitly rather than
    // trusting the default search path.
    let search = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.parent().map(|dir| dir.to_path_buf()));
    let wide: Vec<u16> = search
        .map(|dir| {
            dir.as_os_str()
                .encode_wide()
                .chain(std::iter::once(0))
                .collect()
        })
        .unwrap_or_else(|| vec![0]);
    unsafe {
        if SymInitializeW(
            process,
            windows_core::PCWSTR(wide.as_ptr()),
            // Invade the process so every loaded module (ntdll, the exe and
            // its PDBs) is visible to the walker and the lookup: without it
            // module bases resolve to nothing and the walk derails past the
            // system frames.
            true,
        )
        .is_err()
        {
            wedge_log("stackwatch: SymInitialize failed");
            return;
        }
    }

    let thread = unsafe {
        OpenThread(
            THREAD_SUSPEND_RESUME | THREAD_GET_CONTEXT | THREAD_QUERY_INFORMATION,
            false,
            tid,
        )
    };
    let Ok(thread) = thread else {
        wedge_log("stackwatch: OpenThread failed");
        return;
    };
    wedge_capture_one(process, thread, tid, "preview");

    // Every other thread too: whatever holds the lock the preview thread is
    // wedged on is one of them, and its own stack names where it stands.
    wedge_capture_others(process, tid);

    wedge_log("stackwatch: done");
}

/// Suspend one thread, copy its context, resume it at once, then walk the
/// copied context and log the frames under `label`.
fn wedge_capture_one(
    process: windows::Win32::Foundation::HANDLE,
    thread: windows::Win32::Foundation::HANDLE,
    tid: u32,
    label: &str,
) {
    use windows::Win32::System::Diagnostics::Debug::*;
    use windows::Win32::System::Threading::*;

    // The only two calls made while the thread is held: no allocation in
    // between, so a thread frozen mid-allocator cannot take this one down
    // with it.
    let mut context: CONTEXT = unsafe { std::mem::zeroed() };
    context.ContextFlags = CONTEXT_FULL_AMD64;
    let (suspended, context_ok) = unsafe {
        let suspended = SuspendThread(thread) != u32::MAX;
        let ok = GetThreadContext(thread, &mut context).is_ok();
        (suspended, ok)
    };
    if suspended {
        unsafe {
            ResumeThread(thread);
        }
    }
    if !suspended || !context_ok {
        unsafe {
            let _ = CloseHandle(thread);
        }
        wedge_log("stackwatch: suspend/context failed");
        return;
    }

    let mut frame: STACKFRAME64 = unsafe { std::mem::zeroed() };
    frame.AddrPC.Offset = context.Rip;
    frame.AddrPC.Mode = AddrModeFlat;
    frame.AddrFrame.Offset = context.Rbp;
    frame.AddrFrame.Mode = AddrModeFlat;
    frame.AddrStack.Offset = context.Rsp;
    frame.AddrStack.Mode = AddrModeFlat;

    let mut depth = 0u32;
    while depth < 64 {
        let walked = unsafe {
            StackWalk64(
                0x8664,
                process,
                thread,
                &mut frame,
                &mut context as *mut CONTEXT as *mut std::ffi::c_void,
                None,
                Some(wedge_fn_table_access),
                Some(wedge_get_module_base),
                None,
            )
        };
        if !walked.as_bool() || frame.AddrPC.Offset == 0 {
            break;
        }
        let addr = frame.AddrPC.Offset;
        let mut displacement: u64 = 0;
        let mut buf =
            vec![0u8; std::mem::size_of::<SYMBOL_INFOW>() + 256 * std::mem::size_of::<u16>()];
        let name = unsafe {
            let sym = buf.as_mut_ptr() as *mut SYMBOL_INFOW;
            (*sym).SizeOfStruct = std::mem::size_of::<SYMBOL_INFOW>() as u32;
            (*sym).MaxNameLen = 255;
            if SymFromAddrW(process, addr, Some(&mut displacement), sym).is_ok() {
                let len = (*sym).NameLen as usize;
                String::from_utf16_lossy(std::slice::from_raw_parts(
                    (*sym).Name.as_ptr(),
                    len.min(255),
                ))
            } else {
                String::from("<unknown>")
            }
        };
        wedge_log(&format!("stack [{label}] #{depth}: {name}+{displacement:#x} ({addr:#x})"));
        if frame.AddrReturn.Offset == 0 {
            break;
        }
        depth += 1;
    }
    unsafe {
        let _ = CloseHandle(thread);
    }
    wedge_log(&format!("stackwatch [{label}]: captured {depth} frames"));
}

/// TEMP-WEDGE: capture every other thread of this process under its own
/// label, so the holder of whatever the preview thread is wedged on shows
/// itself. Skips this watchdog thread (suspending it would wedge the watch).
fn wedge_capture_others(process: windows::Win32::Foundation::HANDLE, preview: u32) {
    use windows::Win32::System::Diagnostics::ToolHelp::*;
    use windows::Win32::System::Threading::*;

    let me_myself = unsafe { GetCurrentThreadId() };
    let own_pid = unsafe { GetCurrentProcessId() };
    let snapshot =
        unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0).unwrap_or_default() };
    if snapshot.is_invalid() {
        wedge_log("stackwatch: thread snapshot failed");
        return;
    }

    let mut entry: THREADENTRY32 = unsafe { std::mem::zeroed() };
    entry.dwSize = std::mem::size_of::<THREADENTRY32>() as u32;
    let mut more = unsafe { Thread32First(snapshot, &mut entry).is_ok() };
    while more {
        let tid = entry.th32ThreadID;
        if entry.th32OwnerProcessID == own_pid && tid != me_myself && tid != preview {
            let thread = unsafe {
                OpenThread(
                    THREAD_SUSPEND_RESUME | THREAD_GET_CONTEXT | THREAD_QUERY_INFORMATION,
                    false,
                    tid,
                )
            };
            match thread {
                Ok(thread) => {
                    let label = format!("thread-{tid}");
                    wedge_capture_one(process, thread, tid, &label);
                }
                Err(_) => {
                    wedge_log(&format!("stackwatch [thread-{tid}]: OpenThread failed"));
                }
            }
        }
        more = unsafe { Thread32Next(snapshot, &mut entry).is_ok() };
    }
    unsafe {
        let _ = CloseHandle(snapshot);
    }
}

/// TEMP-WEDGE shims: `StackWalk64` wants plain `extern "system"` callbacks,
/// and the imported DbgHelp helpers are generic wrappers, so these adapt.
unsafe extern "system" fn wedge_fn_table_access(
    hprocess: HANDLE,
    addrbase: u64,
) -> *mut std::ffi::c_void {
    use windows::Win32::System::Diagnostics::Debug::SymFunctionTableAccess64;
    SymFunctionTableAccess64(hprocess, addrbase)
}

/// TEMP-WEDGE: see `wedge_fn_table_access`.
unsafe extern "system" fn wedge_get_module_base(hprocess: HANDLE, address: u64) -> u64 {
    use windows::Win32::System::Diagnostics::Debug::SymGetModuleBase64;
    SymGetModuleBase64(hprocess, address)
}

/// Watch the loop turn and log when it stops while a pin is up: the tail
/// then reads as the heartbeat stopping after the last swap event.
pub(super) fn wedge_spawn_watchdog() {
    std::thread::spawn(|| {
        let mut last = WEDGE_TICK.load(Ordering::Relaxed);
        let mut still = 0u32;
        while RUNNING.load(Ordering::Acquire) {
            std::thread::sleep(Duration::from_secs(1));
            if !pinned() {
                last = WEDGE_TICK.load(Ordering::Relaxed);
                still = 0;
                continue;
            }
            let now = WEDGE_TICK.load(Ordering::Relaxed);
            wedge_log(&format!("heartbeat tick={now}"));
            if now == last {
                still += 1;
                if still == 3 && !WEDGE_STALL_LOGGED.swap(true, Ordering::AcqRel) {
                    wedge_log("TICK STALLED with a pin up — loop is wedged");
                }
            } else {
                still = 0;
                WEDGE_STALL_LOGGED.store(false, Ordering::Release);
            }
            last = now;
        }
    });
}

/// The number of times the window has been taken down, and the lock that makes a
/// take-down and the reveal it races one step rather than two threads writing the
/// window's visibility at once.
///
/// `hide_preview` moves the count there and then, on the Explorer hook's thread,
/// while the frame that puts it up is installed here when a load lands. A
/// load that lands in the moment after the pointer left would otherwise put the
/// preview back up for a file nobody is on any more, to be taken down again by the
/// next tick: the preview that blinks. A load carries the count it was started
/// under, and a hide moves it, so a reveal whose count has moved is refused. It is
/// a comparison and not a wait — nothing here ever holds a preview back from going
/// up, and the window that comes down with the count is posted rather than sent, so
/// a loop that is busy is never a thread the hide waits on (see `hide_preview`).
pub(super) static HIDDEN_EPOCH: Mutex<u64> = Mutex::new(0);

/// The hide count a load starting now is under.
pub(super) fn hidden_epoch() -> u64 {
    match HIDDEN_EPOCH.lock() {
        Ok(epoch) => *epoch,
        Err(poisoned) => *poisoned.into_inner(),
    }
}

/// Whether a load is still the one the pointer asked for, `hidden` being the guard
/// the caller is holding the answer under (see `HIDDEN_EPOCH`): a hide that has run
/// since the load started is the pointer having left it, and the frame it lands
/// with is not to be shown.
pub(super) fn hover_still_wanted(
    hidden: &Option<MutexGuard<'static, u64>>,
    pl: &PendingLoad,
) -> bool {
    hidden
        .as_ref()
        .map(|epoch| **epoch == pl.hide_epoch)
        .unwrap_or(true)
}

/// Whether the media on screen is holding a frame of a video the media engine plays rather
/// than the placeholder a preview of one is loaded with.
///
/// A video preview is loaded with a frame of the size the layout planned and nothing in it:
/// what the engine decodes is written into that frame, one frame at a time, from the first
/// one the engine hands over (see `take_native_video_frame`). Nothing is taken from it at
/// all until the engine has reported that it has one — a surface the engine has not drawn
/// into is not a transparent placeholder but a rectangle of zeros, and a frame taken from
/// one is an opaque picture of nothing (see `video_player::Session::copy_into`) — so the
/// placeholder never reaches a screen, and this is the question whose answer holds the reveal
/// back instead (see `FirstFrameWait`), and what a preview already on screen is asked before
/// it is painted again for one.
pub(super) fn media_holds_a_frame() -> bool {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .map(|media| media.media_type.is_native_video() && media.current_frame_is_opaque())
        })
        .unwrap_or(false)
}

/// How long a carried drag follows a hand that has stopped moving before the tick is given back.
///
/// Long enough that a hand which pauses between movements — which is most of them, and all of them
/// at the moments a window is being lined up — is not mistaken for one that has let go, and short
/// enough that a drag left down over a window nobody is touching is a pause rather than a thread
/// that has stopped turning.
pub(super) const PIN_DRAG_HAND_RESTED: Duration = Duration::from_millis(150);

/// Give the window procedures this thread owns their turn, without blocking on them.
///
/// The loop drains its own queue at the top of every tick, and a drag that is being carried is a
/// tick that may not come round for a while (see `carry_pin_drag_with_the_hand`), so the drain
/// comes in here too. It is the same drain, which is the point: a window being carried must go on
/// answering everything a window answers, since a window that stops answering is the fault this
/// whole arrangement of ends exists to prevent (see `settle_pinned_engine_drag`).
pub(super) fn pump_window_messages() {
    let mut msg = MSG::default();

    unsafe {
        while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

/// How long a carry of a pinned drag may hold the loop before the tick is given it back.
///
/// Long enough that a hand dragging a window across a 4K display is carried the whole way at
/// the pointer's own rate, and short enough that a window left held down over something nobody
/// is touching is a pause rather than a loop that has stopped answering. It is the outer bound
/// the stillness test never provided: that one ends a carry when the hand stops, and says
/// nothing about a hand that never does.
pub(super) const PIN_DRAG_WAIT_MS: u64 = 2_000;

/// Wait for a preview message or for window input, whichever comes first.
///
/// This is what makes a pinned window follow the hand at the pointer's own
/// pace rather than at the pace of the repaint: waiting on the message queue instead of on the
/// channel wakes the moment the pointer moves, and costs nothing extra when idle.
///
/// The wait is on two things at once: the message queue, and the channel's own event. The event
/// is what the slicing this used to do was standing in for. A wait that could not see the channel
/// had to wake on a timer to ask it, so `wait_ms` was cut into eight-millisecond slices and each
/// one was a kernel wait followed by a poll of an empty channel — which for a loop with nothing
/// on screen, waiting `IDLE_WAIT_MS`, was a hundred kernel transitions a second for the whole
/// life of the process, doing nothing. A send now signals the event, so the wait ends where the
/// message lands, and an idle loop costs one wait per tick.
///
/// The event is manual-reset and is reset after the wait rather than at the send, which is what
/// makes the arm-then-send race safe: a message that arrived between the poll above and the wait
/// below has already signalled it, so the wait returns at once instead of sleeping through a
/// message that is in the queue.
pub(super) fn wait_preview_channel(
    rx: &Receiver<PreviewMessage>,
    wait_ms: u64,
) -> Option<PreviewMessage> {
    // A message sent just before the wait is answered without waiting at all.
    if let Ok(message) = rx.try_recv() {
        return Some(message);
    }

    let event = preview_channel_event();
    let handles = [event];
    unsafe {
        let _ = MsgWaitForMultipleObjectsEx(
            Some(&handles),
            wait_ms as u32,
            QS_ALLINPUT,
            MWMO_INPUTAVAILABLE,
        );
        let _ = ResetEvent(event);
    }

    // Woken by input, by a message, or by the timeout. Either way the channel is only ever
    // polled, never blocked on, so the queue is never left unpumped while a drag is going.
    if let Ok(message) = rx.try_recv() {
        return Some(message);
    }

    // Input alone is a reason to go round the loop now rather than to sit out the rest of
    // the tick: the queue holds the move the wait woke for.
    let mut msg = MSG::default();
    unsafe {
        if PeekMessageW(&mut msg, None, 0, 0, PM_NOREMOVE).as_bool() {
            return None;
        }
    }

    None
}

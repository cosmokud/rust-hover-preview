//! What a run costs, when `RHP_HOOK_TRACE` asks for it: the crossings into the shell
//! that were counted, and the file they are written to.
//!
//! The questions that decide whether a probe is worth making at all sit beside them,
//! and so does the gate that admits a file as one this app draws a preview for: a run
//! that paid for a probe of a file no engine is started for has bought nothing.

use super::*;

/// Whether `RHP_HOOK_TRACE` asks for what the probes cost to be written out.
///
/// A run without the variable set writes nothing anywhere. The counters below are
/// kept either way and this is only what says whether they are ever written down:
/// what they are for is a run that is being measured, and what a hover costs is
/// otherwise only reasoned about. One of them in particular — how often the view
/// that answered last is the answer rather than a walk — is the number a measurement
/// settles and an argument does not.
pub(super) static HOOK_TRACE: Lazy<bool> =
    Lazy::new(|| std::env::var_os("RHP_HOOK_TRACE").is_some());

/// The crossings into the shell a probe has made since the last line was written:
/// the window collections walked, the windows those walks passed between them,
/// the probes answered from a set of views already read, the probes whose place or item
/// was answered by the one view the pointer is in, the item walks through the view's
/// provider, the points resolved, and the looks answered from the item under the pointer
/// without asking the shell again.
///
/// The pair to read together is the walks against the sets kept: reading a window's
/// views is one crossing into the shell per Shell window the desktop holds, and a
/// window that holds tabs is one registration per tab — so a set read again on every
/// probe is the cost that grows with how many tabs are open, and a set kept is what
/// says that cost is paid once for a window rather than once for a tick. Beside them,
/// the anchored count against the points resolved says how often the window an item and
/// a place are read in was enough to tell a frame's views apart — a window holding tabs
/// answered by one view rather than by all of them.
pub(super) static PROBE_VIEW_WALKS: AtomicU64 = AtomicU64::new(0);
pub(super) static PROBE_VIEW_WINDOWS: AtomicU64 = AtomicU64::new(0);
/// The probes a frame's views answered without being read again, and the probes the view
/// the pointer is in was asked for rather than every view of the frame — see `frame_views`
/// and `ItemWindow`.
pub(super) static PROBE_VIEW_SETS_KEPT: AtomicU64 = AtomicU64::new(0);
pub(super) static PROBE_VIEW_ANCHORED: AtomicU64 = AtomicU64::new(0);
pub(super) static PROBE_ITEM_WALKS: AtomicU64 = AtomicU64::new(0);
pub(super) static PROBE_POINTER_RESOLUTIONS: AtomicU64 = AtomicU64::new(0);
/// The looks the item under the pointer was answered from, without the shell being
/// asked at all: a hand moving along one row of a list is many points and one item, and
/// this is what says how much of a sweep that answers for — see `AnsweredItem`.
pub(super) static PROBE_ITEM_MEMO_HITS: AtomicU64 = AtomicU64::new(0);

/// The slowest of the view walks since the last line, in milliseconds — what a probe
/// costs in time rather than in calls, which is the number that says whether the
/// shell was held long enough for anyone to feel it.
pub(super) static PROBE_VIEW_SLOWEST_MS: AtomicU64 = AtomicU64::new(0);

/// The slowest of the item walks since the last line, in milliseconds, and timed apart
/// from the walks above on purpose: they are different work against different providers
/// — one walks Explorer's window collection, the other walks its accessibility tree —
/// and one of them being cheap says nothing about the other.
pub(super) static PROBE_ITEM_SLOWEST_MS: AtomicU64 = AtomicU64::new(0);

/// Count one of those crossings.
///
/// Counted whether or not the trace is on. It is one relaxed add to a counter this
/// thread owns, set against the crossings the counters are counting, which are calls
/// into another process — and gating it would buy nothing while making the numbers
/// depend on when the variable happened to be read.
pub(super) fn note_probe(counter: &AtomicU64) {
    counter.fetch_add(1, Ordering::Relaxed);
}

/// Note how long one of those crossings took, keeping the slowest of them.
pub(super) fn note_probe_ms(counter: &AtomicU64, elapsed: Duration) {
    counter.fetch_max(elapsed.as_millis() as u64, Ordering::Relaxed);
}

/// Where the probe counts are written, or nothing at all when `RHP_HOOK_TRACE` did
/// not ask for them. Resolved once, since where a file goes does not change under a
/// run, and read by the loop rather than by every count.
pub(super) fn hook_trace_path() -> Option<PathBuf> {
    HOOK_TRACE.then(|| std::env::temp_dir().join("rhp-hook-trace.log"))
}

/// Write what the probes have cost since the last line to `path`, at most once a
/// second.
///
/// Once a second rather than once a probe, because this writes a file and a file
/// written per probe is the one thing a hover must never do — the counters are here
/// to say what a hover costs, and they must not be what makes it cost more.
pub(super) fn flush_probe_counts(now: Instant, last: &mut Instant, path: &Path) {
    if now.duration_since(*last) < Duration::from_millis(1000) {
        return;
    }
    *last = now;

    let walks = PROBE_VIEW_WALKS.swap(0, Ordering::Relaxed);
    let windows = PROBE_VIEW_WINDOWS.swap(0, Ordering::Relaxed);
    let kept = PROBE_VIEW_SETS_KEPT.swap(0, Ordering::Relaxed);
    let anchored = PROBE_VIEW_ANCHORED.swap(0, Ordering::Relaxed);
    let items = PROBE_ITEM_WALKS.swap(0, Ordering::Relaxed);
    let points = PROBE_POINTER_RESOLUTIONS.swap(0, Ordering::Relaxed);
    let memo = PROBE_ITEM_MEMO_HITS.swap(0, Ordering::Relaxed);
    let view_slowest = PROBE_VIEW_SLOWEST_MS.swap(0, Ordering::Relaxed);
    let item_slowest = PROBE_ITEM_SLOWEST_MS.swap(0, Ordering::Relaxed);

    if walks == 0 && items == 0 && points == 0 {
        return;
    }

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        use std::io::Write;
        let _ = writeln!(
            file,
            "points {points}  item walks {items} (slowest {item_slowest}ms)  view walks {walks} (windows {windows}, slowest {view_slowest}ms)  sets kept {kept}  anchored {anchored}  item memo {memo}"
        );
    }
}

/// Write one line about what a click in the listing did to a pinned window, and write
/// nothing at all where `RHP_HOOK_TRACE` did not ask for it: the line is built inside the
/// guard, so a run with no trace pays nothing for one. A second file rather than more
/// lines in the first, because this is a sequence to read in order and that one is a count
/// to read once.
///
/// This is the question the counters cannot answer: they say how many crossings the shell
/// was asked for, and not whether the click was seen at all, whether the pointer was over
/// a listing, what the lookup made of it, or whether the pin was asked for another file.
/// A click lost to a race reads, in a counter, exactly like a click that never happened.
macro_rules! note_pin_click {
    ($($line:tt)*) => {
        // The gate is read here rather than through `HOOK_TRACE`, so the
        // macro stands on its own wherever it is written from: the preview
        // loop's take of the dismissal ask writes one too, and that is a
        // part of the app this module is not in scope of. The line itself
        // is still built inside the guard, so a run with no trace pays
        // nothing for one.
        if std::env::var_os("RHP_HOOK_TRACE").is_some() {
            if let Ok(mut file) = std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(std::env::temp_dir().join("rhp-pin-click-trace.log"))
            {
                use std::io::Write;
                let _ = writeln!(file, $($line)*);
            }
        }
    };
}

// The pin's own click traces are written from the two parts that watch it, so the macro
// travels with the counter it is written beside.
pub(crate) use note_pin_click;

/// A path as the name of the file in it, for the lines above: the folder is the same on
/// both sides of every question they ask, and a path long enough to be worth reading is a
/// path too long to read.
pub(super) fn trace_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| "-".to_string())
}

/// The class of the window under the pointer, so that a click the shell never heard of
/// says so rather than reading like a click that found nothing.
pub(super) fn window_class_of(window: HWND) -> String {
    if window.is_invalid() {
        return "none".to_string();
    }

    let mut buffer = [0u16; 96];
    let written = unsafe { GetClassNameW(window, &mut buffer) };
    if written <= 0 {
        return "unreadable".to_string();
    }

    String::from_utf16_lossy(&buffer[..written as usize])
}

/// The one line a tick of a click in hand is written as, out of the facts the loop has
/// already read for its own reasons. `clicked` is zero on every tick of the retry, and on
/// every tick of a run in which the press was never read at all — the one reading that says
/// the click was lost before any of this had a chance to answer it.
pub(super) fn trace_click(
    pointer: &PointerTick,
    showing: &Path,
    over_explorer: bool,
    clicked: bool,
    resolved: Option<&PathBuf>,
    offered: bool,
    held_for: Option<Duration>,
) {
    note_pin_click!(
        "click {}  fg {}  over {}  at {},{}  win {}  showing {}  resolved {}  offered {}  held {}",
        clicked as u8,
        is_foreground_explorer() as u8,
        over_explorer as u8,
        pointer.point.x,
        pointer.point.y,
        window_class_of(pointer.window),
        trace_name(showing),
        resolved.map_or_else(|| "-".to_string(), |path| trace_name(path)),
        offered as u8,
        held_for.map_or_else(|| "-".to_string(), |held| format!("{}ms", held.as_millis())),
    );
}

/// End the engines held by a preview loop that has stopped ticking, which is what the
/// loop's own idle tiers would have done had it been running.
///
/// Engines only, and never the player a video preview ran: what ends that is the
/// dismissal that hid it, by id (see `engine_processes`). The one engine left alone is
/// a browser drawing a document — that window *is* the preview while it is up, with
/// this app's own window hidden behind it — so a document being read is not an engine
/// being kept, and a loop that has stopped could not put it back. Nothing here waits
/// on anything, and nothing is written anywhere but the trace.
pub(super) fn end_engines_of_a_stalled_preview(quiet_ms: u64, trace: Option<&Path>) {
    if webview_preview::showing_path().is_some() {
        return;
    }

    engine_processes::terminate_all_owned();
    note_stalled_preview(trace, quiet_ms);
}

/// Note a preview loop that stopped ticking, where a trace is being written: how long
/// it had been quiet, and that the engines it was holding were ended for it. Written
/// once for the stall, like everything else here — the trace is there to say what a
/// hover costs, and must never be what makes it cost more.
pub(super) fn note_stalled_preview(path: Option<&Path>, quiet_ms: u64) {
    let Some(path) = path else {
        return;
    };

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        use std::io::Write;
        let _ = writeln!(
            file,
            "preview loop quiet {quiet_ms}ms: its engines were ended from the hook"
        );
    }
}

/// Drop what describes a view: the folder the last probe remembered for a window,
/// and whether a window is one of Explorer's. The window the pointer is in is
/// resolved again the next time it is asked about.
pub(super) fn clear_shell_view_probe_caches() {
    if let Ok(mut cache) = EXPLORER_WINDOW_CACHE.lock() {
        cache.clear();
    }
    if let Ok(mut cache) = EXPLORER_LAST_REAL_FOLDERS.lock() {
        cache.clear();
    }
}

/// `MONITORINFOF_PRIMARY`, which the Windows bindings do not name: the flag
/// `GetMonitorInfo` sets on the display new windows open on.
pub(super) const MONITORINFOF_PRIMARY_FLAG: u32 = 0x1;

/// One display, as `EnumDisplayMonitors` walks them, added to the list the signature
/// is being built from. The list is what the caller's `LPARAM` points at.
pub(super) unsafe extern "system" fn collect_display(
    monitor: HMONITOR,
    _dc: HDC,
    _rect: *mut RECT,
    data: LPARAM,
) -> BOOL {
    let displays = &mut *(data.0 as *mut Vec<DisplayEntry>);

    let mut info = MONITORINFO {
        cbSize: std::mem::size_of::<MONITORINFO>() as u32,
        ..Default::default()
    };
    if !GetMonitorInfoW(monitor, &mut info).as_bool() {
        return BOOL(1);
    }

    let mut dpi_x = 0u32;
    let mut dpi_y = 0u32;
    let dpi = if GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &mut dpi_x, &mut dpi_y).is_ok() {
        dpi_x
    } else {
        0
    };
    let display = info.rcMonitor;

    displays.push(DisplayEntry {
        rect: (display.left, display.top, display.right, display.bottom),
        dpi,
        primary: info.dwFlags & MONITORINFOF_PRIMARY_FLAG != 0,
    });

    BOOL(1)
}

pub(super) fn current_display_signature() -> Option<DisplaySignature> {
    let mut displays: Vec<DisplayEntry> = Vec::new();

    unsafe {
        let data = LPARAM(&mut displays as *mut Vec<DisplayEntry> as isize);
        if !EnumDisplayMonitors(None, None, Some(collect_display), data).as_bool() {
            return None;
        }
    }

    (!displays.is_empty()).then_some(DisplaySignature { displays })
}

pub(super) fn display_signature_changed(
    previous: Option<&DisplaySignature>,
    current: &DisplaySignature,
) -> bool {
    previous
        .map(|signature| signature != current)
        .unwrap_or(false)
}

pub(super) fn recent_elapsed_within(elapsed: Option<Duration>, limit_ms: u64) -> bool {
    elapsed
        .map(|elapsed| elapsed <= Duration::from_millis(limit_ms))
        .unwrap_or(false)
}

pub(super) fn should_probe_keyboard_focus(recent_navigation_elapsed: Option<Duration>) -> bool {
    recent_elapsed_within(recent_navigation_elapsed, KEYBOARD_FOCUS_INPUT_GRACE_MS)
}

pub(super) fn should_probe_hover_resolver(
    preview_active: bool,
    cursor_moved: bool,
    recent_input_elapsed: Option<Duration>,
) -> bool {
    preview_active
        || cursor_moved
        || recent_elapsed_within(recent_input_elapsed, HOVER_RESOLVER_INPUT_GRACE_MS)
}

pub(super) fn should_probe_stationary_hover(already_probed: bool) -> bool {
    !already_probed
}

/// The pointer probe only matters while a mouse preview can be under the
/// pointer. Keyboard previews and a frozen pointer never trigger it.
pub(super) fn should_probe_preview_hover(
    pointer_frozen: bool,
    mouse_preview_active: bool,
    suppress_until_cursor_leaves: bool,
) -> bool {
    !pointer_frozen && (mouse_preview_active || suppress_until_cursor_leaves)
}

/// Whether the pointer has left the item the preview on screen is about, which is the
/// mouse having moved whatever the distance says.
///
/// How far a hand has come is a measurement of jitter, and one that is deliberately
/// coarse while the keyboard drives — a pointer resting on a desk must not end a
/// keyboard preview — so a pointer can cross to the row below, and another file, and
/// never count as moved at all: the preview of the file it has left stays on screen
/// with nothing taking it down, since the probe that would have is closed behind it.
/// What a preview is about is not a distance but the box the view draws its item in,
/// which the hook publishes with every look at the item under the pointer (see
/// `preview_window::publish_pointer_item_box`); a pointer outside that box has left
/// the item, and the tick that sees it is treated as the move it is — so the preview
/// is taken down and the item now under the pointer is asked about like any other.
///
/// A pointer that the preview itself is holding is not one that has left: a text
/// frame being read, or the spinner a page is being waited behind, is the pointer
/// where it means to be, whatever box the item under it draws.
pub(super) fn pointer_moved_off_the_hovered_item(
    preview_active: bool,
    pointer_hold: bool,
    pointer_on_the_item: bool,
) -> bool {
    preview_active && !pointer_hold && !pointer_on_the_item
}

/// Whether a look that came back with nothing is a look at the item the preview on screen
/// is about.
///
/// The answer a look gives is a file or nothing, and nothing is two different things. One
/// is a read that failed — the walk out through the shell that answers nothing on a volume
/// slow to answer, a view busy drawing the item it was just asked for — where the pointer
/// is still on the file the preview is about, and what that file is owed is the question
/// asked again rather than the preview taken down and put back, which is a blink. The
/// other is an item that is not a file this app previews at all: an application, a folder,
/// a name no kind claims. That look answered properly and there is nothing to show for it,
/// and the preview of the file the pointer has left has to go with it.
///
/// What tells the two apart is the item rather than the answer. Every look publishes the
/// box of the item it found, so a look that found the same item leaves that box exactly
/// where it was, while a look that found another item publishes another box — and an item
/// with no preview of its own is an item all the same, which is the case that makes the
/// published box on its own the wrong question (see `HOVER_POINTER_BOX`). A preview is
/// therefore spared only where the box under the pointer is the box the preview was
/// resolved from *and* still holds the pointer; another item, another box, no box at all
/// are all read as the pointer having left, which is the reading that cannot leave a
/// preview on screen for good.
pub(super) fn read_failure_is_the_same_item(
    hover_box: Option<(i32, i32, i32, i32)>,
    published_box: Option<(i32, i32, i32, i32)>,
    point: POINT,
) -> bool {
    let Some(hover) = hover_box else {
        return false;
    };

    let holds = point.x >= hover.0 && point.x < hover.2 && point.y >= hover.1 && point.y < hover.3;

    published_box == Some(hover) && holds
}

/// Whether a preview may be shown for `path`: the kind of preview it would get,
/// and whether that kind is switched on in the tray's `Preview Types`
/// submenu.
///
/// The kinds are not worked out here: they are asked of the one table that answers for
/// every side of the app, so the hook, the loader and the layout cannot disagree about
/// what a file is (see `formats::routing`). The lists come from the configuration this
/// already holds rather than from the gates' own lookups, which would take the same lock
/// again — and nothing in the router reads the configuration itself.
///
/// What the file's content says it is comes ahead of all of that, where it disagrees
/// with the name: a `.docx` whose bytes are an MP4 is a video, and the engine it is
/// handed to is the one that plays videos (see `content_type`). Both questions are asked
/// with the one copy of the configuration this gate takes — the content question is handed
/// the lists rather than taking them itself, which is what makes a second acquisition on this
/// thread impossible (see `content_type::of`).
pub(super) fn is_media_file(path: &Path) -> bool {
    match crate::formats::head::Facts::read(path) {
        Some(facts) => is_media_file_with_facts(path, &facts),
        None => false,
    }
}

/// The same question asked of a file whose own entry has already been read: what the file's
/// content says it is, decided the way it is decided everywhere else (see
/// `crate::formats::head::Facts`).
pub(super) fn is_media_file_with_facts(path: &Path, facts: &crate::formats::head::Facts) -> bool {
    let Ok(config) = CONFIG.lock() else {
        return false;
    };

    let content = crate::formats::content_type::of_with_facts(path, facts, &config);

    match content {
        crate::formats::content_type::Content::Kind(kind) => return kind.enabled_in(&config),
        // A format no kind of this app previews is no preview at all: nothing is shown
        // and no engine of this app's is started for it.
        crate::formats::content_type::Content::Foreign => return false,
        // Nothing was recognized, or the content and the name agree: the name decides,
        // which is what the lists below are for.
        crate::formats::content_type::Content::Unknown => {}
    }

    // The kind is the router's answer and the router's order is the only one — a video first,
    // then a page, on through the boxes and the documents to the pictures last — so what this
    // gate admits is what the loader will draw (see `formats::routing`). Every list is asked
    // of the copy in hand, and the switch the kind is under is the gate that decides it.
    crate::formats::routing::kind_of(path, &config).is_some_and(|kind| kind.enabled_in(&config))
}

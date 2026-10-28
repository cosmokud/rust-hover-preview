use super::*;

/// What the chrome rule refuses, and the windows it must leave alone: the class name is the
/// only thing that tells a popup from the listing it covers, and the listing's own windows
/// are the ones a press has to keep answering for.
#[test]
fn a_popup_class_is_chrome_and_a_listing_class_is_not() {
    for chrome in [
        "Microsoft.UI.Content.PopupWindowSiteBridge",
        "Microsoft.UI.Content.DesktopChildSiteBridge",
        "#32768",
        "Xaml_WindowedPopupClass",
        "Windows.UI.Core.CoreWindowFlyout",
        "MenuHostWindow",
    ] {
        assert!(
            class_is_menu_popup_chrome(chrome),
            "{chrome} is a menu or a flyout"
        );
    }

    for listing in [
        // The view itself, and the DirectUI frame the toolbar and the rows are drawn in.
        "SysListView32",
        "DirectUIHWND",
        // The frame and the file list, under whatever Explorer is showing.
        "CabinetWClass",
        "ExplorerWClass",
        // This app's own, standing over the listing — a press on it *is* a press on the
        // listing, so a name of our own must not be mistaken for a popup's.
        "RustHoverPreviewWindow",
        // What a tick that read no window, or could not read one, is logged as.
        "none",
        "unreadable",
    ] {
        assert!(
            !class_is_menu_popup_chrome(listing),
            "{listing} is a window a listing can be read under"
        );
    }
}

/// A signature of one display, which is all the tests here need: what is compared
/// is a display's own place and scale, not what the desktop adds up to.
fn signed(displays: &[(u32, bool)]) -> DisplaySignature {
    DisplaySignature {
        displays: displays
            .iter()
            .map(|(dpi, primary)| DisplayEntry {
                rect: (0, 0, 1920, 1080),
                dpi: *dpi,
                primary: *primary,
            })
            .collect(),
    }
}

/// What the signature is for: a display rescaled, or another one made the primary,
/// rearranges everything Explorer is drawing without moving the desktop's own
/// bounds — so both are changes, and a desktop that did not change is not.
#[test]
fn a_display_change_is_a_rescale_or_another_primary_display() {
    let before = signed(&[(96, true), (96, false)]);

    assert!(
        !display_signature_changed(Some(&before), &signed(&[(96, true), (96, false)])),
        "the same desktop is not a change"
    );
    assert!(
        display_signature_changed(Some(&before), &signed(&[(144, true), (96, false)])),
        "the primary display rescaled to 150% is one"
    );
    assert!(
        display_signature_changed(Some(&before), &signed(&[(96, true), (144, false)])),
        "and so is the second display rescaled, which the desktop's bounds do not show"
    );
    assert!(
        display_signature_changed(Some(&before), &signed(&[(96, false), (96, true)])),
        "and another display becoming the primary one"
    );

    assert!(
        !display_signature_changed(None, &before),
        "the first look at a desktop is not a change"
    );
}

/// One state's counts, from the numbers `explorer_state_from_counts` decides by.
fn counts(
    total: usize,
    visible: usize,
    reachable: usize,
    cover: Option<RECT>,
) -> ExplorerWindowCounts {
    ExplorerWindowCounts {
        total,
        visible,
        reachable,
        cover,
    }
}

/// The region a maximized window on the primary display covers, which is what a
/// hidden answer is measured against.
fn primary_display() -> Option<RECT> {
    Some(RECT {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    })
}

/// The case this was fixed for: a maximized window on one display leaves an
/// Explorer window on the display beside it reachable, so the state has to be the
/// one that keeps asking where the cursor is. Answered as hidden, the loop hides
/// the preview and sleeps without ever asking, and a pointer that crossed over to
/// the second display shows nothing at all until the click that focuses Explorer.
#[test]
fn a_window_beside_the_one_in_front_leaves_explorer_reachable() {
    assert_eq!(
        explorer_state_from_counts(&counts(1, 1, 1, primary_display()), false),
        ExplorerState::VisibleNotFocused
    );
}

/// A pin standing over a listing is that listing being worked in, so the loop is paced as it
/// is rather than as a window nobody is in. What it is *not* is an answer for a minimized
/// Explorer or one behind a maximized window: those keep the deep and long rows.
#[test]
fn a_pin_over_a_reachable_listening_is_the_active_arrangement() {
    assert_eq!(
        explorer_state_from_counts(&counts(1, 1, 1, None), true),
        ExplorerState::ActiveFocus,
        "a pin in front of a listing is that listing in use"
    );

    // Which is what the ladder reads as the fast row, and only on the fast row: the two
    // states a pin cannot make anything of are asked before the pin is looked at, so they
    // are the same with and without one.
    let active = explorer_pace(ExplorerState::ActiveFocus, 15);
    assert_eq!(
        explorer_pace(explorer_state_from_counts(&counts(1, 1, 1, None), true), 15),
        active,
        "a pinned tick is asked on the tick's own clock"
    );
    assert_eq!(active.0, 15, "which is the tick setting, not a constant");

    for counts in [
        counts(2, 0, 0, primary_display()),
        counts(2, 2, 0, primary_display()),
    ] {
        assert_eq!(
            explorer_state_from_counts(&counts, true),
            explorer_state_from_counts(&counts, false),
            "minimized, or behind a window that covers it, a pin changes nothing"
        );
    }
}

/// Every window Explorer has showing behind the window in front is the one case
/// that may sleep without asking: there is nothing left for the pointer to reach.
#[test]
fn every_window_behind_the_one_in_front_is_hidden() {
    assert_eq!(
        explorer_state_from_counts(&counts(2, 2, 0, primary_display()), false),
        ExplorerState::HiddenByForeground
    );
}

/// A foreground window that hides nothing — an ordinary one — leaves every
/// Explorer window reachable, whatever display it is on.
#[test]
fn a_window_that_hides_nothing_leaves_explorer_visible() {
    assert_eq!(
        explorer_state_from_counts(&counts(1, 1, 1, None), false),
        ExplorerState::VisibleNotFocused
    );
}

/// No Explorer window at all is answered before anything the window in front
/// does: there is nothing to be hidden and nothing to ask about.
#[test]
fn no_window_is_answered_before_the_window_in_front() {
    assert_eq!(
        explorer_state_from_counts(&counts(0, 0, 0, primary_display()), false),
        ExplorerState::NoExplorerWindows
    );
}

/// Every Explorer window minimized is minimized whatever the window in front is.
#[test]
fn every_window_minimized_is_answered_as_minimized() {
    assert_eq!(
        explorer_state_from_counts(&counts(2, 0, 0, primary_display()), false),
        ExplorerState::AllMinimized
    );
}

/// A region holds a window that is inside it — including one that covers it whole —
/// and does not hold one on the display beside it, or one that reaches past its
/// edge: what a window in front hides is what its own rectangle holds.
#[test]
fn a_region_holds_only_what_is_inside_it() {
    let display = RECT {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let inside = RECT {
        left: 120,
        top: 90,
        right: 900,
        bottom: 700,
    };
    let covering = RECT {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let beside = RECT {
        left: 1920,
        top: 0,
        right: 3840,
        bottom: 1080,
    };
    let reaching = RECT {
        left: 1600,
        top: 0,
        right: 2400,
        bottom: 1080,
    };

    assert!(rect_contains(display, inside));
    assert!(
        rect_contains(display, covering),
        "a window covering the region is on it"
    );
    assert!(
        !rect_contains(display, beside),
        "the display beside it is outside it"
    );
    assert!(
        !rect_contains(display, reaching),
        "a window reaching past its edge is outside it"
    );
}

/// The counts a run is measured by are written where it asked for them, once a
/// second rather than once a probe, and the line carries the numbers themselves
/// rather than only being evidence that something ran — see
/// `flush_probe_counts`.
///
/// One test rather than two: the counters are process-wide, and two tests adding
/// to them at the same time would be asserting about each other's numbers. What
/// is asserted is a floor for the same reason — the probe test beside this one
/// walks a real shell and adds to some of the same counters, so a run holding
/// both would otherwise be reading the other test's totals.
#[test]
fn the_probe_counts_are_written_once_a_second_where_the_trace_asks() {
    let path = std::env::temp_dir().join(format!("rhp-hook-trace-{}.log", std::process::id()));
    let _ = std::fs::remove_file(&path);

    PROBE_VIEW_WALKS.fetch_add(2, Ordering::Relaxed);
    PROBE_VIEW_WINDOWS.fetch_add(7, Ordering::Relaxed);
    PROBE_VIEW_SETS_KEPT.fetch_add(4, Ordering::Relaxed);
    PROBE_VIEW_ANCHORED.fetch_add(3, Ordering::Relaxed);
    PROBE_ITEM_MEMO_HITS.fetch_add(6, Ordering::Relaxed);
    PROBE_POINTER_RESOLUTIONS.fetch_add(1, Ordering::Relaxed);
    note_probe_ms(&PROBE_VIEW_SLOWEST_MS, Duration::from_millis(9));
    note_probe_ms(&PROBE_ITEM_SLOWEST_MS, Duration::from_millis(11));

    let now = Instant::now();
    let mut last = now - Duration::from_secs(5);
    flush_probe_counts(now, &mut last, &path);

    let first = std::fs::read_to_string(&path).expect("the counts are written");

    /// The number written after one label of the line, which is the whole of what
    /// this is reading: the fields are named rather than positional so that one
    /// being added does not silently shift another one's number into its place.
    fn written(line: &str, label: &str) -> u64 {
        line.split(label)
            .nth(1)
            .and_then(|rest| rest.split_whitespace().next())
            .map(|value| value.trim_end_matches([',', ')']))
            .and_then(|value| value.trim_end_matches("ms").parse().ok())
            .unwrap_or_else(|| panic!("the line carries {label:?}: {line}"))
    }

    assert!(written(&first, "points ") >= 1, "points: {first}");
    assert!(written(&first, "view walks ") >= 2, "view walks: {first}");
    assert!(written(&first, "windows ") >= 7, "windows: {first}");
    assert!(written(&first, "sets kept ") >= 4, "sets kept: {first}");
    assert!(written(&first, "anchored ") >= 3, "anchored: {first}");
    assert!(written(&first, "item memo ") >= 6, "item memo: {first}");
    // Two slowest figures, and the first of them belongs to the item walks: read
    // by name so that one field being added cannot shift another's number into
    // its place.
    assert!(
        first.contains("(slowest 11ms)"),
        "the item walks are timed apart from the view walks: {first}"
    );
    assert!(
        first.contains("slowest 9ms)"),
        "the view walks carry their own slowest: {first}"
    );

    // A second flush inside the same second writes nothing more: what a run being
    // measured costs is a line a second, not a line a tick.
    PROBE_VIEW_WALKS.fetch_add(1, Ordering::Relaxed);
    flush_probe_counts(now, &mut last, &path);

    let second = std::fs::read_to_string(&path).expect("the counts are still written");
    let _ = std::fs::remove_file(&path);

    assert_eq!(second.lines().count(), 1, "one line a second: {second}");
}

/// The probe against the shell the machine is actually running — the part of this
/// no unit test reaches: a real window collection, the views a real window holds, and
/// what the view the pointer is in is then asked.
///
/// Ignored because it needs a desktop, which is the same reason the other probes
/// here are. Run by hand: `cargo test -- --ignored the_probe_reads_the_place`.
///
/// What it asserts is about the shape of the probe rather than about the answer,
/// because the answer is whatever this machine happens to have under its pointer:
/// that a place is read out of the views the pointed-at window holds, and that a
/// second probe of the same place reads no view set again — the property the whole
/// arrangement is for, and the one an argument cannot settle.
#[test]
#[ignore = "reads the shell of the desktop it runs on"]
fn the_probe_reads_the_place_out_of_the_window_the_pointer_is_in() {
    // Every call on this path is a Shell object's, so the thread needs an
    // apartment before any of it is asked for — the same one the hook thread
    // takes, which never pumps either.
    if unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }.is_err() {
        println!("no apartment: nothing could be asked of the shell");
        return;
    }

    let mut resolver = ItemResolver::new(None);
    resolver.rebuild_shell();

    let Some(pointer) = read_pointer() else {
        println!("no pointer to ask about");
        return;
    };
    let Some(window) = item_window_of(pointer.window) else {
        println!("no item window to ask about");
        return;
    };

    let kept = PROBE_VIEW_SETS_KEPT.load(Ordering::Relaxed);
    let context = anchored_view_context(&mut resolver, window, true);
    let walks = PROBE_VIEW_WALKS.load(Ordering::Relaxed);

    // An answer names a view, and a view is a window: an answer with no window
    // behind it would be one that came from somewhere no view was ever found.
    if let Some(context) = &context {
        assert_ne!(context.shell_view_hwnd, 0, "an answer names a window");
    }

    // The second probe of an unmoved pointer is answered out of the set the first
    // one read: a window's views are its own until a tab or a window is opened or
    // closed, and reading them again for every probe is the cost that grows with
    // how many tabs are open. What is not kept is a walk that found no view at all —
    // the shell met between two of them — and there the second probe reads again
    // rather than being answered from a set that has nothing in it (see `frame_views`).
    let held_views = resolver
        .window_views
        .as_ref()
        .map(|set| set.views.len())
        .unwrap_or(0);

    let _ = anchored_view_context(&mut resolver, window, true);

    if held_views > 0 {
        assert_eq!(
            PROBE_VIEW_WALKS.load(Ordering::Relaxed),
            walks,
            "a probe of the same place reads no view set again"
        );
        assert!(
            PROBE_VIEW_SETS_KEPT.load(Ordering::Relaxed) > kept,
            "and the set the first probe read is what answered it"
        );
    } else {
        assert!(
            PROBE_VIEW_WALKS.load(Ordering::Relaxed) > walks,
            "a walk that found no view is not an answer to keep"
        );
    }

    // Said out loud rather than only asserted: a run whose pointer is in none of the
    // Shell windows passes the assertions above without describing anything, and the
    // numbers are how that is told apart from a run that answered.
    match &context {
        Some(context) => println!(
            "the pointer is in view {} with folder {:?} (view walks {walks} before it)",
            context.shell_view_hwnd, context.folder_path
        ),
        None => println!("the pointer is in no view a Shell window holds"),
    }

    unsafe { CoUninitialize() };
}

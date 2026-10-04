use super::*;

// WS-H Phase 1: kill superseded relaunches before publish; show only the
// current generation.
//
// A release-then-instant-regrab strands the first gesture's relaunch: it
// lands and is shown at a stale box although a newer gesture superseded it.
// Generation tags every relaunch; the settle shows/places a replacement IFF
// its generation is current; a bump (a newer gesture's park) kills the
// in-flight relaunch before it can publish, and reaps it.
//
// RED (pre-fix): the helpers below do not exist, so this file does not
// compile — the same red the WS-G loop started from. A behavioural red
// through the old fns is impossible: nothing tags a generation yet, so two
// generations are unrepresentable without the fix.

/// A parked pin with a cover, as a gesture's first message leaves it.
fn parked_pin(content: ScreenRegion) -> RecordedPinWindow {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("ws-h-supersede.mkv");
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_pending_pinned_relaunch();
    clear_park_swap_arm();
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (content.0, content.1), true),
        "the first gesture parks: the cover stands before anything relaunches"
    );
    RecordedPinWindow::new(0x1000)
}

/// The state effects of a relaunch, without starting a real player: the pid
/// the record names, the generation tag, and the transport write.
fn fake_relaunch(pid: u32, from: f64, holding: bool) {
    VIDEO_PID.store(pid, Ordering::SeqCst);
    note_pinned_relaunch(pid);
    update_pin_transport(|transport| transport.begun(from, true, holding));
}

/// The swap bound spent, so the settle answers on the window and not on the wait.
fn spend_the_swap_bound() {
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT * 4);
        }
    }
}

fn restore(previous_pin: Option<PinnedPreview>, previous_pid: u32) {
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_pending_pinned_relaunch();
    clear_park_swap_arm();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// A superseded replacement dies unpublished: killed on the newer gesture's
/// bump, reaped, and never shown nor placed — the cover holds.
#[test]
fn superseded_relaunch_dies_unpublished_and_reaped() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let window = parked_pin(content);

    // The first gesture's release: one relaunch behind the cover.
    fake_relaunch(101, 10.0, false);
    assert_eq!(
        pending_pinned_relaunch().map(|pending| pending.pid),
        Some(101),
        "the relaunch is in flight behind the standing cover"
    );

    // The instant regrab: a newer gesture bumps the generation, killing the
    // in-flight relaunch before it can publish.
    park_pinned_player(HWND(0x1000 as *mut _), (100, 80), true);
    bump_pinned_generation();

    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "the superseded player is reaped: no pid left for any show path to find"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "nothing in flight after the kill"
    );

    // Even a window standing in the band swaps nothing: there is no current
    // player to hand the band to.
    assert!(
        !settle_pinned_park_where(&window, true),
        "a superseded relaunch is never shown, even with a window up"
    );
    assert!(
        pin_player_is_parked(),
        "so the cover holds instead of handing the band to a dead player"
    );
    assert_eq!(
        window.calls(),
        Vec::<PinWindowCall>::new(),
        "shown nowhere, placed nowhere: no call on either window"
    );

    restore(previous_pin, previous_pid);
}

/// The settle shows only the current generation, placed at the final box
/// first: the band the pin stands in now, not the box the killed relaunch
/// was begun at.
#[test]
fn settle_shows_only_the_current_generation_at_the_final_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let launched_at = (100, 80, 420, 320);
    let window = parked_pin(launched_at);

    // First gesture's relaunch, superseded by the regrab's bump.
    fake_relaunch(101, 10.0, false);
    park_pinned_player(HWND(0x1000 as *mut _), (100, 80), true);
    bump_pinned_generation();

    // Second gesture's relaunch at the box the hand moved to.
    let final_box = (200, 150, 640, 480);
    if let Some(mut pinned) = pin_state() {
        if let Some(pin) = pinned.pin_mut() {
            pin.content = final_box;
        }
    }
    fake_relaunch(102, 12.0, false);
    spend_the_swap_bound();

    assert!(
        settle_pinned_park_where(&window, true),
        "the current replacement swaps once its window stands in the band"
    );
    assert!(!pin_player_is_parked(), "and the cover comes down with it");
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(final_box)),
            PinWindowCall::Repaint,
        ],
        "placed at the final box before unflagged, repainted in the same \
         tick: never the stale box the killed relaunch was begun at"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        102,
        "exactly one current player"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "the swap consumes the in-flight record"
    );

    restore(previous_pin, previous_pid);
}

/// Instant-regrab down the real road: begin, release (relaunch), regrab
/// (extend + bump kills the in-flight), release (relaunch), one swap — one
/// player, one show, zero lingering windows.
#[test]
fn instant_regrab_ends_with_one_player_one_swap() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let window = parked_pin(content);

    // First release's relaunch lands behind the cover.
    fake_relaunch(101, 10.0, false);

    // Regrab: the extend keeps the cover, the bump kills the in-flight.
    park_pinned_player(HWND(0x1000 as *mut _), (120, 90), true);
    bump_pinned_generation();
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "the first relaunch dies on supersede, before it can publish"
    );

    // Second release's relaunch is the current generation.
    fake_relaunch(102, 14.0, false);
    spend_the_swap_bound();

    assert!(
        settle_pinned_park_where(&window, true),
        "the current replacement swaps"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint,
        ],
        "a single swap: one place-before-show, one repaint"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        102,
        "one player: the killed relaunch leaves no pid behind"
    );

    restore(previous_pin, previous_pid);
}

/// A single gesture is untouched: no bump, no kill — its relaunch swaps at
/// its box with its pid standing.
#[test]
fn single_gesture_relaunch_is_untouched() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let window = parked_pin(content);

    fake_relaunch(101, 10.0, false);
    spend_the_swap_bound();

    assert!(
        settle_pinned_park_where(&window, true),
        "the current relaunch swaps"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "no kill without a supersede: the current relaunch is untouched"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint,
        ],
        "shown at its box, placed before shown"
    );

    restore(previous_pin, previous_pid);
}

/// The kill path reaps: a fake pid is confirmed gone, forgotten, and the
/// record cleared; pid zero is already nothing.
#[test]
fn kill_path_reaps_and_clears() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);

    assert!(
        kill_superseded_player(101),
        "a superseded player is confirmed gone"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "its pid is cleared, so no show path can find it and no audio survives it"
    );
    assert!(
        kill_superseded_player(0),
        "killing nothing is already reaped"
    );

    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// A kill strands the cover over no player: the release's own recovery
/// predicate says so, while a live player and a swapped park say otherwise.
#[test]
fn stranded_cover_predicate() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = parked_pin(content);

    // A live player behind the cover: nothing stranded.
    VIDEO_PID.store(101, Ordering::SeqCst);
    assert!(
        !park_stranded_without_a_player(),
        "a parked player is not stranded"
    );

    // The kill: no pending, no pid, cover standing.
    VIDEO_PID.store(0, Ordering::SeqCst);
    assert!(
        park_stranded_without_a_player(),
        "a cover over a killed in-flight needs its release to relaunch"
    );

    restore(previous_pin, previous_pid);
}

/// A kill not yet taken holds the cover and is retried: the generation moved
/// on while the old player is still dying, so the settle refuses the swap —
/// and the retry reaps it.
#[test]
fn stale_pending_holds_the_cover_and_retries_the_kill() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let window = parked_pin(content);

    fake_relaunch(101, 10.0, false);
    // The bump's terminate is still in flight: the generation moved on
    // without reaping, which is what a slow death looks like.
    PIN_PLAYER_GENERATION.fetch_add(1, Ordering::AcqRel);
    assert!(
        !pending_relaunch_is_current(),
        "the in-flight relaunch is superseded but not yet reaped"
    );

    // No window can be standing for a player being killed: the cover holds
    // and the settle asks the kill again instead.
    assert!(
        !settle_pinned_park_where(&window, false),
        "a superseded generation swaps nothing while its kill is unconfirmed"
    );
    assert!(pin_player_is_parked(), "so the cover holds");
    assert_eq!(
        window.calls(),
        Vec::<PinWindowCall>::new(),
        "shown nowhere, placed nowhere"
    );
    assert_eq!(pending_pinned_relaunch(), None, "the retry reaped the kill");
    assert_eq!(VIDEO_PID.load(Ordering::SeqCst), 0, "and its pid with it");

    restore(previous_pin, previous_pid);
}

/// The generation only moves on a bump, and a relaunch tags the current one.
#[test]
fn generations_tag_relaunches_and_move_only_on_bump() {
    let before = current_pinned_generation();
    let bumped = bump_pinned_generation();
    assert_eq!(bumped, before + 1, "one bump is one generation");
    assert_eq!(
        current_pinned_generation(),
        bumped,
        "the current generation is the last bump"
    );

    note_pinned_relaunch(103);
    assert_eq!(
        pending_pinned_relaunch().map(|pending| (pending.pid, pending.generation)),
        Some((103, bumped)),
        "a relaunch tags the current generation"
    );
    assert!(pending_relaunch_is_current(), "so it is showable");

    note_pinned_relaunch(0);
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "a relaunch that never came up leaves nothing in flight"
    );
    clear_pending_pinned_relaunch();
}

// WS-H Phase 2 (review blockers 1-6): the races the first pass left.
//
// RED (pre-fix): the seams below do not exist, so this section does not
// compile — the same red Phase 1 started from. Behavioural red through the
// old fns is impossible where noted: a single test thread cannot interleave
// a window-thread bump between two adjacent lines, and a fake pid is never
// a live ffplay, so an unconfirmed kill is unrepresentable without the seam.

/// Blocker 1, confirmed path: a teardown with an in-flight relaunch reaps it —
/// killed, pid cleared, record forgotten — rather than clearing the record
/// from under a kill.
#[test]
fn teardown_reaps_a_confirmed_in_flight_relaunch() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = parked_pin(content);
    fake_relaunch(101, 10.0, false);
    assert_eq!(
        pending_pinned_relaunch().map(|pending| pending.pid),
        Some(101),
        "the relaunch is in flight behind the standing cover"
    );

    end_pin_beside_the_state();

    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "the confirmed kill takes the record with it"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and its pid: no orphan ffplay, no zombie audio, nothing to publish"
    );

    restore(previous_pin, previous_pid);
}

/// Blocker 1, unconfirmed path: a kill not yet taken is handed to the orphan
/// reaper that outlives the pin, and the loop's settle retries it to
/// confirmation. A fake pid is never alive, so the stranded half is stood by
/// hand: what the test proves is the reaper half — nothing parked there is
/// ever dropped.
#[test]
fn orphan_reaper_retries_to_confirmation() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);

    // A confirmed kill never reaches the list: reaped on the spot.
    retire_orphaned_player(101);
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "a kill confirmed on the spot clears the pid with it"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "so nothing stranded is parked"
    );

    // A kill still dying is owned by the list, and the loop's settle — which
    // runs with no pin up — drains it once the death confirms.
    orphaned_video_kills_push_for_test(101);
    settle_video_retirement();
    assert!(
        orphaned_video_kills().is_empty(),
        "the reaper drops only confirmed-gone kills, and only by reaping them"
    );

    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// Blocker 5: the supersede-kill is one helper so the relaunch calls it
/// *before* spawning — a stale player must be dying before its replacement
/// exists, never published in the start-to-kill window. The helper clears the
/// record the spawn would otherwise stack behind.
#[test]
fn superseded_kill_clears_the_record_for_the_next_spawn() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = parked_pin(content);
    fake_relaunch(101, 10.0, false);

    assert_eq!(
        reap_superseded_relaunch(),
        Some(101),
        "the unpublished relaunch dies before any replacement is spawned"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "its pid goes with it, so the spawn starts from nothing published"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "and the record is forgotten rather than left for the spawn to stack on"
    );

    restore(previous_pin, previous_pid);
}

/// Blockers 2 and 6: the restart's gate-to-show is one helper — stale refuses
/// the show and retries the kill, current shows and revalidates. A stale
/// generation never places anything.
#[test]
fn gated_show_refuses_a_superseded_generation() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = parked_pin(content);
    fake_relaunch(101, 10.0, false);
    // The bump's terminate is still in flight: the generation moved on without
    // reaping, which is what a slow death looks like.
    PIN_PLAYER_GENERATION.fetch_add(1, Ordering::AcqRel);
    assert!(
        !pending_relaunch_is_current(),
        "the in-flight relaunch is superseded but not yet reaped"
    );

    show_current_replacement(101, content);

    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "the refused show retries the kill rather than dropping it"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and reaps its pid with it"
    );

    restore(previous_pin, previous_pid);
}

/// Blocker 2, second half: a bump landing between the gate and the show cannot
/// abort the single `SetWindowPos` that is both — so the show is revalidated
/// after it, and a condemned window is hidden again rather than left
/// published. Stood by hand: the stale generation with the pid still naming
/// the shown player is exactly what that interleave leaves behind.
#[test]
fn revalidation_hides_a_player_condemned_mid_show() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = parked_pin(content);
    fake_relaunch(101, 10.0, false);
    PIN_PLAYER_GENERATION.fetch_add(1, Ordering::AcqRel);

    revalidate_shown_replacement(101);

    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "the condemned show retries the kill rather than leaving it published"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "its pid goes with it: hidden (a no-op with no window up) and reaped"
    );

    restore(previous_pin, previous_pid);
}

/// Blocker 3: the settle's swap re-checks currency at the show site — a bump
/// landing between the gate and the unpark refuses the swap rather than
/// placing and showing a condemned player.
#[test]
fn unpark_refuses_a_generation_superseded_after_the_gate() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let window = parked_pin(content);
    fake_relaunch(101, 10.0, false);
    // A window-thread bump between the settle's gate and its unpark: the
    // generation moved on, the kill not yet confirmed.
    PIN_PLAYER_GENERATION.fetch_add(1, Ordering::AcqRel);
    assert!(
        !pending_relaunch_is_current(),
        "the in-flight relaunch is superseded but not yet reaped"
    );

    assert!(
        !unpark_pinned_player(&window),
        "a condemned player is never placed nor shown"
    );
    assert!(
        pin_player_is_parked(),
        "so the cover holds instead of handing the band to it"
    );
    assert_eq!(
        window.calls(),
        Vec::<PinWindowCall>::new(),
        "shown nowhere, placed nowhere"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "the refused swap retries the kill rather than dropping it"
    );

    restore(previous_pin, previous_pid);
}

/// Blocker 4: the place-before-show exception is closed — a pin with no media
/// band (collapsed) is still placed, at the box it stands at now, rather than
/// shown at whatever rect its window happens to be at.
#[test]
fn unpark_places_a_pin_with_no_media_band_at_its_current_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let window = parked_pin(content);
    if let Some(mut pinned) = pin_state() {
        if let Some(pin) = pinned.pin_mut() {
            pin.collapsed = true;
        }
    }
    assert!(
        pinned_content().is_none(),
        "a collapsed pin has no media band to place into"
    );

    assert!(
        unpark_pinned_player(&window),
        "the park still ends: placement was the missing half, not the show"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint,
        ],
        "placed at the box the pin stands at now, before shown and repainted"
    );

    restore(previous_pin, previous_pid);
}

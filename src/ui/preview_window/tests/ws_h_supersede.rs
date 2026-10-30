use super::*;

// WS-H: kill a superseded relaunch before it can publish, and show only the current generation.
//
// A release-then-instant-regrab strands the first gesture's relaunch: it lands and is shown at a
// stale box although a newer gesture superseded it. Every relaunch is tagged with a generation, a
// newer one bumps it, the in-flight player is killed before it can publish, and the show places a
// replacement only while its generation is current. **What is superseded is only ever the player
// no picture is drawn from** — the player standing on screen is left to the retirement that ends it
// once its replacement has taken the picture over (see `relaunch_pinned_player`).

/// A pin on a film, standing at `content`: the window every one of these roads acts on.
fn a_film_pin(content: ScreenRegion) -> RecordedPinWindow {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("ws-h-supersede.mkv");
    stand_pin(Some(pin));
    clear_pending_pinned_relaunch();
    RecordedPinWindow::new(0x1000)
}

/// The state effects of a relaunch, without starting a real player: the pid the record names, the
/// generation tag, and the transport write.
fn fake_relaunch(pid: u32, from: f64, holding: bool) {
    VIDEO_PID.store(pid, Ordering::SeqCst);
    note_pinned_relaunch(pid);
    update_pin_transport(|transport| transport.begun(from, true, holding));
}

fn restore(previous_pin: Option<PinnedPreview>, previous_pid: u32) {
    clear_pending_pinned_relaunch();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// The kill path reaps: a fake pid is confirmed gone, forgotten, and the record cleared; pid zero
/// is already nothing.
#[test]
fn kill_path_reaps_and_clears() {
    let _room = the_room_to_itself();
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

/// **A relaunch superseded before it stood is dropped, and the drop is a line.** Nothing of that
/// player ever reached the screen, so nothing of the picture goes with it: the player the host is
/// drawing is not the one the record names, so no kill is drawn onto it (see
/// `reap_superseded_relaunch`).
#[test]
fn a_superseded_relaunch_is_dropped_and_named_in_one_line() {
    let _room = the_room_to_itself();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = a_film_pin(content);
    pin_trace::clear_noted();

    // A replacement begun and never handed the picture: its source is not the one standing in any
    // host, so it is the one a newer relaunch kills.
    fake_relaunch(101, 10.0, false);
    assert_eq!(
        pending_pinned_relaunch().map(|pending| pending.pid),
        Some(101),
        "the replacement is in flight"
    );

    assert_eq!(
        reap_superseded_relaunch(),
        Some(101),
        "the unpublished replacement dies before it can publish"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "its pid goes with it, so no show path can find it"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "and the record is forgotten rather than left to stack the next spawn behind"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line == "relaunch superseded before it stood pid 101"),
        "and the run says which player was dropped, in one line: {:?}",
        pin_trace::noted_lines()
    );

    restore(previous_pin, previous_pid);
}

/// The generation only moves on a relaunch, and a relaunch tags the current one.
#[test]
fn generations_tag_relaunches_and_move_only_on_a_relaunch() {
    let before = current_pinned_generation();

    note_pinned_relaunch(103);
    assert_eq!(
        pending_pinned_relaunch().map(|pending| (pending.pid, pending.generation)),
        Some((103, before)),
        "a relaunch tags the generation that is current"
    );
    assert!(pending_relaunch_is_current(), "so it is showable");

    // The next relaunch's own bump: a new generation, and the older record is not current.
    PIN_PLAYER_GENERATION.fetch_add(1, Ordering::AcqRel);
    assert_ne!(
        current_pinned_generation(),
        before,
        "a relaunch is a new generation"
    );
    assert!(
        !pending_relaunch_is_current(),
        "and what the older one began is superseded by it"
    );

    note_pinned_relaunch(0);
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "a relaunch that never came up leaves nothing in flight"
    );
    clear_pending_pinned_relaunch();
}

/// A teardown with an in-flight relaunch reaps it — killed, pid cleared, record forgotten —
/// rather than clearing the record from under a kill.
#[test]
fn teardown_reaps_a_confirmed_in_flight_relaunch() {
    let _room = the_room_to_itself();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = a_film_pin(content);
    fake_relaunch(101, 10.0, false);
    assert_eq!(
        pending_pinned_relaunch().map(|pending| pending.pid),
        Some(101),
        "the relaunch is in flight"
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

/// A kill not yet taken is handed to the orphan reaper that outlives the pin, and the loop's
/// settle retries it to confirmation. A fake pid is never alive, so the stranded half is stood by
/// hand: what the test proves is the reaper half — nothing handed to it is ever dropped.
#[test]
fn orphan_reaper_retries_to_confirmation() {
    let _room = the_room_to_itself();
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
        "so nothing stranded is handed on"
    );

    // A kill still dying is owned by the list, and the loop's settle — which runs with no pin up —
    // drains it once the death confirms.
    orphaned_video_kills_push_for_test(101);
    settle_video_retirement();
    assert!(
        orphaned_video_kills().is_empty(),
        "the reaper drops only confirmed-gone kills, and only by reaping them"
    );

    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// Blocker 5: the supersede-kill is one helper so a relaunch calls it *before* spawning — a stale
/// player must be dying before its replacement exists, never published in the start-to-kill window.
/// The helper clears the record the spawn would otherwise stack behind.
#[test]
fn superseded_kill_clears_the_record_for_the_next_spawn() {
    let _room = the_room_to_itself();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = a_film_pin(content);
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

/// Blockers 2 and 6: the show is gated on the generation — a stale one refuses the show and retries
/// the kill rather than placing anything.
#[test]
fn gated_show_refuses_a_superseded_generation() {
    let _room = the_room_to_itself();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = a_film_pin(content);
    fake_relaunch(101, 10.0, false);
    // The next relaunch's bump, with its terminate still in flight: the generation moved on
    // without reaping, which is what a slow death looks like.
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

/// Blocker 2, second half: a bump landing between the gate and the show cannot abort the single
/// placement that is both — so the show is revalidated after it, and a condemned window is hidden
/// again rather than left published. Stood by hand: the stale generation with the pid still naming
/// the shown player is exactly what that interleave leaves behind.
#[test]
fn revalidation_hides_a_player_condemned_mid_show() {
    let _room = the_room_to_itself();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let _window = a_film_pin(content);
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

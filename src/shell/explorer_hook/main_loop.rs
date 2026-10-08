//! The loop itself: one reading of the pointer, one look at the shell, and every
//! decision taken between them — what is hovered, what a press was, whether the
//! preview in front is still the file under the pointer, and how long to wait before
//! asking any of it again.
//!
//! It is one function and stays one file. `run_explorer_hook` is the whole of it: the
//! state a tick carries is read and written in the branches of that one loop, so
//! lifting the arms into helpers would be a change to the code rather than a move of
//! it, and the alternative — the same loop with the arms pulled out and handed back a
//! struct — is the rewrite this split is not. Everything that can be lifted out on its
//! own has been: what a tick knows, what place it is over, what the pin is watching
//! and what the keyboard is on are each in a file of their own, and this is what is
//! left.

use super::*;

/// Main loop for explorer hook
pub fn run_explorer_hook() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    // What both paths resolve an item to a file with: one UI Automation client
    // whose property reads are batched into a round trip per element, plus the
    // view that answered last for a window.
    let mut resolver = ItemResolver::new(automation_client());

    let mut last_file: Option<PathBuf> = None;
    // The box the item under the pointer was drawn in where the preview on screen was
    // resolved from it: the item that preview is about, as the view draws it. What it is
    // for is telling a read that failed from a pointer that has moved on — a look at the
    // same item leaves this box where it is, a look at another item publishes another one,
    // and a look that answered nothing is read against it (see
    // `read_failure_is_the_same_item`). A hover that never noted a box holds nothing, and
    // holding nothing is a comparison that fails rather than one that spares.
    let mut hover_item_box: Option<(i32, i32, i32, i32)> = None;
    let mut suppressed = SuppressedHover::default();
    let mut pointer_pause = KeyboardPointerPause::default();
    let mut hover_start: Option<Instant> = None;
    let mut last_cursor_pos = POINT::default();

    // Keyboard hover state
    let mut keyboard_file: Option<PathBuf> = None;
    let mut last_focused_key: Option<FocusedItemKey> = None;
    let mut is_keyboard_hover = false;
    let mut suppress_preview_until_cursor_leaves_preview = false;
    let mut stationary_search_miss_started_at: Option<Instant> = None;
    // Short grace after starting a video preview to avoid instant self-dismiss
    // while ffplay window is still initializing under the cursor.
    let mut video_hover_guard_until: Option<Instant> = None;
    // Folder/input gate state: suppress preview after folder changes until explicit user input.
    // The place the last folder probe found the pointer over, as the facts that probe
    // read — see `HoverLocation`.
    let mut last_cursor_location: Option<HoverLocation> = None;
    let mut hover_resolver_hints = HoverResolverHints::default();
    let mut suspend_preview_until_user_input = false;
    let mut allow_keyboard_preview_on_first_observation = false;
    let mut folder_change_time: Option<Instant> = None;
    let mut suspended_initial_focus: Option<FocusedItemKey> = None;
    // A folder change that follows recent input is user navigation: its gate
    // lifts on its own once the view has settled, so the item under a parked
    // cursor previews without a mouse move.
    let mut folder_change_user_initiated = false;
    // Navigation-key press transitions seen so far, and the value when the
    // current suspension started. Only a later press may lift that
    // suspension, which keeps key state left over from the navigation that
    // opened the folder from counting as new input.
    let mut keyboard_navigation_press_seq: u64 = 0;
    let mut keyboard_press_seq_at_suspend: u64 = 0;
    // Whether the keyboard is the one driving Explorer: set on a navigation key
    // press and kept across the keyboard previews that follow, cleared only by
    // deliberate pointer input — a move past the pointer tolerance or a wheel
    // tick — or by a reset that ends the keyboard's turn outright (previews
    // switched off, a display change, Explorer leaving the foreground, a folder
    // change handing the screen back to the pointer). While it holds — and where the
    // tray's `Prioritize Keyboard` is on — the parked pointer may neither raise a
    // preview nor take one over: a focused item with no preview to give must not hand
    // the pointer the screen, so a key pressed onto a file nothing can show reads as
    // one pressed onto a file that can.
    // What it also holds is the pointer tolerance: for the whole of the keyboard's
    // turn a move has to clear the wider distance before it counts as the mouse
    // taking over, because a keyboard preview is placed beside the focused item
    // and can land under the parked pointer, so jitter must not read as the mouse
    // asking for the screen.
    let mut keyboard_screen_owner = false;
    let mut last_folder_probe = Instant::now();
    let mut last_hover_probe = Instant::now();
    let mut last_keyboard_focus_probe = Instant::now();
    // What this hook watches while a preview is pinned, and nothing at all until one is: the file
    // the pin is showing, where the pointer was the last time it was read, and the item the
    // keyboard is on (see `PinUpdateWatch`).
    let mut pin_watch = PinUpdateWatch::default();
    // Whether Explorer held the foreground the last time a pinned tick asked: the pin takes the
    // focus away from it and gives it back, and what the resolver has cached across that is what
    // a listing that has moved on under the pin still reads as (see the pinned branch below).
    let mut pin_saw_explorer_foreground = is_foreground_explorer();
    let mut last_user_input_at: Option<Instant> = None;
    let mut last_keyboard_navigation_input_at: Option<Instant> = None;
    // Set on a click, Enter or a navigation key press: the folder probe runs at
    // the faster cadence for a moment afterwards, so a folder that press opens
    // is noticed before the user's next key press instead of up to a full idle
    // interval later.
    let mut last_navigation_trigger_at: Option<Instant> = None;
    let mut stationary_hover_probe_done = false;
    // Whether a drag that began on a page the engine is drawing is still down. It stands while
    // a button is and falls on the first tick that finds none, so what it answers is "has the
    // hand let go", which no reading of the pointer alone can answer once an orbit has carried
    // the hand off the page (see `preview_window::note_engine_page_drag`).
    let mut engine_page_drag = false;

    // Safety net for a ffplay process that survived a stop (failed or
    // unconfirmed kill): while nothing is hovered it is re-checked and killed.
    let mut last_video_process_sweep = Instant::now();

    // Wheel scrolling moves the list under a stationary pointer, so the wheel tick
    // counter is the only signal that the hovered item changed (see `wheel_input`).
    // A scroll the cursor has not moved away from is what says a probe that finds
    // nothing under the pointer dismisses rather than waits.
    let mut consumed_wheel_ticks = wheel_input::wheel_tick_count();
    let mut scroll_since_move = false;

    // State for optimized polling
    let mut last_state_check = Instant::now();
    // Read rather than assumed: a state assumed here is one the loop only corrects
    // when its recheck comes round, and the deepest state's recheck is two seconds
    // off — so an app started with Explorer already up spent that long unable to
    // answer the pointer at all. This is the read the first recheck would have made,
    // made before the first tick instead.
    let mut current_state = read_explorer_state();

    // Polling intervals based on state. How fast the loop runs while Explorer has focus
    // is the `tick_ms` setting rather than a constant here — the one number that trades
    // how soon a move is answered against what the app costs while it works (see
    // `DEFAULT_TICK_MS`) — and the stationary probe below is gated by it as well: a file
    // under a parked pointer is read again no sooner than the loop looks. The ladder
    // itself is `explorer_pace`, because a pinned tick is paced by it as well.
    const VIDEO_HOVER_DISMISS_GRACE_MS: u64 = 350;
    const STATIONARY_SEARCH_MISS_HIDE_MS: u64 = 180;
    const VIDEO_PROCESS_SWEEP_MS: u64 = 1000;

    let (mut config_snapshot, mut trigger_key_vk, mut trigger_key_seen) = CONFIG
        .lock()
        .map(|c| {
            let snapshot = (
                c.preview_enabled,
                c.hover_delay_ms,
                c.trigger_key_mode,
                c.same_file_rehover_delay_ms,
                c.settling_delay_ms,
                c.prioritize_keyboard,
                c.trigger_key_enabled,
                c.trigger_key_affect_pin_mode,
                c.tick_ms,
            );
            // Resolved once per config change instead of once per tick: what the tick
            // compares is the spelling, and a spelling that has not changed is a key that
            // has not changed (see the snapshot below).
            let vk = crate::shell::key_input::key_to_vk(&c.trigger_key);
            (snapshot, vk, c.trigger_key.clone())
        })
        .unwrap_or((
            (
                true,
                DEFAULT_HOVER_DELAY_MS,
                TriggerKeyMode::Disable,
                DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
                DEFAULT_SETTLING_DELAY_MS,
                true,
                true,
                DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE,
                DEFAULT_TICK_MS,
            ),
            Some(0x12),
            "alt".to_string(),
        ));
    let mut explorer_probe_backoff_until: Option<Instant> = None;
    let mut last_display_signature = current_display_signature();
    let mut last_display_check = Instant::now();
    // What the Shell objects in hand were built against, and when they were last
    // built: a restart is what invalidates them, and a build that came out missing
    // is tried again rather than left missing for the rest of the run.
    let mut last_explorer_restarts = explorer_restart_count();
    let mut last_shell_build = Instant::now();
    // When a client that could not be bounded was last asked for again, and how long
    // the wait between two asks has grown to.
    let mut last_automation_rebind = Instant::now();
    let mut automation_rebind_interval_ms = UIA_REBIND_RETRY_MS;
    // When the probe counts were last written out, and where they go — nothing at
    // all unless `RHP_HOOK_TRACE` asked for them.
    let mut last_probe_flush = Instant::now();
    let probe_trace_path = hook_trace_path();
    // Whether the engines a preview loop that has stopped ticking was holding have
    // been ended for the stall it is in; cleared the moment it ticks again (see
    // `PREVIEW_STALL_MS`).
    let mut stalled_preview_engines_ended = false;

    while RUNNING.load(Ordering::SeqCst) {
        if let Some(path) = probe_trace_path.as_deref() {
            flush_probe_counts(Instant::now(), &mut last_probe_flush, path);
        }

        // A preview loop that has stopped ticking cannot end what it is holding, so
        // what it keeps warm is ended from here instead — once for the stall, and not
        // again until it ticks (see `PREVIEW_STALL_MS`).
        let preview_quiet_ms = preview_stall_ms();
        if preview_quiet_ms >= PREVIEW_STALL_MS {
            if !stalled_preview_engines_ended {
                stalled_preview_engines_ended = true;
                end_engines_of_a_stalled_preview(preview_quiet_ms, probe_trace_path.as_deref());
            }
        } else {
            stalled_preview_engines_ended = false;
        }
        // Explorer restarting is not something the resolver can recover from by
        // itself: the window collection it holds and the view that answered through
        // it are served by explorer.exe, and a proxy into a process that is gone
        // does not reconnect — its calls fail for good, which is what left every
        // hover after a restart with no answer at all. So the collection is built
        // again when a restart has been counted, and a collection that could not be
        // built when it was first asked for is built again on a slow retry, since
        // leaving it missing answers exactly as a dead one does.
        let explorer_restarts = explorer_restart_count();
        if explorer_restarts != last_explorer_restarts
            || (resolver.shell_windows.is_none()
                && last_shell_build.elapsed() >= Duration::from_millis(SHELL_COLLECTION_RETRY_MS))
        {
            last_explorer_restarts = explorer_restarts;
            last_shell_build = Instant::now();
            resolver.rebuild_shell();
            // Both caches are keyed by window handles, and those belonged to the
            // shell that is gone.
            clear_shell_view_probe_caches();

            // Nothing on screen is about a shell that exists any more, and the
            // windows the pointer could be over are the new shell's — which is
            // still putting them up, so the probes wait a moment before they ask.
            hide_preview();
            last_file = None;
            keyboard_file = None;
            is_keyboard_hover = false;
            last_focused_key = None;
            suppressed.clear();
            pointer_pause.clear();
            stationary_search_miss_started_at = None;
            hover_start = None;
            video_hover_guard_until = None;
            stationary_hover_probe_done = false;
            suspend_preview_until_user_input = true;
            allow_keyboard_preview_on_first_observation = false;
            folder_change_user_initiated = false;
            folder_change_time = Some(Instant::now());
            suspended_initial_focus = None;
            keyboard_press_seq_at_suspend = keyboard_navigation_press_seq;
            keyboard_screen_owner = false;
            hover_resolver_hints = HoverResolverHints::default();
            last_cursor_location = None;
            explorer_probe_backoff_until =
                Some(Instant::now() + Duration::from_millis(EXPLORER_RESTART_BACKOFF_MS));
            current_state = read_explorer_state();
            last_state_check = Instant::now();
        }

        // A resolver that never got a UI Automation client is asked for one again:
        // a client that failed once would otherwise never be asked for again, and
        // every hover after it would have no answer at all. A resolver holding a
        // client that could not be bounded is asked for a bounded one on the same
        // clock, in case the machine can produce one later.
        //
        // The ask is made less and less often. The object either can be created or
        // it cannot, and nothing about that changes between two seconds apart, so a
        // run that cannot be upgraded costs one attempt a minute rather than thirty.
        if !resolver.automation_bounded
            && last_automation_rebind.elapsed()
                >= Duration::from_millis(automation_rebind_interval_ms)
        {
            last_automation_rebind = Instant::now();
            resolver.rebuild_automation();
            automation_rebind_interval_ms =
                (automation_rebind_interval_ms * 2).min(UIA_REBIND_INTERVAL_MAX_MS);
        }

        // The pointer's answer belongs to the tick that produced it: the list under
        // a parked pointer can have moved on by the next one.
        resolver.forget_probe();

        // Nothing is hovered, so any ffplay still alive is a leftover from a
        // stop that did not take effect: kill it before it lingers on screen.
        //
        // A pin is not that. Nothing is hovered while one is up for most of its life, so this
        // sweep came round once a second and ended the sound a pinned window was playing — the
        // player is one process record (`VIDEO_PID`) and the pin's is whichever was started
        // last, so a safety net for a hover's leftovers killed the pin's own playback instead:
        // a sound that played about a second and stopped, a card whose clock went with it, and a
        // seek that began a player the next sweep ended. `pinned()` is the whole of the guard,
        // because a pin's player is taken down with the pin (see `preview_window::pinned`), and
        // the tick after that one this sweep is answering again.
        if last_file.is_none()
            && keyboard_file.is_none()
            && !is_keyboard_hover
            && !pinned()
            && last_video_process_sweep.elapsed() >= Duration::from_millis(VIDEO_PROCESS_SWEEP_MS)
        {
            last_video_process_sweep = Instant::now();
            kill_stray_video_process();
        }

        // Walking every display is not something to do on every tick for a change that
        // is noticed within this of happening anyway, and rebuilding what it
        // invalidates already waits out `DISPLAY_CHANGE_BACKOFF_MS`.
        if last_display_check.elapsed() >= Duration::from_millis(DISPLAY_CHECK_MS) {
            last_display_check = Instant::now();

            if let Some(display_signature) = current_display_signature() {
                if display_signature_changed(last_display_signature.as_ref(), &display_signature) {
                    last_display_signature = Some(display_signature);
                    clear_shell_view_probe_caches();
                    resolver.forget_window_views();
                    // The boxes an item was read from are the ones the display that has
                    // gone drew it in, so the item under the pointer is a question of its
                    // own now rather than an answer to be kept.
                    resolver.forget_item();
                    hide_preview();
                    last_file = None;
                    keyboard_file = None;
                    is_keyboard_hover = false;
                    suppressed.clear();
                    pointer_pause.clear();
                    stationary_search_miss_started_at = None;
                    hover_start = None;
                    video_hover_guard_until = None;
                    stationary_hover_probe_done = false;
                    suspend_preview_until_user_input = true;
                    allow_keyboard_preview_on_first_observation = false;
                    folder_change_user_initiated = false;
                    folder_change_time = Some(Instant::now());
                    suspended_initial_focus = None;
                    keyboard_press_seq_at_suspend = keyboard_navigation_press_seq;
                    keyboard_screen_owner = false;
                    hover_resolver_hints = HoverResolverHints::default();
                    last_cursor_location = None;
                    explorer_probe_backoff_until =
                        Some(Instant::now() + Duration::from_millis(DISPLAY_CHANGE_BACKOFF_MS));
                } else {
                    last_display_signature = Some(display_signature);
                }
            }
        }

        if let Some(until) = explorer_probe_backoff_until {
            if Instant::now() < until {
                // The one thing still watched for is the cursor leaving the file
                // the preview is of: that would leave a preview describing a file
                // the pointer is no longer on, which is worse than no preview. A
                // pointer that stays is answered with what it already has.
                unsafe {
                    let mut cursor_pos = POINT::default();
                    if GetCursorPos(&mut cursor_pos).is_ok() {
                        let dpi = monitor_dpi_from_point(cursor_pos.x, cursor_pos.y);
                        let threshold = pointer_pause.move_threshold_px(false, dpi);

                        if (cursor_pos.x - last_cursor_pos.x).abs() > threshold
                            || (cursor_pos.y - last_cursor_pos.y).abs() > threshold
                        {
                            last_cursor_pos = cursor_pos;

                            if last_file.is_some() {
                                hide_preview();
                                last_file = None;
                                hover_start = None;
                                video_hover_guard_until = None;
                            }
                        }
                    }
                }

                std::thread::sleep(Duration::from_millis(MEDIUM_SLEEP_MS));
                continue;
            }

            explorer_probe_backoff_until = None;
            // The pause is not a place a hover resumes from: whatever the cursor is
            // over when it ends has to be probed as something new.
            hover_start = Some(Instant::now());
            stationary_hover_probe_done = false;
            current_state = read_explorer_state();
            last_state_check = Instant::now();
        }

        if let Ok(config) = CONFIG.lock() {
            config_snapshot = (
                config.preview_enabled,
                config.hover_delay_ms,
                config.trigger_key_mode,
                config.same_file_rehover_delay_ms,
                config.settling_delay_ms,
                config.prioritize_keyboard,
                config.trigger_key_enabled,
                config.trigger_key_affect_pin_mode,
                config.tick_ms,
            );
            // The trigger key is resolved when it is *spelled* differently, not every
            // tick: what the tick does with it is read a key's state, and lower-casing a
            // name and looking it up again is work a setting that has not changed does not
            // need. The comparison is a string compare and costs no allocation; the clone
            // behind it happens once per change.
            if config.trigger_key != trigger_key_seen {
                trigger_key_seen = config.trigger_key.clone();
                trigger_key_vk = crate::shell::key_input::key_to_vk(&config.trigger_key);
            }
        }

        let preview_enabled = config_snapshot.0;
        let hover_delay_ms = config_snapshot.1;
        let trigger_key_mode = config_snapshot.2;
        let same_file_rehover_delay_ms = config_snapshot.3;
        let settling_delay_ms = config_snapshot.4;
        let prioritize_keyboard = config_snapshot.5;
        let trigger_key_enabled = config_snapshot.6;
        let trigger_key_affect_pin_mode = config_snapshot.7;
        let tick_ms = config_snapshot.8;

        // One question, two settings: the key either stops previews while it is
        // held, or is the only thing that lets them happen. Either way, what is left
        // to do when they are not allowed is the same as when they are turned off.
        // A key that is switched off is not asked about, and holds nothing back.
        //
        // A pin is not a hover, and the key that holds hovers back is not read while one is
        // up — up or collapsed into its bubble — unless `Affect Pin Mode` asks for it, which
        // is off where the app starts: a pinned preview is a window the user put there, and
        // its own close button is what takes it down. It is the `Hold to Disable Preview`
        // mode the setting speaks for; the reverse mode is left as it is.
        let trigger_key_muted_by_pin =
            trigger_key_mode == TriggerKeyMode::Disable && pinned() && !trigger_key_affect_pin_mode;
        let trigger_key_down = trigger_key_enabled
            && !trigger_key_muted_by_pin
            && trigger_key_vk.is_some_and(key_is_down);
        let previews_allowed =
            !trigger_key_enabled || trigger_key_mode.allows_previews(trigger_key_down);

        if !previews_allowed || !preview_enabled {
            if last_file.is_some() || keyboard_file.is_some() {
                hide_preview();
                suppressed.clear();
                pointer_pause.clear();
                stationary_search_miss_started_at = None;
                hover_start = None;
            }
            // A pinned preview is a window the user put there, and the two settings above are
            // what says whether previews may be raised at all — which is the question their
            // leaving it standing would answer wrongly. So it comes down with them, by the path
            // its own close button takes (see `request_pin_end`).
            if pinned() {
                request_pin_end();
            }
            keyboard_file = None;
            last_file = None;
            last_focused_key = None;
            is_keyboard_hover = false;
            keyboard_screen_owner = false;
            video_hover_guard_until = None;
            suspend_preview_until_user_input = false;
            allow_keyboard_preview_on_first_observation = false;
            folder_change_user_initiated = false;
            last_cursor_location = None;
            hover_resolver_hints = HoverResolverHints::default();
            folder_change_time = None;
            suspended_initial_focus = None;
            // A held key has to be noticed the moment it is released, so the poll
            // stays quick while the trigger is what is holding previews back, and
            // slows down only when previews are turned off outright.
            std::thread::sleep(Duration::from_millis(if previews_allowed {
                LONG_SLEEP_MS
            } else {
                tick_ms
            }));
            continue;
        }

        // A pin that has just been closed is a pointer that is on something new. The file the
        // pin was of is not a hover this hook has already answered, so none of what a hover
        // leaves behind — the file it was last about, the latch that holds a re-hover of one
        // back, the gate a folder change raises — may be read as still applying to it: a
        // preview of whatever the pointer is on now is due the moment the pin is gone.
        if take_pin_resumed() {
            last_file = None;
            keyboard_file = None;
            last_focused_key = None;
            hover_start = None;
            suppressed.clear();
            pointer_pause.clear();
            video_hover_guard_until = None;
            stationary_search_miss_started_at = None;
            suspend_preview_until_user_input = false;
            is_keyboard_hover = false;
            keyboard_screen_owner = false;
        }

        // A pinned preview is the whole of what this app is showing, and the pin's own promise is
        // that the hover machinery is quiet behind it. The loop stays here, at the tick's own pace,
        // rather than sleeping deeply, and what is not quiet behind it is the pin's own following
        // where `Pin Mode … Update Preview` asks for it (see `PinUpdateWatch`).
        if pinned() {
            // The state is read here too, on the clock the ladder keeps it on, so the one thing
            // this branch does not ask about is not left to go stale behind the pin: the loop's
            // record for the engines' idle timer comes from the window the user is actually in.
            let (_, state_recheck_ms) = explorer_pace(current_state, tick_ms);
            if last_state_check.elapsed() > Duration::from_millis(state_recheck_ms) {
                current_state = read_explorer_state();
                last_state_check = Instant::now();
            }

            // The item and view caches are frozen while a pin is up, and a pin that has just taken
            // the focus, or just given it back, is exactly when the user is about to pick: both go
            // on the transition itself rather than on a timer. This drop cannot cover a click,
            // which is read a tick before the shell has caught up with it (see
            // `PinUpdateWatch::follow`).
            let explorer_fg = is_foreground_explorer();
            if explorer_fg != pin_saw_explorer_foreground {
                resolver.forget_item();
                resolver.forget_window_views();
                pin_saw_explorer_foreground = explorer_fg;
            }

            let (update, on_hover) = pin_update_settings();
            let focus_move = focus_move_input();

            // A press the pin's own window cannot see — it is `WS_EX_NOACTIVATE` and holds no
            // capture, so a hand on another app's window, the tray, or the pin's own bubble never
            // reaches its procedure — is the one thing answered from here: the card's menu, where
            // it is up, is put away by the preview loop on the tick this ask is taken (see
            // `request_pin_menu_dismiss`). Nothing else about the pin is touched, and a press on
            // the pin window itself is not published (see `press_is_outside_the_pin_window`).
            let on_pin_window = focus_move.clicked
                && read_pointer().is_some_and(|pointer| preview_window_is_at(pointer.window));
            if press_is_outside_the_pin_window(focus_move.clicked, on_pin_window) {
                request_pin_menu_dismiss();
            }

            // What the loop itself saw on a pinned tick, before anything is decided about it: a
            // press that never reaches this line was spent before the hook read it, which no
            // reading of anything downstream can say.
            if focus_move.clicked {
                note_pin_click!(
                    "LOOP press  fg {}  update {}  on_hover {}",
                    is_foreground_explorer() as u8,
                    update as u8,
                    on_hover as u8,
                );
            }

            if update {
                // The hover's own delay, and the settling it must outlast as well.
                let delay = hover_delay_ms.max(settling_delay_ms);
                pin_watch.follow(
                    &mut resolver,
                    on_hover,
                    delay,
                    &mut last_keyboard_focus_probe,
                    focus_move,
                );
            } else {
                // A pin that follows nothing watches nothing, so that a setting switched back on
                // begins from what the pin is showing rather than from what the pointer was doing
                // while it was off. The reading above is still made and still dropped, and dropping
                // it is the only thing that spends the press bits.
                pin_watch = PinUpdateWatch::default();
            }

            std::thread::sleep(Duration::from_millis(tick_ms));
            continue;
        }

        let hover_delay = Duration::from_millis(hover_delay_ms);
        let settling_delay = Duration::from_millis(settling_delay_ms);

        // Determine sleep duration and whether to recheck state based on current state
        let (sleep_ms, state_recheck_ms) = explorer_pace(current_state, tick_ms);

        // Periodically re-evaluate the state
        if last_state_check.elapsed() > Duration::from_millis(state_recheck_ms) {
            current_state = read_explorer_state();
            last_state_check = Instant::now();
        }

        // If Explorer is not accessible, hide preview and sleep
        match current_state {
            ExplorerState::NoExplorerWindows
            | ExplorerState::AllMinimized
            | ExplorerState::HiddenByForeground => {
                if last_file.is_some() || keyboard_file.is_some() {
                    hide_preview();
                    last_file = None;
                    stationary_search_miss_started_at = None;
                    hover_start = None;
                    keyboard_file = None;
                    last_focused_key = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                    pointer_pause.clear();
                    keyboard_screen_owner = false;
                }
                std::thread::sleep(Duration::from_millis(sleep_ms));
                continue;
            }
            ExplorerState::VisibleNotFocused => {
                // Explorer is visible but not focused - do a quick cursor check
                // Only activate full polling if cursor is actually over Explorer.
                // A pointer that a preview is holding is not evidence that the
                // user has left: the preview is on top of Explorer, so the check
                // below cannot see Explorer under it.
                //
                // Both questions are about the pointer as it is, and both are asked of
                // one reading of it: what is under the pointer is one window, and a
                // preview holding the pointer is one point in one region (see
                // `PointerTick`).
                let pointer = read_pointer();
                let over_explorer =
                    pointer.is_some_and(|pointer| is_cursor_over_explorer_full(pointer.window));
                let holds_pointer = pointer
                    .is_some_and(|pointer| preview_pointer_hold(pointer.point.x, pointer.point.y));

                if !over_explorer && !holds_pointer {
                    if last_file.is_some() || keyboard_file.is_some() {
                        hide_preview();
                        last_file = None;
                        stationary_search_miss_started_at = None;
                        hover_start = None;
                        keyboard_file = None;
                        last_focused_key = None;
                        is_keyboard_hover = false;
                        video_hover_guard_until = None;
                        pointer_pause.clear();
                        keyboard_screen_owner = false;
                    }
                    std::thread::sleep(Duration::from_millis(sleep_ms));
                    continue;
                }
                // Cursor is over Explorer, switch to active state
                current_state = ExplorerState::ActiveFocus;
            }
            ExplorerState::ActiveFocus => {
                // Continue with active polling below
            }
        }

        // Explorer is active - use the configured tick
        std::thread::sleep(Duration::from_millis(tick_ms));

        unsafe {
            // Get cursor position, as the tick's one reading of the pointer: everything
            // this tick asks about it — where it is, the scale of the display it is on, and
            // the window under it — is a question about one instant, and it is read here
            // rather than again by each caller (see `PointerTick`).
            let mut cursor_pos = POINT::default();
            if GetCursorPos(&mut cursor_pos).is_err() {
                continue;
            }
            let pointer = PointerTick::of(cursor_pos);

            // Whether the pointer is on a preview that holds it: a text preview the
            // user can read or select from — on the preview, or inside the margin
            // around it, which covers the gap it crosses on its way from the file it
            // belongs to — or the spinner a page is being rendered behind, which the
            // pointer reaches only by drifting into its box, and which has no page
            // yet to hand the pointer back to. While that holds, what is on screen
            // is what the user is waiting on rather than something in the way, so it
            // is not dismissed, and the file under the pointer is not resolved, so
            // it cannot be replaced by whatever it covers.
            //
            // Read once here because more than one path in this loop asks: the
            // dismissal below, the hover resolver, the mouse hover delay, and the
            // "Explorer is visible but not focused" branch, where a pointer that
            // is not over Explorer is not a reason to close a preview either.
            let loop_now = Instant::now();

            // The buttons, read before the hold below rather than beside the rest of the
            // key state further down, because what the page rule below reads depends on them
            // and the press bit a click is known by can only be taken once. The two are
            // disjoint sets of keys, so reading the mouse before the keyboard is the same
            // reading whichever order it happens in.
            let MouseButtons {
                active: mouse_button_input,
                pressed: mouse_button_press,
                ..
            } = mouse_buttons();

            // A drag that began on a page the engine is drawing, which is the page's own and
            // has to be held through the hand leaving it: an orbit carries the pointer well
            // outside the rectangle the page was drawn in, and a preview taken down there is a
            // page the user was working on, gone. It is armed by a press inside the engine's
            // own rectangle and stands until a tick finds no button down at all, which is the
            // only reading that says the hand has let go — a drag begun on the listing rather
            // than on the page never arms it, and so never holds a preview the page is not
            // responsible for. Noted before the hold is read, so the tick that begins a drag
            // is the tick the page is already being held by.
            if mouse_button_press {
                engine_page_drag = webview_preview::screen_rect()
                    .is_some_and(|rect| point_in_box(cursor_pos, rect));
            } else if !mouse_button_input {
                engine_page_drag = false;
            }
            note_engine_page_drag(engine_page_drag);

            // Read straight from the published region, with nothing held over from
            // the last tick: the moment the pointer is out of it, the preview is
            // treated the way it was before the pointer ever touched it.
            let pointer_hold = preview_pointer_hold(cursor_pos.x, cursor_pos.y);

            let move_threshold = pointer_pause
                .move_threshold_px(is_keyboard_hover || keyboard_screen_owner, pointer.dpi);
            // How far the hand has come is one reading of "the mouse has moved", and
            // the file the preview on screen is about is another: a pointer that has
            // left that item has moved whatever the threshold says, and the two are
            // taken together so a row crossed under the threshold cannot leave the
            // preview of the file that was left behind (see
            // `pointer_moved_off_the_hovered_item`). The box is only asked about
            // while a preview is up, which is when it means anything at all.
            let moved = (cursor_pos.x - last_cursor_pos.x).abs() > move_threshold
                || (cursor_pos.y - last_cursor_pos.y).abs() > move_threshold
                || pointer_moved_off_the_hovered_item(
                    last_file.is_some(),
                    pointer_hold,
                    pointer_item_holds(cursor_pos.x, cursor_pos.y),
                );
            // Read the navigation keys first: GetAsyncKeyState's "pressed since
            // the previous call" bit goes away with the first read of a key in
            // an iteration, and that fresh press is what a folder change has to
            // tell apart from a held key. The keys a shortcut is made of are read
            // with them, once, for the same reason (see `navigation_input`).
            let navigation = navigation_input();
            let keyboard_navigation_active = navigation.active;
            let keyboard_navigation_press = navigation.pressed;
            let explorer_navigation_shortcut_input = navigation.shortcut;
            let keyboard_navigation_input =
                explorer_navigation_shortcut_input || keyboard_navigation_active;
            let mouse_navigation_input = is_mouse_navigation_button_detected();
            let (activation_key_input, activation_key_press) = activation_key_input_state();

            // A wheel tick only counts while the wheel is driving Explorer: the
            // pointer is over it, or over a keyboard preview that covers the
            // pointer while Explorer still receives the wheel. The counter is
            // always consumed so a tick seen over something else cannot be
            // replayed.
            let wheel_ticks = wheel_input::wheel_tick_count();
            let wheel_tick = wheel_ticks != consumed_wheel_ticks;
            if wheel_tick {
                consumed_wheel_ticks = wheel_ticks;
            }
            let keyboard_owns_pointer = is_keyboard_hover || pointer_pause.freezes_pointer();
            let wheel_scroll = wheel_tick
                && (is_cursor_over_explorer_full(pointer.window)
                    || (keyboard_owns_pointer && cursor_preview_hover().any()));
            if wheel_scroll {
                scroll_since_move = true;
                // The list moves under a parked cursor, so the file under the pointer
                // is a new question: the probe latch is reopened and the hover clock
                // restarted, or a pointer that never moves would never be asked about
                // the file that scrolled under it. The item an answer was read from goes
                // with it: the pointer has not moved, and what it is on has.
                stationary_hover_probe_done = false;
                resolver.forget_item();
                hover_start = Some(loop_now);

                if keyboard_owns_pointer {
                    // The wheel is the mouse taking over from the keyboard: close
                    // the keyboard preview, release the pointer, and end the
                    // recent-keyboard-input window so the keyboard cannot
                    // re-establish its preview while the wheel is driving. The file
                    // the keyboard showed is not latched, so the mouse may preview
                    // it again where the cursor ends up.
                    hide_preview();
                    keyboard_file = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                    pointer_pause.clear();
                    keyboard_screen_owner = false;
                    last_focused_key = None;
                    allow_keyboard_preview_on_first_observation = true;
                    last_keyboard_navigation_input_at = None;
                } else if prioritize_keyboard && keyboard_screen_owner {
                    // A keyboard turn with no preview of its own to close ends the
                    // same way, where `Prioritize Keyboard` is what was holding the
                    // pointer back: the wheel is the pointer taking the screen back, so
                    // the item the scroll brings under it is previewed like any
                    // other. Nothing of the keyboard's is on screen, so nothing is
                    // taken down with it — and the recent-keyboard-input window ends
                    // with the turn, or the focus probe could put a keyboard preview
                    // over the item the wheel has just placed under the pointer.
                    keyboard_screen_owner = false;
                    last_keyboard_navigation_input_at = None;
                }
            }

            if moved
                || keyboard_navigation_input
                || mouse_navigation_input
                || mouse_button_input
                || activation_key_input
                || wheel_scroll
            {
                last_user_input_at = Some(loop_now);
            }
            if keyboard_navigation_input {
                last_keyboard_navigation_input_at = Some(loop_now);
            }
            if keyboard_navigation_press {
                keyboard_navigation_press_seq = keyboard_navigation_press_seq.wrapping_add(1);
                // The item a fresh press selects is the user's own choice, so it
                // must not be swallowed as a fresh baseline, even while no
                // baseline is stored (folder change, mouse move, startup).
                allow_keyboard_preview_on_first_observation = true;
                // The press is also the keyboard taking the screen: the pointer
                // stays parked, and nothing it sits on previews, until it takes
                // its turn back with a move or the wheel.
                keyboard_screen_owner = true;
            }
            if mouse_button_press || activation_key_press || keyboard_navigation_press {
                last_navigation_trigger_at = Some(loop_now);
                // A press is the list being acted on, and what a press can do to a view is
                // move what is under a parked pointer: a click on a column header sorts
                // it. The item an answer was read from is dropped with it.
                resolver.forget_item();
            }

            // Enter opens the item the keyboard is on, so the view is about to change under
            // whatever is on screen the way it changes under a Backspace: the hover goes
            // with the press rather than waiting for the folder probe to notice, which is
            // what used to leave a preview of the file that was being hovered over an empty
            // folder with nothing left to take it down. The folder-change gate below is
            // left to arm itself, so the item that lands under a parked cursor still
            // previews in the new listing without a move: an Enter that opened a folder is
            // navigation, not a reason to hold the pointer off the listing it opened.
            if activation_key_press {
                if last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover {
                    hide_preview();
                }
                // The file that was on screen is latched the way a dismissed hover is,
                // so the hover below cannot put it straight back up in the moment before
                // the view this press opened has changed: a pointer left on it is a user
                // asking for it again once the rehover delay has passed.
                match last_file.clone().or_else(|| keyboard_file.clone()) {
                    Some(file) => suppressed.suppress(file),
                    None => suppressed.clear(),
                }
                last_file = None;
                keyboard_file = None;
                is_keyboard_hover = false;
                pointer_pause.clear();
                stationary_search_miss_started_at = None;
                resolver.forget_item();
                hover_start = None;
                last_focused_key = None;
                video_hover_guard_until = None;
            }

            // A click is the pointer acting on the view, and the one press that can change
            // the folder without the pointer having moved: the address bar's breadcrumbs are
            // clicked where the pointer already is. A keyboard preview cannot be told about
            // that by the folder probe — that probe is the pointer's own and is skipped for
            // as long as a keyboard preview is up — so the press ends the keyboard's turn
            // itself: what is on screen goes, and the new listing is read as the place it is
            // rather than as the listing the preview belongs to. A press on the preview is
            // not this: a text frame is clicked to read it, not to leave it.
            if mouse_button_press
                && (is_keyboard_hover || keyboard_file.is_some())
                && !pointer_hold
                && !cursor_preview_hover().any()
            {
                hide_preview();
                keyboard_file = None;
                is_keyboard_hover = false;
                video_hover_guard_until = None;
                pointer_pause.clear();
            }

            if explorer_navigation_shortcut_input || mouse_navigation_input {
                if last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover {
                    hide_preview();
                }
                last_file = None;
                keyboard_file = None;
                is_keyboard_hover = false;
                suppressed.clear();
                pointer_pause.clear();
                stationary_search_miss_started_at = None;
                resolver.forget_item();
                hover_start = None;
                last_focused_key = None;
                video_hover_guard_until = None;
                suspend_preview_until_user_input = true;
                allow_keyboard_preview_on_first_observation = false;
                folder_change_user_initiated = false;
                // History navigation (Backspace, Alt+arrows, the mouse
                // back/forward buttons) changes the folder too: keep the faster
                // probe cadence through the hold and past the release so the new
                // location is recognized before the user's next key press.
                last_navigation_trigger_at = Some(loop_now);
                folder_change_time = Some(Instant::now());
                suspended_initial_focus = None;
                keyboard_press_seq_at_suspend = keyboard_navigation_press_seq;
                keyboard_screen_owner = false;
                last_cursor_pos = cursor_pos;
                stationary_hover_probe_done = false;
                continue;
            }

            // A keyboard preview is placed next to the focused item, which can put
            // it right over the parked cursor. Decide from the preview's own box
            // whether the pointer is under it, and freeze every pointer-driven
            // trigger until the mouse is moved on purpose.
            if pointer_pause.is_watching() {
                pointer_pause.evaluate_box(cursor_pos, preview_screen_rect());
            }

            // Close as soon as the cursor touches the preview window. Keep
            // suppressing preview until the cursor leaves so a delayed spinner
            // or background load result cannot resurrect a stuck preview under
            // the pointer. Keyboard previews own the screen: they may cover the
            // parked cursor and are never dismissed by it. A page that runs is
            // the third exception and is not answered here at all: it holds the
            // pointer through its own rectangle, and a preview a hold covers is
            // never one this closes.
            //
            // What the pointer can touch is three surfaces, and they are not the same window:
            // this app's own layered preview, the player's window a video is played in, and
            // — for a document or a specimen, or a page that runs — the engine's window,
            // which is asked for by its own handle because a document drawn by a browser
            // is still a preview of this app's (see `cursor_preview_hover`).
            let preview_hover = if should_probe_preview_hover(
                is_keyboard_hover || pointer_pause.freezes_pointer(),
                last_file.is_some(),
                suppress_preview_until_cursor_leaves_preview,
            ) {
                cursor_preview_hover()
            } else {
                PreviewCursorHover::NONE
            };
            let over_image_preview = preview_hover.image;
            let over_engine_preview = preview_hover.engine;
            let over_video_preview = preview_hover.video;
            let over_any_preview = preview_hover.any();

            // A preview that holds the pointer is the exception to that rule: while
            // it does, the preview is kept — see `pointer_hold` above for what
            // holding it covers. A preview the pointer cannot work with keeps the
            // old behaviour, which is to close as soon as it is touched.
            let guard_active = video_hover_guard_until
                .map(|until| Instant::now() < until)
                .unwrap_or(false);
            let should_dismiss_for_preview_hover = ((over_image_preview || over_engine_preview)
                && !pointer_hold)
                || (over_video_preview && !guard_active);

            if should_dismiss_for_preview_hover
                || (suppress_preview_until_cursor_leaves_preview && over_any_preview)
            {
                suppress_preview_until_cursor_leaves_preview = true;
                if let Some(file) = last_file.clone() {
                    suppressed.suppress(file);
                }
                hide_preview();
                last_file = None;
                keyboard_file = None;
                is_keyboard_hover = false;
                stationary_search_miss_started_at = None;
                video_hover_guard_until = None;
                stationary_hover_probe_done = false;
                hover_start = Some(Instant::now());
                continue;
            }

            if suppress_preview_until_cursor_leaves_preview {
                suppress_preview_until_cursor_leaves_preview = false;
                stationary_hover_probe_done = false;
                hover_start = Some(Instant::now());
                last_cursor_pos = cursor_pos;
                continue;
            }

            // Detect folder/navigation changes and suspend preview until user input.
            // Probe at active-poll cadence only while a preview is visible; idle
            // polling keeps the slower cadence to avoid extra COM work. A click,
            // Enter or navigation key press takes that place for a moment: the
            // folder it opens has to be recognized before the user's next key
            // press, which otherwise lands in the folder-change gate below. While
            // the keyboard drives, this cursor-based probe stays off: it resolves
            // the window under the pointer, which is the keyboard preview itself.
            let preview_active =
                last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover;
            let navigation_trigger_active = recent_elapsed_within(
                last_navigation_trigger_at.map(|at| at.elapsed()),
                FOLDER_PROBE_TRIGGER_MS,
            );
            if last_folder_probe.elapsed()
                >= Duration::from_millis(if navigation_trigger_active {
                    FOLDER_PROBE_MS
                } else {
                    folder_probe_interval_ms(preview_active)
                })
                && !is_keyboard_hover
                && !pointer_pause.freezes_pointer()
                && !pointer_hold
                && should_probe_hover_resolver(
                    preview_active,
                    moved,
                    last_user_input_at.map(|at| at.elapsed()),
                )
            {
                last_folder_probe = Instant::now();
                hover_resolver_hints = get_current_hover_resolver_hints(&mut resolver, &pointer);
                let location = HoverLocation::of(&hover_resolver_hints);
                // A look that answered nothing about the place is left alone: it is a
                // question to ask again, not a change to act on. And a fact is only
                // read as a change where both looks answered it, so a folder the shell
                // could not walk out this time is not another place — see
                // `HoverLocation`.
                let differs = location.was_answered()
                    && last_cursor_location
                        .as_ref()
                        .map(|previous| hover_location_changed(previous, &location))
                        // The first look at a place is a change like any other: what
                        // is under the pointer is asked about as a view that has just
                        // been opened rather than one that was always there.
                        .unwrap_or(true);

                if differs {
                    // A change that follows recent input is user navigation:
                    // the file under the parked cursor may preview as soon as
                    // the new view has settled, without a mouse move.
                    let user_navigation = recent_elapsed_within(
                        last_user_input_at.map(|at| at.elapsed()),
                        HOVER_RESOLVER_INPUT_GRACE_MS,
                    );
                    last_cursor_location = Some(location);
                    // The view this location describes is a different one now —
                    // another folder, or the same window searched again — so what
                    // was cached about the last one describes the wrong place: a
                    // folder remembered for the window and the view that answered
                    // for it. A name looked up against those resolves to something
                    // that is not there, or to nothing at all. The answer this tick
                    // already holds was read against the view that has been left, so
                    // it goes with them rather than deciding anything below.
                    clear_shell_view_probe_caches();
                    resolver.forget_window_views();
                    resolver.forget_item();
                    resolver.forget_probe();
                    hover_start = None;
                    last_focused_key = None;
                    // Reset cursor baseline so we don't mistake stale delta for movement.
                    last_cursor_pos = cursor_pos;
                    stationary_hover_probe_done = false;

                    // Whether the file the preview on screen is about is still the one
                    // under the pointer, asked of the view that has just answered. A
                    // place re-read in another spelling, a handle taken again, a walk
                    // that failed once and succeeded the next time are all changes to
                    // the *view*, and the file the pointer stands on is a question of its
                    // own, so what the view says now is what answers it. The box the
                    // preview was resolved from is not asked here: it belongs to the
                    // listing that has been left, and a look that answers nothing against
                    // it is not a read that failed — the file is not there any more.
                    // Read the other way, a folder change with nothing under the pointer
                    // left the preview of the old file on screen with nothing to take it
                    // down, which is what an empty folder and a folder pressed into with
                    // Enter both came to look like (see `read_failure_is_the_same_item`
                    // for the hover the box is still the right question for).
                    let still_on_the_file = last_file.as_ref().is_some_and(|file| {
                        get_file_under_cursor(&mut resolver, &pointer)
                            .is_some_and(|current| same_path(file, &current))
                    });

                    if !still_on_the_file {
                        // Nothing released the gate yet: the view the place describes is
                        // one the pointer's file has not been read in, so the preview
                        // that was up belongs to a file this view does not hold.
                        suspend_preview_until_user_input = true;
                        allow_keyboard_preview_on_first_observation = false;
                        folder_change_user_initiated = user_navigation;
                        folder_change_time = Some(Instant::now());
                        suspended_initial_focus = None;
                        // Drain stale GetAsyncKeyState flags from prior navigation,
                        // then remember the press count: only a later navigation key
                        // press may lift this suspension, so key state left over
                        // from the navigation that opened the folder cannot.
                        let _ = navigation_input();
                        keyboard_press_seq_at_suspend = keyboard_navigation_press_seq;

                        if last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover {
                            hide_preview();
                        }
                        last_file = None;
                        suppressed.clear();
                        pointer_pause.clear();
                        stationary_search_miss_started_at = None;
                        keyboard_file = None;
                        is_keyboard_hover = false;
                        video_hover_guard_until = None;
                    }
                }
            }

            // Hard gate: after folder change, do not preview until explicit user input.
            if suspend_preview_until_user_input {
                // Cooldown: ignore all input for 150ms after folder change to let
                // COM/accessibility settle and to avoid stale keyboard state.
                if let Some(change_time) = folder_change_time {
                    if change_time.elapsed() < Duration::from_millis(150) {
                        continue;
                    }
                }

                let navigation_press =
                    keyboard_navigation_press_seq != keyboard_press_seq_at_suspend;

                // A scroll is deliberate pointer input, so it releases the
                // suspension exactly like a mouse move. The flag is used
                // instead of the raw tick because the cooldown above bails out
                // before this check, which would swallow a one-notch scroll.
                // A folder opened by a click, Enter or a navigation key is user
                // navigation too, so its suspension lifts once the new view has
                // settled and the item under the parked cursor can preview
                // without a mouse move. A navigation key press after the change
                // releases it as well: that press is the user asking for the
                // keyboard preview and must not be swallowed as a baseline.
                if moved || scroll_since_move || folder_change_user_initiated || navigation_press {
                    suspend_preview_until_user_input = false;
                    allow_keyboard_preview_on_first_observation = navigation_press;
                    // A folder change hands the screen back to the pointer: the item
                    // that landed under a parked cursor previews in the new listing
                    // without a mouse move, the way it always has. A press is the
                    // exception — that is the user asking for the keyboard again —
                    // and takes the screen for the new folder with it. Only the turn
                    // `Prioritize Keyboard` keeps is handed over here; with the setting
                    // off the flag is not what holds the pointer back, and what it says
                    // about the pointer's tolerance is left as the keyboard left it.
                    if prioritize_keyboard {
                        keyboard_screen_owner = navigation_press;
                    }
                    folder_change_user_initiated = false;
                    hover_start = Some(Instant::now());
                    stationary_hover_probe_done = false;
                    suspended_initial_focus = None;
                    folder_change_time = None;
                } else {
                    // Nothing released the gate yet. Keep watching UI Automation
                    // focus changes as a fallback for focus moves that no counted
                    // press explains, such as Explorer restoring focus while the
                    // new view is being built.
                    let mut keyboard_unlocked = false;
                    if should_probe_keyboard_focus(
                        last_keyboard_navigation_input_at.map(|at| at.elapsed()),
                    ) && is_foreground_explorer()
                        && last_keyboard_focus_probe.elapsed()
                            >= Duration::from_millis(KEYBOARD_FOCUS_PROBE_MS)
                    {
                        last_keyboard_focus_probe = Instant::now();
                        if let Some(focused_info) = get_focused_explorer_item(&resolver) {
                            let focused_key = FocusedItemKey::new(
                                focused_info.item.name.clone(),
                                &focused_info.item.bounds,
                            );

                            if suspended_initial_focus.is_none() {
                                // Record the auto-focused first item
                                // (set by Windows when folder opens)
                                suspended_initial_focus = Some(focused_key);
                            } else if suspended_initial_focus.as_ref() != Some(&focused_key) {
                                // Focus actually changed — user pressed a navigation key
                                keyboard_unlocked = true;
                            }
                        }
                    }

                    if keyboard_unlocked {
                        suspend_preview_until_user_input = false;
                        allow_keyboard_preview_on_first_observation = true;
                        // No press released the gate, so this is the new view
                        // settling — the focus moving as the listing is built —
                        // rather than the keyboard taking the screen: where the turn
                        // is what holds the pointer back, the folder change stays the
                        // pointer's, and the keyboard's turn comes back with the
                        // user's own next press.
                        if prioritize_keyboard {
                            keyboard_screen_owner = false;
                        }
                        folder_change_user_initiated = false;
                        suspended_initial_focus = None;
                        folder_change_time = None;
                    } else {
                        continue;
                    }
                }
            }

            // A move is "the mouse driving Explorer": resolve the item under the
            // cursor, and drop the preview when that is no longer the file it
            // shows. Another file always takes over, even while the pointer is
            // inside a scrollable preview's region — the region is only about the
            // pointer being on its way to the preview or on it, and the block
            // below tells those two apart by whether anything is under the pointer
            // at all.
            if moved {
                last_cursor_pos = cursor_pos;
                stationary_search_miss_started_at = None;
                stationary_hover_probe_done = false;
                // A real move hands control back to the mouse and ends the scroll the
                // cursor was sitting on.
                pointer_pause.clear();
                scroll_since_move = false;
                keyboard_screen_owner = false;

                // Mouse movement always takes priority - dismiss keyboard hover.
                // The keyboard preview may have been covering the cursor, so the
                // file it showed is latched the way any dismissed hover is: held
                // off the mouse path for the same-file rehover delay, so the
                // handover cannot flash it straight back, and previewable again
                // after that — a pointer left sitting on the file is a user asking
                // for it.
                if is_keyboard_hover {
                    if let Some(file) = keyboard_file.clone() {
                        suppressed.suppress(file);
                    }
                    hide_preview();
                    keyboard_file = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                }
                // A mouse move leaves the keyboard focus baseline unknown, so the
                // next focus observed while the user is driving with the keyboard
                // acts immediately. Recording it as a fresh baseline instead would
                // swallow the first key press and keep the mouse preview on screen.
                last_focused_key = None;
                allow_keyboard_preview_on_first_observation = true;

                if let Some(suppressed_file) = suppressed.file.clone() {
                    if let Some(current_file) = get_file_under_cursor(&mut resolver, &pointer) {
                        if same_path(&suppressed_file, &current_file) {
                            hover_start = Some(Instant::now());
                            continue;
                        }
                        suppressed.clear();
                        stationary_search_miss_started_at = None;
                    }
                }

                // While moving (including list scrolling), avoid heavy accessibility
                // resolution and wait until hover is stable before probing media.
                if last_file.is_some() {
                    let mut keep_while_pointer_held = false;
                    if let Some(current_file) = get_file_under_cursor(&mut resolver, &pointer) {
                        if last_file
                            .as_ref()
                            .map(|last| same_path(last, &current_file))
                            .unwrap_or(false)
                        {
                            hover_start = Some(Instant::now());
                            continue;
                        }
                        // Another file is under the pointer, so this is a hover like
                        // any other and the preview gives way to it — even while the
                        // pointer is inside a scrollable preview's region, which is
                        // why the check below is the "no file at all" case only.
                        suppressed.clear();
                        stationary_search_miss_started_at = None;
                    } else if pointer_hold {
                        // Nothing under the pointer: it is on its way to (or on) the
                        // preview, or on the spinner standing in for one, which is
                        // not the user leaving the file it shows.
                        keep_while_pointer_held = true;
                    } else if read_failure_is_the_same_item(
                        hover_item_box,
                        pointer_item_box(),
                        cursor_pos,
                    ) {
                        // A read that came back with nothing, from the item the preview on
                        // screen is about: what failed is the asking — a walk out through the
                        // shell on a volume slow to answer, a view busy drawing the item it
                        // was just asked for — so the preview is kept with the question asked
                        // again, rather than taken down for a look that failed and put back a
                        // moment later, which is a blink. A read that came back with nothing
                        // from *another* item is a look at an item with no preview of its own,
                        // which is a pointer that has left the file this preview is about: the
                        // preview goes (see `read_failure_is_the_same_item`).
                        keep_while_pointer_held = true;
                    } else if let Some(file) = last_file.clone() {
                        suppressed.suppress(file);
                    }

                    if !keep_while_pointer_held {
                        hide_preview();
                        last_file = None;
                        video_hover_guard_until = None;
                    }
                }
                // The clock the hover below is measured against is restarted by the move,
                // so what the file under the pointer has to outlast starts here. A
                // settling delay is the setting that asks for a pointer which has stopped,
                // and this is where a moving one waits it out: with it on, the hover below
                // is not reached at all. With it off — 0, which is where the app starts —
                // a move falls through instead, so the new file under the cursor has a
                // preview put up for it even while the hand is still on its way to it.
                hover_start = Some(Instant::now());

                if settling_delay_ms > 0 {
                    continue;
                }
            }

            // Mouse is stationary - check for keyboard navigation
            // Only when Explorer is the foreground window (keyboard input goes there).
            // A move that fell through to here is still a move: it has just handed the
            // screen to the mouse and dropped the focus baseline, and the item the
            // keyboard is on is not read again until the pointer has stopped.
            if !moved
                && should_probe_keyboard_focus(
                    last_keyboard_navigation_input_at.map(|at| at.elapsed()),
                )
                && is_foreground_explorer()
                && last_keyboard_focus_probe.elapsed()
                    >= Duration::from_millis(KEYBOARD_FOCUS_PROBE_MS)
            {
                last_keyboard_focus_probe = Instant::now();
                if let Some(focused_info) = get_focused_explorer_item(&resolver) {
                    // The name alone cannot tell two observations apart: a search
                    // can hold the same name in more than one folder, and the box is
                    // what says which of them the keyboard is on.
                    let focused_key = FocusedItemKey::new(
                        focused_info.item.name.clone(),
                        &focused_info.item.bounds,
                    );

                    if last_focused_key.is_none() {
                        if allow_keyboard_preview_on_first_observation {
                            // The focus baseline is unknown — the mouse just moved, or
                            // the user unlocked a folder change with the keyboard — so
                            // this first observed item acts immediately instead of
                            // being recorded and waiting for a second key press.
                            last_focused_key = Some(focused_key);
                            allow_keyboard_preview_on_first_observation = false;

                            // Resolve to a media file and show keyboard preview
                            if let Some(path) =
                                resolve_focused_item_to_path(&mut resolver, &focused_info)
                            {
                                // Dismiss any active mouse hover: the keyboard's own preview
                                // is what is about to replace it.
                                if last_file.is_some() && !is_keyboard_hover {
                                    hide_preview();
                                    last_file = None;
                                    suppressed.clear();
                                    hover_start = None;
                                }

                                if keyboard_file.as_ref() != Some(&path) {
                                    // Hide previous preview before showing new one
                                    if is_keyboard_hover {
                                        hide_preview();
                                    }
                                    keyboard_file = Some(path.clone());
                                    is_keyboard_hover = true;
                                    keyboard_screen_owner = true;
                                    suppress_preview_until_cursor_leaves_preview = false;
                                    pointer_pause.watch_for_box();
                                    video_hover_guard_until = if is_video_file(&path) {
                                        Some(
                                            Instant::now()
                                                + Duration::from_millis(
                                                    VIDEO_HOVER_DISMISS_GRACE_MS,
                                                ),
                                        )
                                    } else {
                                        None
                                    };
                                    show_preview_keyboard(
                                        &path,
                                        focused_info.item.bounds.left,
                                        focused_info.item.bounds.top,
                                        focused_info.item.bounds.right,
                                        focused_info.item.bounds.bottom,
                                        focused_info.item.keyboard_avoid_box(),
                                        focused_info.item.draws_columns(),
                                    );
                                }
                            } else {
                                // Not a media file - hide any keyboard preview
                                if is_keyboard_hover {
                                    hide_preview();
                                }
                                // The pointer's own hover is left where it is: a key pressed
                                // onto an item with nothing to show must not take the preview
                                // of the file under the pointer down and put it straight back
                                // up. It goes only where `Prioritize Keyboard` is what holds
                                // the pointer back — there the keyboard owns the screen and
                                // the pointer waits for its turn (see the guard below).
                                if prioritize_keyboard && last_file.is_some() {
                                    hide_preview();
                                    last_file = None;
                                    suppressed.clear();
                                    hover_start = None;
                                }
                                keyboard_file = None;
                                is_keyboard_hover = false;
                                video_hover_guard_until = None;
                            }
                            continue;
                        }

                        // Nothing to compare against yet and no keyboard input has
                        // claimed this focus, so record it as the baseline only.
                        last_focused_key = Some(focused_key);
                    } else if last_focused_key.as_ref() != Some(&focused_key) {
                        // Focused item changed - keyboard navigation detected
                        last_focused_key = Some(focused_key);
                        allow_keyboard_preview_on_first_observation = false;

                        // Resolve to a media file and show keyboard preview
                        if let Some(path) =
                            resolve_focused_item_to_path(&mut resolver, &focused_info)
                        {
                            // Dismiss any active mouse hover: the keyboard's own preview
                            // is what is about to replace it.
                            if last_file.is_some() && !is_keyboard_hover {
                                hide_preview();
                                last_file = None;
                                suppressed.clear();
                                hover_start = None;
                            }

                            if keyboard_file.as_ref() != Some(&path) {
                                // Hide previous preview before showing new one
                                if is_keyboard_hover {
                                    hide_preview();
                                }
                                keyboard_file = Some(path.clone());
                                is_keyboard_hover = true;
                                keyboard_screen_owner = true;
                                suppress_preview_until_cursor_leaves_preview = false;
                                pointer_pause.watch_for_box();
                                video_hover_guard_until = if is_video_file(&path) {
                                    Some(
                                        Instant::now()
                                            + Duration::from_millis(VIDEO_HOVER_DISMISS_GRACE_MS),
                                    )
                                } else {
                                    None
                                };
                                show_preview_keyboard(
                                    &path,
                                    focused_info.item.bounds.left,
                                    focused_info.item.bounds.top,
                                    focused_info.item.bounds.right,
                                    focused_info.item.bounds.bottom,
                                    focused_info.item.keyboard_avoid_box(),
                                    focused_info.item.draws_columns(),
                                );
                            }
                        } else {
                            // Not a media file - hide any keyboard preview
                            if is_keyboard_hover {
                                hide_preview();
                            }
                            // The pointer's own hover is left where it is, for the reason
                            // the first observation leaves it: a key pressed onto an item
                            // with nothing to show must not take the preview of the file
                            // under the pointer down and put it straight back up. It goes
                            // only where `Prioritize Keyboard` holds the pointer back.
                            if prioritize_keyboard && last_file.is_some() {
                                hide_preview();
                                last_file = None;
                                suppressed.clear();
                                hover_start = None;
                            }
                            keyboard_file = None;
                            is_keyboard_hover = false;
                            video_hover_guard_until = None;
                        }
                        continue;
                    }
                }
            }

            // If keyboard hover is active, or the pointer is frozen under a
            // keyboard preview, skip mouse hover delay logic entirely. The same
            // goes for a pointer held by what is on screen — inside a scrollable
            // text preview, or on the spinner it is waiting behind: the file under
            // it is not what the user is looking at, so nothing is hovered over.
            //
            // A pointer parked since the keyboard last drove is held back by it where
            // `Prioritize Keyboard` asks for that: the keyboard owns the screen, so a file
            // it left behind — or one it landed on that has no preview to give — is not a
            // reason for the pointer to raise a preview of whatever it happens to sit on.
            // That is the answer a key pressed onto a file that can be previewed already
            // gets, and it is what keeps the two previews from fighting over a list the
            // keyboard is walking. The turn ends when the pointer is used on purpose —
            // moved past the tolerance, or given a wheel tick — or when a folder change
            // hands the screen back to it. With the setting off the pointer's own hover
            // wins: a parked pointer previews the file it is on while the keyboard
            // drives, and a key pressed onto an item with no preview to give leaves that
            // hover standing rather than taking it down and putting it back.
            if is_keyboard_hover
                || (prioritize_keyboard && keyboard_screen_owner)
                || pointer_pause.freezes_pointer()
                || pointer_hold
            {
                continue;
            }

            // Check if we've hovered long enough (mouse hover). The settling delay is
            // asked first and it is the pointer's own: it is measured off the same clock,
            // which a move restarts, so it is how long the hand has been still — a
            // pointer that has only just stopped is held back for its length even with no
            // hover delay in front of it, and 0 asks nothing of the pointer at all.
            if let Some(start) = hover_start {
                if start.elapsed() >= hover_delay && start.elapsed() >= settling_delay {
                    if !should_probe_stationary_hover(stationary_hover_probe_done) {
                        if let Some(miss_started) = stationary_search_miss_started_at {
                            if miss_started.elapsed()
                                >= Duration::from_millis(STATIONARY_SEARCH_MISS_HIDE_MS)
                            {
                                match last_file.clone() {
                                    Some(file) => suppressed.suppress(file),
                                    None => suppressed.clear(),
                                }
                                hide_preview();
                                last_file = None;
                                stationary_search_miss_started_at = None;
                                video_hover_guard_until = None;
                            }
                        }
                        continue;
                    }
                    if last_hover_probe.elapsed() < Duration::from_millis(tick_ms) {
                        continue;
                    }
                    last_hover_probe = Instant::now();

                    // A scroll the cursor has not moved away from is what makes a
                    // probe that finds nothing dismiss rather than wait.
                    let scroll_driven = scroll_since_move;

                    // Try to get file under cursor
                    let resolved = get_file_under_cursor(&mut resolver, &pointer);
                    // One probe per parked cursor: this one closes the latch, and the
                    // events that make the file under the cursor a new question — a
                    // move, a wheel tick, a folder change — are what reopen it.
                    stationary_hover_probe_done = true;

                    if let Some(file_path) = resolved {
                        if !last_file
                            .as_ref()
                            .map(|last| same_path(last, &file_path))
                            .unwrap_or(false)
                        {
                            if suppressed.matches(&file_path) {
                                let required_delay = hover_delay_ms.max(same_file_rehover_delay_ms);
                                if !suppressed.rehover_allowed(required_delay) {
                                    // The latch is a delay and not a verdict: the
                                    // probe is left open so the file previews as
                                    // soon as the delay has passed, rather than
                                    // leaving a pointer parked on it unanswered.
                                    stationary_hover_probe_done = false;
                                    continue;
                                }
                            }
                            suppressed.clear();
                            stationary_search_miss_started_at = None;
                            last_file = Some(file_path.clone());
                            // The item this preview is about, as the view drew it: what a
                            // later read that comes back with nothing is compared against
                            // (see `read_failure_is_the_same_item`). The look above is the
                            // one that published it.
                            hover_item_box = pointer_item_box();
                            video_hover_guard_until = if is_video_file(&file_path) {
                                Some(
                                    Instant::now()
                                        + Duration::from_millis(VIDEO_HOVER_DISMISS_GRACE_MS),
                                )
                            } else {
                                None
                            };
                            show_preview(
                                &file_path,
                                cursor_pos.x,
                                cursor_pos.y,
                                avoid_box_under_cursor(&resolver, cursor_pos),
                            );
                        }
                    } else {
                        if scroll_driven {
                            // The scroll left something under the pointer that is not
                            // a media file: drop the preview that scrolled away, the
                            // same way the mouse-move path does.
                            match last_file.clone() {
                                Some(file) => suppressed.suppress(file),
                                None => suppressed.clear(),
                            }
                            hide_preview();
                            last_file = None;
                            stationary_search_miss_started_at = None;
                            video_hover_guard_until = None;
                            hover_start = Some(Instant::now());
                            continue;
                        }

                        let search_view_active = hover_resolver_hints.is_search_view
                            || is_current_search_view_legacy(&mut resolver, &pointer);
                        if search_view_active {
                            let miss_started =
                                stationary_search_miss_started_at.get_or_insert_with(Instant::now);
                            if miss_started.elapsed()
                                >= Duration::from_millis(STATIONARY_SEARCH_MISS_HIDE_MS)
                            {
                                match last_file.clone() {
                                    Some(file) => suppressed.suppress(file),
                                    None => suppressed.clear(),
                                }
                                hide_preview();
                                last_file = None;
                                stationary_search_miss_started_at = None;
                                video_hover_guard_until = None;
                                hover_start = Some(Instant::now());
                                continue;
                            }
                        } else {
                            stationary_search_miss_started_at = None;
                        }
                    }
                }
            } else {
                // Initialize hover_start if not moving
                stationary_search_miss_started_at = None;
                stationary_hover_probe_done = false;
                hover_start = Some(Instant::now());
            }
        }
    }

    // Everything COM handed the loop is released while the apartment that owns it
    // is still initialized. An interface released after `CoUninitialize` belongs to
    // an apartment that has already been torn down, which faults — and the resolver
    // holds the Shell window collection, the batched property request and the view
    // that answered last, not just the automation client.
    drop(resolver);

    unsafe {
        CoUninitialize();
    }
}

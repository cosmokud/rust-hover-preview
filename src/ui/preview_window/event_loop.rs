//! `run_preview_window` itself: the message loop, one tick at a time.
//!
//! This file is over the ceiling every other part of this module is held to, and is left over it
//! on purpose. The loop is one function of about 3,700 lines: each of its arms is one kind of
//! message the window can be sent, each answer ends where the next tick begins, and the order
//! the arms are read in is the order the answers are given in. Splitting it would mean splitting
//! a decision - what a pinned preview does about a file it is still loading, a media that has
//! ended, a pointer that has come back to rest - into pieces that have to be told what the other
//! pieces decided, which is a change to the program rather than a move of it. Nothing is
//! extracted out of it here, and no helper is introduced into it: it is the file as it was.

use super::*;

const PREVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWindow");

pub fn run_preview_window() {
    // Page sizes come from Windows.Data.Pdf and picture sizes from the codec Windows
    // has, so this thread needs an apartment before the first layout asks for one.
    pdf_preview::initialize_apartment();
    wic_image::initialize_apartment();

    let (tx, rx): (Sender<PreviewMessage>, Receiver<PreviewMessage>) = channel();

    // Store sender for other threads to use
    if let Ok(mut sender) = PREVIEW_SENDER.lock() {
        *sender = Some(tx);
    }

    unsafe {
        let hinstance = GetModuleHandleW(None).unwrap();

        // Register window class
        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(window_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance.into(),
            hIcon: Default::default(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hbrBackground: Default::default(),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PREVIEW_CLASS,
            hIconSm: Default::default(),
        };

        RegisterClassExW(&wc);

        // And the class the round bubble is created from: a collapsed pin's window is hidden
        // while the bubble is up, and the bubble is not that window's shape — it is a small
        // circle standing in for a large window, at whatever place on the desktop the user
        // leaves it.
        let bubble_class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(pin_bubble_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance.into(),
            hIcon: Default::default(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hbrBackground: Default::default(),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PIN_BUBBLE_CLASS,
            hIconSm: Default::default(),
        };

        RegisterClassExW(&bubble_class);

        // And the class the taskbar stand-in is created from: a pin put away into the taskbar
        // rather than into the bubble leaves this window behind, and it is the one surface of
        // this app's that is meant to be in the taskbar — an icon and a title, nothing drawn,
        // kept off the screen (see `pin_taskbar`).
        //
        // The icon it is registered under is what the Shell draws its taskbar button with, and a
        // window class takes an icon once, as it registers — so this is the app's own icon rather
        // than whatever the button would otherwise be given (see `app_icon`).
        let taskbar_icon = crate::shell::app_icon::app_icon();
        let taskbar_class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(pin_taskbar_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance.into(),
            hIcon: taskbar_icon,
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hbrBackground: Default::default(),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PIN_TASKBAR_CLASS,
            hIconSm: taskbar_icon,
        };

        RegisterClassExW(&taskbar_class);

        // And the class the card's menu is drawn in: a panel of its own over the screen rather
        // than rows drawn inside the card, which is what keeps a menu taller than the card it
        // came out of from having its rows cut off by the card's own edges (see
        // `pin_menu_window`).
        let menu_class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(pin_menu_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance.into(),
            hIcon: Default::default(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hbrBackground: Default::default(),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PIN_MENU_CLASS,
            hIconSm: Default::default(),
        };

        RegisterClassExW(&menu_class);

        // Create the preview window
        let hwnd = CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE,
            PREVIEW_CLASS,
            w!("Preview"),
            WS_POPUP,
            0,
            0,
            1,
            1,
            None,
            None,
            hinstance,
            None,
        )
        .unwrap();

        // Store HWND as isize
        PREVIEW_HWND.store(hwnd.0 as isize, Ordering::SeqCst);

        // Track current video path to avoid restarting
        let mut current_video_path: Option<PathBuf> = None;
        // The display the hover on screen was measured against: what a sound's card is painted
        // at, and what it is painted at again while its player runs.
        let mut audio_card_dpi = 96u32;
        // When the sound on screen was started, and when its card was last painted. The clock a
        // sound FFmpeg plays is measured from the first, and the second is the cadence its card
        // is drawn at (see `audio_clock` and `AUDIO_CARD_REPAINT`).
        let mut audio_started: Option<Instant> = None;
        // The second of the file the sound was started at, and nothing for one started at its
        // beginning: what a card's clock is measured from where the player reports none.
        let mut audio_start_offset = 0.0f64;
        // The second a key in a pinned window has held a sound at, which is what its card is
        // drawn at while held: the player is ended to pause one, so the card is left standing on
        // the second it stopped at (see `toggle_pinned_audio`).
        let mut audio_paused: Option<f64> = None;
        // A `Volume → Audio Seek` of `Middle` or `Random` asked of a file whose length nothing
        // had read yet — the one start position that cannot be worked out where the sound is
        // started. Asked again the moment a player reports a length (see the tick below).
        let mut audio_share_seek: Option<AudioSeek> = None;
        let mut audio_repaint_at = Instant::now();
        // The beat the card's mark was last drawn with, which is what a fall starts from: a
        // pause hands this loop no pulse to read, and the ripples come home over the repaints
        // that follow instead of standing where the sound left them (see
        // `audio_meter::beat_fade`).
        let mut audio_beat = 0.0f32;
        // The name on screen, scrolled sideways while the card has no room for it, put up with
        // the card it belongs to so a hover that changes begins at the beginning again.
        let mut audio_name_scroll: Option<audio_preview::NameScroll> = None;
        // The hover the preview came from, so a theme or Markdown switch can rebuild it without
        // waiting for the next hover.
        let mut current_show: Option<PreviewMessage> = None;
        // Video position/size, for the periodic topmost re-assertion.
        let mut video_pos: (i32, i32, i32, i32) = (0, 0, 0, 0); // (x, y, w, h)
        let mut last_topmost_check = Instant::now();
        // When the transport bar was last painted: its playhead moves on its own.
        let mut last_pin_repaint = Instant::now();
        // How often a collapsed pin's bubble is painted while its bars move.
        //
        // It is the one window of this app's that is redrawn for no reason but that it is *moving*:
        // a sound's histogram is a shape made of the bands of the sound now, and the bands of a
        // sound that is playing change every few milliseconds — so the bubble is painted a little
        // under thirty times a second, which a hand reads as motion and which is the same order the
        // transport bar's own playhead is painted at (`PIN_TRANSPORT_REPAINT_MS`). A film's bars
        // are the same bands drawn over its own frame, and are on this clock for as long as the
        // film is making sound (see `bubble_bars_are_moving`). Nothing else about a bubble moves, so
        // nothing else is on this clock (see the bubble repaint below).
        const BUBBLE_REPAINT: Duration = Duration::from_millis(33);
        // The play state a collapsed pin's bubble was last painted for, and when it was last
        // painted: what it stands for is asked of the file's own player, and a film paused by a key
        // sends this loop nothing — so the answer is compared against the one the bubble was drawn
        // with rather than waited for (see `bubble_is_playing`).
        let mut last_bubble_playing: Option<bool> = None;
        let mut last_bubble_repaint = Instant::now();

        // Background loading support
        let (load_tx, load_rx): (Sender<LoadResult>, Receiver<LoadResult>) = channel();
        let load_request_slot: LoadRequestSlot = Arc::new((Mutex::new(None), Condvar::new()));
        let load_worker = spawn_load_worker(Arc::clone(&load_request_slot), load_tx);
        // The pin's own slow questions, answered off this thread — a folder walked, a file
        // opened through the Shell — so that a caption being slow does not take its own
        // buttons with it (see `spawn_pin_planner`).
        spawn_pin_planner();
        // And the thread that gives up on a pin this loop has stopped turning, which is the
        // one thing the loop itself cannot do (see `spawn_pin_watchdog`).
        spawn_pin_watchdog();
        let mut current_generation: u64 = 0;
        let mut pending_load: Option<PendingLoad> = None;
        let mut pending_load_cancel: Option<Arc<AtomicBool>> = None;
        let mut last_stream_overlay_repaint = Instant::now();
        // The page an engine owes the preview on screen — Office's render tier where the
        // document is one of its own, the render engine where it is one of `[libre]`'s —
        // and whether one has been asked for and is being waited on.
        let mut page_render_pending: Option<(PathBuf, u64)> = None;
        // The hover an engine has already been asked to come up for, so that the ask is made
        // once for a file the pointer has settled on rather than on every tick (see
        // `warm_engines_for`).
        let mut warmed_generation: Option<u64> = None;
        // A page that arrived for the hover already on screen, and is being loaded
        // to replace what is there rather than to open a new preview.
        let mut page_upgrade: Option<PathBuf> = None;
        // The video whose probe a hover is waiting on, and the generation of the hover
        // that is waiting: the geometry is measured on a thread of its own and the hover
        // is replayed when the answer lands (see `VideoProbed`).
        let mut video_probe: Option<(PathBuf, u64)> = None;
        // The hover a video's probe has just answered for. Its replay is the same wait
        // carried on rather than a new preview: what is on screen — the spinner — stays
        // where it is until the video replaces it, the way a page landing on a spinner
        // behaves (see `upgrading`).
        let mut video_replay: Option<PathBuf> = None;
        // The file the engine was watched failing at: the engine took it and then never drew a
        // frame of it, so it is a file no player here will take — which, now that FFmpeg's player
        // is where every video is played where it is installed at all, is a machine whose only
        // player is the one that failed. The hover that is up is replayed to be answered without
        // a preview, which is what this waits to be done where a replay can be started — the tick
        // that finds the failure out cannot start one itself (see `video_player::mark_unplayable`).
        let mut engine_failed: Option<PathBuf> = None;
        // The hover a box measured off the preview thread has just answered for. Its replay is
        // the same wait carried on rather than a new preview, the way a video's is (see
        // `measure_replay`).
        let mut measure_replay: Option<PathBuf> = None;
        // The file a pinned window is waiting for a box of: what the user picked has no box of its
        // own yet — a page being measured off the preview thread, a video being probed — so what
        // the pick comes to is a wait rather than a swap. The file is picked up again by name when
        // the answer lands, which is a tick of its own, so this outlives the tick the way the
        // replay flags above it do (see `PinPlan::Awaiting`).
        let mut pin_awaiting_box: Option<PathBuf> = None;
        // The return a put-away pin is owed once a file it can be shown is on the window: a pick
        // names a file nothing has read yet, and a window put up on the pick shows the file before it
        // for as long as the read and the decode take — the flash the user sees between the two
        // images. So the return is owed where the file is picked and answered where the file is on
        // the window (see `bring_a_put_away_pin_back` and the `PreviewMessage::Pin` arm, which is the
        // only place this is read).
        //
        // It is the pin the debt is owed to and not a file, because the file a pick names may be one
        // nothing can be shown for: a pin that stays away for it is brought back by the first file
        // the walk lands on that can be shown (see `pin_failure_plan`), which is the answer the user
        // asked for — the pin is the thing missed, not the file it was asked for.
        let mut pin_return_owed = false;
        // The pick a pinned window still owes the front of its band for, because it has just
        // answered one: the click reaches the listing's window after it reaches this loop, and the
        // listing is activated then — over the pin (see `PIN_RAISE_GRACE`, `PinFrontOwed` and
        // `keep_pin_at_the_front`).
        let mut pin_front_owed: Option<PinFrontOwed> = None;
        // The file picked in Explorer while the pin is a bubble: held rather than shown, because a
        // pin that is a bubble has no window to show anything in, and what the user is doing
        // behind it is walking through files. The key is what brings the window back up on it
        // (see `pin_key_restores_bubble`).
        let mut pin_bubble_pick: Option<PathBuf> = None;
        // A file a pinned window is loading, and the wait it is. The load is a thread's work,
        // so the pin keeps the file it is showing until the answer lands, and a wait that has
        // run for `spinner_delay_ms` puts an arc in the middle of the pin's media (see
        // `PinLoad` and `paint_pin_spinner`).
        let mut pin_load: Option<PinLoad> = None;
        // A swap of the pin's file that has been started and is being held until the media
        // engine hands over its first frame: the file the pin is showing stays on screen at the
        // frame it had stopped on, with the arc still turning over it, rather than being taken
        // down for a placeholder that is not a picture of anything (see `PinSwapHold`).
        let mut pin_swap_hold: Option<PinSwapHold> = None;
        // A pinned window's own media being decoded again at a box it has been given, and the
        // wait for it. Held here rather than inside a tick for the reason a load is: the read
        // and the decode are a thread's work, so a window dragged to a new size keeps drawing
        // the frame it had — scaled into the new box — until the frame drawn for that box
        // lands, rather than costing a tick as long as the decode (see `PinRelayout`).
        let mut pin_relayout: Option<PinRelayout> = None;
        // The box a zoomed picture was last laid out again for, so that the file is asked for once per
        // size the picture settles at rather than on every tick it stands there (see
        // `pin_zoom_oversize_box`).
        let mut zoom_sharpened: Option<ScreenRegion> = None;
        // When the last frame of a pinned picture's zoom was taken, which is the clock the travel is
        // measured against: a release is written in seconds, and the loop is the one thing here that
        // knows how long its own last tick was (see `step_pin_zoom`).
        let mut last_zoom_frame = Instant::now();
        // A walk a pin's own caption button stepped, held for as long as the walk has files
        // left to offer rather than for the length of the tick that started it: a file the pin
        // cannot be shown is stepped over rather than stopped at, and the file after it is
        // asked for on a tick of its own, with the walk carried onto it. What the walk stands
        // on is the file to be shown next (see `PinStep`).
        let mut pin_walk: Option<PinStep> = None;
        // When FFmpeg's player behind the file the pin is showing now was started, which is what
        // says a player that is gone was one that refused the file rather than one that was
        // watched to its end or closed by the user. `None` for every other kind, and for a pin
        // that is not up (see `pin_media_failed_before_a_frame`).
        let mut pin_player_started: Option<Instant> = None;
        // A walk the planner is reading the folder for, and the wait it is. It is held here
        // rather than inside a tick for the same reason a load is: the read is on another
        // thread, so the question is still out when this tick ends and the answer to it
        // arrives on one of the next. The wait is painted as the same arc a load's is,
        // because it is the same wait to the person looking at it (see `PinWait`).
        let mut pin_walk_wait: Option<PinWait> = None;
        // Whether an arc was ever put up for a walk, so that the one put up for it comes
        // down when the walk lands or is given up on rather than turning for a question
        // nobody is asking any more.
        let mut pin_walk_wait_seen = false;
        // A listing pick that arrived while a hover load was in flight, held across ticks until
        // nothing is loading rather than dropped with the per-tick slot: the per-tick pick below
        // is consumed where it is taken up and never read again that tick, so a pick held in it
        // would not survive into the tick that takes it up. The walk, if any, rides in the
        // existing outer `pin_walk` (see the `pending_load.is_some()` re-queue arm).
        let mut pin_held_pick: Option<PathBuf> = None;
        // A player that has been started and has not put its window up yet: the wait
        // for a video, which the spinner stands in for until the player's window is
        // there (see `VideoStart`).
        let mut video_start: Option<VideoStart> = None;
        // A video the media engine plays whose window has been held back for its first
        // frame: the preview is put up by the frame's arrival rather than opened on the
        // placeholder its load lands with (see `FirstFrameWait`).
        let mut first_frame_wait: Option<FirstFrameWait> = None;

        // Message loop
        let mut msg = MSG::default();
        // A message the idle wait took off the channel, held for the drain below
        // rather than acted on where it was received.
        let mut carried_preview_msg: Option<PreviewMessage> = None;
        // What a carry of a pinned drag took off the channel while it was inside the tick, in
        // the order it arrived. A carry cannot act on a message itself — the tick owns the media
        // and the windows — so it holds them and the tick takes them up as its own (see `Carry`).
        let mut carried_preview_messages: Vec<PreviewMessage> = Vec::new();
        // How many left-button presses the pin's own press handling has already been given, as a
        // count rather than a latch on the button being down: a press and a release between two
        // ticks of this loop leaves nothing to latch, so a drag begun from one would never begin at
        // all (see `settle_pinned_engine_press`).
        //
        // Reset on every take-up, which is what a local of the loop cannot do for itself: the
        // count is published by the hook for the life of the process, so a snapshot of it
        // outlives the pin it was taken in and swallows the first press of the next one.
        let mut engine_press_seen: u64 = 0;
        // And the same count for the raise a press owes any of the preview's windows — the pin's
        // own, a film in its band, a page the engine draws in one — kept the same way and reset
        // with it (see `settle_pinned_press`).
        let mut pin_press_seen: u64 = 0;
        // Static-tick throttle state (see `STATIC_WAIT_MS`): whether the media on
        // screen was dynamic the last time it was fully looked at, when that look
        // was, and which generation it was for. A swap bumps `current_generation`,
        // which forces the next tick to look fully rather than trust the hint.
        let mut media_dynamic_hint = true;
        let mut last_full_media_check = Instant::now() - Duration::from_secs(10);
        let mut last_checked_generation: u64 = u64::MAX;
        // Pointer-hold publish throttle (see `POINTER_HOLD_HEARTBEAT_MS`): the last
        // time the region was published and the shape it was published for, so a
        // static preview republishes at most twice a second plus on transitions.
        let mut last_hold_publish = Instant::now() - Duration::from_secs(10);
        let mut last_hold_key = (false, false, false, 0u64);
        while RUNNING.load(Ordering::SeqCst) {
            // Every tick is noted, whether it does anything or not: what the note is
            // for is the Explorer hook telling a loop that is working from one that has
            // stopped, and a loop waiting on the channel is working (see
            // `preview_stall_ms`).
            note_preview_alive();

            // And the same note for the pin, which is a narrower question with a much
            // shorter bound: a loop that is not turning while a window the user is
            // looking at is up is a window whose buttons do nothing, and the thread that
            // would notice is the one that is stuck (see `spawn_pin_watchdog`).
            if pinned() {
                note_pin_alive();
            }

            // What the pin asks for this tick, if anything: the key, one of the caption's
            // buttons, or the media behind it having come apart. It is held for the drain
            // below, which is where a hover's own messages are turned into a preview.
            let mut pin_request: Option<PreviewMessage> = None;
            // The file the key asked a bubble to be brought back up on, answered further down the
            // tick by the ordinary swap rather than here, so that a bubble is shown a file the
            // same way a window is (see `pin_bubble_pick`).
            let mut pin_swap_requested: Option<PathBuf> = None;

            // Check for Windows messages
            while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }

            // Detect resume from sleep (WM_POWERBROADCAST handler sets this flag).
            // DWM restarts on resume and destroys the layered window's composition
            // surface, so we must reset all local state to force a fresh start on
            // the next hover.
            if RESUME_FROM_SLEEP.load(Ordering::Acquire) {
                RESUME_FROM_SLEEP.store(false, Ordering::Release);
                current_generation += 1;
                pending_load = None;
                clear_load_request(&load_request_slot);
                if let Some(cancel) = pending_load_cancel.take() {
                    cancel.store(true, Ordering::Release);
                }
                current_video_path = None;
                video_pos = (0, 0, 0, 0);
                // A preview held back for a video's first frame goes with the media the
                // reset above has already dropped: the session it was waiting on is the
                // one the resume let go of, and the frame it was for is not coming.
                first_frame_wait = None;

                // A browser engine is a process that does not survive a suspend in any
                // state worth keeping, so it is let go with everything else and begun
                // again when the next document asks for it.
                webview_preview::shutdown();

                // Re-assert layered window style after DWM restart.
                // DWM is reinitialized during resume and the layered window's
                // per-pixel alpha composition surface may need a fresh anchor.
                //
                // The band is re-asserted as the window's own rather than as the topmost one for
                // every window: a hover is topmost, and a pin is topmost while it holds the front
                // (see `pin_holds_the_front`) — a pin owed nothing is an ordinary window, and this
                // is the one road to the band that does not go through the pin's own raise.
                let banded = !pinned() || pin_holds_the_front();
                let ex_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
                SetWindowLongPtrW(
                    hwnd,
                    GWL_EXSTYLE,
                    if banded {
                        ex_style | WS_EX_LAYERED.0 as isize | WS_EX_TOPMOST.0 as isize
                    } else {
                        (ex_style | WS_EX_LAYERED.0 as isize) & !(WS_EX_TOPMOST.0 as isize)
                    },
                );
                let _ = SetWindowPos(
                    hwnd,
                    if banded { HWND_TOPMOST } else { HWND_NOTOPMOST },
                    0,
                    0,
                    0,
                    0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED,
                );
            }

            // The pin's own business, once a tick: the key that pins and unpins, the three
            // buttons of the caption, and whether the thing the pin is a window onto is still
            // there at all.
            //
            // Snapshot for the static-tick hint (see the bottom of the tick): a load,
            // wait, generation or pin state that moves during this tick may install
            // another kind behind the hint, which is then re-classified rather than
            // trusted until the periodic refresh.
            let tick_had_wait = pending_load.is_some()
                || pin_load.is_some()
                || pin_walk_wait.is_some()
                || video_start.is_some()
                || first_frame_wait.is_some();
            let tick_generation = current_generation;
            let tick_pinned = pinned();
            if PIN_END_REQUESTED.swap(false, Ordering::AcqRel) {
                pin_request = Some(end_pin_state(Reason::Asked));
            }
            // A press the pin's own window never saw — on another app's window, the tray, or
            // the pin's own bubble — is the hook's ask to put the card's menu away: only the
            // menu goes, and the pin itself stands (see `take_pin_menu_dismiss`). The menu's
            // own capture answers these first where it is held; this is the road where it is
            // not, and it lets the pointer go with the menu it puts away.
            if take_pin_menu_dismiss() {
                release_pin_capture(hwnd);
                render_layered_preview(hwnd);
            }

            // Which windows a standing pin takes the keyboard keys for, told to the hook once a
            // tick: a `WH_KEYBOARD_LL` callback may not walk the desktop to find them, so this is
            // the tick publishing what it already knows. Zero for both says there is no such pin,
            // which is what leaves the arrows alone in the listing behind a preview that is not
            // one (see `key_input::pin_owns_caret`).
            let pin_keys_owner = if pinned() { hwnd.0 as u64 } else { 0 };
            let pin_keys_video = if pin_keys_owner != 0 {
                VIDEO_HWND.load(Ordering::SeqCst) as u64
            } else {
                0
            };
            crate::shell::key_input::publish_pin_key_owner(pin_keys_owner, pin_keys_video);
            // And whether that pin is above every other window, which is the same question the
            // wheel asks before it takes a notch: a topmost pin is the window under the pointer
            // wherever the keyboard is (see `pin_owns_foreground`).
            crate::shell::key_input::publish_pin_always_on_top(
                pin_keys_owner != 0 && pin_always_on_top(),
            );
            // And whether the preview is the surface the pointer is standing on, which is what the
            // wheel asks instead of the keyboard's question once a pin is raised over the listing
            // the user is working in: the pointer is over the pin while the keyboard stays behind
            // it, and a notch turned there is the pin's either way (see
            // `wheel_input::publish_preview_under_the_pointer`).
            crate::shell::wheel_input::publish_preview_under_the_pointer(
                cursor_preview_hover().any(),
            );

            // The key is drained whether or not it is watched, so that a press made while the
            // feature was off is not acted on when it comes back on.
            let pin_presses = crate::shell::key_input::take_presses();
            if pin_presses > 0 && pin_enabled() {
                if pinned() {
                    // The key no longer hides a window that is up. A pin the user has not pressed
                    // is a pin they are not in, and a key pressed at a window nobody is in is not a
                    // request to close it — least of all one that reads as a Space thrown at
                    // whatever the user was typing when the pin came up. A window that is up
                    // answers the keyboard itself now, as a window in the foreground does (see
                    // `pinned_key_command`).
                    //
                    // What the key still does is bring back a bubble. A bubble is a window the
                    // user put away on purpose, so the key that put it away is the key that brings
                    // it back, and it is answered only where the keyboard is in Explorer — behind
                    // a bubble is a listing the user is working in, and the file picked there is
                    // what the window should come back up on (see `pin_bubble_pick`).
                    if pin_key_restores_bubble(
                        pin_is_collapsed(),
                        crate::shell::explorer_hook::is_foreground_explorer(),
                    ) {
                        // The window comes back only where the swap has room to be made in the
                        // same tick: a load already in hand is a tick the pick below has no
                        // work for, and a window put back up without the file it was brought
                        // back for is a worse answer than the bubble it stayed in. The file is
                        // left in hand for the next press either way.
                        if pending_load.is_none() {
                            if let Some(path) = pin_bubble_pick.take() {
                                restore_pin(&Win32PinWindow, PinRestore::Key);
                                pin_swap_requested = Some(path);
                                // A file picked in the listing is not a step of the walk, so
                                // nothing is stepped over on its account: it is a file the user
                                // named, and a file the pin cannot show is a pin that keeps what
                                // it has rather than a walk that moves on (see `PinStep`).
                                pin_walk = None;
                            }
                        }
                    }
                } else if let Some(request) =
                    pin_what_is_on_screen(&current_show, pending_load.as_ref())
                {
                    pin_request = Some(request);
                }
            }

            if pinned() {
                // A film that has been watched to its end is begun again here, before anything else
                // in this tick reads the player that ended it as a failure. Three things down here
                // would each close the pin over a player that exited because its film finished
                // rather than because it died, and the loop that replaces them is this one (see
                // `loop_ended_pinned_player`).
                //
                // Not while the cover is owed a player: a supersede-kill leaves the clock near
                // the film's end with no player behind it, and a loop begun at the beginning
                // here would stack a second relaunch onto the replacement already on its way.
                if !pin_park_covers_a_relaunch() {
                    // An advance already on its way is not asked for again (see
                    // `loop_ended_pinned_player`): the wait, the walk and the load are
                    // what say one is.
                    let advance_busy = pin_walk.is_some()
                        || pin_load.is_some()
                        || pin_walk_wait.is_some()
                        || pin_swap_hold.is_some()
                        || pin_awaiting_box.is_some()
                        || pin_held_pick.is_some();
                    loop_ended_pinned_player(
                        pin_walk_anchor(&pin_walk),
                        &mut pin_walk_wait,
                        advance_busy,
                    );
                }

                // A press that landed on the window the engine draws a document in, which is the
                // one thing on a pinned window the window procedure cannot be told about: the
                // band is a browser's window over this one, so the press is read here and the
                // drag it means begun here (see `settle_pinned_engine_press`).
                settle_pinned_engine_press(hwnd, &mut engine_press_seen);

                // And a press on any of the preview's windows is the preview being asked for: the
                // window comes forward whether or not Windows granted it the foreground, and
                // whichever of its windows took the press (see `settle_pinned_press`).
                // The pin's own window, for the same reason the raise above asks for it: the handle
                // this loop holds is the hover's (see `pinned_window`).
                settle_pinned_press(pinned_window(), &mut pin_press_seen);

                // And a drag that press began carried on, and let go of, from what the hook
                // publishes rather than from a message — which is the only end it can have, the
                // press having been taken by the engine's window rather than by this one (see
                // `settle_pinned_engine_drag`). And then followed at the hand's own rate rather
                // than at the tick's, since nothing this window is sent will ever carry it (see
                // `carry_pin_drag_with_the_hand`).
                settle_pinned_engine_drag(hwnd);

                // Followed at the hand's own rate rather than at the tick's, since nothing this
                // window is sent will ever carry it. What the carry found on the channel goes to
                // the front of this tick's own messages: a `Close` on the caption, a media kind
                // switched off in the tray, or an engine that has died are all asked for past the
                // line this loop is inside, so without this a hand that kept moving left every
                // one of them standing — which is the caption going deaf for the length of the
                // drag (see `Carry`).
                carried_preview_messages.extend(carry_pin_drag_with_the_hand(hwnd, &rx));

                // A file the pin was shown and its player cannot draw is not a pin that
                // came apart — it is a pin that is about to be shown something else, and the
                // step off is queued here rather than left to the next tick so that the
                // `navigating` flag below is true while the command that runs there asks
                // whether the pin is still there. Read after the drag and the carry, because a
                // file the engine failed at is one this app has already been told about and
                // nothing on this thread's channel is going to say so.
                if let Some(failing) = pin_media_failed_before_a_frame(pin_player_started) {
                    // The file is written down before anything is stepped over, or the
                    // next visit routes it to the engine again and fails there again for
                    // as long as the pin is up (see `video_player::mark_unplayable`). And
                    // the hover's own flag goes with it: under a pin nothing replays a
                    // hover from it, so it would only sit there being true about a file
                    // this tick has already dealt with.
                    video_player::mark_unplayable(&failing);
                    engine_failed = None;

                    pin_step_off(take_pin_walk(), &mut pin_walk);

                    // And where the walk has nothing left to offer, the mark for the file is
                    // what is shown instead of nothing being queued: a window with no file in
                    // it is a window that comes apart, and this file did not come apart — it
                    // came back from a player. The load is answered before it is handed over,
                    // so it is taken up on this tick rather than the next (see
                    // `show_pin_failure`).
                    //
                    // Not where something else is already on its way, which is the one
                    // condition: a load in hand is a read and a decode that have not been paid
                    // for yet, and putting the mark over it would throw that work away for a
                    // window that is about to be shown the file it was for.
                    if pin_walk.is_none() && pin_load.is_none() {
                        show_pin_failure(&failing, &mut pin_load);
                    }
                }

                // A step the caption's own walk buttons took is a pick like any other, and is
                // held in the walk rather than in the pick slot: the file it stands on is the
                // file to be shown, and the walk is what carries it on when that file turns out
                // to be one the pin cannot be shown (see `PinStep`). Reading the folder is the
                // planner's work now, so what comes back immediately is usually a wait rather
                // than a walk — the walk arrives as its own answer (see `step_pinned_file`).
                // A key a pinned window was given is answered here too, because the player it
                // acts on is this thread's.
                //
                // Whether the pin is still there is asked with `navigating` true wherever a
                // next file is already in hand, because a swap that has been refused leaves
                // nothing behind the window and the player behind that nothing is not a player
                // that has gone (see `pin_media_is_alive`). The mark is a load, so it is in
                // there for the one tick it is in hand; the tick after it, the media answers
                // for itself (see `MediaType::Unplayable`).
                // A key the hook took because the caret was in one of the pin's own two windows,
                // and would otherwise have reached FFmpeg rather than this app. The commands are
                // the same ones the pin's own window procedure queues, so they go on the same
                // queue: a walk of the folder is a walk of the folder whether the arrow arrived as
                // a message to this window or was swallowed before it could become one.
                for taken in crate::shell::key_input::take_pin_key_presses() {
                    if let Some(command) = hook_pin_key_command(taken) {
                        ask_pin(command);
                    }
                }

                // Whether a file is already on its way into the pin, which is what says the media
                // behind the window is not a window onto nothing (see `pin_media_is_alive`), and what
                // says an end of file has already been answered (see `loop_ended_pinned_player`).
                //
                // **The wait for a walk's answer is an advance in flight like any other**, and leaving
                // it out is what closed the pin of a film the media engine had just stopped at its end:
                // the session is let go when the walk is asked for, the swap that replaces it is a
                // folder read away, and a pin whose media is read as *gone* in between is a window that
                // comes down before the next file can land.
                let a_file_is_on_its_way = pin_walk.is_some()
                    || pin_load.is_some()
                    || pin_walk_wait.is_some()
                    || pin_swap_hold.is_some()
                    || pin_awaiting_box.is_some()
                    || pin_held_pick.is_some();

                if let Some(walk) = pin_command_request(
                    hwnd,
                    &mut pin_request,
                    &mut pin_walk_wait,
                    &mut audio_started,
                    &mut audio_start_offset,
                    &mut audio_paused,
                    a_file_is_on_its_way,
                    pin_walk_anchor(&pin_walk),
                ) {
                    pin_walk = Some(walk);
                }

                // The wait is over where its answer landed in this tick, and where it has run
                // for as long as this app will wait for a question it asked: a folder on a
                // share that never answers is a walk that ends rather than an arc that turns
                // for ever. The pin keeps the file it is showing either way (see
                // `PIN_JOB_GIVEUP`).
                //
                // This is read here rather than where the arc is turned because that is
                // before the drain below, where an answer for this walk would be found — a
                // wait given up on in the same tick its answer landed would take the walk
                // with it, and the pin would stop where a question had already been
                // answered.
                if pin_walk_wait
                    .as_ref()
                    .is_some_and(|wait| wait.started.elapsed() >= PIN_JOB_GIVEUP)
                {
                    pin_walk_wait = None;
                }

                // And a wait that is no longer a wait takes its arc down with it, once: a
                // pin left showing its old file is not left showing a spinner for a question
                // nobody is asking.
                if pin_walk_wait_seen && pin_walk_wait.is_none() {
                    pin_walk_wait_seen = false;
                    pin_arc_set(None);
                    render_layered_preview(hwnd);
                }

                // A window that is up holds nothing for the key: what a restore by a click on the
                // bubble shows is the file the window was showing, and a pick made behind a pin
                // nobody asked for one is not owed a swap later (see `pin_bubble_pick`).
                if !pin_is_collapsed() {
                    pin_bubble_pick = None;
                }

                // A press on the bar a sound's card is drawn with, which is the loop's to answer:
                // the clock that card is drawn from is this thread's, and so is a player of this
                // app's. Taken before the collapse below, so that a sound held and then put away
                // is put away at the second the hand asked for (see `settle_pinned_audio_seek`).
                settle_pinned_audio_seek(
                    &mut audio_started,
                    &mut audio_start_offset,
                    &mut audio_paused,
                );

                // The other two doors a card's own buttons open, answered in the same tick and for
                // the same reason: whether a player is going and what second it has got to are
                // this thread's, and so is the player a level is owed to. The play/pause is a
                // click on the card rather than a key, and the level is a knob let go of (see
                // `settle_pinned_audio_toggle` and `settle_pinned_audio_volume`).
                settle_pinned_audio_toggle(
                    &mut audio_started,
                    &mut audio_start_offset,
                    &mut audio_paused,
                );
                settle_pinned_audio_volume(
                    &mut audio_started,
                    &mut audio_start_offset,
                    &mut audio_paused,
                );

                // What a collapse into the bubble holds back and what a restore puts back,
                // read from the two `Pin Mode → Pause on Minimize` switches every tick: a switch
                // thrown while the pin is a bubble is answered on the next one (see
                // `settle_bubble_playback`).
                settle_bubble_playback(&mut audio_started, &mut audio_start_offset);

                // The bubble a collapsed pin leaves is made of what the pin's media has to show —
                // a frame, or the level a sound is being heard at — so it is painted again when any
                // of the three things it is made of has changed: the playback it stands for, a frame
                // grabbed for it out of a film FFmpeg plays, and the clock, while the bars it is
                // drawn with are being read from sound that is moving — a sound's own bands, or the
                // soundtrack of a film that is playing above a level of nothing (see
                // `bubble_bars_are_moving`). Nothing else asks for a bubble repaint: a bubble over a
                // picture is painted once, at the collapse, and a bubble that is not up is never
                // painted at all (see `BUBBLE_REPAINT`).
                //
                // The playback is asked of the file's own player rather than of anything this loop
                // wrote down, because a film paused by a key in its own window sends this thread
                // nothing: what says the bubble is out of date is the answer differing from the one
                // it was drawn with, which is the whole of what the comparison below is for.
                if pin_is_collapsed() {
                    let playing = bubble_is_playing();
                    let animating = playing && bubble_bars_are_moving();
                    let frame_ready = BUBBLE_FRAME_READY.swap(false, Ordering::AcqRel);
                    let changed = last_bubble_playing != Some(playing);
                    let due = animating && last_bubble_repaint.elapsed() >= BUBBLE_REPAINT;

                    // And the frames a film is being drawn by are read beside its player, which is
                    // the one picture this app has to *make*: neither external player can hand a frame
                    // over, and a film this app draws holds its frame here (see `bubble_stream`). The
                    // state of the world is what this is given — where the bubble is standing, and
                    // whether the film is playing — so that a level change, a seek, a loop and a swap
                    // behind the bubble are one answer rather than four hooks.
                    bubble_stream::settle(playing, bubble_box_on_screen());

                    if changed || due || frame_ready {
                        last_bubble_playing = Some(playing);
                        if due {
                            last_bubble_repaint = Instant::now();
                        }

                        // The bubble's own window, from this thread and no other: what is painted
                        // is a layered-window blit out of a surface this thread drew for exactly
                        // that window (see `repaint_pin_bubble`).
                        repaint_pin_bubble();
                    }
                } else {
                    // A pin with its window up has no bubble, and nothing of a bubble outlives the
                    // collapse it was made for: the play state is forgotten, the frame grabbed for it
                    // is dropped, and the frames read beside a player are ended, so a pin collapsed
                    // later starts on its own file's frame rather than on the one before it (see
                    // `BUBBLE_FRAME` and `bubble_stream`).
                    last_bubble_playing = None;
                    BUBBLE_FRAME_READY.store(false, Ordering::Release);
                    bubble_stream::settle(false, None);
                    if let Ok(mut frame) = BUBBLE_FRAME.lock() {
                        *frame = None;
                    }
                }

                // And the same for a bar that is drawn against a player this app started: what
                // it is allowed to claim is settled against the player that is actually there,
                // a player a relaunch has replaced is ended once its replacement has a window to
                // replace it with, and a hold a relaunch could not deliver is delivered now that
                // there is a window to deliver it through. All three are asked of the process
                // rather than of what this app wrote down, because everything a bar says about a
                // video FFmpeg plays is something this app asserted (see `settle_pinned_transport`).
                settle_pinned_transport();
                settle_video_retirement();
                // A pin taken up on a hover's film goes up once its hosted player draws, and comes
                // down where that player never does: the handover is the only visible swap, and
                // the teardown for a wait with nothing to swap to goes down the road a dead film
                // has always gone (see `pin_film_host::settle_pin_handover`).
                if matches!(
                    pin_film_host::settle_pin_handover(),
                    pin_film_host::PinHandover::Failed
                ) && pin_request.is_none()
                {
                    pin_request = Some(end_pin_state(Reason::MediaGone));
                }
                // And the band a drag's parking is holding: it is a hole again only once the
                // player's own window is standing in it, which for a relaunch begun on a resize's
                // release is the first tick that finds the replacement up (see
                // `settle_pinned_park`).
                settle_pinned_park();
                settle_pending_hold();

                // And what the taskbar button's own hover is offering, which is settled by the same
                // tick for the same reason the chrome is: a pin put away into the taskbar can be
                // walked and held from the hover, and what those buttons say is what the pin's
                // transport is doing (see `refresh_pin_media_buttons`). Nothing is sent while the
                // stand-in is not up, so this is two atomics and a comparison on a tick that has no
                // taskbar button to draw.
                //
                // And the form the button is in, which the Shell has gestures of its own for — a
                // second click on the button, `Win+M`, its own `Minimize` — so the pin's place is
                // what it is read back against, once a tick: a stand-in the Shell has minimized
                // under a pin that is on the screen is a hover frozen on the frame it was put down
                // with, and a stand-in the Shell has taken away altogether is a button the user has
                // lost until the app is restarted (see `keep_the_stand_in_up`). Nothing is asked of
                // the Shell here: the window is this app's own, and both reads are of it.
                keep_the_stand_in_up();

                // The picture the hover shows is the same two facts and is asked the same way: what
                // the pin is drawn by, mirrored into the stand-in by the Shell, which is a picture
                // that moves on its own and costs nothing here until what it is a picture of
                // changes (see `pin_taskbar_mirror`).
                let media_kind = current_media_type();
                refresh_pin_media_buttons(media_kind, bubble_is_playing());
                sync_pin_mirror(media_kind);

                // What a key does to a pinned text preview, polled rather than waited for: a pin
                // that the user has not pressed is a window nobody is in, and a window nobody is
                // in is never sent a keystroke. Ctrl+C is answered for every preview by the tick
                // below, which puts what is selected on the clipboard; Ctrl+A is the pin's own
                // answer, because selecting everything in the frame is a thing asked of a window
                // and not of a hover (see `pinned_select_all_requested`).
                if pinned_select_all_requested() {
                    select_all_text_preview(hwnd);
                }

                // A window dragged by an edge owes its media a layout at the box it ended up
                // with: the window procedure owns the drag and asks for it here, because the
                // media is this thread's.
                if pin_request.is_none() {
                    let requested = PIN_BOX_REQUEST
                        .lock()
                        .ok()
                        .and_then(|mut request| request.take());
                    if let Some(content) = requested {
                        pin_request = Some(PreviewMessage::PinBox(content));
                    }
                }

                // Where the player's window belongs while a pin is up: the media band of the
                // pin, which is what the tick's own re-assertion below is handed.
                if let Some(pinned) = pin_state() {
                    if let Some(pin) = pinned.pin() {
                        video_pos = (
                            pin.content.0,
                            pin.content.1,
                            pin.content.2 - pin.content.0,
                            pin.content.3 - pin.content.1,
                        );
                    }
                }

                // A transport bar's playhead moves while its file plays, so the window is painted
                // again at a clock's own pace rather than only where the picture changes.
                let transport_showing = pin_state()
                    .and_then(|pinned| pinned.pin().map(|pin| pin.transport_bar && !pin.collapsed))
                    .unwrap_or(false);

                if transport_showing
                    && last_pin_repaint.elapsed() >= Duration::from_millis(PIN_TRANSPORT_REPAINT_MS)
                {
                    last_pin_repaint = Instant::now();
                    render_layered_preview(hwnd);
                }

                // A window waiting for a file is painted again for the arc: once, when the
                // wait has run for the delay `spinner_delay_ms` names, and then once per turn
                // of it. A window with a transport bar would have been repainted anyway and
                // the arc rides along on that paint; one without has nothing else to repaint
                // for, and this is what a load that outlives the delay is answered with
                // rather than a window frozen for the length of the read (see `PinLoad`).
                if let Some(load) = pin_load.as_mut().filter(|load| load.arc.due()) {
                    // The arc is published as the wait comes due rather than as it starts, so
                    // a load that answers inside the delay never shows one — which is the
                    // whole of what the delay is for.
                    pin_arc_set(Some(load.arc.started.elapsed()));
                    load.arc.spun();

                    // The paint just happened, so the bar's own clock starts again here: a
                    // window with a transport bar would otherwise be painted twice within a
                    // turn, once for the arc and once for the playhead.
                    last_pin_repaint = Instant::now();

                    render_layered_preview(hwnd);
                }

                // A walk the planner is still reading the folder for is the same wait, painted
                // the same way and on the same delay: the pin is showing a file it has not
                // been asked for yet, and the arc is what says the question is still out
                // rather than a window that has simply stopped answering. The wait is read
                // here, where the loop knows what has been asked this tick; whether it is
                // still unanswered is settled further down, where the drain below is.
                if let Some(wait) = pin_walk_wait.as_mut().filter(|wait| wait.due()) {
                    pin_arc_set(Some(wait.started.elapsed()));
                    wait.spun();
                    pin_walk_wait_seen = true;

                    last_pin_repaint = Instant::now();

                    render_layered_preview(hwnd);
                }

                // A swap the pin is being held for is the same wait, painted the same way and on
                // the same arc: the file on screen is the one the pin already had, and what the
                // user is shown while the engine takes its first frame is that file with the
                // arc turning over it. It is the load's own arc carried across the hold rather
                // than a second one, so a swap that is held does not restart the spinner's
                // phase and the arc does not appear to stall where the hold began (see
                // `PinSwapHold`).
                if let Some(hold) = pin_swap_hold.as_mut().filter(|hold| hold.arc.due()) {
                    pin_arc_set(Some(hold.arc.started.elapsed()));
                    hold.arc.spun();

                    last_pin_repaint = Instant::now();

                    render_layered_preview(hwnd);
                }
            }

            // A player that has been started and has not put its window up yet is being
            // waited for: what is on screen is the spinner standing in for the video, at
            // the pointer it is shown at for every other kind of wait, and the wait ends
            // when the player's window is there — the frame the player was started for is
            // handed over then, and this app's window goes with the wait. A player that
            // never came up — one that died, one that took longer than a start ever does
            // — is ended here rather than left playing behind nothing (see `player_wait`).
            if let Some(start) = video_start.take() {
                let window_up = VIDEO_HWND.load(Ordering::SeqCst) != 0
                    && VIDEO_PID.load(Ordering::SeqCst) == start.pid;
                let alive = is_player_pid_alive(start.pid);

                // The hover this player was started for may have moved on before its
                // window was there. The player is still this app's — nobody else holds
                // it — so it is ended below, but what is on screen and what is pending
                // are another hover's by then and are left alone.
                let current = current_video_path.as_deref() == Some(start.path.as_path());
                let outcome = if current {
                    player_wait(window_up, alive, start.started.elapsed())
                } else {
                    Some(PlayerWait::Abandoned)
                };

                match outcome {
                    Some(PlayerWait::Arrived) => {
                        // The video is what is on screen from here: the frame the player
                        // plays into is the preview, and the spinner — this app's window,
                        // which was standing in for it — comes down with the wait.
                        if let Ok(mut media) = CURRENT_MEDIA.lock() {
                            *media = Some(start.media);
                        }

                        // And the film is playing from here, which is what the loop it may be
                        // giving is measured by: the clock was started when the process was
                        // (see `rebase_video_loop_clock`).
                        rebase_video_loop_clock(start.pid);

                        let _ = ShowWindow(hwnd, SW_HIDE);
                        pending_load = None;
                        clear_pointer_hold();
                    }
                    Some(PlayerWait::Abandoned) => {
                        // The player is not coming, or the hover it was started for has
                        // gone: ending it here is what keeps a process this app started
                        // from playing behind a preview that has gone.
                        let mut media = start.media;
                        stop_video_playback(&mut media);

                        if current {
                            let _ = ShowWindow(hwnd, SW_HIDE);
                            pending_load = None;
                            clear_pointer_hold();
                        }
                    }
                    // Still starting: the wait goes on, and the player stays in hand.
                    None => video_start = Some(start),
                }
            }

            // Periodically re-assert topmost on the video window to prevent it
            // from falling behind Explorer or other windows (Bug 2 fix)
            //
            // Split cadence: a pinned video competes with the pin's own window and keeps
            // the tight band, while a hover video gets the slow one — reordering DWM
            // 5x/s for a tooltip is what this used to cost (see `topmost_cadence_ms`).
            //
            // The volume popup is *not* guarded here any more: it is guarded inside the raise, which
            // is the one place a player's window is put on top and is reached from four others that
            // this guard did not cover (see `ensure_video_window_topmost`).
            let topmost_cadence_ms = topmost_cadence_ms(pinned());
            if current_video_path.is_some()
                && !pin_is_collapsed()
                && last_topmost_check.elapsed() >= Duration::from_millis(topmost_cadence_ms)
            {
                last_topmost_check = Instant::now();

                // A volume slider is drawn *into the pin's own window*, in the band where the
                // picture is, and the picture is a window of FFmpeg's own sitting above the pin in
                // the topmost band. So the slider is behind the video unless the pin is raised above
                // it — and the tick that used to raise the player skipped itself entirely while the
                // slider was open, on the reasoning that not re-raising the player was the same
                // thing as keeping it down. It is not: three other places raise the player (a drag,
                // a relaunch, a relayout), and one of them runs on the pointer move that opens the
                // slider's own knob. The symptom is a slider that appears for a frame and is then
                // behind the film again, which reads to the hand as a slider that closes by itself.
                //
                // So while the slider is up the pin is raised authoritatively, on every tick that
                // would otherwise have been spent raising the player. Both windows are topmost, so
                // this is a question of which of the two is nearer the front of one band, and the
                // answer is put in rather than left to whoever raised last.
                //
                // And a player parked for the length of a drag is not raised at all, which this arm
                // is where a hand resting on an edge is answered: the drag's own raises only run
                // while the pointer is moving, so a raise here is the whole of what a still hand
                // would be looking at — and it put the film back over a band being held flat (see
                // `pin_player_is_parked`). The slider is still raised over a parked player, because
                // it is drawn into the pin's own window and that window is the one thing on screen.
                if pin_volume_open() {
                    raise_pinned_window(HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _));
                } else if !pin_player_is_parked() {
                    let _ = ensure_video_window_topmost(
                        video_pos.0,
                        video_pos.1,
                        video_pos.2,
                        video_pos.3,
                    );
                }
            }

            // A player this app is looping itself is begun again at the beginning of its file rather
            // than being given `-loop 0` at the start, because the two cannot be asked for together:
            // with a `-ss` in the same command the player wraps back to the seek rather than to the
            // beginning, so an eight-second film begun at six plays its last two and a half seconds
            // for ever (measured — see `video_launch::loop_is_ours`). It rides the same tick as the
            // topmost re-assertion rather than a timer of its own, because both are questions about
            // the player that is on screen and this is the tick that already asks.
            //
            // The `playing` it is given is the pin's own answer rather than a guess, and it is what
            // holds the loop off a film a gesture has stopped: a hold that did not rebase the loop's
            // clock would keep counting up under the hand, and a film held for a minute near its end
            // would be past its end — and so rewound — the instant it was let go of (see
            // `video_loop_action`).
            //
            // **And it is asked of a pinned film as well as of a hovered one, and of the end of file that
            // film is played under rather than the one it was begun under.** A pin's file is a film
            // this app gives the loop to just as a hover's is, and for mpv — the one player a
            // take-up does not begin again — this tick is the only road its loop has: the clock was
            // noted for the player the *hover* began, so reading the end of file here is what lets
            // `End of File → Loop` loop a film the hover was playing once, and a pin's step or an
            // end of file that is `Pause` stop one the hover was looping (see `pin_video_loops`).
            //
            // **And an end of file that is `Pause` is a rewind too, held at the top it names** (see
            // `pin_video_pauses_at_the_top`): the same wrap a loop is given, so that the press which
            // plays the film again is the plain resume of a held film rather than a seek written
            // against a player that has stopped at the end of its file.
            if let Some(film) = current_video_path.clone().or_else(pinned_path) {
                let _ = video_loop_tick(
                    pin_is_playing_current(),
                    pin_video_loops(&film) || pin_video_pauses_at_the_top(&film),
                );
            }

            // A pin that still believes it holds the keyboard, and that Windows says does not, is
            // asked for it back. The player is begun again on every navigation and every resize, and
            // its window is created activatable and only styled once this app finds it — so an arrow
            // press can be answered by the player rather than by the pin, and FFmpeg's own arrow
            // keys are all seeks (see `pin_ask_keyboard_back`).
            if pinned() {
                let _ = pin_ask_keyboard_back(hwnd);
            }

            // A window being carried or pulled to a new size stops playing for the length of the
            // gesture and picks the film up where it left off on release. A video whose picture is
            // another program's window is re-scaled and re-presented on every one of the pointer's
            // messages, and doing that to a film that is also decoding is what makes a drag stutter;
            // every other kind of pinned window does the same work out of a surface it owns and is
            // not troubled (see `video_drag_hold`).
            //
            // **The gesture's own two ends settle this, so on a gesture of this window's own there
            // is nothing left for the tick to do** — it is here for the gestures that are not: one
            // begun off the hook's published press count is carried on from the pointer's position
            // rather than from a release message, so the tick is what finds it ended (see
            // `settle_pinned_engine_drag`).
            settle_video_drag_hold(pin_is_dragging());

            // Advance animation frames if needed
            let mut needs_repaint = false;
            // Whether this tick takes a full look at the media behind the preview: a
            // swap, a load, a wait or a player forces one, a dynamic hint keeps the fast
            // cadence, and a static hint still re-checks twice a second so late streaming
            // frames are noticed. Named rather than written out inline, because the backstop
            // is the arm most likely to be argued away and the one that cannot be argued away
            // on a machine where nothing happens to stream (see `needs_a_full_media_look`).
            let need_media_check = needs_a_full_media_look(
                media_dynamic_hint,
                pending_load.is_some()
                    || pin_load.is_some()
                    || pin_walk_wait.is_some()
                    || video_start.is_some()
                    || first_frame_wait.is_some(),
                current_generation != last_checked_generation,
                last_full_media_check.elapsed(),
            );

            // The engine draws a document in a window of its own, and that window is put
            // up only once the page has arrived: what is underneath it — the spinner the
            // wait was shown as — comes down then, and the wait comes down with it. What
            // is on screen is a document, and what is shown next is another hover's.
            //
            // What ends a wait is the document the wait is *for*: the engine is handed one
            // file at a time, and a page that lands names the file it was drawn for — so a
            // landing for a hover the loop has already left ends nothing, and the wait goes
            // on for the file the pointer is actually on. Taking any landing as the end of
            // any wait is what put a file the pointer had left on screen and dropped the
            // wait for the one it was on (see `webview_preview::showing_path`).
            if let Some(shown) = webview_preview::showing_path() {
                // The window comes down for the browser's, which draws the whole of what is on
                // screen — except while a preview is pinned, when what this window is drawing
                // is the caption around the browser's own window rather than the document.
                if IsWindowVisible(hwnd).as_bool() && !pinned() {
                    let _ = ShowWindow(hwnd, SW_HIDE);
                }

                let waiting_for_this = pending_load.as_ref().is_some_and(|pl| pl.path == shown);

                if waiting_for_this && pending_load.take().is_some() {
                    if let Some(cancel) = pending_load_cancel.take() {
                        cancel.store(true, Ordering::Release);
                    }
                    clear_pointer_hold();
                }
            }

            // Set inside the look below; read after it to refresh the hint.
            let mut saw_dynamic_this_tick = false;
            // A static tick skips the lock entirely (see `need_media_check`): the
            // guard is only taken when the tick is owed a full look, so a static
            // picture or page of text pays no mutex here at all.
            // A sound's card is painted from the loop's own clock and from the pin's own state, and the
            // pin's half is read here — before the media's lock is taken below rather than inside
            // it, because every repaint below runs with that lock held and a pin's lock asked from
            // under it is the re-entrancy `toggle_pinned_playback` is written down for (see
            // `pinned_audio_chrome`). Nothing at all is asked of a tick with nothing pinned: the
            // published flag answers that in one atomic read.
            let audio_chrome = pinned_audio_chrome(audio_paused);
            // Whether a sound is playing behind the card on screen, which is what its mark beats
            // with: a pinned card's own controls say so, and a hovered sound is the same question
            // put to the players behind it — asked here, beside the pin's own answer and for the
            // same reason, because it is a question that takes the media's lock (see
            // `a_sound_is_playing`). Nothing is playing behind a collapsed pin's bubble, which is
            // the one card this leaves out.
            let sound_playing = match audio_chrome {
                Some(chrome) => chrome.playing,
                None => !pinned() && a_sound_is_playing(),
            };

            let media_lock = need_media_check.then(|| CURRENT_MEDIA.lock());
            if let Some(Ok(mut media_guard)) = media_lock {
                if let Some(ref mut media) = *media_guard {
                    saw_dynamic_this_tick = media.frames.len() > 1
                        || media.is_streaming()
                        || media.media_type.is_native_video()
                        || media.media_type.is_audio()
                        || media.media_type.is_loading()
                        || media.should_draw_streaming_overlay();
                    if media.advance_frame() {
                        needs_repaint = true;
                    }
                    if media.update_loading_frame() {
                        needs_repaint = true;
                    }
                    // A video the media engine plays is a new picture every frame rather
                    // than a file that is decoded once and drawn again, so its frame is
                    // taken here — once a tick, which is as often as one can be shown.
                    // The kind is asked first so that a preview of any other kind pays
                    // for this with one comparison.
                    //
                    // And not while a swap is being held for a video's first frame. What is
                    // in `CURRENT_MEDIA` then is the file the pin is *still* showing, and
                    // the engine is drawing the other one: taking here would pull the new
                    // film's frames into the old file's buffer at the old file's size, which
                    // the window would then draw — a mis-scaled new video inside the old
                    // one, over the frame the hold exists to keep. The held file's own take
                    // is where those frames are wanted, and it is asked once a tick whether
                    // or not this one is (see `PinSwapHold`).
                    if media.media_type.is_native_video() && pin_swap_hold.is_none() {
                        if media.take_native_video_frame() {
                            needs_repaint = true;
                        } else if let Some(failing) = video_player::failing_path() {
                            // A file the engine has had long enough to have handed a
                            // frame over many times and has handed over none is a file the engine
                            // cannot draw: what the probe asked about was a decoder
                            // and a converter, which this file has, and what it cannot speak
                            // for is the pipeline the engine plays through. So the file is
                            // written down as one no player here will take — and because the
                            // engine only holds video files at all where nothing of FFmpeg's
                            // is installed, that is a file with no preview rather than one
                            // handed to another engine. The replay is asked for where a
                            // replay can be started (see `video_player::mark_unplayable`).
                            video_player::mark_unplayable(&failing);
                            engine_failed = Some(failing);
                        }
                    }
                    // A film the media engine plays for the pin ends the way a sound the engine plays
                    // does, and one place reads its end because the readings are one reading: the
                    // walk's next file where the end of file steps on, or the film taken home and held
                    // where it pauses — both of them a tenth of a second before the end of the pass,
                    // off the engine's own clock (see `pinned_native_video_end`).
                    if media.media_type.is_native_video() {
                        if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                            answer_native_pass_end(
                                pinned_native_video_end(path),
                                path,
                                pin_walk_anchor(&pin_walk),
                                &mut pin_walk_wait,
                            );
                        }
                    }
                    // A sound's card is the one painted preview that changes while it is on
                    // screen: the clock and the bar under it are drawn from a player that is
                    // running, a name the card has no room for is scrolled across it, and the
                    // mark beside the name beats with the sound's own bass. The clock's own
                    // seconds are worth watching four times a second and the other two are
                    // not, so a card with something moving on it is painted at the cadence the
                    // spinner's overlay uses and a card that is standing keeps the slower one
                    // (see `AUDIO_CARD_REPAINT` and `AUDIO_NAME_REPAINT`). What a card with no
                    // player behind it — a file nothing will play, at any level — costs is its scroll
                    // and nothing else.
                    if media.media_type.is_audio() {
                        // A sound that was asked to start somewhere other than the beginning
                        // and has not been taken there yet is taken there here, on the first
                        // tick the engine will accept a seek on — which is as soon as it has
                        // read the file's own header rather than the next card repaint a
                        // quarter of a second away, and what is heard of the beginning in
                        // between is that much and no more (see `video_player::apply_seek`).
                        video_player::apply_seek();

                        // A sound FFmpeg plays whose pass has ended is put round to the
                        // beginning of its file here, on the tick the player's own exit is
                        // found rather than on the next card repaint a quarter of a second
                        // away: what stands between two passes is the time it takes to start
                        // a player, and nothing of this side is to be added to it (see
                        // `wrap_audio_player`). A pinned sound whose end of file is `Continue`
                        // is not put round at all — its end asks for the next file of the
                        // folder there, with the walk's own wait — and one whose end of file is
                        // `Pause` stops where it is, the card left held at the beginning with its
                        // play button up.
                        if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                            wrap_audio_player(
                                media,
                                path,
                                &mut audio_started,
                                &mut audio_start_offset,
                                &mut audio_paused,
                                &mut pin_walk_wait,
                                pin_walk_anchor(&pin_walk),
                            );

                            // A sound the engine Windows has was asked to play the pinned
                            // file once where the pin's end of file is not `Loop`, so the end
                            // the engine reports for it is the end of the sound. Where that
                            // end is the next file of the folder, the session is let go — which
                            // is also what stops the end being read again on
                            // the sixty ticks a second this loop turns — and the walk is asked
                            // for, the same ask the caption's own **Next** button makes and the
                            // same wait for its answer. Where it is the card held instead — the
                            // end of file is `Pause` — the session is let go for the same reason and
                            // the card is left standing at the beginning of the file, which is
                            // what its play button now shows and what the press on it starts
                            // the sound again from (see `pinned_native_audio_end`).
                            match pinned_native_audio_end(path) {
                                Some(PinPassEnd::Walk) => {
                                    video_player::stop();
                                    ask_pin_walk(path.to_path_buf(), 1, pin_walk_anchor(&pin_walk));
                                    pin_walk_wait = Some(PinWait::new());
                                }
                                Some(PinPassEnd::Held) => {
                                    video_player::stop();
                                    audio_paused = Some(0.0);
                                    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
                                }
                                // Nothing at all is what a sound that has not ended gets,
                                // and a sound the engine was asked to loop has no end to read
                                // here either (see `pinned_native_audio_end`).
                                Some(PinPassEnd::Wrap) | None => {}
                            }
                        }

                        // The mark beats whenever a sound is playing behind the card — a pinned
                        // one whose player is running, and a hovered one the same way — which is
                        // the third thing a card can be moving by: read from the band the setting
                        // names, and not at all where it is `off` (see `sound_playing`,
                        // `current_beat_hz` and `audio_meter::output_beat`).
                        let beat_hz = if sound_playing {
                            current_beat_hz()
                        } else {
                            None
                        };
                        // A mark a stop left mid-beat is the fourth thing a card can be moving
                        // by: with nothing playing and nothing to read, the ripples are drawn by
                        // the fall until they are home, which is a second or so and then nothing
                        // (see `audio_meter::beat_fade`).
                        let falling = audio_beat > 0.0;
                        let cadence = match &audio_name_scroll {
                            Some(scroll) if scroll.moves() => AUDIO_NAME_REPAINT,
                            _ if sound_playing || falling => AUDIO_NAME_REPAINT,
                            _ => AUDIO_CARD_REPAINT,
                        };

                        if audio_repaint_at.elapsed() >= cadence
                            || AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel)
                        {
                            // How long a fall has been falling by: the time since the paint
                            // the mark was last drawn at, taken before that clock is reset
                            // (see `audio_meter::beat_fade`).
                            let since_paint = audio_repaint_at.elapsed();
                            audio_repaint_at = Instant::now();

                            let name_offset = match audio_name_scroll.as_mut() {
                                Some(scroll) => {
                                    scroll.advance(Instant::now(), cadence);
                                    scroll.offset()
                                }
                                None => 0,
                            };

                            if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                                let (elapsed, duration) = audio_clock(
                                    path,
                                    audio_started,
                                    audio_start_offset,
                                    audio_paused,
                                );

                                // A start position that was a share of a length nothing had
                                // read is asked for here, on the first tick a player says how
                                // long the file is — the ask is one that can be made at any
                                // point in a running sound, since what it is is a seek, and
                                // what it is not is a reason to have left the sound at its
                                // beginning (see `video_player::seek`). A player that says
                                // nothing about its length leaves the ask standing rather than
                                // spending it on a tick it cannot answer.
                                if let (Some(seek), Some(duration)) = (audio_share_seek, duration) {
                                    audio_share_seek = None;

                                    let shared =
                                        audio_seek::start_position(path, seek, Some(duration));
                                    video_player::seek(shared);
                                }

                                // Where the sound had got to is what the mode that resumes one
                                // reads back, so it is written down as the card is repainted —
                                // the only moment anything here knows it. The other three ways
                                // of starting a sound are rules rather than memories and are
                                // not written down at all: a run that is on one of them keeps
                                // nothing, which is what makes the memory the setting's own —
                                // and the setting is the one the sound on screen plays under:
                                // the pin's for a pinned file, the hover's for a hovered one
                                // (see `position_is_remembered`).
                                if position_is_remembered(pinned()) {
                                    if let Some(elapsed) = elapsed {
                                        audio_seek::remember(path, elapsed);
                                    }
                                }

                                // The beat is read from the capture only while a sound is playing
                                // behind the card, whichever card it is: a card with no player
                                // behind it has no pulse to read and pays nothing for the feature
                                // at all (see `audio_meter::output_beat`). A mark a stop left
                                // mid-beat is drawn by the fall instead, timed by the repaint
                                // happening now and counted a step at a time, so a frame the loop
                                // came back late for is a fall still seen coming home and not a
                                // jump to the dot (see `audio_meter::beat_step` and
                                // `audio_meter::beat_fade`).
                                let beat = match beat_hz {
                                    Some(band) => audio_meter::output_beat(band).unwrap_or(0.0),
                                    None if falling => audio_meter::beat_fade(
                                        audio_beat,
                                        audio_meter::beat_step(since_paint),
                                    ),
                                    None => 0.0,
                                };
                                // What the next fall — or the pause that starts one — stands on.
                                audio_beat = beat;

                                if media.refresh_audio_card(
                                    path,
                                    elapsed,
                                    duration,
                                    audio_card_dpi,
                                    name_offset,
                                    audio_chrome,
                                    beat,
                                ) {
                                    needs_repaint = true;
                                }
                            }
                        }
                    }
                    // While streaming first-frame loading, repaint for spinner animation.
                    if media.should_draw_streaming_overlay()
                        && last_stream_overlay_repaint.elapsed() >= Duration::from_millis(83)
                    {
                        last_stream_overlay_repaint = Instant::now();
                        needs_repaint = true;
                    }
                }
            }
            if need_media_check {
                last_full_media_check = Instant::now();
                last_checked_generation = current_generation;
                media_dynamic_hint = saw_dynamic_this_tick
                    || pending_load.is_some()
                    || pin_load.is_some()
                    || pin_walk_wait.is_some()
                    || video_start.is_some()
                    || first_frame_wait.is_some();
            }
            // Whether a video's first frame comes up on this tick, and where: the question is
            // asked before the paint rather than inside the arm that acts on it, because it is
            // what decides whether the paint happens at all. A tick whose repaint is the reveal's
            // own is not painted twice — same frame, same surface, and thirty megabytes of it at
            // the size of a 4K display (see `first_frame_lands`).
            //
            // The hide count is read once for both this and the arm below, so the two cannot be
            // answering about different hovers because a hide landed between them.
            let epoch = hidden_epoch();
            let first_frame_landed = first_frame_wait.as_ref().and_then(|wait| {
                first_frame_lands(
                    wait,
                    current_generation,
                    epoch,
                    pinned(),
                    media_holds_a_frame(),
                )
            });

            if needs_repaint && first_frame_landed.is_none() {
                render_layered_preview(hwnd);
            }

            // A video the engine plays whose preview was held back for its first frame: the
            // engine has handed one over — the repaint above is of it, or the one below is — and
            // what is left of the wait is the window, which is put up here if it is not up
            // already. A window that is on screen needs nothing of this: what it was holding was
            // a wait's frame or another preview's, and the frame that landed is drawn at the
            // place this hover was laid out at rather than at that window's own (see
            // `FirstFrameWait`).
            if let Some(wait) = first_frame_wait {
                if wait.generation != current_generation || wait.epoch != epoch {
                    first_frame_wait = None;
                } else if media_holds_a_frame() {
                    first_frame_wait = None;

                    // The paint the question above already accounted for: this arm's own paint
                    // is at the box this hover was laid out at, and the tick's was at the
                    // window's, so where the two differ it is this one that has to be the
                    // last (see `first_frame_lands`). The question answered that *this* tick
                    // is the reveal's own — a frame in hand on this hover's wait, no pin up —
                    // so it is this paint that runs and the tick's that stands down for it.
                    // The window below is put up on the strength of this paint, and a reveal
                    // that showed it without one would put the window up on the surface it is
                    // still holding, which for a hover that followed another preview is that
                    // preview's frame, one tick early.
                    if !pinned() && first_frame_landed.is_some() {
                        render_layered_preview_at(hwnd, wait.pos.0, wait.pos.1);
                    }

                    if !IsWindowVisible(hwnd).as_bool() {
                        let _ = SetWindowPos(
                            hwnd,
                            HWND_TOPMOST,
                            0,
                            0,
                            0,
                            0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
                        );
                        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                    }
                }
            }

            // A page an engine has finished with, held apart from the hovers: it is
            // not a hover to act on but an answer about the one on screen. Office's
            // tier and the engines beside it send it as a message, a page the render
            // engine draws is read from the folder it lands in — below — and a load
            // that comes back with nothing to draw because the engine has already
            // finished the work has the answer in hand rather than in a message (see
            // the `awaiting_render` arm of the load below).
            let mut page_ready: Option<(PathBuf, u64, bool)> = None;

            // Check for completed background loads
            while let Ok(result) = load_rx.try_recv() {
                // A load the pointer has left since it was started is not this
                // loop's to take up: the hide that took the window down moved the
                // count it carries (see `HIDDEN_EPOCH`), and a pointer that has
                // crossed to another item since the hover was resolved is the same
                // answer one moment earlier, before any hide has been sent for it
                // (see `HOVER_POINTER_BOX`). Either way the answer is dropped here
                // rather than built — no frame installed, no engine window put up,
                // no player started — for a hover that has already gone.
                //
                // What the wait itself put on screen comes down with it, and that is not
                // the same thing as dropping the answer: a spinner left standing is a wait
                // with nothing left to end it — the load that would have answered it is
                // gone from here, and the pointer cannot dismiss it either, because a wait
                // holds the pointer it is waiting for (see `WAITING_PREVIEW_HOLDING`) and
                // the pointer sitting on that spinner is the one thing the hold refuses. A
                // hand that crossed off the file while the load ran would be left with a
                // preview under it that nothing would ever take down.
                //
                // Asked first is whether this is the hover's own load: a keyboard hover
                // carries no placement, a newer hover is another generation, and a stale
                // answer for a load already dropped is nothing to take down twice
                // (see `HoverPlacement`).
                if pending_load
                    .as_ref()
                    .is_some_and(|pl| pl.hide_epoch != hidden_epoch())
                    || !pointer_on_the_hovered_item()
                {
                    let abandoned = pending_load.as_ref().is_some_and(|pl| {
                        pl.generation == result.generation && pl.placement.is_some()
                    });

                    pending_load = None;

                    if abandoned {
                        // The take-down the `Hide` message performs, for the same reason and
                        // in the same order: the wait is over, so its window, its media, the
                        // hold it published and anything it had asked for go with it.
                        current_generation += 1;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        let _ = ShowWindow(hwnd, SW_HIDE);
                        // The pin's own window, where it stands in one, goes down with it.
                        pin_real_window::hide_real_pin_window();
                        clear_pointer_hold();
                        webview_preview::hide();

                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            if let Some(ref mut media) = *current {
                                media.cancel_background_work();
                                stop_video_playback(media);
                            }
                            *current = None;
                        }
                        current_video_path = None;
                        video_pos = (0, 0, 0, 0);

                        if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                            office_render::hover_ended(path);
                        }

                        current_show = None;
                        page_render_pending = None;
                        page_upgrade = None;
                    } else {
                        pending_load_cancel = None;
                    }

                    continue;
                }

                if result.generation == current_generation {
                    match result.media {
                        // A document the engine draws. There is no frame of this app's to
                        // install and no window of this app's to put up, so what happens
                        // here is the handover: the engine is told where the wait has
                        // ended up, and the wait stays armed for the document to land on.
                        Some(media) if media.media_type.is_engine() => {
                            let pending = pending_load.take();
                            pending_load_cancel = None;

                            if let Some(pl) = pending.as_ref() {
                                webview_preview::show(
                                    &pl.path,
                                    webview_preview::Area {
                                        x: pl.pos_x,
                                        y: pl.pos_y,
                                        width: pl.width as i32,
                                        height: pl.height as i32,
                                    },
                                    // The engine draws two kinds — a document and a font
                                    // specimen — and each is composited over a backdrop of
                                    // its own rather than over the picture's.
                                    engine_background(&pl.path),
                                );
                            }

                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut existing) = *current {
                                    existing.cancel_background_work();
                                }
                                // Nothing of this app's goes on screen for a document:
                                // what is up is the engine's window, and what stands in
                                // for it until that arrives is the spinner.
                                *current = None;
                            }

                            // The wait stays armed for the document to land on: an
                            // engine that has to start a browser is a wait like any
                            // other, shown as one once the delay has run, while one
                            // that draws the document in a few milliseconds is over
                            // before there is anything to show (see `spinner_due`).
                            pending_load = pending;
                        }
                        Some(mut media_data) => {
                            // A video the media engine plays is started here, before its
                            // preview is put up: what this window is about to draw is a
                            // frame of it, and an engine that would not start is a file
                            // with no preview rather than a box of the placeholder pixels
                            // a video preview is opened with.
                            let native_video = media_data.media_type.is_native_video();

                            if native_video {
                                let (width, height) =
                                    (media_data.current_width(), media_data.current_height());

                                // A hover that lands on the file already playing leaves it
                                // playing, the same way the FFmpeg path compares the file
                                // it last started.
                                if video_player::playing_path().as_deref()
                                    == Some(result.path.as_path())
                                    && video_player::is_playing()
                                {
                                    video_player::resize(width, height);
                                } else {
                                    video_player::play(
                                        &result.path,
                                        width,
                                        height,
                                        current_video_volume(),
                                        probed_picture(&result.path, width, height),
                                        pin_video_loops(&result.path),
                                    );
                                }

                                if !video_player::is_playing() {
                                    let _ = ShowWindow(hwnd, SW_HIDE);
                                    clear_pointer_hold();
                                    pending_load = None;

                                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                        if let Some(ref mut existing) = *current {
                                            existing.cancel_background_work();
                                        }
                                        *current = None;
                                    }

                                    continue;
                                }

                                // A session that is already playing — the hover is back on
                                // the file, or is a preview of it taken down and asked for
                                // again — has a frame in hand, and it is taken here rather
                                // than waited a tick for: what the window is put up with is
                                // a frame of the file, and an engine that has one has it
                                // now. A session that has just been asked to play has none,
                                // and the reveal below is held for it (see `FirstFrameWait`).
                                media_data.take_native_video_frame();
                            }

                            // Whether this preview has a frame of the video in it: what the
                            // engine had in hand was taken above, and a preview holding none
                            // is one nothing is put up for yet.
                            let video_frame_in_hand =
                                native_video && media_data.current_frame_is_opaque();

                            // A sound is started here, before its card goes up, for the reason
                            // a video's engine is: what the card draws is the clock of a player
                            // that is running. A player that was asked for and did not come up is
                            // a sound with nothing behind it, which is the answer a video's engine
                            // that will not start gets.
                            if media_data.media_type.is_audio() {
                                // Where the sound is dropped in: a question about the file's
                                // own length and the tray's `Volume → Audio Seek`, and one that
                                // is answered here rather than by the player, which knows
                                // neither. A length nothing has read yet — a container that
                                // does not say, on a machine whose engine may still know it —
                                // leaves the two shares of one unanswered until a player
                                // reports one, which the tick below is what asks again.
                                let seek = current_audio_seek();
                                let length = audio_track::playable(&result.path)
                                    .and_then(|track| track.duration);
                                let start = audio_seek::start_position(&result.path, seek, length);

                                audio_share_seek = (start == 0.0
                                    && matches!(seek, AudioSeek::Middle | AudioSeek::Random)
                                    && length.is_none())
                                .then_some(seek);

                                if !start_audio_playback(&result.path, &mut media_data, start) {
                                    let _ = ShowWindow(hwnd, SW_HIDE);
                                    clear_pointer_hold();
                                    pending_load = None;

                                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                        if let Some(ref mut existing) = *current {
                                            existing.cancel_background_work();
                                        }
                                        *current = None;
                                    }

                                    continue;
                                }

                                // The clock a sound FFmpeg plays is this app's own over the
                                // moment the player was started, and there is a player to
                                // measure from exactly where one was started: a decoder that
                                // would not have the file never got one, and a card whose clock
                                // ran anyway would be a sound it says is playing that is not — and
                                // a position this side would write down as one the file had
                                // been left at (see `audio_seek::remember`). A file that has
                                // just been started is not one a key has held, whatever the
                                // file before it was doing.
                                audio_started =
                                    media_data.video_process.is_some().then(Instant::now);
                                audio_start_offset = start;
                                audio_paused = None;
                                audio_repaint_at = Instant::now();
                                // A card that has just been put up is at rest, whatever the
                                // mark of the card before it was doing.
                                audio_beat = 0.0;
                                // The marquee the card's name is drawn with, if it needs one:
                                // what a name is scrolled by is the card's own box, which is
                                // the frame that has just arrived, and a scroll is put up with
                                // the card rather than left over from the hover before it (see
                                // `audio_preview::NameScroll`).
                                audio_name_scroll = Some(audio_preview::NameScroll::of(
                                    &audio_preview::name_of(&result.path),
                                    media_data.current_width(),
                                    audio_card_dpi,
                                    current_audio_options(),
                                ));
                            }

                            let mw = media_data.current_width() as i32;
                            let mh = media_data.current_height() as i32;

                            // A load whose spinner was up is placed like any other:
                            // the wait was shown in the spinner's own box at the
                            // pointer, not in the box the preview arrives in (see
                            // `spinner_pos`), so the frame is installed at the
                            // preview's place — the paint below is what takes the
                            // window from one box to the other.
                            let mut pending = pending_load
                                .take()
                                .filter(|pl| pl.generation == result.generation);
                            pending_load_cancel = None;

                            // Placed once more before it is painted, from the pointer as it is
                            // now: the box this frame lands in was decided when the hover was
                            // asked for, and the hand has had the whole load to move on since.
                            // A preview that arrives under the hand is taken down again by the
                            // touch rule at the next tick — the spawn and the dismissal are one
                            // event for the eye — while one that lands clear of it stays. This
                            // is the placement every tick of the wait already makes (see
                            // `PendingLoad::follow_pointer`), one read before the paint, so what
                            // is painted is the preview of the file under the hand rather than
                            // one under the hand itself. A keyboard hover carries no placement
                            // and is left where the item put it.
                            if let Some(pl) = pending.as_mut() {
                                if let Some(cursor) = cursor_position() {
                                    pl.follow_pointer(cursor, dpi_at(cursor.x, cursor.y));
                                }
                            }

                            // What this load was planned for: the box the layout
                            // came out with, which is the size a slide is
                            // exported at when a page is asked for later.
                            let render_box =
                                pending.as_ref().map(|pl| (pl.width, pl.height)).unwrap_or((
                                    media_data.current_width(),
                                    media_data.current_height(),
                                ));

                            // Move before installing the frame. Crossing between
                            // displays of different scale sends WM_DPICHANGED,
                            // which resets the preview and would otherwise
                            // discard the frame we are about to show, leaving the
                            // other display's image stranded on screen.
                            //
                            // A window that is already on screen is not moved
                            // ahead of the frame, though: a layered window shows
                            // the surface it has at whatever size the window has,
                            // so resizing this one to the page's box before there
                            // is a page to fill it draws the frame it is holding —
                            // the spinner — stretched across that box for as long
                            // as the paint takes. It is moved and resized by the
                            // paint itself, while a hidden window is moved here,
                            // where nothing can show.
                            let visible = IsWindowVisible(hwnd).as_bool();
                            if let Some(ref pl) = pending {
                                if !visible {
                                    let _ = MoveWindow(hwnd, pl.pos_x, pl.pos_y, mw, mh, false);
                                }
                            }

                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut existing) = *current {
                                    existing.cancel_background_work();
                                }
                                *current = Some(media_data);
                            }

                            // A layered window keeps its surface while hidden, so
                            // paint the new frame before revealing the window.
                            // Showing first would flash the previous preview at
                            // the new position and size.
                            //
                            // A video the engine plays is not painted at all until
                            // it has a frame of the file: what its load landed with
                            // is the placeholder frame every video preview is
                            // loaded with, and the window is held back rather than
                            // put up on it — a placeholder is not a picture of
                            // anything, and drawn over the backdrop the tray keeps
                            // for pictures it *is* that backdrop, which is what is
                            // seen as a flash of it at the start of a hover. The
                            // tick that takes the first frame is what puts this
                            // preview up, at the place its layout came out at (see
                            // `FirstFrameWait`).
                            //
                            // The frame and the window it goes into are written
                            // under the hide count's own lock, so a hide cannot
                            // land between the two — and a load the pointer has
                            // left in the meantime is neither painted nor shown
                            // (see `HIDDEN_EPOCH`).
                            {
                                let hidden = HIDDEN_EPOCH.lock().ok();
                                let wanted = pending
                                    .as_ref()
                                    .map(|pl| hover_still_wanted(&hidden, pl))
                                    .unwrap_or(true);

                                if wanted && native_video && !video_frame_in_hand {
                                    first_frame_wait = pending.as_ref().map(|pl| FirstFrameWait {
                                        generation: result.generation,
                                        epoch: pl.hide_epoch,
                                        pos: (pl.pos_x, pl.pos_y),
                                    });
                                } else if wanted {
                                    match pending.as_ref().filter(|_| visible) {
                                        Some(pl) => {
                                            render_layered_preview_at(hwnd, pl.pos_x, pl.pos_y)
                                        }
                                        None => render_layered_preview(hwnd),
                                    }

                                    if let Some(pl) = pending {
                                        let _ = SetWindowPos(
                                            hwnd,
                                            HWND_TOPMOST,
                                            pl.pos_x,
                                            pl.pos_y,
                                            mw,
                                            mh,
                                            SWP_NOACTIVATE | SWP_SHOWWINDOW,
                                        );
                                        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                    }
                                }
                            }

                            // A page for this document is one Office start away,
                            // so it is asked for as soon as the hover is up rather
                            // than after the pointer has rested on it.
                            page_render_pending = request_office_render(
                                &result.path,
                                result.generation,
                                render_box.0,
                                render_box.1,
                            );
                        }
                        None if result.awaiting_render => {
                            // Nothing to draw yet and a page on the way: the
                            // pending load stays armed, so the preview is not
                            // dropped and the page has somewhere to land — and the
                            // wait it is given is the one every other kind of wait
                            // gets, so the spinner goes up once the delay has run
                            // (see `spinner_due`). What is asked for is the room the
                            // display has rather than the box this hover's layout
                            // came out at, which for a wait is the spinner's own box
                            // at the pointer and says nothing about how large the
                            // page will be drawn: a slide is exported at the width
                            // the render is asked for, so the room the display has is
                            // the sharpest page that display can show (see
                            // `PendingLoad::room`).
                            let (width, height) = pending_load
                                .as_ref()
                                .map(|pl| pl.room)
                                .unwrap_or_else(|| office_formats::default_page_size(&result.path));

                            // The wait is on an engine from here, whatever was asked
                            // for below: what the loader could not draw is a page, a
                            // picture or a listing that only an engine can produce, and
                            // that is what the cap on waiting is read against — a wait
                            // whose request is answered as part of another's has nothing
                            // else that would ever end it (see `awaiting_engine`).
                            if let Some(pl) = pending_load.as_mut() {
                                pl.awaiting_engine = true;
                            }

                            // A document the render engine draws is asked for the same
                            // way, and in the same breath: neither page exists until an
                            // engine has drawn it, and this is the one hover that is
                            // waiting for one.
                            //
                            // And a picture the image converter develops is the same wait
                            // once more — the engine's answer arrives as a message rather
                            // than as a page in a folder, and what it writes is a picture
                            // at the size it is shown rather than at the size of the file,
                            // which is what makes its room a ceiling rather than a hint: a
                            // picture developed into a box smaller than the display can
                            // never be drawn any larger than that box.
                            // The engines this hover could be owed a page by, asked in the order
                            // the file's kind names them and only where one of them can answer:
                            // whichever starts work is the one being waited on, and the loop
                            // watches for the page or the picture or the listing it will leave
                            // (see `request_engine_render`).
                            let requested = request_engine_render(
                                &result.path,
                                result.generation,
                                (width, height),
                            );

                            // Nothing was asked for because there is nothing left to ask
                            // about: every engine that could owe this file something has
                            // already produced it. The page, the picture, the listing is
                            // in hand — it landed between the load that came back without
                            // it and this check, which is the one moment the two questions
                            // can disagree about — so the answer is read by having the
                            // file loaded again, and the replay below (`page_upgrade` and
                            // all) is the one a page arriving as a message is given. Left
                            // as a wait with nothing behind it, the hover would be a
                            // spinner that no answer and no cap could ever take down.
                            //
                            // A file the engine turned down while the load ran is answered
                            // by that same replay and by nothing else: the load that reads
                            // it finds no page and no wait to be in, which is the branch
                            // that takes the spinner down rather than leaving it up.
                            if requested.is_none() {
                                page_ready = Some((result.path.clone(), result.generation, true));
                            }

                            page_render_pending = requested;
                        }
                        None => {
                            // Loading failed, which is the answer the mark is for: the wait the
                            // file was in ends as the cross every failed file ends as, in the box
                            // the wait had, rather than by vanishing (see
                            // `take_the_wait_as_failed`).
                            take_the_wait_as_failed(hwnd, &mut pending_load);
                            clear_pointer_hold();
                            pending_load_cancel = None;
                        }
                    }
                }
            }

            // A preview that is still on its way follows the pointer: while a
            // load runs the spinner is the only thing on screen, and a cursor
            // that moves along the item the preview belongs to would otherwise
            // leave it behind. Nothing is measured again — the hover's own size
            // and `Avoid` region are what it is placed with — so this
            // costs a cursor read and a placement per tick, and the window is
            // moved only when the place it comes out at has changed.
            //
            // The spinner follows at its own place rather than the preview's —
            // the arc at the pointer's corner, whatever box the preview will
            // arrive in — so what the hand sees while it waits is the wait at
            // the hand (see `waiting_placement`).
            if let Some(ref mut pl) = pending_load {
                if let Some(cursor) = cursor_position() {
                    let side_before = pl.spinner_side;
                    let dpi = dpi_at(cursor.x, cursor.y);
                    let followed = pl.follow_pointer(cursor, dpi);
                    if followed.spinner && pl.spinner_shown {
                        if pl.spinner_side == side_before {
                            let _ = MoveWindow(
                                hwnd,
                                pl.spinner_pos.0,
                                pl.spinner_pos.1,
                                pl.spinner_side as i32,
                                pl.spinner_side as i32,
                                false,
                            );
                        } else {
                            // The spinner's own box changed size, so its frame is
                            // drawn again at the size it now goes into.
                            show_loading_spinner(hwnd, pl);
                        }
                    }

                    // A wait for an engine-drawn preview takes the engine's window with it:
                    // what the engine is asked for is the box the wait has ended up in, and a
                    // box that moved while the document was on its way is the same document at
                    // the place the hand is — the want is moved rather than the engine being
                    // asked again, which is what keeps a moving pointer from asking for the
                    // same page sixty times a second, and what lets a document that arrives
                    // before its spinner was due land at the hand anyway. Nothing is moved for
                    // an engine that already has it up: that is the preview itself, and it
                    // follows nothing.
                    if followed.preview
                        && engine_kind_of(&pl.path).is_some()
                        && !webview_preview::is_showing()
                    {
                        webview_preview::wanted_here(
                            &pl.path,
                            webview_preview::Area {
                                x: pl.pos_x,
                                y: pl.pos_y,
                                width: pl.width as i32,
                                height: pl.height as i32,
                            },
                        );
                    }
                }
            }

            // Show the loading spinner while a background load runs, once the wait
            // is worth showing: a load that may be about to finish is given the
            // delay `spinner_delay_ms` names (see `spinner_due`).
            if let Some(ref mut pl) = pending_load {
                if pl.spinner_due() {
                    pl.spinner_shown = true;
                    show_loading_spinner(hwnd, pl);
                }
            }

            // And the mark a hover that came to nothing ended as comes down after its own three
            // seconds: it is an answer about a file rather than a preview of it, so nothing moves
            // it along and nothing else would ever take it down (see `failure_mark_due`).
            if failure_mark_due() {
                let _ = ShowWindow(hwnd, SW_HIDE);
                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                    *current = None;
                }
            }

            // A wheel notch over a scrollable text preview belongs to the preview:
            // the wheel hook swallowed it so Explorer does not scroll, and left
            // the ticks here. Rotating the wheel forward scrolls back towards the
            // start of the document, which is the opposite of the sign the
            // message carries.
            let scroll_delta = wheel_input::take_text_scroll_delta();
            if scroll_delta != 0 {
                let notches = (scroll_delta / WHEEL_DELTA) as i64;
                if notches != 0 {
                    let lines = -notches * TEXT_SCROLL_LINES_PER_NOTCH;
                    if let Some(first_line) = text_scroll_target(lines) {
                        scroll_text_preview(hwnd, first_line);
                    }
                }
            }

            // And a selection drag carried past the edge of the box keeps selecting: the page is
            // taken a step the way the drag left by, and the selection grows to the line at the
            // edge it came back with (see `step_text_drag_selection`).
            step_text_drag_selection(hwnd);

            // And a notch turned with `Ctrl` held over a pinned page is that page's own size: the
            // page is laid out again at the step the notches have reached, in the box it already has
            // (see `step_pinned_text_size`).
            let size_delta = wheel_input::take_text_size_delta();
            if size_delta != 0 {
                step_pinned_text_size(hwnd, size_delta.signum());
            }

            // And a notch over a pinned picture is the picture being zoomed: the same hook swallowed it
            // and left the count here. What the notches come to is written down as the size the picture
            // is *going* to rather than the size it is given the instant they are counted, and the
            // travel to it is taken a frame at a time below — a zoom that landed on its size the moment
            // a notch was turned would be a zoom that jumped (see `aim_pin_zoom` and `pin_zoom_step`).
            //
            // The count is in the wheel's own units — 120 to a notch — and it is divided as a fraction
            // rather than truncated to whole notches, because a high-resolution wheel reports less than
            // a notch per message and a truncated count is most of its travel thrown away (see
            // `notch_step`).
            let zoom_delta = wheel_input::take_pin_zoom_delta();
            if zoom_delta != 0 {
                aim_pin_zoom(zoom_delta as f32 / WHEEL_DELTA as f32);
            }

            // And where that travel has got to by now. The clock is read every tick rather than only
            // while there is a travel, so that the frame a notch's own tick takes is measured from the
            // tick before it and not from whenever the last zoom happened to be.
            let zoom_elapsed = last_zoom_frame.elapsed().as_secs_f32();
            last_zoom_frame = Instant::now();
            step_pin_zoom(zoom_elapsed);

            // And the frame it is owed. A picture the wheel has taken past its window is a frame the
            // window is holding scaled up, so the file itself is asked for at the size the picture has
            // settled at — once per size and not on every frame of a travel, which would be a decode
            // per frame, and only once the wheel has stopped asking, since a decode of a box the hand
            // is still passing through is one nothing is looking at by the time it lands (see
            // `pin_zoom_oversize_box`).
            if pin_zoom_in_flight() {
                zoom_sharpened = None;
            } else if let Some(drawn) = pin_zoom_oversize_box() {
                if zoom_sharpened != Some(drawn) {
                    if let Some((path, dpi)) = pinned_media_owner() {
                        zoom_sharpened = Some(drawn);
                        pin_relayout =
                            relayout_pinned_media(&path, drawn, dpi, None, PinRelayoutRoad::Zoom);
                    }
                }
            } else {
                // Nothing is owed a decode: no pin at all, a kind the wheel does not zoom, or a
                // picture that fills its window — so what was asked for last is not a size this pin is
                // standing at any more, and the next one that is will be asked for afresh.
                zoom_sharpened = None;
            }

            // Check for our custom messages. Only the newest hover target matters;
            // collapse stale Show/Hide traffic so we do not spend time computing
            // layouts for files the cursor has already left.
            let mut latest_preview_msg: Option<PreviewMessage> = None;
            let mut refresh_requested = false;
            // The newest file the user has picked while a preview was pinned, if any arrived this
            // tick: what the loop answers by showing the pin that file instead of the one it has
            // (see `PinUpdate`).
            let mut pin_pick: Option<PathBuf> = None;
            // A probe's answer, held apart the same way and for the same reason.
            let mut video_probed: Option<(PathBuf, u64)> = None;
            // And a measure's, which is held apart with the box it answered with: what is
            // waiting on it is a hover, and the box is what that hover is replayed with (see
            // `MeasureProbed`).
            let mut measure_probed: Option<(PathBuf, Option<(u32, u32)>)> = None;
            // What a carry held comes next, oldest first: a pin asked to close while it was
            // being dragged is closed by this tick rather than sitting in the channel until the
            // next one happens to read it.
            let mut next_preview_msg = carried_preview_msg.take();
            if next_preview_msg.is_none() {
                next_preview_msg = if carried_preview_messages.is_empty() {
                    None
                } else {
                    Some(carried_preview_messages.remove(0))
                };
            }
            while let Some(preview_msg) = next_preview_msg.or_else(|| rx.try_recv().ok()) {
                next_preview_msg = None;

                match preview_msg {
                    PreviewMessage::Refresh => {
                        if latest_preview_msg.is_none() {
                            // A painted preview's colors are in its frame — the
                            // colors and glyphs of a text preview and the theme of
                            // an archive listing alike — so a theme switch rebuilds
                            // it from the hover it came from; every other preview
                            // only needs the frame composited again.
                            match (pinned(), current_media_is_painted(), current_show.clone()) {
                                // A pin is not rebuilt from the hover it came from: what is on
                                // screen is a window of this app's own, and what a setting
                                // changed under it is owed is the repaint that composes its
                                // band and its chrome again (see `render_layered_preview`).
                                // The record of what the pin is showing is a `Show` like a
                                // hover's, and taken as one it would put a hover up where the
                                // pin was (see the gate before the match below).
                                //
                                // A page of text is the one kind whose settings are in its own
                                // frame — the theme its glyphs are painted with and the Markdown
                                // it is laid out as — so it is drawn again rather than composed
                                // again, at the size the reader put it to (see
                                // `redraw_pinned_page`).
                                (true, ..) => {
                                    redraw_pinned_page(hwnd, None);
                                    refresh_requested = true
                                }
                                (_, true, Some(show)) => latest_preview_msg = Some(show),
                                // A page the engine was drawing is a window of the engine's
                                // with no media of this app's behind it, so a switch that
                                // stops the file being one the engine draws leaves nothing
                                // to recomposite: the hover is rebuilt from itself instead
                                // (see `refresh_render_html`).
                                (_, false, Some(show))
                                    if current_media_kind().is_none()
                                        && show_path(&show)
                                            .and_then(|path| engine_kind_of(path))
                                            .is_none() =>
                                {
                                    latest_preview_msg = Some(show)
                                }
                                _ => refresh_requested = true,
                            }
                        }
                    }
                    PreviewMessage::RefreshTypes => {
                        // A preview of a kind that was switched off is rebuilt
                        // from the hover it came from, which is what drops it:
                        // the layout finds no size for a file whose kind is off.
                        // Every other preview is left exactly as it is — a
                        // running video is not restarted by a toggle it has
                        // nothing to do with.
                        //
                        // A pinned preview is the exception: there is no hover to rebuild it
                        // from — the pin is not a hover — so one whose kind was switched off
                        // comes down, by the path its own close button takes.
                        if pinned() {
                            let switched_off = current_media_kind()
                                .is_some_and(|kind| !kind.enabled())
                                || current_show
                                    .as_ref()
                                    .and_then(show_path)
                                    .and_then(|path| engine_kind_of(path))
                                    .is_some_and(|kind| !kind.enabled());

                            if switched_off {
                                pin_request = Some(end_pin_state(Reason::SwitchedOff));
                            }
                        } else if latest_preview_msg.is_none() {
                            match (current_media_kind(), current_show.clone()) {
                                (Some(kind), Some(show)) if !kind.enabled() => {
                                    latest_preview_msg = Some(show)
                                }
                                // A preview the engine draws has no media of its own —
                                // the engine's window is the preview — so its kind is
                                // read from the file the hover is about rather than
                                // from what is on screen.
                                (None, Some(show))
                                    if show_path(&show)
                                        .and_then(|path| engine_kind_of(path))
                                        .is_some_and(|kind| !kind.enabled()) =>
                                {
                                    latest_preview_msg = Some(show)
                                }
                                _ => {}
                            }
                        }
                    }
                    // The tray's `Pin Mode → Enable` row, and the configuration behind it having
                    // been reloaded. Whether there is a pin to take down is a question only this
                    // thread can answer, and a pin left standing by a feature that was switched
                    // off is a window nothing would ever take down again — so it comes down the
                    // way its own close button takes it.
                    PreviewMessage::PinChanged => {
                        if pinned() {
                            // A pin that is Always On Top owes the front to the setting, and one
                            // whose row is off never owed it: the flag is what either setting
                            // leaves behind otherwise, and it would be a pin holding a band its
                            // own setting does not name.
                            if pin_always_on_top() || !pin_update_move_to_top() {
                                PIN_FRONT_OWED.store(false, Ordering::SeqCst);
                            }
                            if !pin_enabled() {
                                pin_request = Some(end_pin_state(Reason::SwitchedOff));
                            } else {
                                let insert_after = if pin_holds_the_front() {
                                    HWND_TOPMOST
                                } else {
                                    HWND_NOTOPMOST
                                };
                                let _ = SetWindowPos(
                                    hwnd,
                                    insert_after,
                                    0,
                                    0,
                                    0,
                                    0,
                                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                                );
                                if pin_always_on_top() {
                                    if !pin_is_collapsed() {
                                        hide_pin_taskbar();
                                    }
                                } else {
                                    show_pin_taskbar();
                                }
                                place_pinned_siblings();
                                // And the player is told the same thing where it is mpv's: a film
                                // that held the topmost band for the pin's sake has to give it up
                                // when the setting says so, or it is a topmost window over a pin
                                // that is not (see `settle_pinned_mpv_ontop`).
                                settle_pinned_mpv_ontop();
                                AUDIO_CARD_DIRTY.store(true, Ordering::Release);
                                render_layered_preview(hwnd);
                            }
                        }
                    }
                    // One of the tray's two `Volume` rows was moved, and a pin up on a file of
                    // that half is owed the level at once rather than at the next file: the
                    // pin's player is this thread's, so the level is given here rather than
                    // where the row was clicked (see `tray_volume_moved`).
                    PreviewMessage::TrayVolume { audio } => tray_volume_moved(audio),
                    // The tray's `Text → Font Size` row was picked, and a pin already showing a
                    // page is that page put to the size the row names: a page is laid out on this
                    // thread, so the size is given here rather than where the row was clicked (see
                    // `publish_text_font_size`).
                    PreviewMessage::TrayTextSize { percent } => {
                        redraw_pinned_page(hwnd, Some(percent))
                    }
                    // The file the user picked while a preview was pinned, which the pin is to be
                    // shown instead of the one it has (see the tray's `Pin Mode → Update Preview`).
                    // It is held rather than handled here, the way a hover is: what the loop acts on
                    // is the newest pick, and a key walked down a listing is a pick a tick — each
                    // one a file to decode — so the ones it passes through on its way are not work
                    // anything is owed (see the swap below).
                    PreviewMessage::PinUpdate(path) => {
                        // A newer pick is what the pin is owed an answer to, so the file an older
                        // one is still waiting for a box of stops being the one it waits on: the
                        // answer that comes for it is then no answer to this pick, and the file
                        // the user actually picked last is the one that is shown (see
                        // `pin_awaiting_box`).
                        pin_awaiting_box = None;
                        pin_pick = Some(path);

                        // And whatever the pin was behind, an update is the moment it comes
                        // forward: the file the user just picked is the one they are looking at
                        // next, and a pin left behind the listing it is following is a window they
                        // would have to find before they could see it (see
                        // `Pin Mode → Update Preview → Move To Top`). The raise happens here rather
                        // than at the install so the pin answers the pick at once, with the wait
                        // for the file drawn in it where the read is not instant — and it keeps the
                        // band where a pick's press is holding it, because the click that made the
                        // pick activates the listing after this loop reads it (see `PIN_FRONT_OWED`).
                        // An update with no press behind it — a file walked to with the keys, a
                        // hover — has no activation to be kept out of, so its raise is the band
                        // taken and given straight back.
                        arm_pin_to_the_front();
                        // The window raised is the pin's own, which for a picture is not the handle
                        // this loop holds: the loop's window is the hover's, and a raise that moved it
                        // moved a window nobody can see while the pin stayed behind the listing (see
                        // `pinned_window`).
                        raise_pinned_preview(pinned_window(), pin_front_owed.is_some());
                        extend_pin_front_owed(&mut pin_front_owed);
                    }
                    PreviewMessage::PinPress { window } => {
                        // The click that made the press activates the listing, and that is the
                        // listing being put over the pin — before the shell has said which file
                        // was picked, and before any `PinUpdate` can answer for it. The pin owes
                        // the front from here, and holds the band the activation of the listing
                        // cannot undo (see `PIN_FRONT_OWED` and `keep_pin_at_the_front`); the
                        // pick's answer comes forward again when it lands.
                        //
                        // **The band is held across the activation rather than given back before
                        // it**, which is the whole of why a pick no longer shows the listing for a
                        // tick: this press is read off the hook *before* the click reaches the
                        // listing, so a raise made here is a raise the activation that follows
                        // undoes — and the correction is a tick late, which the user sees as the
                        // listing flashing over the pin.
                        //
                        // The window the grace asks about is the one the press was made in, which
                        // is the listing itself and not the foreground: the foreground is the
                        // window the user came from, the activation still to land, and a window
                        // the pin is already behind (see `PinFrontOwed` and
                        // `pin_press_was_not_a_pick`).
                        arm_pin_to_the_front();
                        raise_pinned_preview(pinned_window(), true);
                        // The two clocks the press begins, from the one moment it was read: how long the
                        // pick is answered for, and how long the band it took is held for. Only the first
                        // is extended, by a file that reaches the window late (see `PinFrontOwed`).
                        let pressed = Instant::now();
                        pin_front_owed = Some(PinFrontOwed {
                            until: pressed + PIN_RAISE_GRACE,
                            band_until: pressed + PIN_RAISE_GRACE,
                            listing: HWND(window as *mut _),
                        });
                    }
                    PreviewMessage::PinNoPick { window } => {
                        // A press that named no file was a click on the folder tree, the gap
                        // between the rows, or the toolbar above them: the user's own click put
                        // the listing in front and the pin belongs behind it. The grace a pick
                        // began is ended with the raise it was holding up, so what the press
                        // above came forward for is given back at once — behind the window the
                        // press was made in, which is the one the user is looking at, and not the
                        // foreground, which can still be the window they came from (see
                        // `pin_press_was_not_a_pick`).
                        //
                        // The window lowered is the pin's own, which for a picture is not the handle
                        // this loop holds: a lowering that moved the loop's window moved a window
                        // nobody can see while the pin stayed over the listing (see `pinned_window`).
                        pin_front_owed = None;
                        lower_pinned_preview(pinned_window(), HWND(window as *mut _));
                    }
                    PreviewMessage::OfficeRenderReady {
                        path,
                        generation,
                        ok,
                    } => {
                        // A page a pin is waiting for is the pin's answer, and it is taken here
                        // rather than by the hover machinery below: what takes it up is the swap,
                        // which lays the file out for the pin's own box, where a hover replayed
                        // for it would be a second preview on screen beside a window that is not
                        // one (see `take_pin_engine_answer`).
                        //
                        // The newest hover wins otherwise, as it does over every other
                        // message: a page that lands in the same tick as a new
                        // hover is not the answer to it.
                        if !take_pin_engine_answer(&path, ok, &mut pin_awaiting_box, &mut pin_pick)
                            && latest_preview_msg.is_none()
                            && page_ready.is_none()
                        {
                            page_ready = Some((path, generation, ok));
                        }
                    }
                    PreviewMessage::VideoProbed { path, generation } => {
                        // A pin waiting for this file's box is picked again the moment the answer
                        // is here, before the hover bookkeeping below — which is about a hover, a
                        // thing a pinned window is not (see `pin_awaiting_box`). A pin that has
                        // gone in the meantime is shown nothing: the wait was the pin's, and a
                        // window that is not up is not owed a file.
                        if pin_awaiting_box.as_deref() == Some(path.as_path()) {
                            pin_awaiting_box = None;
                            if pinned() {
                                pin_pick = Some(path.clone());
                            }
                        } else if latest_preview_msg.is_none() && video_probed.is_none() {
                            // Held apart the way a render's answer is: it is not a hover to
                            // act on but an answer about the one that is waiting. An answer
                            // the pin above has taken is not one of those — what it was for is
                            // the pick, and a hover replayed for it as well would be a second
                            // thing on screen beside the window that took it.
                            video_probed = Some((path, generation));
                        }
                    }
                    PreviewMessage::VideoSubtitlesReady(path) => {
                        // A pinned window showing this film is begun again so the copy that just
                        // landed is drawn by the frame after it — the one reload a hover
                        // deliberately does not have (see `reload_pinned_subtitles`). Nothing
                        // here is hover bookkeeping: the message is not an answer to a hover, and
                        // a hover is never begun again for it.
                        reload_pinned_subtitles(&path);
                    }
                    PreviewMessage::PinAnswered(answer) => {
                        // Both answers are the pin's own, and both are dropped where the
                        // pin is no longer up: the wait was a window's, and a window that
                        // is not there is owed nothing.
                        if pinned() {
                            match answer {
                                // A walk is taken up where a pick is, because that is what a
                                // walk is: a file the pin cannot be shown is stepped over
                                // rather than stopped at (see `PinStep`). The wait it was
                                // asked under is over with it, so the arc comes down on the
                                // tick the walk is acted on rather than turning until it is.
                                //
                                // It is taken only if the pin is still standing where the
                                // walk was asked from. A walk lands some time after the
                                // press, and by then the pin may have been given another
                                // file — a pick in the listing, or a walk the user has
                                // pressed since. Stepping from a file the pin has left is
                                // a window that changes its mind about a file nobody
                                // picked, so the answer is dropped instead.
                                PinPlanned::Walk(walk) => {
                                    pin_walk_wait = None;
                                    if pinned_path().as_deref() == Some(walk.from()) {
                                        // And the playlist it was answered with goes onto the
                                        // pin's own walk now rather than with the install a
                                        // load away: a press made while the file is still
                                        // being read re-anchors to the file on screen in
                                        // this order, so it steps the same playlist instead
                                        // of starting another (see `pin_walk_synced_its_playlist`).
                                        pin_walk_synced_its_playlist(&walk);
                                        pin_walk = Some(walk);
                                    }
                                }
                                // The hand-off button's name, for the file the pin is showing
                                // now. An answer for a file the pin has already left is
                                // dropped rather than shown: a name for the wrong file is
                                // worse than no name.
                                PinPlanned::OpenWith { path, name } => {
                                    if pinned_path().as_deref() == Some(path.as_path()) {
                                        if let Some(mut pinned) = pin_state() {
                                            if let Some(pin) = pinned.pin_mut() {
                                                pin.tooltip.default_app = name;
                                            }
                                        }
                                    }
                                }
                            }
                        } else {
                            pin_walk_wait = None;
                        }
                    }
                    PreviewMessage::MeasureProbed { path, size } => {
                        // The same pick again for a pin, where the measure answered with a box:
                        // that a reader has no box for the file at all is no box to lay anything
                        // out with, so the pin keeps the file it is showing rather than picking
                        // one up again (see `PinPlan::Awaiting`).
                        if pin_awaiting_box.as_deref() == Some(path.as_path()) {
                            pin_awaiting_box = None;
                            if size.is_some() && pinned() {
                                pin_pick = Some(path.clone());
                            }
                        } else if latest_preview_msg.is_none() && measure_probed.is_none() {
                            // And a measured box, held apart with the box itself: what is waiting
                            // on it is a hover, and the box is what that hover is laid out with
                            // (see `measure_probed`). A box the pin above has taken is spent on
                            // the pick rather than on a hover, on the terms the probe's is.
                            measure_probed = Some((path, size));
                        }
                    }
                    PreviewMessage::MagickReady {
                        path,
                        generation,
                        ok,
                    } => {
                        // A pinned window's picture first, on the terms the render tier's page
                        // above it is taken (see `take_pin_engine_answer`).
                        //
                        // The engine's answer, held apart for the reason the render tier's
                        // is: what the hover that asked is waiting for is a picture to be
                        // placed with, not another hover — and an engine that will not draw
                        // the file is the same wait answered, with nothing in it.
                        if !take_pin_engine_answer(&path, ok, &mut pin_awaiting_box, &mut pin_pick)
                            && latest_preview_msg.is_none()
                            && page_ready.is_none()
                        {
                            page_ready = Some((path, generation, ok));
                        }
                    }
                    PreviewMessage::PeazipReady {
                        path,
                        generation,
                        ok,
                    } => {
                        // And a listing, which is the same answer once more, taken for a pin the
                        // same way: what the hover is waiting for is a table of contents to draw
                        // a page from, and an archive the engine will not list is that wait
                        // answered with nothing.
                        if !take_pin_engine_answer(&path, ok, &mut pin_awaiting_box, &mut pin_pick)
                            && latest_preview_msg.is_none()
                            && page_ready.is_none()
                        {
                            page_ready = Some((path, generation, ok));
                        }
                    }
                    other => {
                        latest_preview_msg = Some(other);
                        refresh_requested = false;
                    }
                }
            }

            // A pinned window's chrome is asked for on this loop's clock: the pointer is read from
            // the cursor rather than waited on, because a strip that has gone is not a region the
            // mouse can be over — and an answer that has changed owes the window a repaint, which
            // is the whole of what showing and hiding it costs (see `refresh_pin_chrome`). Nothing
            // is asked of a pin that is not up, and nothing of one whose chrome is not drawn over
            // its media, which is the answer both of those questions are asked through.
            if pinned() {
                let changed = pin_state().and_then(|mut pinned| {
                    let pin = pinned.pin_mut()?;
                    let now = Instant::now();
                    Some(refresh_pin_chrome(pin, now, cursor_screen_point()))
                });

                if changed == Some(true) {
                    refresh_requested = true;
                    // The strips are the chrome now, so a change answers them here rather than
                    // waiting for the repaint below: the paint carries their pixels, and the sync
                    // carries the paint onto the screen (see `sync_the_chrome_strips`).
                    pin_chrome_strips::sync_the_chrome_strips();
                }

                // A strip of the pin's chrome stands over whatever is in the band — a hosted film,
                // the cover, or FFmpeg's player — and that window is the one on top for as long as
                // no strip is showing. Both of its raises are held off from here on (see
                // `pin_chrome_up`), so the raise the pin owes is the pin's own: on the way up,
                // which is a moment a hand is looking for a close button, and again for as long
                // as a strip is showing, because the film stands itself back up over the pin on
                // its own band (see `player_stands_over_the_pin`). The way out needs none: the
                // first raise through either caller once the last strip has gone puts the player
                // back on top. What keeps this one raise from being a raise per tick is that it
                // is asked of the order that is standing rather than of the strips that are up —
                // the player *under* the pin is the answer both raisers put in, and the one this
                // leaves alone (see `place_player_window`).
                if pin_chrome_up() {
                    // A topmost pin is put back over the film whenever the film has stood itself
                    // back up on its band. A pin that is not Always On Top has no such fight to
                    // win and no business being raised at all — what it owes over the film is the
                    // *film's* place, and only when the film is not already under it (see
                    // `raise_pinned_window`, `put_player_under_the_pin`).
                    if pin_always_on_top() {
                        if player_stands_over_the_pin(hwnd) {
                            raise_pinned_window(hwnd);
                        }
                    } else if !player_is_under_the_pin(hwnd) {
                        put_player_under_the_pin();
                    }
                }

                // And the card's menu, which is a window of its own over the screen: the
                // raise above has to be answered with the menu's own, and the menu's own
                // paint is the one thing the card's repaint does not cover (see
                // `render_pin_menu`). Asked whether anything changed or not, because a
                // menu is a panel a hand is looking at and the price of being sure it is
                // where it stands is a repaint no bigger than its own rows.
                render_pin_menu();
            }

            // A page the render engine draws is not messaged about the way an Office page
            // is: what that engine writes is a file under the app's own folder, so whether
            // the page has arrived — or whether the engine has answered that it will not
            // draw the document at all — is a read of that folder rather than a message
            // from a thread. What comes of it is the answer an Office page gives, and the
            // same code below takes it up: it is the same wait, in the same box. The page a book
            // is converted into is the second engine that answers this way, and it is read here
            // for the same reason.
            //
            // Which documents are watched for is asked of the place the request was made
            // from, so a hover is never watched for a page nothing was asked to draw: an
            // Office document whose own application is here has a page asked of that
            // application's tier, and is answered by a message rather than by this read (see
            // `libre_render_is_due`).
            if page_ready.is_none() {
                if let Some((path, generation)) = page_render_pending.as_ref() {
                    if let Some(drawn) = engine_page_answer(path) {
                        page_ready = Some((path.clone(), *generation, drawn));
                    }
                }
            }

            // A pinned window waits on the same two engines the same way, and it is watched for
            // here for the same reason a hover is: neither engine messages when it is done, so
            // what says a page is there is the page. The pin's wait is its own — what takes the
            // answer up is the swap rather than a hover replayed — so it is asked apart from the
            // request above, which is a hover's (see `take_pin_engine_answer` and `PinPlan::Awaiting`).
            if let Some(path) = pin_awaiting_box.clone() {
                // Either answer ends the wait: a page that has landed is picked up, and a refusal
                // leaves the pin showing the file it has, as a refusal always does. Nothing to
                // say at all — which is every Office document and every picture the image
                // converter develops, those two answering by message instead — leaves the wait
                // standing until its answer arrives.
                if let Some(drawn) = engine_page_answer(&path) {
                    take_pin_engine_answer(&path, drawn, &mut pin_awaiting_box, &mut pin_pick);
                }
            }

            // A page the render tier has finished with, for the hover that asked
            // for it: that hover is replayed, which measures the page itself and
            // draws it in place of the spinner it supersedes. A payload from an
            // older hover is dropped here — the page it wrote is kept as far as the
            // cache budget allows, and no further.
            if let Some((ready_path, ready_generation, ready_ok)) = page_ready {
                let shown = current_show.as_ref().and_then(show_path);
                let hovered = ready_generation == current_generation
                    && shown.map(|path| path.as_path()) == Some(ready_path.as_path());

                // Which request the answer belongs to, so that a page landing for a hover
                // that has gone cannot clear a wait another hover is still in.
                let answers_the_request =
                    page_render_pending
                        .as_ref()
                        .is_some_and(|(path, generation)| {
                            *path == ready_path && *generation == ready_generation
                        });

                // An answer for the file the hover on screen is waiting on, naming the
                // hover before it, is that wait's own answer: the request it made was
                // folded into the work this one belongs to, and what it answers is the
                // page, picture or listing the wait is for (see `answer_belongs_to_the_wait`).
                let waiting_for_this_file = !hovered
                    && answer_belongs_to_the_wait(
                        &ready_path,
                        ready_ok,
                        shown.map(|path| path.as_path()),
                        pending_load.as_ref(),
                    );

                if (hovered || waiting_for_this_file) && ready_ok {
                    if latest_preview_msg.is_none() {
                        // The wait is over, so the request stops being the pending
                        // one here — where the page is actually taken up.
                        page_render_pending = None;
                        // The page is there, so the hover is replayed: that measures
                        // the page itself, moves the window to its size and loads it.
                        // It is an upgrade rather than a new preview, though, and what
                        // is on screen stays while it happens — hiding the spinner for
                        // the second a large picture takes to decode is a blink the
                        // user sees and reads as the preview failing.
                        page_upgrade = Some(ready_path.clone());
                        // A mouse hover is replayed where the pointer is now: the
                        // spinner it replaces was kept with the pointer while the
                        // render ran, and a page that jumped back to where the
                        // hover started would jump away from where it was waited
                        // for.
                        latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
                    }
                    // A newer message was in hand, so the page is not shown now. The
                    // wait it was rendered for is left standing rather than cleared
                    // with it: what is being waited on is still that page, so the cap
                    // on waiting keeps something to measure and the spinner comes down
                    // when its time is up instead of hanging there for good.
                } else if hovered {
                    // Nothing was drawn and nothing is coming. A preview that is
                    // not a spinner is kept — it is a preview like any other —
                    // while a spinner has nothing left to stand in for.
                    let showing_spinner = CURRENT_MEDIA
                        .lock()
                        .map(|media| {
                            media
                                .as_ref()
                                .map(|media| media.media_type.is_loading())
                                .unwrap_or(false)
                        })
                        .unwrap_or(false);
                    if showing_spinner {
                        let _ = ShowWindow(hwnd, SW_HIDE);
                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            *current = None;
                        }
                    }
                    // The wait is over whether or not its spinner had gone up: an engine
                    // that turns a file down in less than the spinner's delay — a raw
                    // file whose length works out to no picture, a converter that finds
                    // nothing to read — answers before there is anything to take down,
                    // and the pending load is what the spinner would go up for. Left
                    // armed it goes up a moment later for a preview that is not coming
                    // and stays up, which is the one thing the entry above cannot undo:
                    // the answer it belonged to has been read already.
                    pending_load = None;
                    page_render_pending = None;
                } else {
                    // The hover this page was rendered for is over: it landed after
                    // the pointer had moved on, so nothing is waiting for it. What was
                    // rendered is kept as far as the budget allows and no further — at
                    // a size of nothing it is dropped here rather than held for a
                    // hover that has already gone.
                    //
                    // Only the wait this answer belongs to is cleared with it: a hover
                    // that is still waiting for a page of its own keeps the request it
                    // was made for, which is what the cap on waiting is read against.
                    if answers_the_request {
                        page_render_pending = None;
                    }
                    office_render::hover_ended(&ready_path);
                }
            }

            // A video's probe has answered for the hover that was waiting on it: that
            // hover is replayed — this time the layout has the shape the probe cached, so
            // the hover goes on as the video it is — and the wait it is in goes on as the
            // video's own rather than as a new preview (see `video_probe` and
            // `video_replay`). An answer for a hover that has gone is dropped here: what
            // the probe measured is held for the next hover of the file either way.
            if let Some((probed_path, probed_generation)) = video_probed {
                let waiting = video_probe.as_ref().is_some_and(|(path, generation)| {
                    *path == probed_path && *generation == probed_generation
                });

                if waiting {
                    video_probe = None;

                    if probed_generation == current_generation && latest_preview_msg.is_none() {
                        // The hover is replayed where the pointer is now, the way a page
                        // landing on a spinner is: the wait was kept with the pointer
                        // while the probe ran, so the video should be too.
                        video_replay = Some(probed_path);
                        latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
                    }
                }
            }

            // A box measured off the preview thread has answered for the hover that was waiting
            // on it: that hover is replayed — this time the measure reads the box off the table
            // rather than out of the file, so the hover goes on as the preview it is — and the
            // wait it is in goes on as the preview's own rather than as a new one (see
            // `measured_off_the_tick` and `measure_replay`). An answer for a hover that has gone
            // is dropped: what the measure answered is held for the next hover of the file
            // either way.
            if let Some((measured_path, size)) = measure_probed {
                let waiting = current_show.as_ref().and_then(show_path) == Some(&measured_path);

                if waiting && latest_preview_msg.is_none() {
                    if size.is_some() {
                        measure_replay = Some(measured_path);
                        latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
                    } else {
                        // The reader has nothing for this file, which is not a wait that can be
                        // answered: the wait ends as the mark every failed file ends as rather
                        // than the spinner standing over nothing until the pointer moves (see
                        // `take_the_wait_as_failed`).
                        let waiting_on_it = pending_load
                            .as_ref()
                            .is_some_and(|pl| pl.path == measured_path);

                        if waiting_on_it {
                            take_the_wait_as_failed(hwnd, &mut pending_load);
                            clear_pointer_hold();
                            pending_load_cancel = None;
                        }
                    }
                }
            }

            // The display under the preview changed and the frame that was on screen
            // went with it. The window proc could only discard what was drawn; the
            // hover it came from is what knows how to draw it again, at the scale of
            // the display the pointer is on now. A newer message in hand is left to
            // speak for itself, and the flag waits for a tick where none does.
            if latest_preview_msg.is_none() && DISPLAY_RESET.swap(false, Ordering::AcqRel) {
                if pinned() {
                    // A pin is not replayed from a hover: it is put back on the display it has
                    // to be on now, at the box that display can show of it.
                    pin_request = replace_pinned_window();
                } else {
                    latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
                }
            }

            // And a file the engine was watched failing at: the engine took it — every part of
            // the question the probe asks answered yes — and then never drew a frame of it, so
            // nothing here will play it and the hover is replayed to be answered without a
            // preview (see `video_player::mark_unplayable`). It waits for the tick where the hover
            // of the file is the one that can be replayed, which is why it is a flag rather than
            // something the tick does itself.
            //
            // A pin is not replayed from a hover, so a file that fails under a pin keeps what it
            // has: the mark is what the *next* pin of that file takes, and what plays it from then
            // on is nothing at all.
            let failed_file_is_shown = engine_failed.as_ref().is_some_and(|failed| {
                !pinned() && current_show.as_ref().and_then(show_path) == Some(failed)
            });

            if latest_preview_msg.is_none() && failed_file_is_shown {
                engine_failed = None;
                latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
            }

            // The engine has something to answer for: it could not be had at all, or it
            // could not put the file of the hover that is up into its window. Either way
            // this app draws none of it itself, so there is nothing to fall back to and
            // nothing left to wait for — the wait goes rather than standing as a spinner
            // over nothing.
            if latest_preview_msg.is_none() && webview_preview::take_failure_notice() {
                let waiting_on_the_engine = pending_load
                    .as_ref()
                    .is_some_and(|pl| engine_kind_of(&pl.path).is_some());

                if waiting_on_the_engine && !webview_preview::is_showing() {
                    pending_load = None;
                    pending_load_cancel = None;
                    let _ = ShowWindow(hwnd, SW_HIDE);
                    clear_pointer_hold();

                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        *current = None;
                    }
                }

                // A pin is the same answer with one difference: a hover whose wait has gone is
                // nothing on screen, while a pin whose browser has failed is a window standing
                // over an empty band — this app draws none of a document itself, so there is
                // nothing to fall back to and it comes down the way its own close button takes
                // it. What a failed *navigation* leaves behind is the document that was asked
                // for and never arrived, which is exactly this case: the wait is over and what
                // it was waiting for is not there (see `note_document_failed`).
                let pinned_engine_failed = pinned()
                    && matches!(
                        current_media_type(),
                        Some(MediaType::EngineSvg) | Some(MediaType::EngineFont)
                    )
                    && !webview_preview::is_showing();

                if pinned_engine_failed {
                    pin_request = Some(end_pin_state(Reason::MediaGone));
                }
            }

            // A file the user picked while a preview is pinned, which the pin is to be shown
            // instead of the one it has (see the tray's `Pin Mode → Update Preview`). What is on
            // screen is not a hover — there is no layout to make and no generation for an answer to
            // be matched against — so what is swapped is the media, here, and the take-up asked for
            // below is the pin's own: the box is the one the pin already has with the new file's
            // shape fitted inside it, and what comes of it is the same window showing something
            // else (see `PreviewMessage::Pin`).
            //
            // Nothing is swapped where the pin has gone in the meantime, where the setting was
            // switched off behind it, where a wait is already in hand, or where the file cannot be
            // shown at all: a pin that is up can only keep the file it is showing.
            //
            // A pin that has gone waits for nothing either: the slot a wait is held in outlives
            // the window it belongs to, and an entry left behind by a pin that has been closed
            // would take the landing of a hover's wait for its own (see `pin_awaiting_box`).
            if !pinned() {
                pin_awaiting_box = None;
                // A file a pin that is over was loading is not waited for: the window it was
                // being shown in has gone, and the answer landing for it would be installed
                // into a pin that is not there (see `PinLoad`).
                pin_load = None;
                // And neither is a swap it was being held for: the engine behind that one is
                // decoding a file for a window that has closed, and the file it is held over
                // is going with the window, so there is nothing left to wait for it with.
                abandon_pin_swap(&mut pin_swap_hold);
                pin_arc_set(None);
                // A walk a pin that is over was stepping is not stepped on: a caption's next
                // belongs to the window it was pressed on, and the pin that follows is shown
                // whatever the user picks rather than the rest of a walk behind it.
                pin_walk = None;
                pin_player_started = None;
                // And the walk it was reading the folder for is a wait on a window that is
                // gone: it is not waited for, and the arc it may have put up goes down with
                // it (see `PinWait`).
                pin_walk_wait = None;
                pin_walk_wait_seen = false;
                // A file picked behind a pin that is over is not owed to the pin that follows it:
                // what the key brings a bubble back on is what was picked while *that* bubble was
                // down, and a file nobody is holding any more is a swap nothing asked for (see
                // `pin_bubble_pick`). It is cleared here rather than with the file that is on
                // screen, because a pin ended and another taken up in the same tick is a tick this
                // is the only place to notice.
                pin_bubble_pick = None;
                // And a return owed a pin that is over is not owed to the pin that follows it: the
                // file it was owed for belongs to the window that has gone (see `pin_return_owed`).
                pin_return_owed = false;
            }

            if let Some((path, walk, user_named)) = take_pin_pick(
                &mut pin_pick,
                &mut pin_held_pick,
                &mut pin_swap_requested,
                &mut pin_walk,
                pin_awaiting_box.is_some(),
            ) {
                // A pinned window being shown a file it is not already showing is the other
                // place a preview is built, so the engine choices are read here — before the
                // box is measured, the probe started or the file loaded — and everything this
                // pick is asked after this is answered with what is read here (see
                // `latch_engines_for_the_preview`). A pick of the file the pin is already
                // showing is not one: it is a step of a walk that landed back on it, or a pick
                // this block holds for a window that has not come back yet, and nothing is
                // built for either.
                //
                // shortcut: the read lands a tick before the window is given the file, so a
                // box change in exactly that gap relaunches the outgoing film under the new
                // choice; the read belongs in the install if that is ever seen.
                if pinned_path().as_deref() != Some(path.as_path()) {
                    latch_engines_for_the_preview();
                }

                pin_awaiting_box = None;
                // Whatever was loading is a load for a file nobody is asking for any more, and
                // the answer it lands with is dropped with it: what the pin shows next is
                // planned here, for this file, and taking up an answer for another one would
                // install a frame the take-up below was not asked for.
                pin_load = None;
                // And so is a swap being held for a video's first frame, which is the same
                // answer to the same question one step later: this pick is what the pin is
                // being shown next, the engine behind the held one is started for a file
                // nobody is going to look at, and the frame it has been drawing into goes
                // with it. The second load this pick starts plans the engine again from
                // nothing, which is what `video_player::play` does to whatever was playing
                // (see `abandon_pin_swap`).
                abandon_pin_swap(&mut pin_swap_hold);
                pin_arc_set(None);
                // A pin that is a bubble holds the file it is showing for the key rather than
                // swapping in a file picked in the listing — but a step of the pin's own walk
                // is swapped in even while collapsed, so that an automatic advance (a loop-end
                // step, a wrap, a caption press held) starts its player while minimized rather
                // than stalling silent: what the user is doing behind a bubble is picking, and
                // a bubble that took the pin down on the first pick would be a window lost for
                // a file they did not ask to see (see `pin_bubble_pick`).
                //
                // **Unless the update is the pin coming back.** `Pin Mode → Update Preview → Move
                // To Top` is the row that says an update brings the pin in front of the listing it
                // follows, and a pin that is away has no window in any band to bring forward — so
                // with that row on, the file the pick names is the file the pin comes back up on
                // and is put in front of the listing by, rather than one held for the key (see
                // `bring_a_put_away_pin_back`). It is asked here, before the file is planned,
                // because a file held for the key is not one the pin is being shown; and only where
                // nothing is loading, which is the rule the key's own restore keeps for the bubble it
                // brings up — a load already in hand is a tick the pick below has no work for, and a
                // window put back up without the file it was brought back for is a worse answer than
                // the bubble it stayed in (see `pin_bubble_pick`).
                let the_pin_comes_back = walk.is_none()
                    && pending_load.is_none()
                    && an_update_brings_a_put_away_pin_back();

                if pinned() && pin_is_collapsed() && walk.is_none() && !the_pin_comes_back {
                    pin_bubble_pick = Some(path);
                } else if pinned() && (pin_update_enabled() || !user_named) {
                    if pending_load.is_some() {
                        // A hover somewhere behind the pin is loading, and a pick that arrived
                        // while it did was the user's first pick after the pin took the focus —
                        // the one whose click or key press is a single press bit the hook never
                        // reads twice. It is held here and taken up on a tick when nothing is
                        // loading, rather than dropped: the walk is carried with it, so a caption
                        // step held this way is still a step of the walk it was a part of.
                        pin_walk = walk;
                        pin_held_pick = Some(path);
                    } else {
                        // The setting is the gate on a *pick in the listing* — a file the pointer
                        // or the keyboard chose behind the window — and not on a step of the pin's
                        // own walk, which is a thing the user pressed on the window and is asked
                        // for by no setting: a caption's next and previous did nothing at all with
                        // `Pin Mode → Update Preview` off, which is the same "the navigation
                        // stopped" a file with no preview causes, asked for by no one (see
                        // `step_pinned_file`). Whether the plan answered that the file cannot be
                        // shown at all, as opposed to a box or an engine that has not answered yet.
                        // Only the first of the two is a file the walk steps over: a wait is a
                        // wait, and what ends it is the answer it was asked for.
                        let mut refused = false;
                        let update = match pin_update_plan(&path) {
                            Some(PinPlan::Show(update)) => Some(update),
                            // The file has no box of its own yet — what was measured for it is the
                            // wait for a read, for a probe that has not answered, or for a page an
                            // engine still owes it — so the pin keeps the file it is showing and
                            // picks this one up again when the answer lands. A box read and a
                            // sound's probe are started where they are taken, which is
                            // `media_dimensions`; a video's probe and an engine's page, picture or
                            // listing are asked for here (see `pin_swap_awaits`, `audio_box` and
                            // `request_pin_engine_render`).
                            Some(PinPlan::Awaiting) => {
                                let hover = HoverFacts::read(&path);
                                if video_probe_due(&hover) {
                                    spawn_video_probe(path.clone(), current_generation);
                                }

                                // What nothing has asked for yet is what an engine owes the file.
                                // It is asked only where an engine really can be asked, so a file
                                // no engine here reaches keeps the rule a swap has always had: the
                                // pin keeps the file it is showing.
                                let asked = page_is_on_the_way(&path)
                                    && request_pin_engine_render(&path, current_generation)
                                        .is_some();

                                // What the pin is waiting for, if anything: a read or a probe in
                                // flight, or an engine that has just been asked. Nothing at all is a
                                // pick to make again rather than a wait — an answer that landed
                                // between the plan and the ask is one no engine will announce (see
                                // the hover's own `requested.is_none` arm) — while a file no engine
                                // can be asked about is left showing what it has.
                                let outstanding =
                                    measure_waiting(&path) || video_probe_due(&hover) || asked;

                                if outstanding {
                                    pin_awaiting_box = Some(path.clone());
                                    // The walk is parked, not dropped: the box it is a step of
                                    // is still waited on, and the pick retried when the answer
                                    // lands is a step of the same playlist rather than the
                                    // start of another (see `take_pin_pick`).
                                    pin_walk = walk.clone();
                                    None
                                } else if page_is_on_the_way(&path) {
                                    // A file the layout reads as waiting on an engine that no
                                    // reader of its kind can be asked for — a tier this machine
                                    // has not got, a kind whose engine was switched off — has
                                    // nothing coming and nothing to keep: the pin keeps the file
                                    // it is showing, which is the only thing a window that is
                                    // already up can do for a file it cannot show (the hover comes
                                    // down on the same answer, having nothing to keep).
                                    refused = true;
                                    None
                                } else {
                                    // Nothing was on its way after all — an answer that landed
                                    // between the plan and the ask is one no engine will announce
                                    // — so the file is planned once more, and this time there is a
                                    // box to lay out.
                                    match pin_update_plan(&path) {
                                        Some(PinPlan::Show(update)) => Some(update),
                                        _ => {
                                            refused = true;
                                            None
                                        }
                                    }
                                }
                            }
                            None => {
                                // A file no kind previews, and a file the pin is already showing —
                                // the walk's own list of one steps onto it. The first is stepped
                                // over, because a walk that stopped on a file with no preview
                                // would stop on every one of them; the second is not, because there
                                // is nothing to show it again and nowhere to step on to.
                                refused = walk.is_some()
                                    && pinned_path().as_deref() != Some(path.as_path());
                                None
                            }
                        };

                        if let Some(update) = update {
                            // The read and the decode are a thread's work, so a file whose load is
                            // slow is a wait the pin is painted for rather than a window frozen
                            // for as long as the disk takes (see `PinLoad`).
                            pin_load = Some(PinLoad::start(&path, update, walk));
                        } else if refused {
                            // A file the pin cannot be shown is stepped over rather than stopped
                            // at: one press of a caption button is one gesture, and the gesture
                            // is the next file there is to look at (see `PinStep`).
                            //
                            // The mark is not put up for it. Nothing has failed here — the pin was
                            // never shown this file, and the one it is showing is still on screen
                            // and still good — so a cross would be drawn over a picture that is
                            // working, in answer to a file the user is not looking at. What the
                            // walk runs out on is the other case, and the mark is the answer to
                            // that (see `show_pin_failure`).
                            pin_step_off(walk, &mut pin_walk);
                        }

                        // **And the pin comes back up on the file it was brought back for, which is
                        // not here.** It is owed here, after the plan, for the reason the key's own
                        // restore waits for a pick to be in hand: a window put back up without the
                        // file it was brought back for is a worse answer than the bubble it stayed in
                        // (see `pin_bubble_pick`).
                        //
                        // And it is answered at the take-up of the file rather than here, because a
                        // pick names a file nothing has read yet: a window put up now comes up on the
                        // file before it and swaps the new one in a tick or more later, which is the
                        // flash between the two images the user reported (see
                        // `bring_a_put_away_pin_back`). The one file there is nothing to wait for is
                        // the file the pin is already showing — it is on the window now — and the door
                        // answers for that one alone, every other answer being the take-up's below.
                        if the_pin_comes_back {
                            // The door answers for the file the pin is already showing, which is the
                            // one pick there is nothing to wait for: a pin left owing a return for it
                            // would wait on a load that is never asked for, since a swap to the file
                            // already on the window is no swap at all (see `pin_update_plan`).
                            pin_return_owed = !bring_a_put_away_pin_back(&path);
                        }
                    }
                } else if pinned() {
                    // A pick the setting refused leaves a parked walk parked rather than
                    // dropping it: the setting gates a file the user named, not the playlist
                    // behind it, so the next press still steps the order it holds.
                    pin_walk = walk;
                }
            }

            // A relayout of a pinned window's own media has answered, and is installed here on this
            // thread. It is asked of before the load below, and that order is the only thing
            // that matters: a swap installs its own media and a relayout left over from the box
            // before it would put a frame of the file just left behind over the file just
            // arrived — which the path check in `take_pin_relayout` refuses, but which is
            // better not to be asked about at all.
            //
            // And it is not asked of at all while a swap is being held. The path check in there
            // is what keeps a relayout from installing over a file the pin has left, and it asks
            // whether the pin is still standing on the file the relayout was for — which during
            // a hold it is, and which is no longer the answer worth acting on: the engine behind
            // the pin is already playing the *next* file, so what a relayout would install is a
            // decode of the file being held over, for a window that is about to be shown
            // something else. The answer waits in the slot until the hold is over and the path
            // check has an honest question to refuse it with (see `PinSwapHold`).
            if pin_swap_hold.is_none() && take_pin_relayout(&mut pin_relayout) {
                // A frame that has just been laid out again is a frame the window has not drawn: what
                // is on screen is the one before it, scaled into the box the new one was decoded for.
                // Nothing else on this tick is looking for a paint — a picture that is not moving has
                // no frame to advance and no clock to repaint for — so this is where the sharper frame
                // goes up (see `take_pin_relayout`).
                render_layered_preview(hwnd);
            }

            let swap = settle_pin_swap(&mut pin_swap_hold, &mut pin_load);

            if let Some(swap) = swap {
                // The loop's own record of what the pin is showing, borrowed for whichever of
                // the two things below is about to rewrite it. It is borrowed here rather than
                // inside the arms because there is one list of what a swap touches and it is
                // `pin_install`'s (see `PinInstall`).
                let install = pin_install(
                    &mut pin_walk,
                    &mut pin_load,
                    &mut pin_player_started,
                    &mut audio_started,
                    &mut audio_start_offset,
                    &mut audio_share_seek,
                    &mut audio_paused,
                    &mut audio_repaint_at,
                    &mut audio_card_dpi,
                    &mut audio_name_scroll,
                    &mut video_pos,
                    &mut current_video_path,
                    &mut current_show,
                    &mut pin_request,
                );

                match swap {
                    // The engine is started and its first frame is not in yet. Nothing of the
                    // window changes: the file on screen is the one it had, the arc the load
                    // put up is left turning over it, and the loop's own per-tick take is kept
                    // off the standing file until the frame lands (see `PinSwapHold`).
                    PinSwap::Holding(hold) => pin_swap_hold = Some(hold),
                    // A file this app cannot read, a file still in the cloud, a player that
                    // would not start, or a wait that has run out of reasons to go on: nothing
                    // of it to show, so the walk is asked for the next file rather than left on
                    // one that cannot be shown — and where the walk has nothing left to ask
                    // for, the mark for the file is what the pin is left standing over (see
                    // `pin_step_off` and `show_pin_failure`).
                    //
                    // A pin that is away stays away for it: the file was never put on the window,
                    // so there is nothing for it to come back on — and the mark is no file either
                    // (see `pin_failure_plan`). What brings it back is the first file in the walk
                    // that can be shown (see `pin_return_owed`).
                    PinSwap::Refused { path, walk } => refuse_pinned_media(install, path, walk),
                    PinSwap::Ready(file) => {
                        install_pinned_media(install, file);

                        // The pin has just been put up with the file the user picked, and what the
                        // click that picked it did to the z-order may have landed while the file
                        // was being read: the front the pin owes is owed again here, where the
                        // window that is to stand over the listing is what is on screen (see
                        // `PIN_FRONT_OWED`).
                        arm_pin_to_the_front();
                        extend_pin_front_owed(&mut pin_front_owed);

                        // And the band a step's cover is holding, where this install came out of
                        // a hold: the park's own settle ran earlier in this tick, before the pin
                        // had the box it has just been given, so it is asked again now the file
                        // the window was held for is in it. A cover left standing over a band
                        // that has moved is the outgoing frame stretched into the incoming
                        // file's shape, which is the whole of what the hold was for (see
                        // `PinSwapHold` and `settle_pinned_park`).
                        settle_pinned_park();
                    }
                }
            }

            // The pin's window is this app's own to keep on the screen, where Windows is not already
            // keeping it. The Shell's `Win+D` takes every ordinary window away and gives them all back on
            // the second press, and a pin that stands in a window of its own is one of them — so what is
            // left to this road is the pin that stands in the hover's window, which is no window of the
            // Shell's list: it is put up here at the first tick the user is back in a window rather than
            // looking at the desktop (see `the_pin_owes_a_show` for the whole of the question, and
            // `foreground_is_the_desktop` for the one answer it refuses).
            if the_pin_owes_a_show(pinned_window(), GetForegroundWindow()) {
                show_pinned_window(pinned_window());
            }

            // A pin that has just answered a pick keeps answering for the front of its band for a
            // moment after it: the click is this loop's before it is the listing's, and the window the
            // listing is in is activated after it — which is that window being put over the pin again
            // where the band is not held across it (see `PIN_RAISE_GRACE` and `keep_pin_at_the_front`).
            keep_pin_at_the_front(
                pinned_window(),
                &mut pin_front_owed,
                Instant::now(),
                GetForegroundWindow(),
            );

            // A pin that owes the front owes it while the window it follows is the one the user is
            // in. Another program in front is the user working elsewhere, so the band goes back to
            // the ordinary one and the pin goes below that program — the desktop included, except
            // that the desktop is the bottom of the desktop's own list and the pin keeps the place
            // it stands in rather than going under the wallpaper (see `pin_holds_the_front` and
            // `foreground_is_the_desktop`). Without this, a pin that holds the band would stay over
            // every other program until the next press that named no file, which is one click too
            // late — and with `Always On Top` off that band is not the setting's to keep. A pick's
            // grace is left alone: the click it answers may not have activated the listing yet, and
            // the window in front at that moment is the one the user is coming *from*.
            if pin_front_owed.is_none() && !pin_always_on_top() && pin_holds_the_front() {
                let foreground = GetForegroundWindow();
                if !foreground.is_invalid()
                    && !foreground_is_ours(foreground)
                    && !crate::shell::explorer_hook::is_foreground_explorer()
                {
                    if foreground_is_the_desktop(foreground) {
                        lower_pinned_preview(pinned_window(), HWND(ptr::null_mut()));
                    } else {
                        lower_pinned_preview(pinned_window(), foreground);
                    }
                }
            }

            // A pin that owes nothing and holds the topmost band is a window in the wrong band,
            // whatever road put it there, and the band is given back here: with `Always On Top` off
            // nothing else in this app gives it back on its own, so the pin would stand over every
            // program the user worked in until the first press that happened to raise it — the one
            // road that takes the band off a window — which is one click too late. `HWND_NOTOPMOST`
            // is the whole repair: the window keeps the place it stands among the ordinary ones and
            // lands in front of them, which is where a pin with the setting off belongs while the
            // user is in it (see `band_place`, the same question asked where a window is placed).
            //
            // **The pin's own window is asked too, and it is the one that needs it most.** A pick's
            // press holds that band across the listing's activation, and the grace that began it is
            // what gives it back; a debt that ended by any other road — the flag cleared by a press on
            // the pin's chrome, a setting turned off under it — would otherwise leave the pin over
            // every program the user works in, with nothing left ticking to look for it (see
            // `PinFrontOwed` and `give_the_band_back`).
            if pinned()
                && pin_front_owed.is_none()
                && !pin_always_on_top()
                && !pin_holds_the_front()
            {
                // And the band is nobody's to hold any more, so the next placement asks the setting
                // again rather than keeping a window in a band its hold is over (see `PIN_BAND_HELD`).
                PIN_BAND_HELD.store(false, Ordering::SeqCst);
                for window in [hwnd, pinned_window()] {
                    if window_holds_the_topmost_band(window) {
                        let _ = SetWindowPos(
                            window,
                            HWND_NOTOPMOST,
                            0,
                            0,
                            0,
                            0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                        );
                    }
                }
            }

            // The pin's request, if it made one, is the newest thing that happened and speaks
            // for what is on screen: a key that was pressed, a button that was clicked, or the
            // media behind the pin having come apart. It is put to the same machinery a hover
            // is, which is what puts the window up and takes it down (see `PreviewMessage::Pin`).
            if let Some(request) = pin_request.take() {
                latest_preview_msg = Some(request);
            }

            // While a preview is pinned, nothing is shown as a hover — including the hovers this
            // loop makes for itself out of the record of what is pinned, which is why the gate is
            // here and not only at the hook's own doors (see `hover_is_shown`). The messages a pin
            // answers for itself are not hovers and go on to the match below as they always did.
            if let Some(message) = latest_preview_msg.as_ref() {
                if !hover_is_shown(message, pinned()) {
                    latest_preview_msg = None;
                }
            }

            // Anchored one read before the layout it decides, so the box is placed clear of
            // the hand that asked for this hover rather than of the one it was measured from
            // (see `replay_where_the_pointer_is`).
            if let Some(preview_msg) = replay_where_the_pointer_is(latest_preview_msg) {
                // Common variables for Show/ShowKeyboard - set in match, used after
                let mut show_path: Option<PathBuf> = None;
                let mut show_layout: Option<PreviewLayout> = None;
                let mut show_spinner_layout: Option<PreviewLayout> = None;
                let mut show_placement: Option<HoverPlacement> = None;
                let mut show_is_video: bool = false;
                // Whether this hover is waiting on a video's probe rather than on the
                // video itself (see `video_probe_due`).
                let mut show_video_probe: bool = false;
                // Whether this hover is waiting on a measure that reads the file rather than on
                // the file itself (see `measure_waiting`).
                let mut show_measure_probe: bool = false;
                let mut show_requested = false;
                let mut preview_scale = current_hover_scales().picture;
                let mut show_dpi = 96u32;
                // The room the display the hover is on has, which is what an engine that
                // has to draw the preview before the file can be measured is asked for.
                // It is the display's room rather than the one this hover's layout comes
                // out at, which for a hover waiting on an engine is the corner of the
                // display the spinner was put in (see `PendingLoad::room`).
                let mut show_room: Option<(u32, u32)> = None;
                // A hover is a preview built from nothing — the file is about to be laid out,
                // loaded and played, or the wait for it is — so this is where the engine
                // choices are read, and every question this hover asks after this is answered
                // with what is read here rather than with a row moved since (see
                // `latch_engines_for_the_preview`). Nothing already showing a film comes
                // through here: a hover is refused while a pin is up, and a pin is taken up
                // over a hover that keeps the player it started (see `show_preview`).
                let begins_a_hover = matches!(
                    preview_msg,
                    PreviewMessage::Show(..) | PreviewMessage::ShowKeyboard(..)
                );
                if begins_a_hover {
                    latch_engines_for_the_preview();
                }

                let show_snapshot = begins_a_hover.then(|| preview_msg.clone());

                match preview_msg {
                    PreviewMessage::Show(path, x, y, avoid) => {
                        show_requested = true;
                        // Remember where this preview was opened from: the region
                        // that keeps a scrollable preview alive stretches from
                        // here to the preview, so the pointer can travel between
                        // the two without losing it.
                        set_text_scroll_anchor(x, y);

                        let bounds = work_area_at(x, y);
                        let dpi = dpi_at(x, y);
                        // One reading of the file and one of the configuration for the whole
                        // of this arm: the six questions below are asked of what comes back
                        // rather than of the path, which is what took about twenty
                        // `fs::metadata` calls per hover down to one (see `HoverFacts`).
                        let hover = HoverFacts::read(&path);
                        let follow_cursor = hover.follow_cursor;
                        preview_scale = hover_preview_scale_of(&hover, hover.scales);

                        // A document with no page rendered for it yet has nothing to
                        // measure but the wait, so its preview is laid out as the
                        // spinner's own box (see `office_preview::measure` and the
                        // `libre` arm of `get_media_dimensions`) and that box is placed
                        // at the pointer's own corner rather than a preview's way out
                        // beside it: the page it waits for is laid out again by the replay
                        // that arrives with it, so until then the hover is the wait for the
                        // file under the hand, which belongs at the hand.
                        let waiting_spinner = page_is_on_the_way(&path);

                        // A video whose shape the probe has not answered for yet is the
                        // same kind of wait, and for the same reason: there is nothing to
                        // lay out as a video until the probe answers, so the hover is the
                        // wait for it — the spinner at the pointer's own corner — and is
                        // replayed when the answer lands (see `video_probe_due`).
                        let probing = video_probe_due(&hover);

                        if let Some(orig_dims) = media_dimensions_of(&hover, &path, bounds, dpi) {
                            // A box that is being read is placed at the pointer's own corner the
                            // way every other wait is: what is on screen is the spinner for a
                            // measure the layout has just started, and it belongs at the hand
                            // that asked (see `measure_waiting`).
                            let measuring = measure_waiting(&path);
                            let is_video = hover.is_video();
                            let mut placement = HoverPlacement {
                                orig_dims,
                                avoid,
                                follow_cursor,
                                preview_scale,
                                at_the_pointer_corner: waiting_spinner || probing || measuring,
                            };
                            let placed = compute_mouse_layout(x, y, placement, bounds, dpi);
                            if let Some(layout) = placed {
                                let (layout, text_size) = text_preview_layout(
                                    &hover,
                                    &path,
                                    layout,
                                    bounds,
                                    dpi,
                                    |size| {
                                        compute_mouse_layout(
                                            x,
                                            y,
                                            HoverPlacement {
                                                orig_dims: size,
                                                ..placement
                                            },
                                            bounds,
                                            dpi,
                                        )
                                    },
                                );
                                // A text preview is placed again at the width its box came
                                // out with, and the frame that lands is taller by the rows a
                                // long line wraps into there. The wait re-places from the size
                                // kept here, so it keeps the re-measured one: the display's own
                                // measurement would step the wait around the name for a height
                                // the frame does not have (see `text_preview_layout`).
                                if let Some(size) = text_size {
                                    placement.orig_dims = size;
                                }
                                show_is_video = is_video;
                                show_video_probe = probing;
                                show_measure_probe = measuring;
                                // The display the hover is on, which is what a card of a sound
                                // is painted at and what its clock repaints it at.
                                audio_card_dpi = dpi;
                                show_layout = Some(layout);
                                show_placement = Some(placement);
                                // The wait for this hover is the spinner's own box at
                                // the pointer, whatever the preview's own place came
                                // out at (see `waiting_placement`).
                                show_spinner_layout = compute_mouse_layout(
                                    x,
                                    y,
                                    waiting_placement(placement),
                                    bounds,
                                    dpi,
                                );
                                show_path = Some(path);
                                show_dpi = dpi;
                                show_room = Some(bounds.room());
                            }
                        }
                    }
                    PreviewMessage::ShowKeyboard(path, il, it, ir, ib, avoid, columns) => {
                        show_requested = true;
                        // The focused item lives inside the Explorer window, so
                        // its center resolves to that window's monitor.
                        let center = ((il + ir) / 2, (it + ib) / 2);
                        set_text_scroll_anchor(center.0, center.1);

                        let bounds = work_area_at(center.0, center.1);
                        let dpi = dpi_at(center.0, center.1);
                        // One reading of the file and one of the configuration for the whole of
                        // this arm, as the pointer's arm above does (see `HoverFacts`).
                        let hover = HoverFacts::read(&path);
                        let follow_cursor = hover.follow_cursor;
                        preview_scale = hover_preview_scale_of(&hover, hover.scales);

                        if let Some(orig_dims) = media_dimensions_of(&hover, &path, bounds, dpi) {
                            let is_video = hover.is_video();
                            let placement = KeyboardPlacement {
                                item_rect: (il, it, ir, ib),
                                avoid,
                                columns,
                                orig_dims,
                                follow_cursor,
                                preview_scale,
                            };
                            if let Some(layout) = compute_keyboard_layout(placement, bounds, dpi) {
                                let (layout, _) = text_preview_layout(
                                    &hover,
                                    &path,
                                    layout,
                                    bounds,
                                    dpi,
                                    |size| {
                                        compute_keyboard_layout(
                                            KeyboardPlacement {
                                                orig_dims: size,
                                                ..placement
                                            },
                                            bounds,
                                            dpi,
                                        )
                                    },
                                );
                                show_is_video = is_video;
                                show_video_probe = video_probe_due(&hover);
                                show_measure_probe = measure_waiting(&path);
                                audio_card_dpi = dpi;
                                show_layout = Some(layout);
                                // A keyboard hover has no pointer for a wait to be
                                // placed at, so its spinner is the arc's own box
                                // beside the item, the way its preview is.
                                show_spinner_layout = compute_keyboard_layout(
                                    KeyboardPlacement {
                                        orig_dims: (
                                            office_preview::WAITING_BOX,
                                            office_preview::WAITING_BOX,
                                        ),
                                        preview_scale: PreviewScale::Percent(100),
                                        ..placement
                                    },
                                    bounds,
                                    dpi,
                                );
                                show_path = Some(path);
                                show_dpi = dpi;
                                show_room = Some(bounds.room());
                            }
                        }
                    }
                    PreviewMessage::Hide => {
                        // Invalidate any pending background loads
                        current_generation += 1;
                        pending_load = None;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        let _ = ShowWindow(hwnd, SW_HIDE);
                        // The pin's own window, where it stands in one, goes down with it. The flag
                        // that said which window the pin was in is cleared by the end of the pin,
                        // which happened before this message arrived - so the hide is asked of the
                        // handle itself rather than of the pin's state (see `hide_real_pin_window`).
                        pin_real_window::hide_real_pin_window();
                        clear_pointer_hold();

                        // A document the engine is playing goes with the preview: its
                        // window is its own, so nothing else here takes it down.
                        webview_preview::hide();

                        // Stop video playback if any
                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            if let Some(ref mut media) = *current {
                                media.cancel_background_work();
                                stop_video_playback(media);
                            }
                            *current = None;
                        }
                        current_video_path = None;
                        video_pos = (0, 0, 0, 0);

                        // The page held for the preview that is going away is no
                        // longer being waited on: at a budget of nothing it is
                        // dropped here rather than kept until something else is
                        // rendered to make room for. What is on screen is read from
                        // the show this message takes down — the message itself is a
                        // hide, and `show_path` names this message's own path here.
                        if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                            office_render::hover_ended(path);
                        }

                        current_show = None;
                        page_render_pending = None;
                        page_upgrade = None;
                    }
                    PreviewMessage::Refresh => {
                        render_layered_preview(hwnd);
                    }
                    // A type toggle is answered by the receive loop above, which
                    // is where the kind of the preview on screen is known; it
                    // replays the hover instead of arriving here as itself.
                    PreviewMessage::RefreshTypes => {}
                    // And a row that was switched off, which is answered by the receive loop
                    // above the way a type toggle is: only that side reads what the
                    // configuration now says, and what it does with the answer is take a pin
                    // down where there is one (see `PinChanged`).
                    PreviewMessage::PinChanged => {}
                    // Likewise answered above: a level moved in one of the tray's two `Volume`
                    // halves is given to the pin that is up there, and it is not a hover — a
                    // level is not a thing a layout is made for (see
                    // `PreviewMessage::TrayVolume`).
                    PreviewMessage::TrayVolume { .. } => {}
                    // And a size picked in the tray's `Text → Font Size`, which is a page of text
                    // laid out on the receive loop's own thread for the same reason (see
                    // `PreviewMessage::TrayTextSize`).
                    PreviewMessage::TrayTextSize { .. } => {}
                    // A preview that stopped being a hover: the media stays exactly where it
                    // is, and the window grows around it — the caption above it, and the
                    // transport bar below it where the kind has one. A kind whose chrome is
                    // drawn over its media needs no room for any of it: the window *is* the
                    // media's box, and what is drawn over it is asked for rather than always
                    // there (see `pin_overlay_chrome`).
                    PreviewMessage::Pin { path, rect } => {
                        // A window that has just been taken up is waiting for nothing: a file an
                        // earlier pin was still waiting for a box of is that pin's wait, and this
                        // window is not it (see `pin_awaiting_box`).
                        pin_awaiting_box = None;

                        // The press tracking starts with this pin rather than with the loop.
                        //
                        // It is a local of the loop, so without this it survived every pin
                        // boundary in the life of the process: a press the hook counted while
                        // one pin was up was still "already seen" on the tick after the next
                        // one was taken up, so the first press on a new pin was swallowed and
                        // a drag begun from it never began. Set to what the hook has published
                        // rather than to zero, because zero is a count no press has reached —
                        // asking for a pin that has just come up to handle a press from the
                        // window it replaced is the same bug from the other end.
                        //
                        // The Shell side fixed this exact shape on its own side, where the
                        // equivalent is a destructive drain rather than a snapshot, so that
                        // there is no state to go stale between two things.
                        engine_press_seen = pin_media_press_count();
                        pin_press_seen = pin_media_press_count();

                        // And it is owed no hover's wait either, which is the one it inherits.
                        // A pin is taken up out of a hover, and a hover that was still loading
                        // when the key was pressed leaves its wait standing here — where nothing
                        // will ever take it up, because a pin refuses hovers at every door and
                        // this loop drops their messages (see `hover_is_shown`). Left armed it
                        // is not merely useless: a file the user picks in the listing is held
                        // behind it and put back every tick, so `Pin Mode → Update Preview` looks
                        // broken rather than absent, and the first click after every take-up is
                        // the one it eats.
                        //
                        // The generation moves with it so an answer still on its way is dropped
                        // rather than installed over a window that is now a pin — the same
                        // teardown a `Hide` performs, which is what this is: the hover is over.
                        current_generation += 1;
                        pending_load = None;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        // A document, a specimen and a page of HTML are on screen in the engine's
                        // own window with no media of this app's behind them, so what the hover
                        // left in the slot is nothing at all — or the spinner that stood in for
                        // the page while it was on its way, which its landing never clears. A pin
                        // built on either would be a window with no kind in it: no frame, no
                        // transport, and a liveness read that takes the window straight back down
                        // (see `pin_media_is_alive`). So the kind is installed here, before it is
                        // read, which is the same handover the swap path performs for the same
                        // file (see `swap_pinned_media`) and the loop's own handover performs for
                        // a hover. It goes in only where the engine draws the file, and only where
                        // the slot does not already say so: a file this app draws has a frame of
                        // its own there, and a swap has just installed the very kind.
                        if !current_media_type().is_some_and(|kind| kind.is_engine()) {
                            if let Some(media) = pinned_engine_media(&path) {
                                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                    *current = Some(media);
                                }
                            }
                        }

                        // Whatever the pin was being pressed for, it is being pressed for no
                        // longer: the state below replaces it whole, and the pointer a press on
                        // this window took goes with the drag it was taken for. It is let go
                        // here rather than left to the release that is never coming, and it is
                        // let go before the pin's own lock is taken rather than inside it,
                        // because `ReleaseCapture` delivers `WM_CAPTURECHANGED` and the window
                        // procedure asks for that same lock (see the note in
                        // `take_up_pinned_window`, and `pinned_release` for the release this
                        // stands in for).
                        release_pin_capture(hwnd);

                        // The pin itself, built whole — including the caption's own room, the
                        // frame, and the transport the new file is to be played by. Both roads
                        // that reach this arm reach it through the same call, so a window shown
                        // another file is the same window as one shown its first (see
                        // `take_up_pinned_window`). Whether a pin is already up tells the second
                        // road from the first: a swap installs through here with the pin standing,
                        // and a take-up from a hover arrives with none (see `settle_pin_handover`).
                        let was_up = pinned();
                        let pin = take_up_pinned_window(&path, rect);
                        let content = pin.content;

                        // The kind on screen, read back rather than carried out of the take-up:
                        // it is the two kinds below whose media has to be laid out again now
                        // that the pin is up, and it is the kind that file was installed as.
                        let kind = CURRENT_MEDIA
                            .lock()
                            .ok()
                            .and_then(|media| media.as_ref().map(|media| media.media_type));

                        // Whether this pin stands in a window of its own is decided here, once,
                        // from the kind it was taken up on — or from the file's route where the
                        // media is a placeholder — and never asked of the media again: every kind
                        // this app draws the whole way is in a window Windows activates and
                        // lists, and the kinds whose band holds somebody else's window keep the
                        // preview's (see `pin_should_stand_in_its_own_window`).
                        pin_stands_in_its_own_window(pin_should_stand_in_its_own_window(
                            kind, &path,
                        ));

                        // The pin is published as up inside this, after the state it publishes
                        // is written (see `pin_window::install`).
                        install(pin);

                        // A pin that is taken up owes the front from the moment it is up: the
                        // click that picks the next file activates the listing, and a pin in the
                        // ordinary band is that window over it for a frame however fast the pick
                        // is answered (see `PIN_FRONT_OWED`).
                        arm_pin_to_the_front();

                        // The button a collapsed pin left says which file the pin is of, and the
                        // file has just changed: a stand-in already up is not shown again by the
                        // take-up below (`show_pinned_window` is refused for a collapsed pin), so
                        // the name is written here or the button names the file before the step
                        // (see `pin_taskbar::refresh_pin_taskbar_title`).
                        pin_taskbar::refresh_pin_taskbar_title();

                        // A pin comes up under a pointer that may be anywhere, including inside a
                        // name being renamed, so it may take the focus but is not given it until
                        // the hand presses it (see `pin_set_focusable`). The pin holds no keyboard
                        // at all until then, so there is nothing here to clear.
                        //
                        // **And a pin that stands in a window of its own leaves the preview's alone.**
                        // The hover's window must never be activatable — it appears under a pointer
                        // that may be inside a name being renamed — and the pin's own window was made
                        // with the styles that let Windows activate it (see `pin_real_window`).
                        if !pin_is_its_own_window() {
                            pin_set_focusable(hwnd, true);
                        }

                        // **A player adopted without its film's copied subtitles is begun
                        // again here**, where the copy is ready now and the adopted player is
                        // one the hover began before it existed — the corner the ready-message
                        // cannot reach, because it has already been and gone (see
                        // `reload_adopted_subtitles`). A pin taken up on a film that is still
                        // being copied is not this: the message will find the pin when it
                        // lands, and there is nothing to draw yet in any case.
                        reload_adopted_subtitles(&path);

                        // A text preview is the one kind a pin *changes* rather than frames: it
                        // comes up in full mode, which is the scrollbar, the selection, and the
                        // keys that copy it all out (see `current_text_options`). That is the
                        // media laid out again, and it can only be asked for once the pin is up,
                        // because being pinned is the whole of what full mode is read from.
                        //
                        // A sound's card is the other: a pin is measured again at the take-up (see
                        // `pinned_audio_card_box`) and the card has to be drawn into the box that
                        // came of it rather than into the hover's, or the window is one box and
                        // the card another.
                        if kind == Some(MediaType::Text) {
                            if let Some((path, dpi)) = pinned_media_owner() {
                                // A take-up is a box this window has not been given before, so
                                // the relayout asked for here is asked for the file the pin is
                                // now showing, and any older one is dropped with it.
                                pin_relayout = relayout_pinned_media(
                                    &path,
                                    content,
                                    dpi,
                                    None,
                                    PinRelayoutRoad::TakeUp,
                                );
                            }
                        } else if kind == Some(MediaType::Audio) {
                            if let Some((path, _)) = pinned_media_owner() {
                                // A sound a pin comes up on is the sound the hover was already
                                // playing: no player is begun for it here, so where the rule it
                                // goes round by is not the pin's own the pin begins it again
                                // where it is — a player that takes the file round on its own has
                                // no pass end for the pin's `End of File` to be read at (see
                                // `take_up_pinned_audio`).
                                take_up_pinned_audio(
                                    &path,
                                    &mut audio_started,
                                    &mut audio_start_offset,
                                    audio_paused,
                                );

                                let card = AudioCardClock {
                                    started: audio_started,
                                    from: audio_start_offset,
                                    paused: audio_paused,
                                    name_offset: audio_name_scroll
                                        .as_ref()
                                        .map(|scroll| scroll.offset())
                                        .unwrap_or(0),
                                    dpi: audio_card_dpi,
                                    chrome: pinned_audio_chrome(audio_paused),
                                };

                                pin_relayout = relayout_pinned_media(
                                    &path,
                                    content,
                                    audio_card_dpi,
                                    Some(card),
                                    PinRelayoutRoad::TakeUp,
                                );
                            }
                        }

                        // **A pin that was away comes back here, where the file is on the window.**
                        // The take-up above has just given the window the box the file's own kind
                        // asks for, and the show behind this paints before it appears — so the return
                        // owed at the pick is answered here and nowhere earlier, and what comes up is
                        // the file the update named rather than the one before it (see
                        // `bring_a_put_away_pin_back` and `pin_return_owed`).
                        //
                        // Everything else that reaches this arm — a file a walk landed on, a pin that
                        // was never away, a first take-up — is owed the show alone, which is the
                        // paint of the new file. The debt is read once and let go of either way: this
                        // is the take-up it was owed for.
                        let came_back = pin_return_owed && bring_a_put_away_pin_back(&path);
                        pin_return_owed = false;
                        if !came_back {
                            show_pinned_window(hwnd);
                        }

                        // A film adopted from a hover is told whose setting governs its window
                        // now: the hover's topmost band is the preview's, and the pin's own
                        // `Always On Top` is the answer from here on. Asked after the show rather
                        // than at the take-up, because a take-up that defers its showing keeps the
                        // hover up and playing behind it — and a hover still on screen keeps the
                        // band it was given (see `pin_film_host::settle_pin_handover`).
                        settle_pinned_mpv_ontop();

                        // And a re-take-up that lands on a hosted film already drawing is a step's
                        // handover rather than an arrival: the swap held the outgoing frame with
                        // the chrome hidden until the new player's first frame, so the bars start
                        // now, over the new film (see `PinSwapHold`).
                        if was_up
                            && pin_is_its_own_window()
                            && pin_film_host::the_pinned_player_was_begun_in_the_host()
                            && pin_film_host::the_hosted_player_is_drawing(
                                VIDEO_PID.load(Ordering::SeqCst),
                            )
                        {
                            with_pin(|pin| pin.chrome = PinChrome::on_arrival(Instant::now()));
                        }

                        place_pinned_siblings();
                        publish_pointer_hold(pinned_window());
                    }
                    // The pinned window was given another box — maximized, restored, resized,
                    // or put back on a display that changed: the media is laid out again at the
                    // size it is now drawn at, and the window is put up around the result.
                    PreviewMessage::PinBox(content) => {
                        // A window given another box has had its media laid out again under it, and
                        // a volume popup floating over that media belongs to the box it was opened
                        // in: it is put away rather than left over a picture that has moved out from
                        // under it.
                        close_pin_volume();

                        let card = AudioCardClock {
                            started: audio_started,
                            from: audio_start_offset,
                            paused: audio_paused,
                            name_offset: audio_name_scroll
                                .as_ref()
                                .map(|scroll| scroll.offset())
                                .unwrap_or(0),
                            dpi: audio_card_dpi,
                            chrome: pinned_audio_chrome(audio_paused),
                        };

                        if let Some((path, dpi)) = pinned_media_owner() {
                            // A box change is what a relayout is asked for, so this is where one
                            // starts — and where an older one is dropped: the newest box is the
                            // only one worth decoding for, and the frame already decoded is
                            // drawn scaled into this one meanwhile.
                            pin_relayout = relayout_pinned_media(
                                &path,
                                content,
                                dpi,
                                Some(card),
                                PinRelayoutRoad::BoxChange,
                            );
                        }
                        // A page is the one kind a box is *laid out* for rather than scaled into, so
                        // the box it has just been given is the box it is drawn for now — ahead of the
                        // window being put up around it, and ahead of the answer the load above will
                        // bring. It is the kind that answers that, and it answers for no other, which
                        // leaves every other kind drawn into the new box from the frame in hand (see
                        // `relayout_text_for_the_box`).
                        relayout_text_for_the_box(content);
                        show_pinned_window(hwnd);
                        place_pinned_siblings();
                        publish_pointer_hold(hwnd);
                    }
                    // Likewise answered above: a rendered page replays the hover
                    // it belongs to rather than being handled as a message here.
                    PreviewMessage::OfficeRenderReady { .. } => {}
                    // And a probe's answer, which replays the hover that was waiting
                    // on it the same way.
                    PreviewMessage::VideoProbed { .. } => {}
                    // And a measured box, which replays the hover that was waiting on it in
                    // the same way — or takes its wait down, where the answer is that there is
                    // no box to be had.
                    PreviewMessage::MeasureProbed { .. } => {}
                    // And an engine's, which is the page-shaped answer of a picture
                    // rather than of a page: it replays the hover that was waiting on
                    // it exactly as the render tier's answer does.
                    PreviewMessage::MagickReady { .. } => {}
                    // And the listing an engine produced, which is the same shape of answer
                    // once more: a page's worth of content arriving for the hover that asked
                    // for it, replayed rather than handled as a hover here.
                    PreviewMessage::PeazipReady { .. } => {}
                    // And a film's copied subtitles arriving, which a pinned window is begun
                    // again for above rather than a hover being laid out for (see
                    // `PreviewMessage::VideoSubtitlesReady`).
                    PreviewMessage::VideoSubtitlesReady(_) => {}
                    // And the file a pinned window is to be shown instead of the one it has,
                    // which is answered above: it is not a hover, so nothing here has a layout
                    // to make for it (see `PreviewMessage::PinUpdate`).
                    PreviewMessage::PinUpdate(_) => {}
                    PreviewMessage::PinPress { .. } | PreviewMessage::PinNoPick { .. } => {}
                    // And what the planner answered, which is the pin's own walk and its
                    // hand-off name rather than a hover — taken where a pin is acted on, not
                    // here (see `PreviewMessage::PinAnswered`).
                    PreviewMessage::PinAnswered(_) => {}
                }

                // Shared load/display logic for Show and ShowKeyboard
                if let (Some(path), Some(layout)) = (show_path, show_layout) {
                    let pos_x = layout.pos_x;
                    let pos_y = layout.pos_y;
                    let media_width = layout.preview_w as i32;
                    let media_height = layout.preview_h as i32;
                    let max_width = layout.max_width;
                    let max_height = layout.max_height;
                    let preview_w = layout.preview_w;
                    let preview_h = layout.preview_h;

                    // The room an engine-drawn preview is asked for is the room the display
                    // has rather than the one this layout came out at: a hover waiting on an
                    // engine is laid out as the spinner's own box at the pointer, and the
                    // room that layout comes out at is the corner the spinner was put in —
                    // which for a picture developed at the size it is shown at would be a
                    // ceiling on the size it could ever be drawn at (see `PendingLoad::room`).
                    let room = show_room.unwrap_or((max_width, max_height));

                    // The box the wait goes in while the load runs: the spinner's own
                    // place, which is not the preview's. A hover with no room for the
                    // spinner anywhere is answered with the preview's own box rather
                    // than with no spinner at all.
                    let (spinner_x, spinner_y, spinner_side) = match show_spinner_layout {
                        Some(spinner) => (spinner.pos_x, spinner.pos_y, spinner.preview_w),
                        None => (pos_x, pos_y, preview_w),
                    };

                    // A page that arrived for the hover already on screen is an
                    // upgrade: what is there — the spinner — stays up while the page
                    // is loaded, and is replaced when it lands. A hover replayed for a
                    // video's probe, or for a box measured off this thread, is the same thing
                    // reached another way: it is the wait it was already in, carried on.
                    let upgrading = page_upgrade.as_deref() == Some(path.as_path())
                        || video_replay.as_deref() == Some(path.as_path())
                        || measure_replay.as_deref() == Some(path.as_path());
                    page_upgrade = None;
                    video_replay = None;
                    measure_replay = None;

                    // A text or archive preview is rendered at the size the
                    // layout planned for it: it is painted at a fixed font
                    // size, so the planned box is the box it draws into rather
                    // than a space to be scaled within — and what the window
                    // is sized to is the frame that comes back. Every other
                    // format is loaded against the free space it may be
                    // scaled within.
                    let painted = page_is_painted(&path);
                    let (load_width, load_height) = if painted {
                        (preview_w, preview_h)
                    } else {
                        (max_width, max_height)
                    };

                    if show_snapshot.is_some() {
                        current_show = show_snapshot.clone();
                    }

                    // Which road a video takes is the route, and it is asked at the fork below
                    // rather than here: this branch is the wait for its probe, and the probe is
                    // what asks the route on a thread of its own (see `spawn_video_probe`) — so by
                    // the time the fork is reached the file-shaped half of the answer is held, and
                    // asking it here would open the media engine over the file on the thread that
                    // draws, for a hover that is only going to wait.
                    //
                    // The two roads are the two they always were: the engine's frames come back
                    // through the ordinary load and are drawn by this app's own window — which is
                    // what a pin of one is resized, maximized and dragged by — while a player in a
                    // window of its own is the branch below, which is both of the players this app
                    // can run: FFmpeg's and mpv's (see `video_route`). A video neither of them
                    // will take is answered by that load with no media at all, rather than by a
                    // third branch here (see `load_video_thumbnail`).
                    if show_video_probe || show_measure_probe {
                        // The hover is waiting on a probe: nothing of the file can be
                        // laid out or loaded until there is a shape or a box to lay it out with,
                        // so what is put up is the wait every other preview is given —
                        // the spinner at the pointer — and the hover it came from is
                        // replayed when the answer lands, which is when there is a video
                        // to load or a box to lay out (see `video_probe_due`, `video_probe`
                        // and `measure_waiting`). The probe itself runs on a thread of its
                        // own: for a video it is two external processes, and for a box it is a
                        // read of the file, and this is the thread that draws the wait.
                        //
                        // A replay carries the wait it was already in rather than opening
                        // a new one: the same clock — so a probe answered inside
                        // `spinner_delay_ms` does not start that delay over — and the same
                        // spinner, which is on screen already.
                        let waiting = pending_load.take();
                        let (started, spinner_shown, upgrade) = match (upgrading, waiting) {
                            (true, Some(pl)) => (pl.started, pl.spinner_shown, true),
                            _ => (Instant::now(), false, false),
                        };

                        current_generation += 1;
                        let gen = current_generation;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        if !upgrade {
                            // Another preview takes over here, so what is on screen goes
                            // as the wait for the probe goes up — the frame this app's
                            // window holds, the player a previous hover started, and the
                            // window a document the engine draws was put in.
                            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut media) = *media_guard {
                                    media.cancel_background_work();
                                    stop_video_playback(media);
                                }
                                // Clear immediately so old pixels never flash while
                                // the new target is being measured.
                                *media_guard = None;
                            }

                            if current_video_path.is_some() {
                                current_video_path = None;
                                video_pos = (0, 0, 0, 0);
                            }

                            webview_preview::hide();

                            let _ = ShowWindow(hwnd, SW_HIDE);
                        }

                        pending_load = Some(PendingLoad {
                            generation: gen,
                            hide_epoch: hidden_epoch(),
                            path: path.clone(),
                            started,
                            pos_x,
                            pos_y,
                            width: preview_w,
                            height: preview_h,
                            room,
                            spinner_shown,
                            spinner_delay: load_spinner_delay(),
                            spinner_pos: (spinner_x, spinner_y),
                            spinner_side,
                            placement: show_placement,
                            upgrade,
                            // The probe a video waits on is marked as an engine's wait is:
                            // what it is waiting for is outside this side, and what is read
                            // against this flag is the cap on a wait that nothing else ends
                            // (see `awaiting_engine`). A box measured off this thread is
                            // not marked — what it waits for is a read that finishes
                            // (see `measured_off_the_tick`).
                            awaiting_engine: show_video_probe,
                        });
                        // The probe a hover is waiting on is the one this message asked for: a
                        // measured box started its own thread where it was measured, so what is
                        // left to start here is a video's (see `measured_off_the_tick`).
                        if show_video_probe {
                            video_probe = Some((path.clone(), gen));
                            spawn_video_probe(path, gen);
                        }
                    } else if show_is_video
                        && matches!(video_route(&path), VideoRoute::Ffplay | VideoRoute::Mpv)
                    {
                        // Cancel any in-flight image load before switching to video.
                        current_generation += 1;
                        pending_load = None;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        let no_cancel = Arc::new(AtomicBool::new(false));
                        if let Some(media_data) = load_media(
                            &path,
                            load_width,
                            load_height,
                            preview_scale,
                            show_dpi,
                            // A hover's own load: no page's reader has sized this one, so what it is
                            // drawn at is what the configuration says (see `load_media`).
                            None,
                            no_cancel,
                        ) {
                            // Nothing of this app's is on screen for a video — the
                            // player draws it in a window of its own — so what is put
                            // up while the player starts is the wait every other kind
                            // of preview is given: the spinner at the pointer, and it
                            // goes the moment the player's window is there (see
                            // `video_start` and `player_wait`). A hover replayed for a
                            // probe is already that wait, and it stays where it is.
                            if !upgrading {
                                let _ = ShowWindow(hwnd, SW_HIDE);
                            }

                            let process_running = is_video_process_running();
                            let should_start =
                                current_video_path.as_ref() != Some(&path) || !process_running;

                            if should_start {
                                if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                    if let Some(ref mut media) = *media_guard {
                                        media.cancel_background_work();
                                        stop_video_playback(media);
                                    }
                                }

                                // The previous ffplay may have survived its stop
                                // (dropped handle or an unconfirmed kill): kill it
                                // before a new one takes the screen.
                                kill_stray_video_process();

                                let video_process = start_video_playback(
                                    &path,
                                    pos_x,
                                    pos_y,
                                    media_width,
                                    media_height,
                                    0.0,
                                    false,
                                    current_video_volume(),
                                    // A hover shows the file the way the player chooses to, which
                                    // is its first subtitle stream and the one any relaunch
                                    // reaches again by the same route. Only a pinned window is
                                    // remembering a track of its own (see `next_subtitle`).
                                    None,
                                );
                                let pid = video_process
                                    .as_ref()
                                    .map(|player| player.id())
                                    .unwrap_or(0);

                                current_video_path = Some(path.clone());
                                video_pos = (pos_x, pos_y, media_width, media_height);
                                let _ = ensure_video_window_topmost(
                                    pos_x,
                                    pos_y,
                                    media_width,
                                    media_height,
                                );

                                let mut data = media_data;
                                data.video_process = video_process;

                                if pid != 0 {
                                    // The player is a process with a window to create
                                    // before anything of the file is on screen, and
                                    // there is nothing under the spinner that could
                                    // arrive sooner — a player that starts instantly is
                                    // the only start this wait is not seen for — so it
                                    // is shown from the first tick rather than after
                                    // `spinner_delay_ms`: a frame of the spinner is what
                                    // a video hover would otherwise spend showing the
                                    // desktop.
                                    video_start = Some(VideoStart {
                                        media: data,
                                        path: path.clone(),
                                        pid,
                                        started: Instant::now(),
                                    });
                                    pending_load = Some(PendingLoad {
                                        generation: current_generation,
                                        hide_epoch: hidden_epoch(),
                                        path: path.clone(),
                                        started: Instant::now(),
                                        pos_x,
                                        pos_y,
                                        width: media_width as u32,
                                        height: media_height as u32,
                                        room,
                                        spinner_shown: false,
                                        spinner_delay: Duration::ZERO,
                                        spinner_pos: (spinner_x, spinner_y),
                                        spinner_side,
                                        placement: show_placement,
                                        upgrade: false,
                                        awaiting_engine: false,
                                    });
                                } else if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                    // A player that would not start is no preview:
                                    // nothing of this app's goes up for one.
                                    *current = Some(data);
                                }
                            } else {
                                video_pos = (pos_x, pos_y, media_width, media_height);
                                let _ = ensure_video_window_topmost(
                                    pos_x,
                                    pos_y,
                                    media_width,
                                    media_height,
                                );
                            }
                        }
                    } else {
                        // For images/animations, load async
                        if current_video_path.is_some() {
                            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut media) = *media_guard {
                                    media.cancel_background_work();
                                    stop_video_playback(media);
                                }
                            }
                            current_video_path = None;
                            video_pos = (0, 0, 0, 0);
                        }

                        // Any preview the engine does not draw is drawn here, so its window
                        // — if one is still up — comes down as this one goes up.
                        if engine_kind_of(&path).is_none() {
                            webview_preview::hide();
                        }

                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }
                        if !upgrading {
                            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut media) = *media_guard {
                                    media.cancel_background_work();
                                    // What is being replaced goes with it, and for a sound
                                    // that is more than bookkeeping: the engine playing one
                                    // is this app's own and this thread's alone, so a media
                                    // dropped here without this call plays on with nothing on
                                    // screen that could stop it. The take-down the `Hide`
                                    // would have performed is not a second chance at it: a
                                    // keyboard preview replacing a hovered one sends its
                                    // `Hide` and its show in the same breath, and the drain
                                    // above keeps only the newest of them (see the collapse
                                    // there). Every other branch that replaces a preview
                                    // stops what it replaces; this one did not.
                                    stop_video_playback(media);
                                }
                                // Clear immediately so old pixels never flash while
                                // the new target is being decoded.
                                *media_guard = None;
                            }
                            let _ = ShowWindow(hwnd, SW_HIDE);
                        }

                        // Start background load; the spinner follows if the wait
                        // turns out to be worth showing (see `spinner_due`).
                        current_generation += 1;
                        let gen = current_generation;
                        let load_cancel = Arc::new(AtomicBool::new(false));
                        pending_load_cancel = Some(Arc::clone(&load_cancel));
                        pending_load = Some(PendingLoad {
                            generation: gen,
                            hide_epoch: hidden_epoch(),
                            path: path.clone(),
                            started: Instant::now(),
                            pos_x,
                            pos_y,
                            width: preview_w,
                            height: preview_h,
                            room,
                            spinner_shown: false,
                            spinner_delay: load_spinner_delay(),
                            spinner_pos: (spinner_x, spinner_y),
                            spinner_side,
                            placement: show_placement,
                            upgrade: upgrading,
                            awaiting_engine: false,
                        });

                        queue_load_request(
                            &load_request_slot,
                            LoadRequest {
                                generation: gen,
                                path,
                                max_width: load_width,
                                max_height: load_height,
                                preview_scale,
                                dpi: show_dpi,
                                cancel: Arc::clone(&load_cancel),
                            },
                        );
                    }
                } else if show_requested {
                    // A newer hover target could not produce a layout/path. Treat it
                    // like a hide so stale async loads cannot resurrect old previews.
                    current_generation += 1;
                    pending_load = None;
                    clear_load_request(&load_request_slot);
                    if let Some(cancel) = pending_load_cancel.take() {
                        cancel.store(true, Ordering::Release);
                    }

                    let _ = ShowWindow(hwnd, SW_HIDE);
                    webview_preview::hide();

                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        if let Some(ref mut media) = *current {
                            media.cancel_background_work();
                            stop_video_playback(media);
                        }
                        *current = None;
                    }
                    current_video_path = None;
                    video_pos = (0, 0, 0, 0);
                    current_show = None;
                    page_render_pending = None;
                    page_upgrade = None;
                }
            } else if refresh_requested {
                render_layered_preview(hwnd);
            }

            // Ctrl+C over a text preview copies what is selected in it. The key
            // is polled rather than waited for: a preview the user has not pressed
            // is a window nobody is in, so it would never receive the keystroke as
            // a message. Selecting everything is the menu's `Select All`, not a key
            // of its own.
            if text_preview_copy_requested() {
                copy_text_preview(hwnd);
            }

            // The cap on waiting for a page is read against the hover on screen
            // rather than remembered: a render that outlives its hover is dropped
            // here and its page left in the cache — for Office's tier as much as for
            // the render engine, whose conversion runs on and is kept the same way.
            let shown_path = current_show.as_ref().and_then(show_path).cloned();

            // How long the hover on screen has been waiting, which is what the cap is
            // read of, and whether what it is waiting on is an engine at all.
            let (waited, awaiting_engine) = pending_load
                .as_ref()
                .map(|pl| {
                    (
                        pl.started.elapsed() >= Duration::from_secs(OFFICE_RENDER_WAIT_SECS),
                        pl.awaiting_engine,
                    )
                })
                .unwrap_or((false, false));

            // An engine that draws the hover's page is asked to come up once the pointer has
            // settled on the file: what the ask overlaps is the rest of the wait — a launch is
            // a second or more — and a file the pointer is merely crossing costs nothing, since
            // the ask is not made until the hover has outlasted the settle (see `WARM_SETTLE_MS`
            // and `warm_engines_for`). Once per hover: an engine is asked for the file under the
            // hand, and nothing is gained by asking again on every tick.
            if warmed_generation != Some(current_generation) {
                if let Some(pl) = pending_load
                    .as_ref()
                    .filter(|pl| !pl.upgrade && pl.started.elapsed() >= WARM_SETTLE_MS)
                {
                    warmed_generation = Some(current_generation);
                    warm_engines_for(&pl.path);
                }
            }

            let render_wait = page_render_pending.as_ref().map(|(path, generation)| {
                *generation == current_generation && shown_path.as_deref() == Some(path.as_path())
            });

            // A request made for a hover that has gone is not waited for any longer: the
            // page it produces is kept by the engine and the next hover of the file reads
            // it. A wait the hover on screen is still in is left standing.
            if render_wait == Some(false) {
                page_render_pending = None;
            }

            // The preview has waited as long as it waits — a page that has not arrived by
            // now may still be coming, a very large document takes as long as it takes —
            // so the spinner comes down and the hover is left to itself. The render is not
            // abandoned with it: it runs on and its page is cached, so the next hover of
            // that file shows it. Only the engine itself can say a render failed, and it
            // remembers that for the file.
            //
            // What is read here is the wait rather than the request. An engine answers a
            // file once, so a hover whose request was folded into one already in flight is
            // never answered on its own — its request has been read and left, and a cap
            // that was read of the request would be no cap at all: the spinner would stand
            // there for good (see `awaiting_engine`).
            if waited && awaiting_engine {
                page_render_pending = None;
                // The cap is the loop giving up rather than the engine answering, so the wait
                // ends as the mark every failure ends as: a page still coming is the file the
                // next hover reads, and what the hand is left with is an answer about this one
                // (see `take_the_wait_as_failed`).
                take_the_wait_as_failed(hwnd, &mut pending_load);
                if let Some(cancel) = pending_load_cancel.take() {
                    cancel.store(true, Ordering::Release);
                }
            }

            // Keep the pointer region in step with the window rather than only
            // with the paints: the window is moved when a frame is installed, and
            // the Explorer hook reads this on every one of its own ticks. Throttled
            // (see `POINTER_HOLD_HEARTBEAT_MS`): the region only changes on
            // show/move/resize/swap/hide — a swap bumps `current_generation`, which
            // is part of the key — so a slow heartbeat plus an immediate publish on
            // transitions is the whole of what the hook needs, without the per-tick
            // locks and `GetWindowRect`.
            let hold_key = (
                current_show.is_some(),
                pending_load.is_some(),
                pinned(),
                current_generation,
            );
            if hold_key != last_hold_key
                || last_hold_publish.elapsed() >= Duration::from_millis(POINTER_HOLD_HEARTBEAT_MS)
            {
                publish_pointer_hold(hwnd);
                last_hold_publish = Instant::now();
                last_hold_key = hold_key;
            }

            if current_show.is_none() && pending_load.is_none() {
                // Nothing is on screen, which is where this thread used to spend
                // its whole life waking sixty times a second to drain a queue that
                // was empty and publish a region that did not exist. There is
                // nothing to animate, nothing to repaint and nothing left to keep
                // in step, so the wait becomes the preview channel itself: a hover
                // is answered as it arrives rather than on the next tick, and the
                // interval is only a ceiling on how long a window message — a
                // resume, a display change — waits to be noticed. The wait wakes
                // on window input at once, so a drag never queues behind it (see
                // `wait_preview_channel`).
                carried_preview_msg =
                    wait_preview_channel(&rx, wait_before_the_next_tick(true, false, pinned()));
            } else {
                // Something is on screen, and which band of the cadence it lands in is
                // `wait_before_the_next_tick`'s question — a tick that has something to
                // animate keeps the frame cadence, a static picture waits on the channel, and a
                // pinned static one keeps a shorter ceiling so the caption's buttons stay
                // snappy. A drag is dispatched at the pointer's pace either way since the wait
                // wakes on input (see `wait_preview_channel`).
                //
                // A wait that ended, a generation that moved or a pin that came up
                // or down may have installed another kind behind the hint: classify
                // it again on the next tick rather than waiting for the periodic
                // refresh, so an animation never starts half a second late.
                let tick_has_wait = pending_load.is_some()
                    || pin_load.is_some()
                    || pin_walk_wait.is_some()
                    || video_start.is_some()
                    || first_frame_wait.is_some();
                if (tick_had_wait && !tick_has_wait)
                    || current_generation != tick_generation
                    || pinned() != tick_pinned
                {
                    media_dynamic_hint = true;
                }
                // A travel of the wheel's is a picture with something to animate, so it keeps the
                // frame cadence for as long as it lasts and the loop drops back to its own wait the
                // tick after it has arrived (see `pin_zoom_in_flight`).
                let travelling = pin_zoom_in_flight();
                let wait_ms = wait_before_the_next_tick(
                    false,
                    media_dynamic_hint || tick_has_wait || travelling,
                    pinned(),
                );
                // And a travel is paced by the compositor rather than by that wait, because what its
                // frames are worth is the display's own period and not a timer's: a sleep of
                // `FRAME_WAIT_MS` comes back in 27.8 ms — 36 frames a second — where a flush of the
                // compositor comes back in the 6.94 ms of a 144 Hz display's frame (see
                // `wait_for_the_compositor`).
                carried_preview_msg = if travelling {
                    wait_for_the_compositor(&rx, wait_ms)
                } else {
                    wait_preview_channel(&rx, wait_ms)
                };
            }
        }

        // Signal the dedicated loader worker to stop and wait for shutdown.
        if let Some(cancel) = pending_load_cancel.take() {
            cancel.store(true, Ordering::Release);
        }
        clear_load_request(&load_request_slot);
        let (_, cvar) = &*load_request_slot;
        cvar.notify_all();
        let _ = load_worker.join();
    }
}

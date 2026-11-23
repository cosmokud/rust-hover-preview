//! Owned per-loop session state, RAII COM apartment guard, and generation-cached
//! configuration snapshot for the Explorer hook loop.

use super::*;
use crate::config::config::AppConfig;

/// RAII guard for the Explorer hook thread's STA COM apartment.
/// Calls `CoUninitialize` on drop only when `CoInitializeEx` succeeded (`S_OK` or `S_FALSE`).
pub(super) struct ComGuard {
    initialized: bool,
}

impl ComGuard {
    /// Returns true iff `CoInitializeEx` succeeded and owns an apartment reference to release.
    pub(super) fn should_uninitialize(hr: i32) -> bool {
        hr >= 0
    }

    pub(super) fn init_sta() -> Self {
        let hr = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };
        Self {
            initialized: Self::should_uninitialize(hr.0),
        }
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        if self.initialized {
            unsafe {
                CoUninitialize();
            }
        }
    }
}

/// Named configuration fields consumed by the Explorer hook loop and `PinUpdateWatch`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct HookConfigSnapshot {
    pub(super) preview_enabled: bool,
    pub(super) hover_delay_ms: u64,
    pub(super) trigger_key_mode: TriggerKeyMode,
    pub(super) same_file_rehover_delay_ms: u64,
    pub(super) settling_delay_ms: u64,
    pub(super) prioritize_keyboard: bool,
    pub(super) trigger_key_enabled: bool,
    pub(super) trigger_key_affect_pin_mode: bool,
    pub(super) tick_ms: u64,
    pub(super) pin_enabled: bool,
    pub(super) pin_update_enabled: bool,
    pub(super) pin_update_on_hover: bool,
}

impl HookConfigSnapshot {
    pub(super) fn from_config(c: &AppConfig) -> Self {
        Self {
            preview_enabled: c.preview_enabled,
            hover_delay_ms: c.hover_delay_ms,
            trigger_key_mode: c.trigger_key_mode,
            same_file_rehover_delay_ms: c.same_file_rehover_delay_ms,
            settling_delay_ms: c.settling_delay_ms,
            prioritize_keyboard: c.prioritize_keyboard,
            trigger_key_enabled: c.trigger_key_enabled,
            trigger_key_affect_pin_mode: c.trigger_key_affect_pin_mode,
            tick_ms: c.tick_ms,
            pin_enabled: c.pin_enabled,
            pin_update_enabled: c.pin_update_enabled,
            pin_update_on_hover: c.pin_update_on_hover,
        }
    }
}

fn publish_listing_switches(c: &AppConfig) {
    // Low-level hooks read these lock-free atomics instead of locking `CONFIG`.
    crate::app::seam::publish_browse_dialog_previews(c.browse_dialog_preview_enabled);
    crate::app::seam::publish_everything_previews(c.everything_previews);
    crate::app::seam::publish_q_dir_previews(c.q_dir_previews);
}

/// Generation-cached `HookConfigSnapshot`: zero locks on steady-state ticks; refreshes once
/// whenever `seam::config_generation()` advances.
pub(super) struct CachedHookConfig {
    generation: u64,
    snapshot: HookConfigSnapshot,
}

impl CachedHookConfig {
    pub(super) fn capture_initial() -> Self {
        let generation = crate::app::seam::config_generation();
        let cfg = crate::config::snapshot();
        publish_listing_switches(&cfg);
        Self {
            generation,
            snapshot: HookConfigSnapshot::from_config(&cfg),
        }
    }

    pub(super) fn current(&mut self) -> HookConfigSnapshot {
        let generation = crate::app::seam::config_generation();
        if generation != self.generation {
            let cfg = crate::config::snapshot();
            publish_listing_switches(&cfg);
            self.snapshot = HookConfigSnapshot::from_config(&cfg);
            self.generation = generation;
        }
        self.snapshot
    }
}

/// Pure owned state machine for `run_explorer_hook` hover, keyboard, folder-gate, scroll,
/// and engine-page-drag state. Transition methods clear all dependent fields atomically so
/// stale state cannot survive a reset.
pub(super) struct HookSession {
    pub(super) last_file: Option<PathBuf>,
    pub(super) hover_item_box: Option<(i32, i32, i32, i32)>,
    pub(super) suppressed: SuppressedHover,
    pub(super) pointer_pause: KeyboardPointerPause,
    pub(super) hover_start: Option<Instant>,
    pub(super) last_cursor_pos: POINT,

    pub(super) keyboard_file: Option<PathBuf>,
    pub(super) last_focused_key: Option<FocusedItemKey>,
    pub(super) last_focused_frame: Option<isize>,
    pub(super) is_keyboard_hover: bool,
    pub(super) suppress_preview_until_cursor_leaves_preview: bool,
    pub(super) stationary_search_miss_started_at: Option<Instant>,
    pub(super) video_hover_guard_until: Option<Instant>,

    pub(super) last_cursor_location: Option<HoverLocation>,
    pub(super) hover_resolver_hints: HoverResolverHints,
    pub(super) suspend_preview_until_user_input: bool,
    pub(super) allow_keyboard_preview_on_first_observation: bool,
    pub(super) folder_change_time: Option<Instant>,
    pub(super) suspended_initial_focus: Option<FocusedItemKey>,
    pub(super) folder_change_user_initiated: bool,
    pub(super) keyboard_navigation_press_seq: u64,
    pub(super) keyboard_press_seq_at_suspend: u64,
    pub(super) keyboard_screen_owner: bool,

    pub(super) last_folder_probe: Instant,
    pub(super) last_hover_probe: Instant,
    pub(super) last_keyboard_focus_probe: Instant,
    pub(super) last_user_input_at: Option<Instant>,
    pub(super) last_keyboard_navigation_input_at: Option<Instant>,
    pub(super) last_navigation_trigger_at: Option<Instant>,
    pub(super) stationary_hover_probe_done: bool,
    pub(super) engine_page_drag: bool,

    pub(super) consumed_wheel_ticks: u64,
    pub(super) scroll_since_move: bool,

    pub(super) last_state_check: Instant,
    pub(super) current_state: ExplorerState,
    pub(super) explorer_probe_backoff_until: Option<Instant>,
}

impl HookSession {
    pub(super) fn new(now: Instant, wheel_ticks: u64, initial_state: ExplorerState) -> Self {
        Self {
            last_file: None,
            hover_item_box: None,
            suppressed: SuppressedHover::default(),
            pointer_pause: KeyboardPointerPause::default(),
            hover_start: None,
            last_cursor_pos: POINT::default(),
            keyboard_file: None,
            last_focused_key: None,
            last_focused_frame: None,
            is_keyboard_hover: false,
            suppress_preview_until_cursor_leaves_preview: false,
            stationary_search_miss_started_at: None,
            video_hover_guard_until: None,
            last_cursor_location: None,
            hover_resolver_hints: HoverResolverHints::default(),
            suspend_preview_until_user_input: false,
            allow_keyboard_preview_on_first_observation: false,
            folder_change_time: None,
            suspended_initial_focus: None,
            folder_change_user_initiated: false,
            keyboard_navigation_press_seq: 0,
            keyboard_press_seq_at_suspend: 0,
            keyboard_screen_owner: false,
            last_folder_probe: now,
            last_hover_probe: now,
            last_keyboard_focus_probe: now,
            last_user_input_at: None,
            last_keyboard_navigation_input_at: None,
            last_navigation_trigger_at: None,
            stationary_hover_probe_done: false,
            engine_page_drag: false,
            consumed_wheel_ticks: wheel_ticks,
            scroll_since_move: false,
            last_state_check: now,
            current_state: initial_state,
            explorer_probe_backoff_until: None,
        }
    }

    pub(super) fn has_active_preview(&self) -> bool {
        self.last_file.is_some() || self.keyboard_file.is_some() || self.is_keyboard_hover
    }

    pub(super) fn clear_engine_page_drag(&mut self) {
        self.engine_page_drag = false;
    }

    /// Clear active mouse-hover fields together (always clears `hover_item_box`).
    pub(super) fn clear_mouse_hover(&mut self) {
        self.last_file = None;
        self.hover_item_box = None;
        self.video_hover_guard_until = None;
    }

    /// Clear active keyboard-hover fields together.
    pub(super) fn clear_keyboard_hover(&mut self) {
        self.keyboard_file = None;
        self.is_keyboard_hover = false;
        self.video_hover_guard_until = None;
    }

    /// Clear active hover state (mouse + keyboard) without touching folder-gate state.
    pub(super) fn reset_hover(&mut self) {
        self.clear_mouse_hover();
        self.clear_keyboard_hover();
        self.stationary_search_miss_started_at = None;
        self.hover_start = None;
        self.pointer_pause.clear();
    }

    /// Atomic reset on Explorer restart (`clear_focus_baseline = true`) or display change.
    pub(super) fn reset_for_env_change(
        &mut self,
        now: Instant,
        backoff_ms: u64,
        clear_focus_baseline: bool,
    ) {
        self.reset_hover();
        if clear_focus_baseline {
            self.last_focused_key = None;
            self.last_focused_frame = None;
        }
        self.suppressed.clear();
        self.stationary_hover_probe_done = false;
        self.suspend_preview_until_user_input = true;
        self.allow_keyboard_preview_on_first_observation = false;
        self.folder_change_user_initiated = false;
        self.folder_change_time = Some(now);
        self.suspended_initial_focus = None;
        self.keyboard_press_seq_at_suspend = self.keyboard_navigation_press_seq;
        self.keyboard_screen_owner = false;
        self.hover_resolver_hints = HoverResolverHints::default();
        self.last_cursor_location = None;
        self.scroll_since_move = false;
        self.engine_page_drag = false;
        self.explorer_probe_backoff_until = Some(now + Duration::from_millis(backoff_ms));
    }

    /// Atomic reset when previews are held back (disabled or trigger-key blocked).
    /// Returns true if a preview was showing and must be hidden.
    pub(super) fn reset_for_hold_back(&mut self) -> bool {
        let should_hide = self.last_file.is_some() || self.keyboard_file.is_some();
        if should_hide {
            self.suppressed.clear();
            self.pointer_pause.clear();
            self.stationary_search_miss_started_at = None;
            self.hover_start = None;
        }
        self.clear_mouse_hover();
        self.clear_keyboard_hover();
        self.last_focused_key = None;
        self.last_focused_frame = None;
        self.keyboard_screen_owner = false;
        self.suspend_preview_until_user_input = false;
        self.allow_keyboard_preview_on_first_observation = false;
        self.folder_change_user_initiated = false;
        self.last_cursor_location = None;
        self.hover_resolver_hints = HoverResolverHints::default();
        self.folder_change_time = None;
        self.suspended_initial_focus = None;
        self.scroll_since_move = false;
        self.engine_page_drag = false;
        should_hide
    }

    /// Atomic reset when a pin closes and hovers resume.
    pub(super) fn reset_for_pin_resume(&mut self, handed_off: bool, cursor_pos: Option<POINT>) {
        self.reset_hover();
        self.last_focused_key = None;
        self.last_focused_frame = None;
        self.suppressed.clear();
        self.suspend_preview_until_user_input = false;
        self.keyboard_screen_owner = false;
        self.scroll_since_move = false;
        self.engine_page_drag = false;

        if handed_off {
            if let Some(pos) = cursor_pos {
                self.last_cursor_pos = pos;
            }
            self.suspend_preview_until_user_input = true;
            self.folder_change_time = None;
        }
    }

    /// Atomic reset when Explorer is unavailable or unfocused (no hover reachable).
    /// Returns true if a preview was showing and must be hidden.
    pub(super) fn reset_for_focus_loss(&mut self) -> bool {
        self.scroll_since_move = false;
        self.engine_page_drag = false;
        self.hover_item_box = None;

        let should_hide = self.last_file.is_some() || self.keyboard_file.is_some();
        if should_hide {
            self.reset_hover();
            self.last_focused_key = None;
            self.last_focused_frame = None;
            self.keyboard_screen_owner = false;
        }
        should_hide
    }

    /// Atomic reset when the folder probe detects a new location and the hovered file is gone.
    /// Returns true if a preview was showing and must be hidden.
    pub(super) fn reset_for_folder_change(&mut self, now: Instant, user_navigation: bool) -> bool {
        let should_hide = self.has_active_preview();
        self.suspend_preview_until_user_input = true;
        self.allow_keyboard_preview_on_first_observation = false;
        self.folder_change_user_initiated = user_navigation;
        self.folder_change_time = Some(now);
        self.suspended_initial_focus = None;
        self.keyboard_press_seq_at_suspend = self.keyboard_navigation_press_seq;

        self.clear_mouse_hover();
        self.clear_keyboard_hover();
        self.suppressed.clear();
        self.pointer_pause.clear();
        self.stationary_search_miss_started_at = None;
        self.scroll_since_move = false;
        should_hide
    }

    /// Atomic reset for history navigation (Backspace, Alt+arrows, mouse back/forward).
    /// Returns true if a preview was showing and must be hidden.
    pub(super) fn reset_for_history_nav(&mut self, now: Instant, cursor_pos: POINT) -> bool {
        let should_hide = self.has_active_preview();
        self.reset_hover();
        self.suppressed.clear();
        self.last_focused_key = None;
        self.last_focused_frame = None;
        self.suspend_preview_until_user_input = true;
        self.allow_keyboard_preview_on_first_observation = false;
        self.folder_change_user_initiated = false;
        self.last_navigation_trigger_at = Some(now);
        self.folder_change_time = Some(now);
        self.suspended_initial_focus = None;
        self.keyboard_press_seq_at_suspend = self.keyboard_navigation_press_seq;
        self.keyboard_screen_owner = false;
        self.last_cursor_pos = cursor_pos;
        self.stationary_hover_probe_done = false;
        self.scroll_since_move = false;
        self.engine_page_drag = false;
        should_hide
    }

    /// Atomic reset for Enter / activation key press.
    /// Returns true if a preview was showing and must be hidden.
    pub(super) fn reset_for_activation_key(&mut self) -> bool {
        let should_hide = self.has_active_preview();
        match self
            .last_file
            .clone()
            .or_else(|| self.keyboard_file.clone())
        {
            Some(file) => self.suppressed.suppress(file),
            None => self.suppressed.clear(),
        }
        self.reset_hover();
        self.last_focused_key = None;
        self.last_focused_frame = None;
        should_hide
    }

    /// Atomic reset when the cursor touches a non-holding preview window.
    pub(super) fn reset_for_preview_hover_dismiss(&mut self, now: Instant) {
        self.suppress_preview_until_cursor_leaves_preview = true;
        if let Some(file) = self.last_file.clone() {
            self.suppressed.suppress(file);
        }
        self.clear_mouse_hover();
        self.clear_keyboard_hover();
        self.stationary_search_miss_started_at = None;
        self.stationary_hover_probe_done = false;
        self.hover_start = Some(now);
    }

    /// Dismiss mouse hover and suppress its file (scroll-driven miss or search-view miss).
    pub(super) fn dismiss_mouse_hover_with_suppress(&mut self, restart_hover_at: Option<Instant>) {
        match self.last_file.clone() {
            Some(file) => self.suppressed.suppress(file),
            None => self.suppressed.clear(),
        }
        self.clear_mouse_hover();
        self.stationary_search_miss_started_at = None;
        if let Some(at) = restart_hover_at {
            self.hover_start = Some(at);
        }
    }

    /// Dismiss active mouse hover when keyboard navigation takes over.
    pub(super) fn dismiss_mouse_hover_for_keyboard(&mut self) {
        self.clear_mouse_hover();
        self.suppressed.clear();
        self.hover_start = None;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dirty_session(now: Instant) -> HookSession {
        let mut s = HookSession::new(now, 10, ExplorerState::ActiveFocus);
        s.last_file = Some(PathBuf::from("a.png"));
        s.hover_item_box = Some((10, 20, 100, 40));
        s.suppressed.suppress(PathBuf::from("b.png"));
        s.pointer_pause.watch_for_box();
        s.hover_start = Some(now);
        s.keyboard_file = Some(PathBuf::from("c.png"));
        s.last_focused_key = Some(FocusedItemKey::new(
            "c.png".to_string(),
            &RECT {
                left: 1,
                top: 2,
                right: 3,
                bottom: 4,
            },
        ));
        s.last_focused_frame = Some(42);
        s.is_keyboard_hover = true;
        s.stationary_search_miss_started_at = Some(now);
        s.video_hover_guard_until = Some(now + Duration::from_millis(350));
        s.last_cursor_location = Some(HoverLocation {
            folder: Some("folder".to_string()),
            search_root: None,
            location_url: None,
            view_hwnd: Some(1),
        });
        s.keyboard_screen_owner = true;
        s.stationary_hover_probe_done = true;
        s.scroll_since_move = true;
        s.engine_page_drag = true;
        s
    }

    #[test]
    fn com_guard_uninitializes_only_on_success_hresults() {
        assert!(ComGuard::should_uninitialize(0), "S_OK must uninitialize");
        assert!(
            ComGuard::should_uninitialize(1),
            "S_FALSE must uninitialize"
        );
        assert!(
            !ComGuard::should_uninitialize(0x8001_0106u32 as i32),
            "RPC_E_CHANGED_MODE must NOT uninitialize"
        );
        assert!(
            !ComGuard::should_uninitialize(0x8007_000Eu32 as i32),
            "E_OUTOFMEMORY must NOT uninitialize"
        );
    }

    #[test]
    fn reset_for_env_change_clears_stale_hover_box_scroll_and_drag() {
        let now = Instant::now();
        let mut s = dirty_session(now);
        s.reset_for_env_change(now, 500, true);

        assert!(s.last_file.is_none());
        assert!(s.hover_item_box.is_none());
        assert!(s.keyboard_file.is_none());
        assert!(!s.is_keyboard_hover);
        assert!(s.last_focused_key.is_none());
        assert!(s.last_focused_frame.is_none());
        assert!(!s.scroll_since_move);
        assert!(!s.engine_page_drag);
        assert!(s.suspend_preview_until_user_input);
        assert!(s.explorer_probe_backoff_until.is_some());
        assert!(s.last_cursor_location.is_none());

        // Display change preserves focus baseline when requested, still clears all three stale leaks.
        let mut d = dirty_session(now);
        d.reset_for_env_change(now, 500, false);
        assert!(d.hover_item_box.is_none());
        assert!(!d.scroll_since_move);
        assert!(!d.engine_page_drag);
        assert!(d.last_focused_key.is_some());
    }

    #[test]
    fn reset_for_hold_back_clears_hover_box_scroll_and_drag() {
        let now = Instant::now();
        let mut s = dirty_session(now);
        assert!(s.reset_for_hold_back());

        assert!(s.last_file.is_none());
        assert!(s.hover_item_box.is_none());
        assert!(s.keyboard_file.is_none());
        assert!(!s.scroll_since_move);
        assert!(!s.engine_page_drag);
        assert!(!s.suspend_preview_until_user_input);
    }

    #[test]
    fn reset_for_pin_resume_clears_hover_box_scroll_and_drag() {
        let now = Instant::now();
        let mut s = dirty_session(now);
        s.reset_for_pin_resume(true, Some(POINT { x: 77, y: 88 }));

        assert!(s.last_file.is_none());
        assert!(s.hover_item_box.is_none());
        assert!(s.keyboard_file.is_none());
        assert!(!s.scroll_since_move);
        assert!(!s.engine_page_drag);
        assert!(s.suspend_preview_until_user_input);
        assert_eq!(s.last_cursor_pos.x, 77);
        assert_eq!(s.last_cursor_pos.y, 88);
    }

    #[test]
    fn reset_for_focus_loss_always_clears_hover_box_scroll_and_drag() {
        let now = Instant::now();
        // With active preview
        let mut s = dirty_session(now);
        assert!(s.reset_for_focus_loss());
        assert!(s.last_file.is_none());
        assert!(s.hover_item_box.is_none());
        assert!(!s.scroll_since_move);
        assert!(!s.engine_page_drag);

        // Even when no preview is active, stale scroll/drag/box must still be cleared.
        let mut idle = HookSession::new(now, 0, ExplorerState::HiddenByForeground);
        idle.hover_item_box = Some((1, 2, 3, 4));
        idle.scroll_since_move = true;
        idle.engine_page_drag = true;
        assert!(!idle.reset_for_focus_loss());
        assert!(idle.hover_item_box.is_none());
        assert!(!idle.scroll_since_move);
        assert!(!idle.engine_page_drag);
    }

    #[test]
    fn reset_for_folder_change_and_history_nav_clear_hover_box_and_scroll() {
        let now = Instant::now();
        let mut s = dirty_session(now);
        assert!(s.reset_for_folder_change(now, true));
        assert!(s.hover_item_box.is_none());
        assert!(!s.scroll_since_move);
        assert!(s.folder_change_user_initiated);
        assert!(s.suspend_preview_until_user_input);

        let mut h = dirty_session(now);
        assert!(h.reset_for_history_nav(now, POINT { x: 5, y: 6 }));
        assert!(h.hover_item_box.is_none());
        assert!(!h.scroll_since_move);
        assert!(!h.engine_page_drag);
        assert_eq!(h.last_cursor_pos.x, 5);
    }

    #[test]
    fn clear_mouse_hover_and_dismiss_variants_always_clear_hover_item_box() {
        let now = Instant::now();
        let mut s = dirty_session(now);
        s.clear_mouse_hover();
        assert!(s.last_file.is_none());
        assert!(s.hover_item_box.is_none());

        let mut s2 = dirty_session(now);
        assert!(s2.reset_for_activation_key());
        assert!(s2.hover_item_box.is_none());

        let mut s3 = dirty_session(now);
        s3.reset_for_preview_hover_dismiss(now);
        assert!(s3.hover_item_box.is_none());
        assert!(s3.suppress_preview_until_cursor_leaves_preview);

        let mut s4 = dirty_session(now);
        s4.dismiss_mouse_hover_with_suppress(Some(now));
        assert!(s4.hover_item_box.is_none());
    }

    #[test]
    fn cached_hook_config_refreshes_when_generation_advances() {
        let _folder = crate::config::config::ConfigFolderOfItsOwn::take();
        let mut cached = CachedHookConfig::capture_initial();
        let initial_delay = cached.current().hover_delay_ms;

        let new_delay = if initial_delay == 123 { 456 } else { 123 };
        crate::config::commit_config(|cfg| {
            cfg.hover_delay_ms = new_delay;
            cfg.pin_update_enabled = false;
            cfg.pin_update_on_hover = true;
        });

        let updated = cached.current();
        assert_eq!(updated.hover_delay_ms, new_delay);
        assert!(!updated.pin_update_enabled);
        assert!(updated.pin_update_on_hover);
    }
}

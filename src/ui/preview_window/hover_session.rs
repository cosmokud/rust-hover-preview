//! Hover and preview session state machine.
//!
//! Owns the active hover target, its resolved [`HoverFacts`] / [`HoverScales`] snapshot, and the
//! replay/wait/failure flags that span ticks so repeated `Show`, `Hide`, file switches,
//! `ConfigChanged`, and pin handoffs obey one set of invariants without stale-flag drift.

use super::*;

/// Result of transitioning [`HoverSession`] on a `Show` / `ShowKeyboard` message.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ShowOutcome {
    /// Already showing or replaying this file: keep window/frame/spinner in place and upgrade.
    KeepAndUpgrade,
    /// New file or fresh preview: replace whatever was previously on screen.
    Replace,
}

impl ShowOutcome {
    pub(super) fn is_upgrade(self) -> bool {
        self == Self::KeepAndUpgrade
    }
}

/// Owned per-loop hover and pin-wait session state.
pub(super) struct HoverSession {
    current_show: Option<PreviewMessage>,
    facts: Option<Arc<HoverFacts>>,
    scales: HoverScales,
    page_upgrade: Option<PathBuf>,
    video_replay: Option<PathBuf>,
    measure_replay: Option<PathBuf>,
    engine_failed: Option<PathBuf>,
    pin_awaiting_box: Option<PathBuf>,
    pin_cold_wait: Option<PathBuf>,
    pin_page_shown: Option<PathBuf>,
}

impl HoverSession {
    pub(super) fn new(scales: HoverScales) -> Self {
        Self {
            current_show: None,
            facts: None,
            scales,
            page_upgrade: None,
            video_replay: None,
            measure_replay: None,
            engine_failed: None,
            pin_awaiting_box: None,
            pin_cold_wait: None,
            pin_page_shown: None,
        }
    }

    pub(super) fn current_show(&self) -> &Option<PreviewMessage> {
        &self.current_show
    }

    pub(super) fn current_show_mut(&mut self) -> &mut Option<PreviewMessage> {
        &mut self.current_show
    }

    pub(super) fn shown_path(&self) -> Option<&Path> {
        self.current_show
            .as_ref()
            .and_then(show_path)
            .map(PathBuf::as_path)
    }

    pub(super) fn is_showing(&self) -> bool {
        self.current_show.is_some()
    }

    #[cfg(test)]
    pub(super) fn scales(&self) -> HoverScales {
        self.scales
    }

    /// Inherit the active hover's resolved `HoverFacts` if it matches `path`.
    pub(super) fn facts_for(&self, path: &Path) -> Option<Arc<HoverFacts>> {
        self.facts
            .as_ref()
            .filter(|facts| facts.probe.path() == path)
            .cloned()
    }

    /// Return the active hover's `HoverFacts` when `path` matches, or read from disk as fallback.
    #[cfg(test)]
    pub(super) fn facts_or_read(&self, path: &Path) -> Arc<HoverFacts> {
        self.facts_for(path)
            .unwrap_or_else(|| Arc::new(HoverFacts::read(path)))
    }

    /// Transition on `Show` / `ShowKeyboard` after layout succeeds.
    ///
    /// Repeated `Show` for the file already on screen (or matching a pending replay/upgrade flag)
    /// returns [`ShowOutcome::KeepAndUpgrade`]; switching to another file clears stale failure
    /// state and returns [`ShowOutcome::Replace`]. Replay flags are always consumed atomically.
    pub(super) fn on_show(
        &mut self,
        snapshot: Option<PreviewMessage>,
        facts: Option<Arc<HoverFacts>>,
        path: &Path,
        pending_load_path: Option<&Path>,
        media_is_ready: bool,
    ) -> ShowOutcome {
        let same_file = self.shown_path() == Some(path);
        let showing_this_file =
            same_file && pending_load_path.map_or(media_is_ready, |pending| pending == path);
        let upgrading = showing_this_file
            || self.page_upgrade.as_deref() == Some(path)
            || self.video_replay.as_deref() == Some(path)
            || self.measure_replay.as_deref() == Some(path);

        self.page_upgrade = None;
        self.video_replay = None;
        self.measure_replay = None;
        if !same_file {
            self.engine_failed = None;
        }

        if snapshot.is_some() {
            self.current_show = snapshot;
        }
        if let Some(facts) = facts {
            self.scales = facts.scales;
            self.facts = Some(facts);
        } else if !same_file {
            self.facts = None;
        }

        if upgrading {
            ShowOutcome::KeepAndUpgrade
        } else {
            ShowOutcome::Replace
        }
    }

    /// Transition on `Hide`, abandoned load, or unpreviewable `Show`.
    ///
    /// Atomically clears hover target, cached `HoverFacts`, and all hover replay/failure flags,
    /// returning the departed path (if any) so callers can notify engine render caches.
    pub(super) fn on_hide(&mut self) -> Option<PathBuf> {
        let departed = self.shown_path().map(Path::to_path_buf);
        self.current_show = None;
        self.facts = None;
        self.page_upgrade = None;
        self.video_replay = None;
        self.measure_replay = None;
        self.engine_failed = None;
        departed
    }

    /// Transition on live config reload (`ConfigChanged` or generation bump).
    ///
    /// Invalidates cached `HoverFacts` when scales changed so pin take-ups never inherit stale
    /// scales, and returns the active `Show` message to replay if a non-pinned hover replay is in
    /// flight.
    pub(super) fn on_config_changed(
        &mut self,
        scales_changed: bool,
        new_scales: HoverScales,
        is_pinned: bool,
    ) -> Option<PreviewMessage> {
        self.scales = new_scales;
        if !scales_changed {
            return None;
        }
        self.facts = None;
        if !is_pinned && self.has_pending_replay() {
            self.current_show.clone()
        } else {
            None
        }
    }

    pub(super) fn has_pending_replay(&self) -> bool {
        self.video_replay.is_some() || self.measure_replay.is_some() || self.page_upgrade.is_some()
    }

    pub(super) fn mark_page_upgrade(&mut self, path: PathBuf) {
        self.page_upgrade = Some(path);
    }

    pub(super) fn mark_video_replay(&mut self, path: PathBuf) {
        self.video_replay = Some(path);
    }

    pub(super) fn mark_measure_replay(&mut self, path: PathBuf) {
        self.measure_replay = Some(path);
    }

    pub(super) fn on_engine_failed(&mut self, path: PathBuf) {
        self.engine_failed = Some(path);
    }

    pub(super) fn clear_engine_failed(&mut self) {
        self.engine_failed = None;
    }

    /// Consume `engine_failed` when it matches the active non-pinned hover, returning its `Show`
    /// message for replay.
    pub(super) fn take_engine_failure_replay(&mut self, is_pinned: bool) -> Option<PreviewMessage> {
        let matches_shown = !is_pinned
            && self
                .engine_failed
                .as_deref()
                .is_some_and(|failed| self.shown_path() == Some(failed));
        if matches_shown {
            self.engine_failed = None;
            self.current_show.clone()
        } else {
            None
        }
    }

    pub(super) fn pin_awaiting_box(&self) -> Option<&Path> {
        self.pin_awaiting_box.as_deref()
    }

    pub(super) fn has_pin_awaiting_box(&self) -> bool {
        self.pin_awaiting_box.is_some()
    }

    pub(super) fn mark_pin_awaiting_box(&mut self, path: PathBuf) {
        self.pin_awaiting_box = Some(path);
    }

    pub(super) fn clear_pin_awaiting_box(&mut self) {
        self.pin_awaiting_box = None;
    }

    pub(super) fn pin_awaiting_box_mut(&mut self) -> &mut Option<PathBuf> {
        &mut self.pin_awaiting_box
    }

    pub(super) fn take_pin_awaiting_box_if(&mut self, path: &Path) -> bool {
        if self.pin_awaiting_box.as_deref() == Some(path) {
            self.pin_awaiting_box = None;
            true
        } else {
            false
        }
    }

    pub(super) fn pin_cold_wait(&self) -> Option<&Path> {
        self.pin_cold_wait.as_deref()
    }

    pub(super) fn pin_cold_wait_mut(&mut self) -> &mut Option<PathBuf> {
        &mut self.pin_cold_wait
    }

    /// Record that `path` was raised on the pin window; returns `true` only on transition.
    pub(super) fn note_pin_page_shown(&mut self, path: PathBuf) -> bool {
        if self.pin_page_shown.as_deref() != Some(path.as_path()) {
            self.pin_page_shown = Some(path);
            true
        } else {
            false
        }
    }

    pub(super) fn clear_pin_page_shown(&mut self) {
        self.pin_page_shown = None;
    }

    /// Clear all pin wait/page flags when no pin is standing.
    pub(super) fn clear_pin_waits(&mut self) {
        self.pin_awaiting_box = None;
        self.pin_cold_wait = None;
        self.pin_page_shown = None;
    }

    /// Clear outstanding box/cold waits when a new pin pick supersedes them.
    pub(super) fn clear_pin_waits_for_pick(&mut self) {
        self.pin_awaiting_box = None;
        self.pin_cold_wait = None;
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::config::{AppConfig, PreviewScale};

    fn show_msg(path: &Path) -> PreviewMessage {
        PreviewMessage::Show(path.to_path_buf(), 100, 200, None)
    }

    fn default_scales() -> HoverScales {
        HoverScales::of(&AppConfig::default())
    }

    #[test]
    fn repeated_show_for_already_shown_file_preserves_state_without_teardown() {
        let mut session = HoverSession::new(default_scales());
        let file = PathBuf::from("clips/demo.mp4");
        let facts = Arc::new(HoverFacts::read(&file));

        // First show replaces empty screen.
        let first = session.on_show(
            Some(show_msg(&file)),
            Some(Arc::clone(&facts)),
            &file,
            None,
            false,
        );
        assert_eq!(first, ShowOutcome::Replace);
        assert_eq!(session.shown_path(), Some(file.as_path()));
        assert!(session.facts_for(&file).is_some());

        // Repeated Show while media is ready keeps state (no teardown).
        let repeated_ready = session.on_show(
            Some(show_msg(&file)),
            Some(Arc::clone(&facts)),
            &file,
            None,
            true,
        );
        assert_eq!(repeated_ready, ShowOutcome::KeepAndUpgrade);
        assert_eq!(session.shown_path(), Some(file.as_path()));

        // Repeated Show while pending load is for the same file also keeps state.
        let repeated_pending = session.on_show(
            Some(show_msg(&file)),
            Some(Arc::clone(&facts)),
            &file,
            Some(&file),
            false,
        );
        assert_eq!(repeated_pending, ShowOutcome::KeepAndUpgrade);
    }

    #[test]
    fn switching_files_and_hide_atomically_clear_all_replay_and_failure_flags() {
        let mut session = HoverSession::new(default_scales());
        let file_a = PathBuf::from("docs/a.pdf");
        let file_b = PathBuf::from("docs/b.pdf");
        let facts_a = Arc::new(HoverFacts::read(&file_a));
        let facts_b = Arc::new(HoverFacts::read(&file_b));

        session.on_show(
            Some(show_msg(&file_a)),
            Some(Arc::clone(&facts_a)),
            &file_a,
            None,
            false,
        );
        session.mark_video_replay(file_a.clone());
        session.mark_measure_replay(file_a.clone());
        session.mark_page_upgrade(file_a.clone());
        session.on_engine_failed(file_a.clone());

        // Switching to file B replaces and clears all stale flags from A.
        let outcome = session.on_show(
            Some(show_msg(&file_b)),
            Some(Arc::clone(&facts_b)),
            &file_b,
            Some(&file_a),
            true,
        );
        assert_eq!(outcome, ShowOutcome::Replace);
        assert!(!session.has_pending_replay());
        assert!(session.take_engine_failure_replay(false).is_none());
        assert!(session.facts_for(&file_a).is_none());
        assert!(session.facts_for(&file_b).is_some());

        // Arm flags on B, then Hide: every hover flag + facts must clear atomically.
        session.mark_video_replay(file_b.clone());
        session.mark_measure_replay(file_b.clone());
        session.mark_page_upgrade(file_b.clone());
        session.on_engine_failed(file_b.clone());

        let departed = session.on_hide();
        assert_eq!(departed.as_deref(), Some(file_b.as_path()));
        assert!(!session.is_showing());
        assert!(session.shown_path().is_none());
        assert!(!session.has_pending_replay());
        assert!(session.facts_for(&file_b).is_none());
        assert!(session.take_engine_failure_replay(false).is_none());
    }

    #[test]
    fn config_changed_scales_invalidates_facts_and_replays_active_hover_replay() {
        let initial = default_scales();
        let mut updated = initial;
        updated.video = PreviewScale::Percent(50);

        let mut session = HoverSession::new(initial);
        let file = PathBuf::from("clips/movie.mp4");
        let facts = Arc::new(HoverFacts::read(&file));

        session.on_show(
            Some(show_msg(&file)),
            Some(Arc::clone(&facts)),
            &file,
            None,
            false,
        );

        // No scale change: facts stay valid, no replay triggered.
        assert!(session.on_config_changed(false, initial, false).is_none());
        assert!(session.facts_for(&file).is_some());

        // Scale change without pending replay: invalidates stale facts, does not rebuild static hover.
        assert!(session.on_config_changed(true, updated, false).is_none());
        assert!(session.facts_for(&file).is_none());
        assert_eq!(session.scales(), updated);

        // Scale change with pending replay while unpinned: triggers replay of current_show.
        session.mark_video_replay(file.clone());
        let replay = session.on_config_changed(true, updated, false);
        assert!(replay.is_some());
        assert_eq!(replay.as_ref().and_then(show_path), Some(&file));

        // While pinned: scale change does not trigger hover replay.
        assert!(session.on_config_changed(true, updated, true).is_none());
    }

    #[test]
    fn pin_takeup_inherits_matching_hover_facts_and_falls_back_on_different_path() {
        let mut session = HoverSession::new(default_scales());
        let hovered = PathBuf::from("pics/photo.png");
        let swapped = PathBuf::from("pics/other.png");
        let hovered_facts = Arc::new(HoverFacts::read(&hovered));

        session.on_show(
            Some(show_msg(&hovered)),
            Some(Arc::clone(&hovered_facts)),
            &hovered,
            None,
            false,
        );

        // Matching path inherits the exact Arc<HoverFacts> without re-reading disk.
        let inherited = session.facts_or_read(&hovered);
        assert!(Arc::ptr_eq(&inherited, &hovered_facts));

        // Different path (pin swap / cold pin) misses snapshot and falls back to fresh read.
        assert!(session.facts_for(&swapped).is_none());
        let fallback = session.facts_or_read(&swapped);
        assert!(!Arc::ptr_eq(&fallback, &hovered_facts));
        assert_eq!(fallback.probe.path(), swapped.as_path());
    }

    #[test]
    fn pin_wait_flags_clear_atomically_on_unpin_and_new_pick() {
        let mut session = HoverSession::new(default_scales());
        let a = PathBuf::from("docs/a.docx");
        let b = PathBuf::from("docs/b.docx");

        session.mark_pin_awaiting_box(a.clone());
        *session.pin_cold_wait_mut() = Some(a.clone());
        assert!(session.note_pin_page_shown(a.clone()));
        assert!(!session.note_pin_page_shown(a.clone()));

        // New pick clears box & cold waits while leaving page_shown intact until page changes.
        session.clear_pin_waits_for_pick();
        assert!(!session.has_pin_awaiting_box());
        assert!(session.pin_cold_wait().is_none());

        // Unpin clears all three pin flags together.
        session.mark_pin_awaiting_box(b.clone());
        *session.pin_cold_wait_mut() = Some(b.clone());
        session.clear_pin_waits();
        assert!(!session.has_pin_awaiting_box());
        assert!(session.pin_cold_wait().is_none());
        assert!(session.note_pin_page_shown(a));
    }
}

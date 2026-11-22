// The settings module is reached as `config::config` from every reader and engine,
// which is the path this app has always named it by: renaming it would touch a
// hundred import sites to answer a lint about the name itself.
#[allow(clippy::module_inception)]
pub(crate) mod config;
pub(crate) mod theme_files;

use crate::app::seam::{self, PreviewMessage};

/// Consistent point-in-time copy of `AppConfig` under one lock acquisition (recovers if poisoned).
pub(crate) fn snapshot() -> config::AppConfig {
    crate::CONFIG
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .clone()
}

/// Mutate `CONFIG`, persist `config.ini`, bump cross-process config generation, and dispatch
/// `PreviewMessage::ConfigChanged` (waking the viewer loop). Lock drops before seam notifications.
pub(crate) fn commit_config<R>(mutate: impl FnOnce(&mut config::AppConfig) -> R) -> R {
    let result = {
        let mut cfg = crate::CONFIG.lock().unwrap_or_else(|e| e.into_inner());
        let out = mutate(&mut cfg);
        cfg.save();
        out
    };
    seam::bump_config_generation();
    seam::send_preview(PreviewMessage::ConfigChanged);
    result
}

/// Commit a config mutation and dispatch follow-up preview messages strictly AFTER `ConfigChanged`.
pub(crate) fn publish_config_change<R>(
    mutate: impl FnOnce(&mut config::AppConfig) -> R,
    follow_ups: &[PreviewMessage],
) -> R {
    let result = commit_config(mutate);
    for msg in follow_ups {
        seam::send_preview(msg.clone());
    }
    result
}

#[cfg(test)]
mod publication_tests {
    use super::*;
    use std::sync::mpsc::channel;

    #[test]
    fn publish_config_change_sends_config_changed_before_follow_ups_and_bumps_generation() {
        let _folder = config::ConfigFolderOfItsOwn::take();
        let (tx, rx) = channel();
        let prev_sender = seam::PREVIEW_SENDER
            .lock()
            .unwrap_or_else(|e| e.into_inner())
            .replace(tx);

        let gen_before = seam::config_generation();
        let follow_ups = [
            PreviewMessage::TrimCaches,
            PreviewMessage::TrayVolume { audio: false },
            PreviewMessage::PinChanged,
            PreviewMessage::RefreshTypes,
            PreviewMessage::ClosePdfPins,
        ];
        publish_config_change(
            |cfg| {
                cfg.image_cache_mb = 256;
            },
            &follow_ups,
        );
        let gen_after = seam::config_generation();

        *seam::PREVIEW_SENDER
            .lock()
            .unwrap_or_else(|e| e.into_inner()) = prev_sender;

        assert!(
            gen_after > gen_before,
            "generation must advance on config publication"
        );
        let drained: Vec<PreviewMessage> = rx.try_iter().collect();
        assert_eq!(drained.len(), 1 + follow_ups.len());
        assert!(
            matches!(drained[0], PreviewMessage::ConfigChanged),
            "ConfigChanged must precede all follow-up commands"
        );
        assert!(matches!(drained[1], PreviewMessage::TrimCaches));
        assert!(matches!(
            drained[2],
            PreviewMessage::TrayVolume { audio: false }
        ));
        assert!(matches!(drained[3], PreviewMessage::PinChanged));
        assert!(matches!(drained[4], PreviewMessage::RefreshTypes));
        assert!(matches!(drained[5], PreviewMessage::ClosePdfPins));
    }

    #[test]
    fn snapshot_returns_consistent_copy_after_commit() {
        let _folder = config::ConfigFolderOfItsOwn::take();
        commit_config(|cfg| {
            cfg.video_volume = 70;
            cfg.audio_volume = 40;
            cfg.pin_always_on_top = false;
        });
        let snap = snapshot();
        assert_eq!(snap.video_volume, 70);
        assert_eq!(snap.audio_volume, 40);
        assert!(!snap.pin_always_on_top);
    }
}

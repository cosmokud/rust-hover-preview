use directories::ProjectDirs;
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AppConfig {
    pub run_at_startup: bool,
    pub preview_size: u32,
    pub preview_offset: i32,
    pub hover_delay_ms: u64,
    pub preview_enabled: bool,
    pub follow_cursor: bool,
    #[serde(default)]
    pub video_volume: u32,
}

impl Default for AppConfig {
    fn default() -> Self {
        Self {
            run_at_startup: false,
            preview_size: 300,
            preview_offset: 20,
            hover_delay_ms: 0,
            preview_enabled: true,
            follow_cursor: false,
            video_volume: 0, // Mute by default
        }
    }
}

impl AppConfig {
    fn config_path() -> Option<PathBuf> {
        ProjectDirs::from("com", "RustHoverPreview", "RustHoverPreview")
            .map(|dirs| dirs.config_dir().join("config.json"))
    }

    pub fn load() -> Self {
        let config: Self = Self::config_path()
            .and_then(|path| fs::read_to_string(path).ok())
            .and_then(|content| serde_json::from_str(&content).ok())
            .unwrap_or_default();
        
        // Always save to ensure new fields are written to config file
        config.save();
        config
    }

    pub fn save(&self) {
        if let Some(path) = Self::config_path() {
            if let Some(parent) = path.parent() {
                let _ = fs::create_dir_all(parent);
            }
            if let Ok(content) = serde_json::to_string_pretty(self) {
                let _ = fs::write(path, content);
            }
        }
    }
}

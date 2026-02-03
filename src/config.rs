use configparser::ini::Ini;
use std::env;
use std::fs;
use std::path::PathBuf;

const CONFIG_SECTION: &str = "settings";

#[derive(Debug, Clone)]
pub struct AppConfig {
    pub run_at_startup: bool,
    pub hover_delay_ms: u64,
    pub preview_enabled: bool,
    pub follow_cursor: bool,
    pub video_volume: u32,
}

impl Default for AppConfig {
    fn default() -> Self {
        Self {
            run_at_startup: true,
            hover_delay_ms: 0,
            preview_enabled: true,
            follow_cursor: false,
            video_volume: 0, // Mute by default
        }
    }
}

impl AppConfig {
    pub fn config_path() -> Option<PathBuf> {
        env::var_os("APPDATA")
            .map(PathBuf::from)
            .map(|base| base.join("rust-hover-preview").join("config.ini"))
    }

    pub fn load() -> Self {
        let mut config = Self::default();

        if let Some(path) = Self::config_path() {
            let mut ini = Ini::new();
            if ini.load(path.to_string_lossy().as_ref()).is_ok() {
                config.apply_ini(&ini);
            }
        }
        
        // Always save to ensure new fields are written to config file
        config.save();
        config
    }

    pub fn save(&self) {
        if let Some(path) = Self::config_path() {
            if let Some(parent) = path.parent() {
                let _ = fs::create_dir_all(parent);
            }
            let mut ini = Ini::new();
            ini.set(CONFIG_SECTION, "run_at_startup", Some(self.run_at_startup.to_string()));
            ini.set(CONFIG_SECTION, "hover_delay_ms", Some(self.hover_delay_ms.to_string()));
            ini.set(CONFIG_SECTION, "preview_enabled", Some(self.preview_enabled.to_string()));
            ini.set(CONFIG_SECTION, "follow_cursor", Some(self.follow_cursor.to_string()));
            ini.set(CONFIG_SECTION, "video_volume", Some(self.video_volume.to_string()));
            let _ = ini.write(path.to_string_lossy().as_ref());
        }
    }

    fn apply_ini(&mut self, ini: &Ini) {
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "run_at_startup") {
            self.run_at_startup = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "hover_delay_ms") {
            self.hover_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "preview_enabled") {
            self.preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "follow_cursor") {
            self.follow_cursor = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "video_volume") {
            if let Ok(value) = u32::try_from(value) {
                self.video_volume = value;
            }
        }
    }
}

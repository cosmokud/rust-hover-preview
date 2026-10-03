use super::*;

/// A file as the app reads one: what it has wrong or missing is put right first, and then
/// what the configuration reads is what the file says.
fn read_file(ini: &mut Ini) -> AppConfig {
    lists::repair_older_lists(ini);

    let mut config = AppConfig::default();
    config.apply_ini(ini);
    config
}

/// A file holding everything this app writes, as `save` would write it: what a file on disk
/// is once it has been through the app, and what the tests below change one key of.
fn written_file() -> Ini {
    let mut ini = Ini::new();
    for (section, keys) in AppConfig::default().to_ini().get_map_ref() {
        for (key, value) in keys {
            ini.set(section, key, value.clone());
        }
    }

    ini
}

/// And the same file as a build before this one left it: everything this app writes, less the
/// sections that build did not have. It is how a file that is missing a whole list is written
/// down for the tests — a `[ffmpeg]` a file has never had is one the app adds as it writes,
/// and a section set to nothing is not a section that is not there.
fn written_file_before_this_build(new_sections: &[&str]) -> Ini {
    let mut ini = Ini::new();
    for (section, keys) in AppConfig::default().to_ini().get_map_ref() {
        if new_sections.contains(&section.as_str()) {
            continue;
        }

        for (key, value) in keys {
            ini.set(section, key, value.clone());
        }
    }

    ini
}

mod extension_lists;
mod persistence;
mod scales;
mod settings;

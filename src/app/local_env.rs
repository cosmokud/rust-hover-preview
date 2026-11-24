use std::path::PathBuf;

/// Path to `.env.local` beside the running executable (`<exe_dir>\.env.local`).
pub fn exe_local_env_path() -> Option<PathBuf> {
    std::env::current_exe()
        .ok()?
        .parent()
        .map(|dir| dir.join(".env.local"))
}

/// Path to `.env.local` inside local AppData or portable data directory.
pub fn local_env_path() -> Option<PathBuf> {
    if let Some(portable) = crate::paths::data_root() {
        return Some(portable.join(".env.local"));
    }
    directories::BaseDirs::new()
        .map(|dirs| dirs.data_local_dir().join("rust-hover-preview"))
        .or_else(|| {
            std::env::var_os("LOCALAPPDATA")
                .map(|dir| PathBuf::from(dir).join("rust-hover-preview"))
        })
        .map(|dir| dir.join(".env.local"))
}

fn same_env_path(a: &std::path::Path, b: &std::path::Path) -> bool {
    if cfg!(windows) {
        a.to_string_lossy()
            .eq_ignore_ascii_case(&b.to_string_lossy())
    } else {
        a == b
    }
}

fn candidate_env_paths(exe_env: Option<PathBuf>, data_env: Option<PathBuf>) -> Vec<PathBuf> {
    let mut paths = Vec::with_capacity(2);
    if let Some(path) = exe_env {
        paths.push(path);
    }
    if let Some(path) = data_env {
        if paths.iter().all(|existing| !same_env_path(existing, &path)) {
            paths.push(path);
        }
    }
    paths
}

/// Parses key-value pairs from lines in `.env` syntax.
///
/// Ignores comments (`#`), empty lines, and supports optional `export ` prefix.
/// Strips surrounding matching single or double quotes from values.
pub fn parse_env_lines(contents: &str) -> Vec<(String, String)> {
    let mut entries = Vec::new();

    for line in contents.lines() {
        let line = line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        let line = line.strip_prefix("export ").unwrap_or(line).trim();

        if let Some((key, val)) = line.split_once('=') {
            let key = key.trim();
            if key.is_empty() {
                continue;
            }

            let mut val = val.trim();
            if (val.starts_with('"') && val.ends_with('"') && val.len() >= 2)
                || (val.starts_with('\'') && val.ends_with('\'') && val.len() >= 2)
            {
                val = &val[1..val.len() - 1];
            }

            entries.push((key.to_string(), val.to_string()));
        }
    }

    entries
}

fn load_from_paths(paths: &[PathBuf]) {
    for path in paths {
        let Ok(content) = std::fs::read_to_string(path) else {
            continue;
        };
        for (key, val) in parse_env_lines(&content) {
            if std::env::var_os(&key).is_none() {
                std::env::set_var(&key, &val);
            }
        }
    }
}

/// Reads `<exe_dir>\.env.local` and `local_env_path()` (skipping duplicates) if present,
/// populating any environment variables not already set in the current process.
pub fn load() {
    let paths = candidate_env_paths(exe_local_env_path(), local_env_path());
    load_from_paths(&paths);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_env_lines_extracts_keys_and_values() {
        let sample = r#"
        # Comment line
        RHP_DISABLE_FFMPEG=1
        export OTHER_VAR="hello world"
        SINGLE_QUOTED='single'
        UNQUOTED=some_value
        EMPTY_VAL=
        "#;

        let parsed = parse_env_lines(sample);
        assert_eq!(
            parsed,
            vec![
                ("RHP_DISABLE_FFMPEG".to_string(), "1".to_string()),
                ("OTHER_VAR".to_string(), "hello world".to_string()),
                ("SINGLE_QUOTED".to_string(), "single".to_string()),
                ("UNQUOTED".to_string(), "some_value".to_string()),
                ("EMPTY_VAL".to_string(), "".to_string()),
            ]
        );
    }

    #[test]
    fn local_env_path_resolves() {
        let path = local_env_path();
        if let Some(p) = path {
            assert!(p.ends_with(std::path::Path::new("rust-hover-preview").join(".env.local")));
        }
        if let Some(exe_env) = exe_local_env_path() {
            assert_eq!(
                exe_env.file_name().and_then(|n| n.to_str()),
                Some(".env.local")
            );
        }
    }

    #[test]
    fn candidate_env_paths_dedupes_and_load_preserves_precedence() {
        let same_a = PathBuf::from("app").join(".env.local");
        let same_b = PathBuf::from("APP").join(".env.local");
        let other = PathBuf::from("data").join(".env.local");
        if cfg!(windows) {
            assert_eq!(
                candidate_env_paths(Some(same_a.clone()), Some(same_b)),
                vec![same_a.clone()]
            );
        }
        assert_eq!(
            candidate_env_paths(Some(same_a.clone()), Some(other.clone())),
            vec![same_a, other]
        );

        let dir = std::env::temp_dir().join(format!("rhp_local_env_test_{}", std::process::id()));
        let exe_dir = dir.join("exe");
        let data_dir = dir.join("data");
        let _ = std::fs::create_dir_all(&exe_dir);
        let _ = std::fs::create_dir_all(&data_dir);
        let exe_env = exe_dir.join(".env.local");
        let data_env = data_dir.join(".env.local");

        let k_exe = format!("RHP_TEST_EXE_{}", std::process::id());
        let k_data = format!("RHP_TEST_DATA_{}", std::process::id());
        let k_preset = format!("RHP_TEST_PRESET_{}", std::process::id());
        std::env::remove_var(&k_exe);
        std::env::remove_var(&k_data);
        std::env::set_var(&k_preset, "already_set");

        let _ = std::fs::write(&exe_env, format!("{k_exe}=from_exe\n{k_preset}=from_exe\n"));
        let _ = std::fs::write(
            &data_env,
            format!("{k_exe}=from_data\n{k_data}=from_data\n"),
        );

        let paths = candidate_env_paths(Some(exe_env), Some(data_env));
        load_from_paths(&paths);

        assert_eq!(std::env::var(&k_exe).as_deref(), Ok("from_exe"));
        assert_eq!(std::env::var(&k_data).as_deref(), Ok("from_data"));
        assert_eq!(std::env::var(&k_preset).as_deref(), Ok("already_set"));

        std::env::remove_var(&k_exe);
        std::env::remove_var(&k_data);
        std::env::remove_var(&k_preset);
        let _ = std::fs::remove_dir_all(&dir);
    }
}

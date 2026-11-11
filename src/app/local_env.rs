use std::path::PathBuf;

/// Path to `.env.local` inside local AppData directory (`%LOCALAPPDATA%\rust-hover-preview\.env.local`).
pub fn local_env_path() -> Option<PathBuf> {
    directories::BaseDirs::new()
        .map(|dirs| dirs.data_local_dir().join("rust-hover-preview"))
        .or_else(|| {
            std::env::var_os("LOCALAPPDATA")
                .map(|dir| PathBuf::from(dir).join("rust-hover-preview"))
        })
        .map(|dir| dir.join(".env.local"))
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

/// Reads `%LOCALAPPDATA%\rust-hover-preview\.env.local` if it exists and populates
/// any environment variables that are not already set in the current process.
pub fn load() {
    let Some(path) = local_env_path() else {
        return;
    };

    let Ok(content) = std::fs::read_to_string(&path) else {
        return;
    };

    for (key, val) in parse_env_lines(&content) {
        if std::env::var_os(&key).is_none() {
            std::env::set_var(&key, &val);
        }
    }
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
    }
}

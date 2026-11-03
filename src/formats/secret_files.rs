use std::path::Path;

/// Whether `path` is recognized as a secret, credential, or sensitive file.
///
/// Matches the sensitive file reference rules from cc-safety-net:
/// - Sensitive file names (.env, .npmrc, .pypirc, .netrc, .git-credentials, id_rsa, id_ed25519, id_ecdsa, credentials)
/// - Sensitive file name patterns (.env.*, extensionless *_(rsa|dsa|ed25519|ecdsa))
/// - Key file name variants (id_rsa-*, id_rsa_*, id_rsa.bak, id_rsa.key, credentials.pem, etc.)
/// - Sensitive file extensions (.pem, .key, .ovpn, .p12, .kdbx, etc.)
/// - Sensitive home folders & coding agent credentials (~/.ssh, ~/.aws, ~/.docker/config.json, ~/.gemini, etc.)
///
/// Exempts template env files (.env.example, .env.sample, .env.template, .env.tpl, .env.defaults)
/// and public key files (*.pub).
pub(crate) fn is_secret_file(path: &Path) -> bool {
    let Some(file_name) = path.file_name().and_then(|n| n.to_str()) else {
        return false;
    };
    let file_name_lower = file_name.to_ascii_lowercase();

    // 1. Check exemptions first:
    // Public keys (*.pub for key prefixes)
    if file_name_lower.ends_with(".pub") {
        let stem = &file_name_lower[..file_name_lower.len() - 4];
        if is_key_stem(stem) {
            return false;
        }
    }

    // .env exemptions
    if file_name_lower == ".env.defaults" {
        return false;
    }
    if file_name_lower.starts_with(".env.") {
        const ENV_TEMPLATE_SUFFIXES: &[&str] = &[".example", ".sample", ".template", ".tpl"];
        if ENV_TEMPLATE_SUFFIXES
            .iter()
            .any(|s| file_name_lower.ends_with(s))
        {
            return false;
        }
        if file_name_lower.starts_with(".env.example.")
            || file_name_lower.starts_with(".env.sample.")
        {
            return false;
        }
    }

    // 2. Sensitive exact basenames
    const SENSITIVE_BASENAMES: &[&str] = &[
        ".env",
        ".npmrc",
        ".pypirc",
        ".netrc",
        ".git-credentials",
        "id_rsa",
        "id_ed25519",
        "id_ecdsa",
        "credentials",
    ];
    if SENSITIVE_BASENAMES.iter().any(|b| file_name_lower == *b) {
        return true;
    }

    // 3. Sensitive file name patterns
    // .env.* (already checked exemptions above)
    if file_name_lower.starts_with(".env.") {
        return true;
    }

    // Extensionless SSH key-like basenames: *_(rsa|dsa|ed25519|ecdsa)
    if !file_name_lower.contains('.') {
        const SSH_KEY_SUFFIXES: &[&str] = &["_rsa", "_dsa", "_ed25519", "_ecdsa"];
        if SSH_KEY_SUFFIXES
            .iter()
            .any(|s| file_name_lower.ends_with(s))
        {
            return true;
        }
    }

    // 4. Key file name variants for prefixes: id_rsa, id_dsa, id_ed25519, id_ecdsa, credentials
    const KEY_PREFIXES: &[&str] = &["id_rsa", "id_dsa", "id_ed25519", "id_ecdsa", "credentials"];
    for prefix in KEY_PREFIXES {
        if let Some(rest) = file_name_lower.strip_prefix(prefix) {
            // id_rsa-* or id_rsa_*
            if rest.starts_with('-') || rest.starts_with('_') {
                return true;
            }
            // id_rsa.<suffix>
            const DOT_SUFFIXES: &[&str] = &[
                ".bak",
                ".backup",
                ".copy",
                ".disabled",
                ".key",
                ".old",
                ".orig",
                ".pem",
                ".save",
                ".tmp",
            ];
            if DOT_SUFFIXES.contains(&rest) {
                return true;
            }
        }
    }

    // 5. Sensitive file extensions
    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase());
    if let Some(ref ext) = extension {
        const SENSITIVE_EXTENSIONS: &[&str] = &[
            "agilekeychain",
            "asc",
            "bek",
            "cscfg",
            "fve",
            "gnucash",
            "jks",
            "keychain",
            "kwallet",
            "mdf",
            "ovpn",
            "p12",
            "pcap",
            "pem",
            "pfx",
            "pkcs12",
            "psafe3",
            "rdp",
            "sdf",
            "tblk",
            "tpm",
        ];
        if SENSITIVE_EXTENSIONS.iter().any(|e| ext == e) {
            return true;
        }

        // Sensitive extension patterns
        if ext == "key" || ext == "keypair" {
            return true;
        }
        if ext == "keystore" || ext == "keyring" {
            return true;
        }
        if ext == "kdb" || ext == "kdbx" {
            return true;
        }
    }

    // 6. Sensitive home folders & coding agent credentials in path
    let normalized = format!(
        "/{}",
        path.to_string_lossy()
            .trim_start_matches(['/', '\\'])
            .to_ascii_lowercase()
            .replace('\\', "/")
    );

    // Any path inside .ssh, .aws, or .gcp folders
    if normalized.contains("/.ssh/") || normalized.ends_with("/.ssh") {
        return true;
    }
    if normalized.contains("/.aws/") || normalized.ends_with("/.aws") {
        return true;
    }
    if normalized.contains("/.gcp/") || normalized.ends_with("/.gcp") {
        return true;
    }

    // Home config folders & backup variants
    if normalized.contains("/.config/gcloud/") || normalized.ends_with("/.config/gcloud") {
        return true;
    }
    if normalized.ends_with("/.config/gh/hosts.yml") {
        return true;
    }

    // ~/.kube/config (+ variants)
    const BACKUP_SUFFIXES: &[&str] = &[
        "",
        ".bak",
        ".backup",
        ".copy",
        ".disabled",
        ".old",
        ".orig",
        ".save",
        ".tmp",
    ];
    for suffix in BACKUP_SUFFIXES {
        let kube_pattern = format!("/.kube/config{suffix}");
        if normalized.ends_with(&kube_pattern) {
            return true;
        }
        let docker_pattern = format!("/.docker/config.json{suffix}");
        if normalized.ends_with(&docker_pattern) {
            return true;
        }
    }

    // 7. Coding agent credentials
    const CODING_CLI_CREDENTIAL_PATHS: &[&str] = &[
        "/.claude/.credentials.json",
        "/.codex/auth.json",
        "/.codex/.credentials.json",
        "/.codex/secrets",
        "/.codex/.sandbox-secrets",
        "/.gemini/oauth_creds.json",
        "/.gemini/mcp-oauth-tokens.json",
        "/.gemini/a2a-oauth-tokens.json",
        "/.gemini/gemini-credentials.json",
        "/.copilot/config.json",
        "/.copilot/mcp-oauth-config",
        "/.copilot/mcp-secrets",
        "/.kimi-code/server.token",
        "/.kimi-code/credentials",
        "/.kimi/credentials",
        "/.kimi/mcp-oauth",
        "/.local/share/opencode/auth.json",
        "/.local/share/opencode/mcp-auth.json",
        "/.local/share/opencode/opencode.db",
        "/.local/share/amp/secrets.json",
        "/.amp/oauth",
        "/.cursor/auth.json",
        "/.config/cursor/auth.json",
        "/.grok/auth.json",
        "/.grok/mcp_credentials.json",
        "/.factory/auth.encrypted",
        "/.factory/auth.v2.file",
        "/.factory/auth.v2.key",
        "/.factory/auth.v2.loginkeychain",
        "/.local/share/devin/credentials.toml",
        "/.local/share/devin/mcp/oauth",
    ];
    if CODING_CLI_CREDENTIAL_PATHS
        .iter()
        .any(|p| normalized.ends_with(p))
    {
        return true;
    }
    // Cursor project mcp-auth.json
    if normalized.contains("/.cursor/projects/") && normalized.ends_with("/mcp-auth.json") {
        return true;
    }

    false
}

fn is_key_stem(stem: &str) -> bool {
    const KEY_PREFIXES: &[&str] = &["id_rsa", "id_dsa", "id_ed25519", "id_ecdsa"];
    KEY_PREFIXES.iter().any(|prefix| {
        stem == *prefix
            || stem.starts_with(&format!("{prefix}-"))
            || stem.starts_with(&format!("{prefix}_"))
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    #[test]
    fn sensitive_exact_basenames_are_detected() {
        assert!(is_secret_file(Path::new(".env")));
        assert!(is_secret_file(Path::new(".npmrc")));
        assert!(is_secret_file(Path::new(".pypirc")));
        assert!(is_secret_file(Path::new(".netrc")));
        assert!(is_secret_file(Path::new(".git-credentials")));
        assert!(is_secret_file(Path::new("id_rsa")));
        assert!(is_secret_file(Path::new("id_ed25519")));
        assert!(is_secret_file(Path::new("id_ecdsa")));
        assert!(is_secret_file(Path::new("credentials")));
    }

    #[test]
    fn env_variants_and_exemptions() {
        assert!(is_secret_file(Path::new(".env.local")));
        assert!(is_secret_file(Path::new(".env.production")));
        assert!(is_secret_file(Path::new(".env.staging.secrets")));

        // Exemptions:
        assert!(!is_secret_file(Path::new(".env.example")));
        assert!(!is_secret_file(Path::new(".env.sample")));
        assert!(!is_secret_file(Path::new(".env.template")));
        assert!(!is_secret_file(Path::new(".env.tpl")));
        assert!(!is_secret_file(Path::new(".env.defaults")));
        assert!(!is_secret_file(Path::new(".env.example.local")));
        assert!(!is_secret_file(Path::new(".env.sample.dev")));
        assert!(!is_secret_file(Path::new(".env.test.template")));
    }

    #[test]
    fn ssh_key_basenames_and_exemptions() {
        assert!(is_secret_file(Path::new("my_rsa")));
        assert!(is_secret_file(Path::new("deploy_ed25519")));
        assert!(is_secret_file(Path::new("server_ecdsa")));
        assert!(is_secret_file(Path::new("legacy_dsa")));

        // With extension, not extensionless key:
        assert!(!is_secret_file(Path::new("my_rsa.txt")));
        assert!(!is_secret_file(Path::new("report_ed25519.pdf")));

        // Public keys exempt:
        assert!(!is_secret_file(Path::new("id_rsa.pub")));
        assert!(!is_secret_file(Path::new("id_ed25519.pub")));
        assert!(!is_secret_file(Path::new("id_ecdsa.pub")));
        assert!(!is_secret_file(Path::new("id_rsa-work.pub")));
    }

    #[test]
    fn key_file_variants() {
        assert!(is_secret_file(Path::new("id_rsa-work")));
        assert!(is_secret_file(Path::new("id_rsa_backup")));
        assert!(is_secret_file(Path::new("id_rsa.bak")));
        assert!(is_secret_file(Path::new("id_rsa.key")));
        assert!(is_secret_file(Path::new("id_rsa.pem")));
        assert!(is_secret_file(Path::new("credentials-prod")));
        assert!(is_secret_file(Path::new("credentials_aws")));
        assert!(is_secret_file(Path::new("credentials.old")));
    }

    #[test]
    fn sensitive_extensions_and_patterns() {
        assert!(is_secret_file(Path::new("cert.pem")));
        assert!(is_secret_file(Path::new("vpn.ovpn")));
        assert!(is_secret_file(Path::new("client.p12")));
        assert!(is_secret_file(Path::new("traffic.pcap")));
        assert!(is_secret_file(Path::new("bundle.pfx")));
        assert!(is_secret_file(Path::new("remote.rdp")));
        assert!(is_secret_file(Path::new("wallet.kdbx")));
        assert!(is_secret_file(Path::new("legacy.kdb")));
        assert!(is_secret_file(Path::new("server.key")));
        assert!(is_secret_file(Path::new("client.keypair")));
        assert!(is_secret_file(Path::new("store.keystore")));
        assert!(is_secret_file(Path::new("login.keyring")));
    }

    #[test]
    fn home_folders_and_coding_credentials() {
        let ssh_file = PathBuf::from("dummy/.ssh/config");
        assert!(is_secret_file(&ssh_file));

        let aws_file = PathBuf::from("user/.aws/config");
        assert!(is_secret_file(&aws_file));

        let gcloud_file = PathBuf::from("user/.config/gcloud/credentials.db");
        assert!(is_secret_file(&gcloud_file));

        let gh_hosts = PathBuf::from("user/.config/gh/hosts.yml");
        assert!(is_secret_file(&gh_hosts));

        let kube_config = PathBuf::from("user/.kube/config");
        assert!(is_secret_file(&kube_config));

        let kube_backup = PathBuf::from("user/.kube/config.bak");
        assert!(is_secret_file(&kube_backup));

        let docker_cfg = PathBuf::from("user/.docker/config.json");
        assert!(is_secret_file(&docker_cfg));

        let gemini_token = PathBuf::from("user/.gemini/oauth_creds.json");
        assert!(is_secret_file(&gemini_token));

        let claude_cred = PathBuf::from("user/.claude/.credentials.json");
        assert!(is_secret_file(&claude_cred));

        let cursor_auth = PathBuf::from("user/.cursor/auth.json");
        assert!(is_secret_file(&cursor_auth));

        let cursor_proj = PathBuf::from("user/.cursor/projects/demo/mcp-auth.json");
        assert!(is_secret_file(&cursor_proj));
    }

    #[test]
    fn ordinary_files_are_not_secrets() {
        assert!(!is_secret_file(Path::new("main.rs")));
        assert!(!is_secret_file(Path::new("README.md")));
        assert!(!is_secret_file(Path::new("package.json")));
        assert!(!is_secret_file(Path::new("notes.txt")));
        assert!(!is_secret_file(Path::new("style.css")));
    }
}

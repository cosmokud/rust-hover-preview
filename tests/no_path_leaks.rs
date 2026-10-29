//! A repository that writes down where it was built, or who built it, has published
//! something it cannot take back: a path typed into a fixture is a path in every clone,
//! every fork, and every cached diff of every commit that ever carried it.
//!
//! This reads the files a clone receives — what git tracks and what it would add — and fails
//! on the shapes that name an *account's own* directory.
//! A drive letter by itself is not a leak — this repository is full of made-up fixtures
//! like `C:\docs\` and `D:/Pictures/`, and they are what the tests are written against.
//! What is a leak is a path rooted at a profile directory, because that directory is
//! named after the account that owns it.

use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

/// A file this size is a build product or an asset, not source that was written by hand.
const WORTH_READING: u64 = 2 << 20;

fn is_skipped(name: &str) -> bool {
    // `.codegraph` is a code-graph tool's own state — an index and the daemon's log of the
    // runs that wrote it — and the log names this checkout by its absolute path the moment the
    // daemon starts. It is not source and it is never committed (the folder ignores itself
    // outright, which is what the root `.gitignore` says one of these tools needed after it
    // had already published where the repository lived), so there is nothing here for this
    // walk to guard: what it would flag is written by a tool on every run and would have to
    // be deleted before every run to keep it quiet.
    matches!(name, ".git" | "target" | ".scratch" | ".tmp" | ".codegraph")
}

/// The shapes that name an account's own directory: a drive letter followed by an account
/// directory, and the two POSIX spellings of a home. Held as fragments so that the
/// literals in this file do not match the patterns it searches for.
fn shapes() -> Vec<Vec<u8>> {
    vec![
        [b":\\Us".as_slice(), b"ers\\".as_slice()].concat(),
        [b":/Us".as_slice(), b"ers/".as_slice()].concat(),
        [b"/hom".as_slice(), b"e/".as_slice()].concat(),
        [b"/Us".as_slice(), b"ers/".as_slice()].concat(),
    ]
}

/// Every run of backslashes collapsed to one, so that a raw string and an escaped one are
/// read as the same path rather than as two spellings to match separately.
fn collapse(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(bytes.len());
    let mut after_backslash = false;
    for &byte in bytes {
        if byte == b'\\' && after_backslash {
            continue;
        }
        after_backslash = byte == b'\\';
        out.push(byte);
    }
    out
}

fn is_separator(byte: u8) -> bool {
    byte == b'/' || byte == b'\\'
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() || haystack.len() < needle.len() {
        return None;
    }
    haystack.windows(needle.len()).position(|w| w == needle)
}

/// A shape only counts where a name follows it, so that the bare directory on its own —
/// which is what this file's own fragments spell — is not a finding.
fn is_a_leak(line: &[u8], shapes: &[Vec<u8>]) -> bool {
    shapes.iter().any(|shape| {
        find(line, shape)
            .and_then(|at| line.get(at + shape.len()))
            .is_some_and(|next| !is_separator(*next))
    })
}

fn check(path: &Path, root: &Path, shapes: &[Vec<u8>], hits: &mut Vec<String>) {
    if fs::metadata(path).is_ok_and(|meta| meta.len() > WORTH_READING) {
        return;
    }
    let Ok(bytes) = fs::read(path) else { return };
    let text = String::from_utf8_lossy(&bytes);

    for (number, line) in text.lines().enumerate() {
        if is_a_leak(&collapse(line.as_bytes()), shapes) {
            let relative = path.strip_prefix(root).unwrap_or(path).display();
            hits.push(format!("{relative}:{}", number + 1));
        }
    }
}

/// The files a clone receives: what git tracks, plus what it would add, which is the set a
/// commit can carry without asking for it by name. A file the repository ignores is
/// machine-local and cannot be published, so it is not read — the walk below is what a
/// source export falls back to, where there is no checkout to ask.
fn published_files(root: &Path) -> Option<Vec<PathBuf>> {
    let output = Command::new("git")
        .args([
            "ls-files",
            "-z",
            "--cached",
            "--others",
            "--exclude-standard",
        ])
        .current_dir(root)
        .output()
        .ok()?;
    if !output.status.success() {
        return None;
    }

    Some(
        String::from_utf8_lossy(&output.stdout)
            .split('\0')
            .filter(|name| !name.is_empty())
            .map(|name| root.join(name))
            .collect(),
    )
}

fn walk(dir: &Path, root: &Path, shapes: &[Vec<u8>], hits: &mut Vec<String>) {
    let Ok(entries) = fs::read_dir(dir) else {
        return;
    };

    for entry in entries.flatten() {
        let Ok(kind) = entry.file_type() else {
            continue;
        };
        let Some(name) = entry.file_name().to_str().map(str::to_owned) else {
            continue;
        };
        let path = entry.path();

        if kind.is_dir() {
            if !is_skipped(&name) {
                walk(&path, root, shapes, hits);
            }
        } else if kind.is_file() {
            check(&path, root, shapes, hits);
        }
    }
}

/// Every hit in the files a clone receives. A source export has no checkout to ask, so the
/// walk answers instead with everything it can read.
fn hits_in(root: &Path) -> Vec<String> {
    let shapes = shapes();
    let mut hits = Vec::new();

    match published_files(root) {
        Some(files) => {
            for file in &files {
                check(file, root, &shapes, &mut hits);
            }
        }
        None => walk(root, root, &shapes, &mut hits),
    }

    hits
}

#[test]
fn no_file_names_an_account_directory() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let hits = hits_in(root);

    assert!(
        hits.is_empty(),
        "these lines name a profile directory, which is named after the account that owns \
         it:\n{}\nWrite the path with a made-up directory instead — `C:\\downloads\\`, \
         `D:/Pictures/`, `/srv/media/` — or ask for it from the environment.",
        hits.join("\n")
    );
}

/// A file the repository ignores is machine-local by definition, and the one this repository
/// writes itself names the checkout by an absolute path: `post-checkout` rewrites
/// `.cargo/target-dir.local.toml` with `RHP_WORKTREE` on every switch. Reading the directory
/// instead of what git publishes therefore fails every push on a machine that set
/// `rhp.targetDir`, which is the workflow CONTRIBUTING.md asks for.
#[test]
fn a_file_git_ignores_is_not_read() {
    let root = std::env::temp_dir().join(format!("no-path-leaks-{}", std::process::id()));
    let _ = fs::remove_dir_all(&root);
    fs::create_dir_all(&root).unwrap();

    let Ok(init) = Command::new("git")
        .args(["init", "-q"])
        .current_dir(&root)
        .output()
    else {
        let _ = fs::remove_dir_all(&root);
        return;
    };
    if !init.status.success() {
        let _ = fs::remove_dir_all(&root);
        return;
    }

    // Spelled from the fragments the patterns are made of, so this file carries no path of
    // its own for the other test to find.
    let leak = format!(
        "dir = \"C{}someone\"\n",
        String::from_utf8_lossy(&shapes()[1])
    );

    fs::write(root.join(".gitignore"), "local.toml\n").unwrap();
    fs::write(root.join("local.toml"), &leak).unwrap();
    fs::write(root.join("tracked.rs"), &leak).unwrap();

    let hits = hits_in(&root);

    let _ = fs::remove_dir_all(&root);

    assert_eq!(
        hits,
        vec!["tracked.rs:1".to_owned()],
        "only what a clone receives is read: the ignored file is this checkout's own and \
         never leaves it, and the tracked one is what a push publishes"
    );
}

/// What a commit carries besides its files: the message, and the names it was written under.
/// A path in any of those reaches every clone the way a path in a fixture does, and a pass
/// over the files cannot see one because a message is not a file.
///
/// This asks only about the checked-out branch. The commits that carried the paths are not its
/// ancestors, so no depth of history makes them appear here, and the tags that do reach them
/// are not walked at all — scanning every ref would be the stronger check and also the one
/// that could never be made to pass, since an immutable tag is not a fix that is available.
#[test]
fn no_commit_message_names_an_account_directory() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let shapes = shapes();

    let Ok(history) = Command::new("git")
        .args([
            "log",
            "HEAD",
            "--format=%h%x00%B%x00%an%x00%ae%x00%cn%x00%ce%x00",
        ])
        .current_dir(root)
        .output()
    else {
        // Not a checkout: a source export, a tarball. There is no history to read.
        return;
    };
    if !history.status.success() {
        // No HEAD to walk. Nothing is being asserted about a history that is not there.
        return;
    }

    let text = String::from_utf8_lossy(&history.stdout);
    let fields: Vec<&str> = text.split('\0').collect();
    let mut commits = Vec::new();

    for commit in fields.chunks(7) {
        if commit.len() < 7 {
            continue;
        }
        let named = commit[1..]
            .iter()
            .flat_map(|field| field.lines())
            .any(|line| is_a_leak(&collapse(line.as_bytes()), &shapes));
        if named {
            commits.push(commit[0]);
        }
    }

    assert!(
        commits.is_empty(),
        "these commits carry a profile directory in their message or in a name they were \
         written under, which every clone of this branch inherits:\n{}\n\
         `git log -1 <commit>` to see the line. A message is not a file, so nothing else in \
         this repository is watching it.",
        commits.join("\n")
    );
}

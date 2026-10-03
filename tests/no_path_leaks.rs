//! A repository that writes down where it was built, or who built it, has published
//! something it cannot take back: a path typed into a fixture is a path in every clone,
//! every fork, and every cached diff of every commit that ever carried it.
//!
//! This walks the tree and fails on the shapes that name an *account's own* directory.
//! A drive letter by itself is not a leak — this repository is full of made-up fixtures
//! like `C:\docs\` and `D:/Pictures/`, and they are what the tests are written against.
//! What is a leak is a path rooted at a profile directory, because that directory is
//! named after the account that owns it.

use std::fs;
use std::path::Path;

/// A file this size is a build product or an asset, not source that was written by hand.
const WORTH_READING: u64 = 2 << 20;

fn is_skipped(name: &str) -> bool {
    matches!(name, ".git" | "target" | ".scratch" | ".tmp")
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

#[test]
fn no_file_names_an_account_directory() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let shapes = shapes();
    let mut hits = Vec::new();

    walk(root, root, &shapes, &mut hits);

    assert!(
        hits.is_empty(),
        "these lines name a profile directory, which is named after the account that owns \
         it:\n{}\nWrite the path with a made-up directory instead — `C:\\downloads\\`, \
         `D:/Pictures/`, `/srv/media/` — or ask for it from the environment.",
        hits.join("\n")
    );
}

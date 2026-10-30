//! A test of this suite that puts a window on the desktop is a window the user can see, and a
//! window *created* visible is one that takes their keyboard with it: `cargo test` typed into a
//! terminal is a run that must not take the caret out of it, and a run of the suite must not leave
//! rectangles of its own over the desk.
//!
//! So the suite makes its windows one way, and this reads the source of every test module under
//! `src/ui/preview_window/` to hold it to that:
//!
//! - **no window is created with `WS_VISIBLE`**, because `CreateWindowExW` activates a window made
//!   visible wherever it stands — which, and not the box it is put at, is what takes the caret out
//!   of the user's terminal. A window that has to be on the desktop is created hidden, made
//!   composited-as-nothing, and then shown unactivated (see `a_window_of_the_suite` in
//!   `src/ui/preview_window/tests.rs`). A `WS_CHILD` window is the one exception: a child cannot
//!   activate, so a child made visible takes nothing from anybody.
//! - **no window is shown by an activating call**: `SW_SHOW`, `SW_SHOWNORMAL` and the rest of the
//!   family raise the window and take the keyboard with it, so a show here is `SW_SHOWNOACTIVATE`,
//!   or one of the three that are not a show at all — a hide, a minimize and a restore, which are
//!   the behaviour under test rather than a way to put a window up.
//!
//! **What this does not read** is whether a window that is up is drawn: the suite's windows are
//! layered at no alpha by construction (`a_window_of_the_suite`), and the one road that must be
//! *composited* rather than absent — the stand-ins the app's own picture is read off the desk, in
//! `pin_taskbar_window.rs` — is made directly for that reason. Those are shown unactivated, so they
//! take no caret, and they are a rectangle of the suite's for as long as they are up: that is the
//! price of a picture read off the screen, and it is not a thing a layered window can pay.

use std::fs;
use std::path::{Path, PathBuf};

/// The one part of the tree the suite's own windows are made in.
const PREVIEW_WINDOW: &str = "src/ui/preview_window";

/// The call a window is made with, and the one it is put up with.
const CREATED: &str = "CreateWindowExW(";
const SHOWN: &str = "ShowWindow(";

/// The shows that do not activate anything: the one a window of the suite's own is put up with, and
/// the three that are the opposite of putting one up.
const SHOWS_THAT_TAKE_NOTHING: [&str; 4] = [
    "SW_SHOWNOACTIVATE",
    "SW_HIDE",
    "SW_MINIMIZE",
    "SW_RESTORE",
];

/// A module of the suite rather than of the app: `tests.rs` beside the part it tests, or any file
/// under a directory named `tests`. The app's own windows are made by the roads that own them, and
/// what a shipped window is allowed to be is not this file's question.
fn is_a_test_module(path: &Path) -> bool {
    let named_tests = path.file_name().is_some_and(|name| name == "tests.rs");
    let under_tests = path
        .parent()
        .and_then(Path::file_name)
        .is_some_and(|name| name == "tests");

    named_tests || under_tests
}

/// Every `.rs` file under `dir`, walked in the order a filesystem gives them: the walk is a set
/// rather than a sequence, so the order a finding is reported in is sorted below.
fn rust_files(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(dir) else {
        return;
    };

    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            rust_files(&path, out);
        } else if path.extension().is_some_and(|kind| kind == "rs") {
            out.push(path);
        }
    }
}

/// The text of a call's arguments, from its opening bracket to the one that closes it, with string
/// literals stepped over so that a bracket or a word inside one is not read as the call's own.
fn arguments_of(text: &str, from: usize) -> &str {
    let bytes = text.as_bytes();
    let mut depth = 0usize;
    let mut index = from;
    let mut in_string = false;

    while index < bytes.len() {
        let byte = bytes[index];
        if in_string {
            if byte == b'\\' {
                index += 1;
            } else if byte == b'"' {
                in_string = false;
            }
        } else {
            match byte {
                b'"' => in_string = true,
                b'(' => depth += 1,
                b')' => {
                    depth -= 1;
                    if depth == 0 {
                        return &text[from..=index];
                    }
                }
                _ => {}
            }
        }
        index += 1;
    }

    &text[from..]
}

/// Every call of `name` in `text`, as the line it is on and the arguments it is given.
fn calls<'a>(text: &'a str, name: &str) -> Vec<(usize, &'a str)> {
    let mut found = Vec::new();
    let mut at = 0;

    while let Some(offset) = text[at..].find(name) {
        let start = at + offset + name.len() - 1;
        let line = text[..start].lines().count() + 1;
        found.push((line, arguments_of(text, start)));
        at = start + 1;
    }

    found
}

/// Every test module under the preview window, as the path it is read from and the text of it.
fn the_suites_own_modules(root: &Path) -> Vec<(PathBuf, String)> {
    let mut files = Vec::new();
    rust_files(&root.join(PREVIEW_WINDOW), &mut files);
    files.sort();
    files.retain(|path| is_a_test_module(path));

    files
        .into_iter()
        .filter_map(|path| {
            let text = fs::read_to_string(&path).ok()?;
            Some((path, text))
        })
        .collect()
}

/// Where a module is, as this file names it: relative to the checkout, never by its own root.
fn named(root: &Path, path: &Path) -> String {
    path.strip_prefix(root)
        .unwrap_or(path)
        .display()
        .to_string()
}

/// The windows made visible at the moment they are made, which is what activates them.
fn created_visible(modules: &[(PathBuf, String)], root: &Path) -> Vec<String> {
    let mut hits = Vec::new();

    for (path, text) in modules {
        for (line, arguments) in calls(text, CREATED) {
            // A child is the exception: it has no activation of its own to take.
            if arguments.contains("WS_VISIBLE") && !arguments.contains("WS_CHILD") {
                hits.push(format!("{}:{line}", named(root, path)));
            }
        }
    }

    hits
}

/// The windows put up by a call that raises them and takes the keyboard with them.
fn shown_by_an_activating_call(modules: &[(PathBuf, String)], root: &Path) -> Vec<String> {
    let mut hits = Vec::new();

    for (path, text) in modules {
        for (line, arguments) in calls(text, SHOWN) {
            let takes_nothing = SHOWS_THAT_TAKE_NOTHING
                .iter()
                .any(|show| arguments.contains(show));
            if !takes_nothing {
                hits.push(format!("{}:{line}", named(root, path)));
            }
        }
    }

    hits
}

#[test]
fn no_window_of_the_suite_is_created_visible() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let modules = the_suites_own_modules(root);

    assert!(
        modules
            .iter()
            .any(|(path, _)| path.file_name().is_some_and(|name| name == "tests.rs")),
        "the walk found no test module under {PREVIEW_WINDOW}, so it is guarding nothing: it reads \
         the suite as `tests.rs` beside the part it tests and every file under a directory named \
         `tests`, and neither was there"
    );

    let hits = created_visible(&modules, root);

    assert!(
        hits.is_empty(),
        "these windows are created with `WS_VISIBLE`, which activates them wherever they stand — \
         that is what takes the caret out of whatever the user is typing in while the suite runs:\n{}\n\
         Make the window hidden, make it composited-as-nothing and show it unactivated instead — \
         `a_window_of_the_suite` in `src/ui/preview_window/tests.rs` is that road, and it keeps the \
         window on the desktop for the tests that ask the desktop about it.",
        hits.join("\n")
    );
}

#[test]
fn no_window_of_the_suite_is_shown_by_an_activating_call() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let modules = the_suites_own_modules(root);
    let hits = shown_by_an_activating_call(&modules, root);

    assert!(
        hits.is_empty(),
        "these windows are shown by a call that activates them, which takes the keyboard from the \
         user the way a create in `WS_VISIBLE` does:\n{}\n\
         Show them with `SW_SHOWNOACTIVATE` — or `show_a_window_of_the_suite` — and let the hide, \
         the minimize and the restore that are the behaviour under test be the only others.",
        hits.join("\n")
    );
}

/// The scan is only a guard if it is reading the suite: a walk that found no window at all, or that
/// read no module, would pass both of the tests above and hold nothing to anything.
#[test]
fn the_scan_reads_the_windows_the_suite_makes() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let modules = the_suites_own_modules(root);

    let made = modules
        .iter()
        .map(|(_, text)| calls(text, CREATED).len())
        .sum::<usize>();
    assert!(
        made > 0,
        "no window of the suite's own was found under {PREVIEW_WINDOW}: the modules were read \
         ({}) and none of them makes one",
        modules.len()
    );

    let beside_the_part = root.join(PREVIEW_WINDOW);
    let factory = modules
        .iter()
        .find(|(path, _)| path.parent() == Some(beside_the_part.as_path()))
        .map(|(_, text)| text.contains("a_window_of_the_suite"))
        .unwrap_or(false);
    assert!(
        factory,
        "`tests.rs` beside the part the suite tests is where the suite's windows are made and where \
         the road that makes one unseen lives, and this walk did not find it: the walk is reading \
         the wrong tree"
    );
}

//! Both halves of the wrapper `.cargo/config.toml` hands every rustc call to, run the way
//! cargo runs it: through cmd.exe, with the compiler and its arguments after the wrapper.
//!
//! The half that needs the check is the quiet one. A wrapper that took the same path on
//! every machine would still compile, so a tree that builds is not what says it is right:
//! a contributor with sccache would build uncached without ever noticing, and one without
//! it would not build at all.

#![cfg(windows)]

use std::path::Path;
use std::process::Command;

/// The wrapper, under the name `.cargo/config.toml` gives it — a path read relative to the
/// config file's own directory, which is why the run below is rooted at the manifest, and
/// a backslash because `/` in a command position is a switch as far as cmd.exe is concerned.
const WRAPPER: &str = r"tools\sccache-or-rustc.cmd";

/// Runs `cmd /c <wrapper> <arguments…>`, answering with what the wrapper left on stdout and
/// whether it succeeded.
///
/// `path` stands in for PATH when the machine's own is not the one to answer the question
/// with: the directory cmd.exe itself lives in is a PATH carrying the Windows directory and
/// nothing else, which is what a machine without sccache on it looks like.
fn run_wrapper(path: Option<&str>, arguments: &[&str]) -> (String, bool) {
    let mut command = Command::new("cmd");
    command
        .arg("/d")
        .arg("/c")
        .arg(WRAPPER)
        .args(arguments)
        .current_dir(env!("CARGO_MANIFEST_DIR"));
    if let Some(path) = path {
        command.env("PATH", path);
    }

    let output = command
        .output()
        .expect("cmd.exe runs the wrapper on Windows");
    (
        String::from_utf8_lossy(&output.stdout).into_owned(),
        output.status.success(),
    )
}

/// A machine that has sccache compiles through it: nothing asked for, nothing to remember,
/// and no rebuild served without the cache.
///
/// Only asserted where sccache is installed. The machine that never installed it is the
/// fallback below, whose own assertion is the one that covers it.
#[test]
fn the_wrapper_sends_the_compiler_through_sccache_when_it_is_installed() {
    let on_path = Command::new("cmd")
        .args(["/d", "/c", "where sccache"])
        .output()
        .expect("cmd.exe answers whether sccache is on PATH")
        .status
        .success();
    if !on_path {
        eprintln!("sccache is not on PATH: nothing to assert about the machine that has it");
        return;
    }

    let (stdout, success) = run_wrapper(None, &["--version"]);
    assert!(success, "the wrapper ran `sccache --version` itself");
    assert!(
        stdout.starts_with("sccache"),
        "`sccache --version` answers with its own name, and not with {stdout:?}"
    );
}

/// A machine that has no sccache still builds: the compiler runs wrapped exactly as it
/// would have run unwrapped, which is what keeps a contributor who never installed it out
/// of a build that fails before it starts.
#[test]
fn the_wrapper_runs_the_compiler_itself_when_sccache_is_not_installed() {
    // cmd.exe by its own absolute path, which needs no PATH entry at all: the "compiler"
    // handed to the wrapper is `echo`, and what it prints is what a wrapper that went the
    // other way could never print — sccache would be handed `echo` instead, and would fail
    // on a binary it cannot find.
    let cmd = std::env::var("COMSPEC").expect("COMSPEC names cmd.exe");
    let system = Path::new(&cmd)
        .parent()
        .expect("cmd.exe lives in a directory")
        .to_string_lossy()
        .into_owned();

    let (stdout, success) = run_wrapper(Some(&system), &[&cmd, "/c", "echo", "compiler ran"]);
    assert!(success, "the wrapper ran the command it was handed");
    assert!(
        stdout.contains("compiler ran"),
        "the plain branch printed {stdout:?}, which is not what it ran"
    );
}

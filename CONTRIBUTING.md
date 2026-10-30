# Contributing to Rust Hover Preview

Thank you for your interest in contributing. This guide gets you set up and tells you what the
checks cost and what they do not cover.

## Code of Conduct

This project follows the [Contributor Covenant Code of Conduct](CODE_OF_CONDUCT.md).
By participating, you agree to follow it. Please report unacceptable behavior
by opening an issue at https://github.com/cosmokud/rust-hover-preview/issues.

## Ways to contribute

You don't need to write code to help. Useful contributions include:

- Reporting bugs (something doesn't work as expected).
- Suggesting new features or improvements.
- Improving documentation (README, comments, guides).
- Fixing bugs or implementing features.
- Testing preview builds and reporting what you find.

## Before you start

1. Check the [open issues](https://github.com/cosmokud/rust-hover-preview/issues)
   to see if your bug or idea is already reported. If it is, add a comment
   instead of opening a duplicate.
2. For a new feature or a large change, open an issue first and describe your
   idea. This avoids duplicated work and lets maintainers give feedback early.
3. If you want to work on an existing issue, leave a comment saying so.

## Getting set up

Requirements: Windows 11, Rust 1.98.1 (the release `rust-toolchain.toml` pins),
Visual Studio Build Tools (MSVC / C++), and the Windows SDK.

```bash
git clone https://github.com/cosmokud/rust-hover-preview.git
cd rust-hover-preview
cargo build
cargo test
```

The release binary is written to `target/release/rust-hover-preview.exe` by `cargo build --release`.

`cargo test` runs against folders of the tests' own rather than the machine's: a test build of
the app keeps its `config.ini`, its engine records and its browser profiles under
`%TEMP%\rust-hover-preview-tests`, so a test run never touches the settings an app you already
have running is watching.

To share one build cache between checkouts — and between worktrees — set
`git config rhp.targetDir <path>`: the repo's `post-checkout` hook writes it into a git-ignored
`.cargo/target-dir.local.toml`, which `.cargo/config.toml` includes, and `git worktree add` runs
that hook, so a fresh worktree gets it with no manual step. Every checkout then builds into that
one directory, so a dependency is compiled and stored once. What a shared directory cannot share
is this package: cargo's artifact and fingerprint names carry no part of a checkout's path, so a
build in one checkout can be called fresh against another's recorded inputs and leave that tree's
binary behind. The hook therefore writes the checkout's own path beside the directory
(`RHP_WORKTREE`), which `build.rs` watches, so a build in a different checkout recompiles this
package — and nothing else. With nothing set, cargo uses the checkout's own `target/`. The hook
only runs where the repo's hooks are the configured ones (`git config core.hooksPath .githooks`).

See [README.md](README.md) for app usage and [ARCHITECTURE.md](ARCHITECTURE.md)
for a full system overview.

## Making changes

1. Create a branch from `main` with a short, descriptive name:

   ```bash
   git checkout -b fix/short-description
   ```

2. Make your changes. Follow the existing code style and keep changes focused:
   one issue per branch. Leave the formatting of code you did not touch alone,
   unless the formatting is the change itself: a whole-tree sweep is a branch of
   its own, one `cargo fmt --all` commit with nothing else in it.
3. Format, check, and test before you push:

   ```bash
   cargo fmt --all -- --check
   cargo clippy --all-targets -- -D warnings
   cargo test
   ```

   If a check fails, fix it before opening a pull request. `rust-toolchain.toml` pins
   the toolchain, so cargo, rustfmt and clippy answer the same on every machine as they
   do in CI — `manual-ci.yml`, which runs these three commands plus
   `cargo test --test no_path_leaks` and is started by hand from the Actions tab.
4. Write a clear commit message in the present tense, for example
   `Fix preview placement on second monitor` rather than `Fixed...`.
5. Push your branch and open a pull request against `main`.

## Checks, and what they cost

Measured warm, on the pinned toolchain: a clean build is minutes, because the 194 dependencies in
the lock include native C and C++ in the image and archive crates. Then `cargo check --all-targets`
is about 5 s, `cargo clippy --all-targets -- -D warnings` 0.3 s, `cargo fmt --all -- --check` 1.5 s,
and `cargo test` about 10 s of test time.

## Running one test

The suite is over 1,400 tests, so filtering is how you iterate on one of them: cargo matches a
filter against a test's path, so a module name is a filter, and `--test <name>` picks one file
under `tests/` rather than the whole target:

```bash
cargo test pin_zoom              # 49 tests in 0.4 s
cargo test config                # 106 tests in 0.36 s
cargo test --test no_path_leaks  # one file under tests/, 0.8 s
```

## The suite and the desktop

The default suite creates real windows, so it wants an interactive Windows session: one helper,
`src/ui/preview_window/tests.rs::a_window_of_the_suite`, makes each of them as a layered window at
no alpha, shown unactivated — on the desktop, drawn as nothing. A headless session has nothing for
those windows to answer, so verify the data layers instead: `formats`, `config`, `readers` and
`text` read files, parse bytes and answer values. `preview_window` and `explorer_hook` get no
signal there — their tests are written against a desktop, a window made or Explorer's view read —
so a failure is the session answering rather than the change.

## The tests that do not run

37 tests carry `#[ignore]`: they start a third-party program, ask the network, or read a desktop a
run of the suite must not take over — a window shown, a pointer driven, the composed desk. Each
message names its input; `grep -rn '#[ignore' src tests` finds all of them with those reasons.

The engine files are handed in one variable each — `RHP_OFFICE_PROBE`, `RHP_LIBRE_PROBE`,
`RHP_PEAZIP_PROBE`, `RHP_CALIBRE_PROBE`, `RHP_COMIC_PROBE`, `RHP_CONTENT_PROBE`, `RHP_PDF_PROBE`,
`RHP_AUDIO_PROBE`, `RHP_AUDIO_FACTS_PROBE`, `RHP_VIDEO_PROBE`, `RHP_WEBVIEW_PROBE` — and a film
played at the display's own size is `RHP_VIDEO_PERF`, with `RHP_VIDEO_PERF_BOX` for its box. The
rest take no variable: the network ask in `app::updates`, the real mpv in `mpv_ipc`, the audio
stream in `tests/audio_bands`, the pointer probes in `tests/probes`, and the four composed-desktop
reads in `tests/pin_taskbar_window`, which cannot be made headless. `cargo test -- --ignored` runs
them.

## Hooks

`git config core.hooksPath .githooks` is what makes these run:

- `prepare-commit-msg` refuses a message that names a profile directory, which reaches every clone
  the way a path in a fixture does; `--no-verify` does not skip it.
- `pre-push` runs `cargo test --test no_path_leaks` — the tracked tree and the messages in its
  history — and stops the push on a failure, building on a drive rather than a Windows-hosted share
  because a share gives cargo's session directory no file locking.
- `pre-commit` runs `cargo fmt --all -- --check`, then `cargo clippy --all-targets -- -D warnings` when
  the cargo target directory already holds a build — a first commit in a fresh clone skips the lint
  rather than paying for a 194-dependency compile — and stops the commit on a failure, printing the
  command that fixes what failed. It runs no test, so it cannot hang.
  `RHP_SKIP_PRECOMMIT=1 git commit ...` skips both checks for that one commit and prints that it did.

## Opening a pull request

- Fill in the pull request template (what you changed, why, and how you
  tested it).
- Link the issue it fixes, if there is one (for example, `Fixes #12`).
- Keep the pull request focused. If you notice something unrelated, open a
  separate issue or pull request for it.
- A maintainer will review your pull request. You may be asked to make
  changes; just push new commits to the same branch.

## Reporting bugs

Use the bug report issue template and include:

- What you expected to happen and what happened instead.
- Steps to reproduce the problem.
- App version, Windows version, and anything unusual about your setup.
- Screenshots, logs, or sample files if they help (remove private data).

## Suggesting features

Use the feature request issue template and describe the problem your idea
solves, what you would like to happen, and any alternatives you considered.

## Questions

If you are unsure about anything, just ask in an issue. Beginners are welcome,
and no question is too small.

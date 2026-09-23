# Contributing to Rust Hover Preview

Thank you for your interest in contributing! This guide will help you get
started, even if this is your first time contributing to an open-source
project.

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

Requirements: Windows 11, Rust 1.98.1+, Visual Studio Build Tools (MSVC / C++),
and the Windows SDK.

```bash
git clone https://github.com/cosmokud/rust-hover-preview.git
cd rust-hover-preview
cargo build
cargo test
```

The release binary is written to `target/release/rust-hover-preview.exe`:

```bash
cargo build --release
```

See [README.md](README.md) for app usage and [ARCHITECTURE.md](ARCHITECTURE.md)
for a full system overview.

## Making changes

1. Create a branch from `main` with a short, descriptive name:

   ```bash
   git checkout -b fix/short-description
   ```

2. Make your changes. Follow the existing code style and keep changes focused:
   one issue per branch.
3. Format, check, and test before you push:

   ```bash
   cargo fmt --check
   cargo clippy --all-targets -- -D warnings
   cargo test
   ```

   If a check fails, fix it before opening a pull request.
4. Write a clear commit message in the present tense, for example
   `Fix preview placement on second monitor` rather than `Fixed...`.
5. Push your branch and open a pull request against `main`.

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

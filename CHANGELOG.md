# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [0.1.10] - 2026-04-29

### Changed

- Restored the default NSIS older-version prompt while keeping the installer-side process shutdown, config default completion, and old-folder migration logic.
- Synced the `run_at_startup` setting to the Windows startup registry on app launch so autorun is enabled by default on first run.
- Updated release metadata for v0.1.10.

## [0.1.9] - 2026-04-29

### Added

- Added transparent-background preview modes for PNG/WebP transparency: transparent, black, white, and checkerboard.
- Added configurable same-file rehover delay with a 750 ms default and tray controls.
- Added installer migration cleanup for the old `%LOCALAPPDATA%\Rust Hover Preview` install folder.

### Changed

- Switched animated WebP playback to Google's libwebp via `webp-animation` for smoother startup and playback.
- Moved per-user installer output to `%LOCALAPPDATA%\rust-hover-preview` while keeping the display name `Rust Hover Preview`.
- Switched configuration path resolution to the `directories` crate while keeping config at `%APPDATA%\rust-hover-preview\config.ini`.
- Updated the tray off-trigger label to show the configured key, such as `Enable Off Trigger Key (Alt)`.
- Updated the tray icon tooltip to show the full `Rust Hover Preview` app name.

### Fixed

- Updated NSIS and WiX installers to stop a running `rust-hover-preview.exe` before replacing installed files during upgrades.
- Added installer-side config default completion so missing `config.ini` parameters are added without changing existing user values.
- Hardened preview hide behavior so stuck video previews are hidden and ffplay is stopped more aggressively.
- Added Explorer COM/UIA slow-probe backoff and cache clearing to reduce runaway polling after Explorer gets sluggish.
- Fixed repeated same-file hover retrigger timing after the preview self-dismisses.
- Fixed transparent PNG/WebP rendering artifacts by compositing alpha explicitly for each transparency mode.
- Removed unused Explorer helper functions that produced dead-code warnings.

## [0.1.8] - 2026-04-29

### Added

- Added optional off-trigger key support so previews can be suppressed while a configured modifier key is held.
- Added active Explorer Shell view detection and Shell view media indexing for more reliable file resolution.
- Added media indexing from configured search roots to broaden Explorer hover matching.

### Changed

- Improved Explorer media path normalization, URL handling, and URL decoding.
- Cached Explorer real folder lookups to reduce repeated Shell resolution work.
- Updated Windows installer build orchestration to use the PowerShell helper in CI.

### Fixed

- Removed branch-triggered release deployment so installer publishing only follows the intended release flow.

## [0.1.7] - 2026-04-28

### Added

- Added Windows installer packaging via `cargo packager`, producing release `.exe` and `.msi` assets in CI.
- Added `build-installers.ps1` helper flow for local installer generation with artifact validation.

### Changed

- Updated deploy workflow release packaging to publish `target/packager/*.exe` and `target/packager/*.msi` on `v*.*.*` tags.
- Refactored preview window, media loading, and video-dimension internals for clearer structure and maintainability.
- Improved installer build orchestration with a WiX fallback path for constrained local sessions.

### Fixed

- Added explicit Windows Installer (`msiserver`) availability checks before MSI generation.
- Improved installer-build error handling and output validation to fail fast when packaging artifacts are missing.

## [0.1.6] - 2026-04-06

### Changed

- Optimized animated GIF/WebP streaming by switching to queue-based frame transfer, reducing duplicate frame copying and memory churn.
- Limited preview animation decoding to the first 300 streamed frames and standardized animation timing to a 30 FPS ceiling for more stable CPU usage.
- Added a new tray option, `Confirm File Type` (off by default), to optionally validate image headers for mislabeled file extensions.

### Fixed

- Reduced severe CPU spikes while hovering animated WebP files by tightening streaming overlay redraw conditions and easing decoder contention.
- Improved rapid-hover behavior across multiple animated files by making decoder workers more cooperative during long streams.
- Fixed hover preview triggering for `.jpeg` files by treating JPEG aliases (`.jpg`, `.jpeg`, `.jpe`, `.jfif`) consistently during media resolution.
- With `Confirm File Type` enabled, mislabeled images (for example PNG content with a JPG/JPEG extension) now load with the correct decoder.
- Prevented the previous hovered image from flashing briefly when switching to a new file.

## [0.1.5] - 2026-04-01

### Changed

- Improved Explorer folder caching to reduce repeated directory lookups.
- Optimized media file resolution timing in the Explorer hook with folder indexing and faster lookups.

## [0.1.4] - 2026-03-03

### Changed

- Refined folder-navigation input gate to use UI Automation focus changes and a brief cooldown, eliminating stale `GetAsyncKeyState` triggers.

### Fixed

- Eliminated residual previews caused by leftover keyboard state when opening folders; preview now only appears after actual user movement or navigation.

## [0.1.3] - 2026-03-03

### Added

- Post-folder-navigation input gate that suspends preview until explicit user input is detected.

### Changed

- Mouse hover target resolution now validates accessibility hit-testing against the actual cursor position.
- Prioritized media file resolution in the Explorer folder currently under the cursor before global folder fallback.

### Fixed

- Prevented automatic preview of Explorer's first auto-selected item when opening a folder with a stationary cursor.
- Keyboard preview now starts only after real navigation input (arrow/Home/End/PageUp/PageDown) following folder changes.
- Removed dead-code warnings from unused helper functions to keep builds warning-free.

## [0.1.2] - 2026-03-03

### Added

- Separate cursor-over detection for image preview and ffplay video preview windows.
- Short startup grace period for video previews to avoid immediate dismissal while ffplay initializes.

### Changed

- Refined preview dismissal logic to dismiss on mouse movement over preview windows, preserving keyboard navigation when the cursor is stationary.
- Improved ffplay window discovery by preferring visible, top-level windows and selecting the largest valid candidate.
- Reasserted topmost and no-activate window styles during video playback to keep the preview above Explorer.
- Updated ffprobe/ffplay invocation flags to better tolerate corrupt frames and missing timestamps.

### Fixed

- Prevented premature video preview dismissal during ffplay startup race conditions.
- Reduced false preview closures when no hovered file is detected while the mouse is not moving.

## [0.1.1] - 2026-03-03

### Added

- Keyboard preview support when navigating focused items in Explorer (arrow keys/tab).
- Support for animated WebP files and improved GIF/WebP streaming with background decoding.
- Loading spinner overlay displayed while animated media frames load in the background.
- New system tray and configuration options: follow cursor positioning and video volume control.
- GitHub Actions workflows for deployment and nightly builds.

### Changed

- Refactored `preview_window` module and media streaming logic for smoother animation and reduced coupling.
- Enhanced animated media playback: multiple-frame skipping, accurate timing, and background frame buffering.
- Optimized CPU usage by varying polling rates based on Explorer window state (hidden/minimized/active).
- Configuration now uses INI format with automatic save-on-load; added new fields (`follow_cursor`, `video_volume`).

### Fixed

- Prevent preview window from hiding while keyboard-based hover is active.
- Corrected wording/formatting of preview delay options in the context menu.
- Improved frame decoding reliability for animated formats.

### Miscellaneous

- Updated README to document new features and configuration settings.
- Added initial configuration file generation on first run.

## [0.1.0] - 2026-02-03

### Added

- Image preview on hover for JPG, JPEG, PNG, GIF, BMP, ICO, TIFF/TIF, and WebP
- Animated GIF and animated WebP playback
- Video preview on hover for MP4, WebM, MKV, AVI, MOV, WMV, FLV, and M4V (via FFmpeg tools)
- System tray controls for preview enable/disable, startup toggle, preview position, and video volume
- INI configuration file in %APPDATA%\rust-hover-preview\config.ini for hover delay, preview enablement, and playback settings

# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [0.1.6] - 2026-04-06

### Changed

- Optimized animated GIF/WebP streaming by switching to queue-based frame transfer, reducing duplicate frame copying and memory churn.
- Limited preview animation decoding to the first 300 streamed frames and standardized animation timing to a 30 FPS ceiling for more stable CPU usage.

### Fixed

- Reduced severe CPU spikes while hovering animated WebP files by tightening streaming overlay redraw conditions and easing decoder contention.
- Improved rapid-hover behavior across multiple animated files by making decoder workers more cooperative during long streams.
- Fixed hover preview triggering for `.jpeg` files by treating JPEG aliases (`.jpg`, `.jpeg`, `.jpe`, `.jfif`) consistently during media resolution.
- Added quick header-based image format detection so mislabeled files (for example PNG content with a JPG/JPEG extension) still load with the correct decoder.

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

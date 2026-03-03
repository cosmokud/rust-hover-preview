# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

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

# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.70+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

Rust Hover Preview is a Windows 11 system tray app that shows instant image and video previews in File Explorer when you hover files with the mouse or navigate with the keyboard.

Inspired by QTTabBar (QuizoApps) hover preview.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer
- Static image previews plus animated GIF playback and libwebp-backed animated WebP playback
- Video previews through FFmpeg (`ffplay` + `ffprobe`)
- Tray controls for enable/disable, delay, positioning, scaling, startup, off-trigger key, and volume
- Explorer Shell view detection, folder caching, and path normalization for reliable hover matching
- Topmost, non-activating preview windows designed to avoid focus stealing
- Per-monitor DPI awareness to reduce scaling artifacts on high-DPI displays
- Display-aware placement that keeps the preview inside the monitor under the cursor or focused item, with frames painted before the window is revealed so a new hover never flashes the previous image
- Scale-aware sizing from 25% to 400% or fit-to-screen, always clamped to the space available on the display so a scaled-up preview is never clipped
- EnumWindows-based Explorer detection with CabinetWClass/ExplorerWClass class matching to keep idle polling light and avoid Explorer-side COM allocations, plus input-grace helpers that throttle hover and keyboard focus probes to recent user activity
- Keyboard previews take priority over the pointer: a preview opened with the arrow keys stays on top when a parked cursor sits under it, mouse-driven triggers stay frozen until the mouse is moved on purpose, and the cursor-over-preview check is skipped instead of running on every poll
- Single-instance enforcement: launching the app again exits immediately instead of adding a duplicate tray icon, Explorer hook, and preview window

## Supported Formats

### Images

`jpg`, `jpeg`, `png`, `gif`, `bmp`, `ico`, `tiff`, `tif`, `webp`

### Videos (FFmpeg required)

`mp4`, `webm`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `m4v`

## Installation (Recommended)

Each release provides two asset options:

- **`rust-hover-preview_<version>_x64-setup.exe`** — the NSIS installer. Run it to install to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- **`rust-hover-preview.exe`** — the standalone portable binary. Place it in any folder on your PC (for example: `C:\Tools\RustHoverPreview`) and run it directly. No installation needed.

1. Open [Releases](../../releases)
2. Download your preferred asset
3. Run the installer or place the portable binary wherever you like
4. Launch Rust Hover Preview

No Rust toolchain is needed when installing from Releases.

> **Note for existing users upgrading from an earlier version:**  
> The installer handles upgrades automatically, cleaning up the old `%LOCALAPPDATA%\Rust Hover Preview` folder if present.

## Optional: Enable Video Preview (FFmpeg)

Video previews require `ffplay` and `ffprobe` available in `PATH`.

### Option A: Install with winget

```powershell
winget install --id Gyan.FFmpeg -e
```

Then reopen your terminal and verify:

```powershell
ffplay -version
ffprobe -version
```

### Option B: Manual install

1. Download a Windows FFmpeg build from https://ffmpeg.org/download.html
2. Extract it to a location such as `C:\ffmpeg`
3. Add `C:\ffmpeg\bin` to your user `PATH`
4. Open a new terminal and run:

```powershell
ffplay -version
ffprobe -version
```

## Usage

1. Start the app (tray icon appears)
2. Hover media files in Explorer to preview them
3. Use keyboard navigation in Explorer (arrow keys/tab) to trigger focused-item previews
4. Right-click the tray icon to configure behavior

## System Tray Menu

- **Enable Preview**: Turn previews on or off
- **Preview Delay**: `Instant (0 ms)`, `Fast (200 ms)`, `Medium (500 ms)`, `Relaxed (750 ms)`, `Slow (1000 ms)`
- **Same File Rehover Delay**: `Instant (0 ms)`, `Fast (200 ms)`, `Medium (500 ms)`, `Relaxed (750 ms)`, `Slow (1000 ms)` — delay before the same file can preview again after the preview self-dismisses
- **Video Volume**: `Max (100%)`, `High (80%)`, `Medium (50%)`, `Low (25%)`, `Very Low (10%)`, `Mute (0%)`
- **Preview Position**: `Follow Cursor` or `Best Position`
- **Preview Scaling**: `Fit to Screen`, `400%`, `300%`, `200%`, `150%`, `100% (Default)`, `50%`, `25%` — sizes the preview relative to the image or video's native resolution instead of always showing it verbatim. `Fit to Screen` enlarges the preview as much as the display allows. Any scale that would extend past the screen edge is reduced to fit, so the preview is never clipped, in both `Follow Cursor` and `Best Position` modes.
- **Transparent Background**: `Transparent`, `Black`, `White`, or `Checkerboard`
- **Enable Off Trigger Key**: Temporarily suppress previews while the displayed configured key is held
- **Confirm File Type**: When enabled, validates file content signatures (magic bytes) against the extension to avoid loading mislabeled files. If previews don't appear for certain files that should be supported, try enabling this option — the app will attempt to decode them by their true content type rather than relying solely on the file extension.
- **Run at Startup**: Add/remove startup entry in Windows
- **Edit Config.ini**: Open configuration file in your default editor
- **Exit**: Close the application

## Configuration

Settings are stored at:

```text
%APPDATA%\rust-hover-preview\config.ini
```

Example:

```ini
[settings]
run_at_startup=true
hover_delay_ms=0
same_file_rehover_delay_ms=750
preview_enabled=true
enable_off_trigger_key=true
off_trigger_key=alt
confirm_file_type=false
follow_cursor=false
transparent_background=black
webp_playback_fps=90
video_volume=0
preview_scale=100
```

- When `enable_off_trigger_key` is enabled, hold the configured `off_trigger_key` to keep previews hidden while browsing Explorer.
- When `confirm_file_type` is enabled, the app validates file content signatures (magic bytes) against the extension — useful for files with incorrect extensions.
- `webp_playback_fps` controls the maximum playback speed for animated WebP files (1–90 FPS; 0 resets to the default of 90).
- `preview_scale` controls the preview size relative to the media's native resolution. Use `fit` to scale the preview as large as the display area allows, or a percentage between 1 and 1000 written as `200`, `200%`, or `75`. A scale larger than the available space is reduced to fit so the preview cannot be clipped; `0` resets to the default of 100.

## Build from Source

### Requirements

- Windows 11
- Rust toolchain 1.70+
- Visual Studio Build Tools (MSVC)
- Windows SDK

### Build Commands

```bash
# Debug
cargo build

# Release
cargo build --release
```

Release binary output:

```text
target/release/rust-hover-preview.exe
```

## Architecture Notes

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview.

- Uses Windows accessibility APIs (MSAA + UI Automation) to resolve hovered/focused Explorer items
- Uses Shell COM APIs to identify active Explorer windows and folders
- Uses GDI for image rendering in a layered topmost preview window
- Uses Google's libwebp through `webp-animation` for animated WebP decoding
- Uses `directories` for Windows roaming configuration paths
- Uses `ffprobe` for video dimensions and `ffplay` for video playback
- Sets per-monitor DPI awareness (v2 with fallback) on startup to prevent scaling artifacts on layered windows
- Bounds the preview to the monitor under the cursor or focused item via `MonitorFromPoint`/`GetMonitorInfoW` (with a virtual-screen fallback) and repositions the window before installing a frame, so cross-display hops at different scale never strand a stale preview
- Uses the registry (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`) for startup control
- Enforces a single running instance with a session-local named mutex, exiting before COM setup, hooks, or background threads start
- Counts and classifies Explorer browser windows via EnumWindows and CabinetWClass/ExplorerWClass class matching, so idle polling never spins up Explorer's shell automation providers
- Gates hover and keyboard focus probes behind input-grace windows (recent_elapsed_within, should_probe_keyboard_focus, should_probe_hover_resolver, should_probe_stationary_hover) and a stationary_hover_probe_done latch to avoid redundant accessibility work for a parked cursor

## License

MIT. See [LICENSE](LICENSE).

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
- Tray controls for enable/disable, delay, positioning, startup, off-trigger key, and volume
- Explorer Shell view detection, folder caching, and path normalization for reliable hover matching
- Topmost, non-activating preview windows designed to avoid focus stealing

## Supported Formats

### Images

`jpg`, `jpeg`, `png`, `gif`, `bmp`, `ico`, `tiff`, `tif`, `webp`

### Videos (FFmpeg required)

`mp4`, `webm`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `m4v`

## Installation (Recommended)

Download prebuilt assets from the repository **Releases** tab:

1. Open [Releases](../../releases)
2. Download one of the latest assets:
   - `.msi` installer (recommended for most users)
   - `.exe` portable build
3. If you downloaded `.msi`, run the installer. It installs to `%LOCALAPPDATA%\rust-hover-preview`. If you downloaded `.exe`, place it in any folder you prefer (for example: `C:\Tools\RustHoverPreview`)
4. Launch Rust Hover Preview

No Rust toolchain is needed when installing from Releases.

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
- **Same File Rehover Delay**: delay before the same file can preview again after the preview self-dismisses
- **Video Volume**: `Max (100%)`, `High (80%)`, `Medium (50%)`, `Low (25%)`, `Very Low (10%)`, `Mute (0%)`
- **Preview Position**: `Follow Cursor` or `Best Position`
- **Transparent Background**: `Transparent`, `Black`, `White`, or `Checkerboard`
- **Enable Off Trigger Key**: Temporarily suppress previews while the displayed configured key is held
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
follow_cursor=false
transparent_background=transparent
video_volume=0
```

When `enable_off_trigger_key` is enabled, hold the configured `off_trigger_key` to keep previews hidden while browsing Explorer.

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

- Uses Windows accessibility APIs (MSAA + UI Automation) to resolve hovered/focused Explorer items
- Uses Shell COM APIs to identify active Explorer windows and folders
- Uses GDI for image rendering in a layered topmost preview window
- Uses Google's libwebp through `webp-animation` for animated WebP decoding
- Uses `directories` for Windows roaming configuration paths
- Uses `ffprobe` for video dimensions and `ffplay` for video playback
- Uses the registry (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`) for startup control

## License

MIT. See [LICENSE](LICENSE).

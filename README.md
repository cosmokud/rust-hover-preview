# Rust Hover Preview

Rust Hover Preview is a Windows 11 system tray app that shows instant image and video previews in File Explorer when you hover files with the mouse or navigate with the keyboard.

Inspired by QTTabBar (QuizoApps) hover preview.

![Showcase](https://github.com/user-attachments/assets/3f1b3360-dfa4-47c0-bb4a-a2971d729cd8)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer
- Static image previews plus animated GIF/WebP playback
- Video previews through FFmpeg (`ffplay` + `ffprobe`)
- Tray controls for enable/disable, delay, positioning, startup, and volume
- Topmost, non-activating preview windows designed to avoid focus stealing

## Supported Formats

### Images

`jpg`, `jpeg`, `png`, `gif`, `bmp`, `ico`, `tiff`, `tif`, `webp`

### Videos (FFmpeg required)

`mp4`, `webm`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `m4v`

## Installation (Recommended)

Download the prebuilt executable from the repository **Releases** tab:

1. Open [Releases](../../releases)
2. Download the latest `rust-hover-preview.exe` asset
3. Place it in any folder you prefer (for example: `C:\Tools\RustHoverPreview`)
4. Launch `rust-hover-preview.exe`

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
- **Preview Delay**: `Instant (0 ms)`, `Fast (200 ms)`, `Medium (500 ms)`, `Slow (1000 ms)`
- **Video Volume**: `Max (100%)`, `High (80%)`, `Medium (50%)`, `Low (25%)`, `Very Low (10%)`, `Mute (0%)`
- **Preview Position**: `Follow Cursor` or `Best Position`
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
preview_enabled=true
follow_cursor=false
video_volume=0
```

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
- Uses `ffprobe` for video dimensions and `ffplay` for video playback
- Uses the registry (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`) for startup control

## License

MIT. See [LICENSE](LICENSE).

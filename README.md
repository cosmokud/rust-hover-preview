# Rust Hover Preview

Windows 11 system tray application that shows image and video previews when hovering over files in Windows Explorer. Inspired by QTTabBar (QuizoApps) hover preview.

[hover_preview_2.webm](https://github.com/user-attachments/assets/3f1b3360-dfa4-47c0-bb4a-a2971d729cd8)

## Features

- **Image preview on hover** for common formats (JPG, JPEG, PNG, GIF, BMP, ICO, TIFF/TIF, WebP)
- **Animated GIF** and **animated WebP** playback
- **Video preview on hover** (MP4, WebM, MKV, AVI, MOV, WMV, FLV, M4V) using FFmpeg tools
- **System tray controls** for enabling/disabling previews, startup, video volume, preview position, and preview delay
- **Configurable hover delay**

## System Tray Menu

Right-click the system tray icon to access:

- **Enable Preview**: Toggle all previews on/off
- **Preview Delay**: Instant (0 ms), Fast (200 ms), Medium (500 ms), Slow (1000 ms)
- **Video Volume**: Set preview volume from 0–100%
- **Preview Position**: Follow Cursor or Best Position
- **Run at Startup**: Toggle automatic startup with Windows
- **Edit Config.ini**: Open `config.ini` in the default text editor
- **Exit**: Close the application

## Requirements

- Windows 11
- Rust toolchain (1.70+)
- Windows 11 SDK
- Visual Studio Build Tools (for Windows linking)
- **Optional:** FFmpeg in PATH (for video previews)
  - Required tools: `ffplay` and `ffprobe`

## Install FFmpeg in PATH (Windows)

### Option A: Install with winget (recommended)

```powershell
winget install --id Gyan.FFmpeg -e
```

Close and reopen your terminal, then verify:

```powershell
ffplay -version
ffprobe -version
```

### Option B: Manual install

1. Download a Windows FFmpeg build from https://ffmpeg.org/download.html.
2. Extract the archive to a folder such as `C:\ffmpeg`.
3. Add `C:\ffmpeg\bin` to your user PATH:
   - Open **Edit environment variables for your account**
   - Select **Path** → **Edit** → **New**
   - Add `C:\ffmpeg\bin` and save
4. Open a new terminal and run:

```powershell
ffplay -version
ffprobe -version
```

## Building

```bash
# Debug build
cargo build

# Release build (optimized)
cargo build --release
```

### Custom Icon

To add a custom application icon:

1. Place your `.ico` file at `assets/icon.ico`
2. Rebuild the application

## Installation

1. Build the release version: `cargo build --release`
2. Copy `target/release/rust-hover-preview.exe` to your preferred location
3. Run the application
4. (Optional) Enable "Run at Startup" from the system tray menu

## Configuration

Settings are stored in:

```
%APPDATA%\rust-hover-preview\config.ini
```

Example configuration:

```ini
[settings]
run_at_startup=false
hover_delay_ms=0
preview_enabled=true
follow_cursor=false
video_volume=0
```

## How It Works

The application uses several Windows APIs:

- **UI Accessibility (MSAA)** to detect the hovered item in Explorer
- **Shell Windows API** to resolve active Explorer windows and folders
- **GDI** for rendering image previews without stealing focus
- **Registry API** for managing startup entries
- **FFmpeg (ffplay/ffprobe)** for video preview playback and sizing

## Notes

- If FFmpeg is not installed or not in PATH, video previews are skipped.
- The preview window is layered, topmost, and does not steal focus.
- No telemetry is collected.

## License

MIT License - See LICENSE for details

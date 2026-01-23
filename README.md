# Rust Hover Preview

A Windows 11 system tray application that shows image previews when hovering over images in Windows Explorer.

## Features

- **Image Preview on Hover**: When you hover over an image file in Windows Explorer, a preview window appears at the bottom-right of your cursor position.
- **System Tray Integration**: Runs quietly in the system tray with minimal resource usage.
- **Run at Startup**: Toggle option to automatically start the application when Windows boots.
- **Supported Formats**: JPG, JPEG, PNG, GIF, BMP, ICO, TIFF, TIF, WebP, SVG

## System Tray Menu

Right-click the system tray icon to access:

- **Run at Startup**: Toggle to enable/disable automatic startup with Windows
- **Exit**: Close the application

## Building

### Prerequisites

- Rust toolchain (1.70+)
- Windows 11 SDK
- Visual Studio Build Tools (for Windows linking)

### Build Commands

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

## How It Works

The application uses several Windows APIs:

- **UI Accessibility (MSAA)**: To detect what file the cursor is hovering over in Explorer
- **Shell Windows API**: To interact with Windows Explorer windows
- **GDI**: For rendering the image preview window
- **Registry API**: For managing startup entries

## Configuration

Settings are stored in:

```
%APPDATA%\RustHoverPreview\RustHoverPreview\config.json
```

## Technical Notes

- The preview window is a layered, topmost window that doesn't steal focus
- Image loading is done using the `image` crate with automatic resizing
- The application uses polling (50ms interval) to check cursor position
- Preview appears after 300ms hover delay to avoid flickering

## License

MIT License - See LICENSE file for details

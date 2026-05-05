# Architecture

## Overview

Rust Hover Preview is a Windows 11 tray application that watches File Explorer focus and hover state, then renders a non-activating preview window near the cursor or focused item. It is split into a tray UI, an Explorer hook, and a preview renderer that communicate through lightweight shared state and message passing.

## Runtime Topology

- Main thread initializes COM, config, DPI awareness, then runs the tray event loop.
- Preview thread owns the layered preview window and media decoding/rendering.
- Explorer hook thread polls Explorer state with UI Automation/MSAA and Shell COM APIs.
- Config watcher thread reloads `config.ini` when it changes on disk.

## Core Modules

- `main.rs`: process startup, COM lifecycle, DPI awareness, thread orchestration.
- `explorer_hook.rs`: resolves hovered/focused Explorer items, handles path normalization, and sends preview messages.
- `preview_window.rs`: layered window rendering, animation streaming for GIF/WebP, and FFmpeg-backed video playback.
- `tray.rs`: tray icon and menu, configuration toggles, and exit flow.
- `config.rs`: INI-backed configuration with defaults and input sanitization.
- `startup.rs`: registry integration for the Run-at-startup setting.

## Media Pipeline

- Images (static, GIF, WebP) are decoded in the preview thread, with animated formats streaming frames into a shared queue.
- The preview window uses GDI and `UpdateLayeredWindow` to draw to a topmost, no-activate surface.
- Video previews launch `ffplay` for playback and query `ffprobe` for video geometry.

## Explorer Hook Flow

1. Detect active Explorer window and focused or hovered item.
2. Normalize the resolved path and validate the file extension.
3. Send `Show` or `Hide` messages to the preview thread via channel.
4. Cache folder and Shell view data to reduce repeated COM work.

## DPI Awareness

The app sets per-monitor DPI awareness v2 on startup, with a fallback to per-monitor DPI awareness if v2 is unavailable. This avoids scaling artifacts on layered windows when Windows UI scaling is above 100%.

## Configuration

Configuration is stored at `%APPDATA%\rust-hover-preview\config.ini`. Changes are detected by the config watcher thread and applied without restarting the app.

## Build And Packaging Notes

- The MSVC toolchain uses `rust-lld` via .cargo/config.toml for faster, more consistent linking.
- Windows resources are set in `build.rs` through `winres`.
- Release installers are produced by `cargo packager` (NSIS and WiX).

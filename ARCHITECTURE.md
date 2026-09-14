# Architecture

## Overview

Rust Hover Preview is a Windows 11 tray application that watches File Explorer focus and hover state, then renders a non-activating preview window near the cursor or focused item. It is split into a tray UI, an Explorer hook, and a preview renderer that communicate through lightweight shared state and message passing.

## Runtime Topology

- Main thread claims the single-instance mutex, initializes COM, config, DPI awareness, then runs the tray event loop.
- Preview thread owns the layered preview window and media decoding/rendering.
- Explorer hook thread polls Explorer state with UI Automation/MSAA and Shell COM APIs, and uses EnumWindows with CabinetWClass/ExplorerWClass class matching to count and classify Explorer browser windows so idle polling never spins up Explorer's shell automation providers.
- Config watcher thread reloads `config.ini` when it changes on disk.

## Single Instance

The app runs as a single instance per user session. `main` claims a session-local named mutex (`Local\rust-hover-preview-single-instance`) as its first step, before DPI awareness, COM initialization, config loading, and any thread is started. A second launch — from a desktop shortcut, the Start menu, the startup entry, or the `.exe` directly — finds the name already taken, closes its own handle, and returns from `main` without creating a tray icon, Explorer hook, or preview window. The guard handle is held for the lifetime of the process, so the name is released when the app exits, including after a crash or a forced kill, and the next launch becomes the primary instance. The `Local\` prefix scopes the guard to the signed-in session, so a second user signed in over Remote Desktop still gets their own instance and tray icon.

## Core Modules

- `main.rs`: single-instance guard, process startup, COM lifecycle, DPI awareness, thread orchestration.
- `single_instance.rs`: named-mutex guard that limits the app to one running instance per user session.
- `explorer_hook.rs`: resolves hovered/focused Explorer items, handles path normalization, and sends preview messages.
- `preview_window.rs`: layered window rendering, scale-aware sizing, animation streaming for GIF/WebP, FFmpeg-backed video playback, monitor-bounded placement with paint-before-show presentation, and WM_POWERBROADCAST handling to reset state on system resume and clean up on suspend.
- `tray.rs`: tray icon and menu, configuration toggles, exit flow, and WM_POWERBROADCAST handling to re-add the icon after DWM/Explorer restart on resume.
- `config.rs`: INI-backed configuration with defaults and input sanitization.
- `startup.rs`: registry integration for the Run-at-startup setting.

## Media Pipeline

- Images (static, GIF, WebP) are decoded in the preview thread, with animated formats streaming frames into a shared queue.
- The preview window uses GDI and `UpdateLayeredWindow` to draw to a topmost, no-activate surface.
- Video previews launch `ffplay` for playback and query `ffprobe` for video geometry.

## Explorer Hook Flow

1. Detect active Explorer window and focused or hovered item via HWND/class matching and UI Automation, with input-grace helpers (should_probe_keyboard_focus, should_probe_hover_resolver, should_probe_stationary_hover) throttling probes to recent user activity and a stationary_hover_probe_done latch capping stationary-hover work to a single probe per parked cursor.
2. Normalize the resolved path and validate the file extension.
3. Send `Show` or `Hide` messages to the preview thread via channel.
4. Cache folder and Shell view data to reduce repeated COM work.
5. Keep keyboard and mouse precedence explicit: an arrow-key preview owns the screen and is never dismissed by the parked pointer, while a mouse preview is still dismissed the moment the cursor touches it. Each keyboard spawn measures the preview's own on-screen box (`preview_screen_rect`) and, when that box covers the cursor, freezes the pointer-driven triggers — `should_probe_keyboard_focus`'s companion `should_probe_preview_hover`, the hover resolver/folder probe, and mouse hover previews — until the cursor is moved more than the pointer tolerance (20 px), so jitter cannot end a keyboard preview. The focused-item baseline (`last_focused_name`) is cleared on mouse movement, and the focus observed after that acts immediately, so the first navigation key switches to the keyboard preview instead of being swallowed as a fresh baseline. When the mouse does take over, the file that was shown stays latched (sticky `SuppressedHover`) until the cursor resolves a different file or the keyboard is used again.

## DPI Awareness

The app sets per-monitor DPI awareness v2 on startup, with a fallback to per-monitor DPI awareness if v2 is unavailable. This avoids scaling artifacts on layered windows when Windows UI scaling is above 100%.

## Preview Window Placement

The preview is bounded to the display nearest the hovered cursor or focused item rather than the whole virtual desktop. Bounds are resolved through `MonitorFromPoint`/`GetMonitorInfoW`, falling back to the virtual screen when the monitor lookup fails, so the window no longer spills onto a neighboring display.

Each update repositions the window before installing the new frame or spinner. Crossing between displays of different scale raises `WM_DPICHANGED`, which resets the preview, so installing the content first would discard it and leave the previous display's image stranded on screen. For the same reason, the layered surface is painted before the window is revealed.

Preview size is derived from the media's native dimensions and the `preview_scale` setting, which is either a percentage of the native size or `fit` for the largest size the display area allows. The requested scale is always capped by the free space on the chosen side or quadrant, so a preview can never be clipped by the display edge — a large image at a small scale shrinks to fit, and a small image at 400% or fit-to-screen is reduced to whatever the display can hold. The same scaling is applied to the decode request, so image, GIF/WebP, and video previews all land at the computed size.

Keyboard previews are placed next to the focused item, which can put them over a parked pointer. That is allowed: while a keyboard preview is up it stays topmost and the pointer is not treated as hovering it, so it cannot be cancelled or made to blink by a cursor that happens to sit underneath. The hook confirms this from the preview's own box (`preview_screen_rect`, taken from the layered window or the `ffplay` window once visible) rather than from a cursor position test, and holds the pointer off until the mouse is moved on purpose.

## Configuration

Configuration is stored at `%APPDATA%\rust-hover-preview\config.ini`. Changes are detected by the config watcher thread and applied without restarting the app.

## Build And Packaging Notes

- The MSVC toolchain uses `rust-lld` via .cargo/config.toml for faster, more consistent linking.
- Windows resources are set in `build.rs` through `winres`.
- Release installers are produced by `cargo packager` (NSIS `.exe` setup).

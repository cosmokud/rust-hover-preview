# Architecture

## Overview

Rust Hover Preview is a Windows 11 tray application that watches File Explorer focus and hover state, then renders a non-activating preview window near the cursor or focused item. It is split into a tray UI, an Explorer hook, and a preview renderer that communicate through lightweight shared state and message passing.

## Runtime Topology

- Main thread claims the single-instance mutex, initializes COM, config, DPI awareness, then runs the tray event loop.
- Preview thread owns the layered preview window and media decoding/rendering.
- Explorer hook thread polls Explorer state with UI Automation/MSAA and Shell COM APIs, and uses EnumWindows with CabinetWClass/ExplorerWClass class matching to count and classify Explorer browser windows so idle polling never spins up Explorer's shell automation providers.
- Wheel watcher thread installs a system-wide low-level mouse hook (`WH_MOUSE_LL`) and pumps the messages it needs, publishing a wheel-tick counter that the Explorer hook consumes so wheel scrolling refreshes the hovered item.
- Config watcher thread reloads `config.ini` when it changes on disk.

## Single Instance

The app runs as a single instance per user session. `main` claims a session-local named mutex (`Local\rust-hover-preview-single-instance`) as its first step, before DPI awareness, COM initialization, config loading, and any thread is started. A second launch — from a desktop shortcut, the Start menu, the startup entry, or the `.exe` directly — finds the name already taken, closes its own handle, and returns from `main` without creating a tray icon, Explorer hook, or preview window. The guard handle is held for the lifetime of the process, so the name is released when the app exits, including after a crash or a forced kill, and the next launch becomes the primary instance. The `Local\` prefix scopes the guard to the signed-in session, so a second user signed in over Remote Desktop still gets their own instance and tray icon.

## Core Modules

- `main.rs`: single-instance guard, process startup, COM lifecycle, DPI awareness, thread orchestration.
- `single_instance.rs`: named-mutex guard that limits the app to one running instance per user session.
- `explorer_hook.rs`: resolves hovered/focused Explorer items, handles path normalization, and sends preview messages.
- `wheel_input.rs`: system-wide mouse-wheel hook thread that publishes a tick counter, so a scroll that moves the list under a parked pointer is visible to the polling loop.
- `preview_window.rs`: layered window rendering, scale-aware sizing, animation streaming for GIF/WebP, FFmpeg-backed video playback with verified-PID process supervision (see Video Process Lifecycle), monitor-bounded placement with paint-before-show presentation, and WM_POWERBROADCAST handling to reset state on system resume and clean up on suspend.
- `pdf_preview.rs`: first-page PDF rendering through the PDF engine that ships with Windows (`Windows.Data.Pdf`), with the page size cached per file and the file's own header confirmed before a path reaches the OS renderer.
- `tray.rs`: tray icon and menu, configuration toggles, exit flow, and WM_POWERBROADCAST handling to re-add the icon after DWM/Explorer restart on resume.
- `config.rs`: INI-backed configuration with defaults and input sanitization.
- `startup.rs`: registry integration for the Run-at-startup setting.

## Media Pipeline

- Images (static, GIF, WebP) are decoded in the preview thread, with animated formats streaming frames into a shared queue.
- The preview window uses GDI and `UpdateLayeredWindow` to draw to a topmost, no-activate surface.
- Video previews launch `ffplay` for playback and query `ffprobe` for video geometry; the player's lifetime is supervised as described in Video Process Lifecycle.
- PDF previews render page 1 through `Windows.Data.Pdf`; see PDF Previews for the split between measuring a page and rendering it.

## PDF Previews

A PDF preview shows page 1, rendered by `Windows.Data.Pdf` — the PDF engine that ships with Windows — so the app bundles no renderer and the user installs nothing. Two threads do this work, and both initialize a multithreaded apartment (`CoInitializeEx(COINIT_MULTITHREADED)`) before their first call, which is what WinRT requires of the calling thread.

The preview thread measures the page before anything is decoded, because the preview is positioned and sized first: the page size the engine reports (DIPs, 96 per inch) is what the layout is computed from, the same way video geometry is resolved up front. That probe is cached per path, failures included, so a file is parsed once and an unreadable file is not re-parsed on every hover. Before a path is handed to the renderer it has to carry a `%PDF-` header in its first kilobyte, which is what keeps a mislabeled file out of the OS parser.

The decode worker then renders page 1 into exactly the pixel size the layout asked for, after fitting the page's own aspect ratio into that box, so a page that is not the shape the layout assumed is letterboxed instead of stretched. Rendering at the target size is also what keeps a scaled-up preview sharp: `100%` is a 96 DPI page, while 200% or fit-to-screen is rendered larger rather than enlarged afterwards. The renderer is asked for its BMP encoder, which turns the decode on this side into a header parse instead of a PNG inflate, and the page is painted on an opaque white background so its text stays readable over any `transparent_background` mode. From there the result is an ordinary static frame: the existing spinner covers the wait, and a render that finishes after the cursor has moved on is dropped by the same generation check as every other load.

## Video Process Lifecycle

`ffplay` is the only long-lived external process the app spawns. It is started per hover, its own window is used as the preview surface while the layered window stays hidden, and it is stopped when the hover ends — switching to another video, switching to an image, leaving the file, disabling previews, a display change, or system suspend.

Stopping is kill-only: the app requests termination and never waits for exit, because a process stuck in kernel I/O cannot be terminated by anyone in user mode, and an unbounded `wait` on the Explorer hook thread would freeze hover tracking for good. The PID of the last spawned player is therefore kept until the process is confirmed gone rather than cleared on the spot, and two checks recover a player that survived its stop: the Explorer hook re-checks the recorded PID once a second while no file is hovered and no keyboard preview is up, and the preview thread re-checks it before spawning a new `ffplay`, so a new preview cannot stack a second player next to a stalled one. Both checks terminate only when the recorded PID still belongs to `ffplay.exe`, read from the same process handle that is terminated, so a recycled PID can never hit an unrelated process; the record is cleared only after the process is gone, through a compare-exchange so a freshly spawned player's PID cannot be wiped by a stale check.

## Explorer Hook Flow

1. Detect active Explorer window and focused or hovered item via HWND/class matching and UI Automation, with input-grace helpers (should_probe_keyboard_focus, should_probe_hover_resolver, should_probe_stationary_hover) throttling probes to recent user activity and a stationary_hover_probe_done latch capping stationary-hover work to a single probe per parked cursor.
2. Treat a wheel tick as input when the wheel is driving Explorer (the pointer is over it, or over a preview window that covers the pointer while Explorer still receives the wheel). The wheel moves the list under a stationary cursor, which changes the hovered item without any mouse movement: the stability window restarts while the wheel turns and the latch is reopened, so when the list settles the item that landed under the cursor is resolved and previewed (or the preview is dropped when that item is not media). A `ScrollSettleProbe` only acts on an item that survives two consecutive probes, because Explorer can still be animating the scroll. A scroll also ends keyboard ownership: the keyboard preview is closed, the pointer freeze is released, and the recent-keyboard-input window ends, so the mouse preview takes over the item under the cursor.
3. Treat clicks, Enter and the navigation keys as user navigation. A folder probe runs at the active cadence for a moment after such a press, so a folder that press opens is recognized before the user's next key press rather than up to a full idle interval later. Once the new view has settled the post-folder-change suspension lifts on its own, so the item that landed under a parked cursor previews without a mouse move, and the first navigation key press after the change previews the item it selects — the gate lifts on a fresh press transition (never on key-down state left over from the navigation that opened the folder), so that press is no longer swallowed as a focus baseline. A folder change that no input precedes (a programmatic renavigation, a network refresh) still waits for user input, and the auto-focused first item still never previews without a key press.
4. Normalize the resolved path and validate the file extension. The `.ts` and `.mts` extensions are shared with TypeScript sources, so those two are additionally confirmed by a short MPEG-TS sync-byte probe (`src/video_formats.rs`) before a video preview is started.
5. Send `Show` or `Hide` messages to the preview thread via channel.
6. Cache folder and Shell view data to reduce repeated COM work.
7. Keep keyboard and mouse precedence explicit: an arrow-key preview owns the screen and is never dismissed by the parked pointer (a wheel scroll does end it, handing the screen back to the mouse), while a mouse preview is still dismissed the moment the cursor touches it. Each keyboard spawn measures the preview's own on-screen box (`preview_screen_rect`) and, when that box covers the cursor, freezes the pointer-driven triggers — `should_probe_keyboard_focus`'s companion `should_probe_preview_hover`, the hover resolver/folder probe, and mouse hover previews — until the cursor is moved more than the pointer tolerance (20 px), so jitter cannot end a keyboard preview. The focused-item baseline (`last_focused_name`) is cleared on mouse movement, and the focus observed after that acts immediately, so the first navigation key switches to the keyboard preview instead of being swallowed as a fresh baseline. When the mouse does take over, the file that was shown stays latched (sticky `SuppressedHover`) until the cursor resolves a different file or the keyboard is used again.

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

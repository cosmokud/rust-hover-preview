# Changelog

## [0.2.0] - 2026-09-16

### Added

- PDF previews show the first page of a `.pdf` file, rendered by the PDF engine that ships with Windows (`Windows.Data.Pdf`) rather than by a bundled renderer: this adds no crate, ships no native library in the installer, and asks the user to install nothing, which is what made it the cheaper pick over `pdfium` (a new crate plus roughly 10 MB of native library) and over a pure-Rust PDF rasterizer (a new crate and its rendering stack). The page is rendered at exactly the pixel size the layout asks for, so it is drawn once at the size it is shown instead of being enlarged afterwards, and the size the engine reports for the page (DIPs, 96 per inch) is what the preview is placed and sized from, so a page that is not A4 or Letter lands correctly instead of at a guessed shape. The engine is asked to encode BMP so the decode on this side is a header parse instead of a PNG inflate, the page is painted on an opaque white background so its text stays readable under every `transparent_background` mode, and the finished frame then takes the existing static-image path, spinner included.
- PDF previews ignore the `Preview Scaling` setting and always take the largest size the display area allows. A PDF page is a vector and it is rendered at exactly the preview's size rather than enlarged afterwards, so a bigger preview is sharper text instead of a blurrier one — there is nothing to gain from holding a page at 25% or 100%. The decision lives in one place (`effective_preview_scale`) so the layout and the render always agree on the size, it covers `Best Position` and `Follow Cursor` and mouse and keyboard previews alike, and the fitted size is still clamped to the free space on the chosen side, so a page still cannot run off the display edge.
- A page is measured once. The preview is positioned and sized before anything is decoded, so page 1's size is resolved when the hover arrives and cached per path afterwards — failures included, so a file that cannot be read is not re-parsed on every hover — and that cache is bounded the way the video geometry cache is. Both the sizing probe and the render initialize a multithreaded apartment on their thread first, which is what WinRT requires, so the apartment is claimed before either is used rather than left to fail the first call.
- A `.pdf` file is confirmed to carry a `%PDF-` header within its first kilobyte before the path reaches the OS renderer, so a mislabeled file never gets parsed by it, and a PDF that cannot be rendered — password-protected, damaged, or mislabeled — is simply not previewed instead of showing an error or a placeholder.

### Changed

- Documentation is up to date with the release: README.md lists PDF under supported formats and states what is and is not previewed, and ARCHITECTURE.md documents the PDF path, the apartment both rendering threads initialize, the page-size cache, and the header check.
- Bumped version to 0.2.0 in Cargo.toml and Cargo.lock

### Fixed

- PDF previews appear when a PDF is hovered in a normal folder view. The document was loaded through `StorageFile`, which rejects the paths the Explorer hook hands over: canonicalizing a shell path returns the verbatim form (`\\?\G:\...`) and `StorageFile.GetFileFromPathAsync` fails on it with `ERROR_BAD_PATHNAME`, so the page size could not be read, no layout was computed, and the hover showed nothing at all — no preview and no spinner. The same file previewed from a search-results view because that path is taken from the accessibility value and never canonicalized, and images were never affected because `image::open` reads verbatim paths without complaint. The file is now read with `std::fs::read` and handed to the engine as an in-memory stream, so the WinRT boundary accepts every path form the rest of the app produces — verbatim, long and UNC paths included — and the document only has to be in memory while it is parsed. Measured on the same PDFs, opening through the stream costs what the storage-file route did (about 5 ms for a 1.4 MB manual, under 1 ms for a 29 KB invoice), and the explicit `Storage` feature that route needed is gone from the manifest.

## [0.1.14] - 2026-09-15

### Added

- Image previews now cover the rest of the formats the bundled `image` decoder already reads instead of stopping at the eleven hardcoded extensions: APNG (`.apng`), Targa (`.tga`), the portable anymap family (`.pbm`, `.pgm`, `.ppm`, `.pam`, `.pnm`), Radiance HDR (`.hdr`), OpenEXR (`.exr`), QOI (`.qoi`) and farbfeld (`.ff`). They all decode through the same `image` crate that already handled PNG and JPEG, so this adds no dependency, no startup cost and no idle cost — the codecs were already compiled into the binary and only the extension gate kept them out of previews.
- APNG previews animate like GIFs when the file has more than one frame and fall back to the static image path when it does not. An animated PNG is recognized from the file's content rather than its extension: a `.png` file's chunk list is walked for an `acTL` chunk ahead of its first image data — a few dozen header bytes, no decoding — so a `.png` animation takes the animation path even though its extension is not `.apng`, while every other `.png` keeps the static path it had before. A single-frame animation still shows statically, and a malformed or truncated chunk list is treated as static.
- Video previews now cover the formats FFmpeg natively demuxes instead of the eight hardcoded extensions: the whole MPEG-TS family, MPEG-PS and elementary streams, Windows and TiVo recordings, raw codec streams, ISO base media variants, RealMedia, Flash, Ogg, MXF/GXF, DV, NUT, NSV, IVF, Y4M, MJPEG, and the game and camcorder containers FFmpeg reads. `.ts` and `.mts` are shared with TypeScript sources, so those two are no longer decided by name alone: the file is probed with a 2 KB read that requires an MPEG-TS sync byte (`0x47`) at a 188-, 192- or 204-byte packet stride and at the alignment the file starts at, so a TypeScript file no longer starts `ffplay`.
- Added a `Preview Scaling` tray submenu with `Fit to Screen`, `400%`, `300%`, `200%`, `150%`, `100% (Default)`, `50%`, and `25%` options, persisted through the new `preview_scale` config key. That key also accepts hand-edited values: `fit` (or `fit to screen`) plus any percentage from 1 to 1000 written as `200`, `200%`, or `75`, while `0` resets to the 100% default.
- Added single-instance enforcement: launching the app while it is already running — desktop shortcut, Start menu, startup entry, or the `.exe` directly — detects the running copy through a session-local named mutex and exits immediately instead of adding a second tray icon with its own Explorer hook and preview window. The guard is claimed before DPI setup, COM initialization, config loading, and every background thread, so a duplicate launch does no work at all, and the mutex is released when the process ends — including after a crash or a forced kill — so the next launch becomes the primary instance again.
- Added `src/wheel_input.rs`: a system-wide low-level mouse hook (`WH_MOUSE_LL`) installed on its own thread with a message pump, publishing a monotonic wheel-tick counter that the Explorer hook consumes, giving the polling loop a signal for wheel input without touching Explorer or its accessibility providers.
- Added system sleep/resume resilience: the preview window handles `WM_POWERBROADCAST` to reset the layered window's composition surface (which is destroyed when DWM restarts) and re-assert topmost/layered styles so previews keep working after wake, the tray window re-adds its icon after a DWM/Explorer restart, and video playback and background decoding are cleaned up on suspend/standby to avoid resource leaks.

### Changed

- Preview sizing is derived from the requested scale instead of always rendering at the media's native resolution, and every requested scale is capped by the space available beside the cursor or focused item — in both `Follow Cursor` and `Best Position` modes, and for mouse and keyboard previews — so a scaled-up preview is reduced to fit rather than clipped by the display edge. `Fit to Screen` scales the preview as large as the display area allows, including enlarging images smaller than the screen, and enlarged GIF and animated WebP frames use a smooth filter instead of nearest-neighbor.
- Preview layout is bounded to the display nearest the hovered cursor or focused item instead of the whole virtual desktop, so the preview no longer clips onto a neighboring monitor when multiple displays are attached or their configuration changes.
- The layered preview window keeps one memory DC and one DIB section for its lifetime instead of building and tearing both down on every repaint, and frames are composed straight into that surface. An animated preview used to allocate, fill and free a full `width * height * 4` buffer plus four GDI objects per frame, and the composed frame existed twice — the per-frame buffer and the copy that moved it into the DIB are both gone, and the composition loop walks the frame row by row so each pixel's position is a counter instead of a division. The surface is rebuilt only when the preview size changes, which is exactly the case the old code had no choice but to build a new one for, and every byte written is the byte the previous path produced, including short-frame zero padding and the transparent, black, white and checkerboard background formulas.
- Video playback no longer re-scans the desktop for the `ffplay` window once it is on screen. The style monitor and the preview loop used to run a full top-level window enumeration every 5-100 ms and every 200 ms respectively for the whole of a playback; they now reuse the window found by the first scan while it still exists and still belongs to the recorded player process, and fall back to a full enumeration exactly when that handle stops being a window or no longer belongs to the player. The styles applied to the window and the PID-verified stop path are unchanged.
- Explorer detection switched from ShellWindows COM enumeration to EnumWindows with class-based matching (`CabinetWClass`/`ExplorerWClass`), so idle polling no longer triggers Explorer-side COM provider allocations, and the cursor-over-Explorer check is a pure HWND/class walk now that ShellWindows is never called during cursor polling.
- Hover and keyboard focus probing is input-aware: input-grace helpers (`recent_elapsed_within`, `should_probe_keyboard_focus`, `should_probe_hover_resolver`, `should_probe_stationary_hover`) run probes only when input is recent or a preview is already active, and a `stationary_hover_probe_done` latch caps stationary-hover work to one probe per parked cursor. The cursor-over-preview check is skipped entirely while a keyboard preview is on screen or the pointer is frozen, instead of running on every poll.
- Wheel scrolling refreshes the preview without moving the mouse: a wheel tick counts as user input, the hover stability window restarts while the wheel turns, and once the list stops moving the item that landed under the parked cursor is resolved and previewed — or the preview is dropped when that item is not media, exactly like a mouse move off a file. A scroll-driven probe only acts on an item that survives two consecutive probes, a scroll that leaves the same file under the cursor keeps the current preview instead of restarting it, and a tick is only counted while the wheel is driving Explorer (the pointer is over it, or over a preview window that covers the pointer while Explorer still receives the wheel), so scrolling a menu, a browser, or the desktop changes nothing.
- Keyboard previews take priority over the pointer: a preview opened with the arrow keys stays topmost over a parked cursor, and the pointer-driven triggers — hover resolver, folder probe, cursor-over-preview check and mouse hover previews — stay frozen until the cursor is moved more than 20 px on purpose or the wheel is turned. When the wheel or a deliberate mouse move takes over, the keyboard preview is closed, the pointer freeze is released, and the screen is handed back to the mouse, so the item under the cursor is previewed instead of the keyboard preview staying frozen while the list scrolls underneath it.
- The Explorer hook no longer allocates on every poll tick: the per-tick configuration snapshot no longer clones the off-trigger key string, the off-trigger virtual key is resolved once when the configuration changes instead of being trimmed, lowercased and parsed every tick, the Explorer class check compares `CabinetWClass`/`ExplorerWClass` in place instead of lowercasing every top-level window's class name on every enumeration, and the 250 ms Explorer window and folder snapshot hands back a reference-counted list instead of cloning the vector and every folder path. When the snapshot is rebuilt did not change: the same 250 ms boundary and the same cache-clearing events.
- The video extension list moved into a single `src/video_formats.rs` module exposing `is_video_file`, replacing the identical eight-extension constant and extension check that were duplicated in `explorer_hook.rs` and `preview_window.rs`.
- Documentation is up to date with the release: README.md covers the expanded image and video format support, preview scaling, display-aware placement, flicker-free presentation, wheel-scroll refresh, keyboard preview priority and single-instance behavior, ARCHITECTURE.md documents single instance, preview placement and the video process lifecycle, and the NSIS installer asset name in README.md was corrected to `rust-hover-preview_<version>_x64-setup.exe` to match the files published with each release.
- Bumped version to 0.1.14 in Cargo.toml and Cargo.lock

### Fixed

- Animated WebP, APNG and GIF previews no longer stop part-way and start over before reaching the end. The streaming decoder used to stop for good once the frames it had decoded passed 256 MB, and it also stopped after 300 frames; the player then treated the animation as finished and looped back to the first frame. At 1024x1664 and 20 fps a decoded frame is about 6.5 MB, so that limit landed under two seconds into a seven-second animation — the longer or larger the file, the earlier it restarted. Playback now runs through a sliding window instead of a hard cap: the decoder waits while it is a few frames ahead of what is on screen, the player releases frames it has already shown once 16 MB of them have piled up behind the playhead, and when the file ends the decoder starts it over so the animation loops. Memory is bounded by the window rather than by the length of the animation, the frame cap is gone so the whole file plays however many frames it holds, and an animation small enough to fit is still decoded once and loops from memory exactly as before.
- Video previews no longer risk leaving a stray `ffplay` process, and its frozen window, behind when a stop does not take effect. Stopping playback is now kill-only — the app requests termination and never blocks in `wait`, which could freeze the Explorer hook thread indefinitely on a process stuck in kernel I/O — and the recorded PID is kept until the process is confirmed gone instead of being cleared on the spot. Two checks make that recoverable: the Explorer hook re-checks the recorded PID once a second while nothing is hovered, and the preview thread re-checks it before spawning a new `ffplay`, so a new preview can no longer stack a second player next to one that survived its stop. Both terminate only when the PID still belongs to `ffplay.exe`, read from the same process handle that is terminated so a recycled PID can never hit an unrelated process, and the record is cleared through a compare-exchange only once the process is gone, so a fresh spawn cannot have its PID wiped by a stale check. A player that still refuses to die is retried on the next sweep instead of being forgotten.
- A single navigation key press now opens the keyboard preview right after a folder change. The post-folder-change gate compared the focused item against a baseline recorded only once a key was already pressed, so that first press was swallowed as the baseline and the preview only appeared on the second one. The gate now also lifts on a navigation key press seen after the change — the press transition, not the key-down state, so key state left over from the navigation that opened the folder cannot lift it — and the item that press selects previews immediately. After a mouse-hover preview the first key press likewise switches straight to the keyboard preview instead of being consumed as a fresh focus baseline.
- Entering a folder no longer waits for a mouse move before previewing the item under the cursor. Clicks (left/right/middle), Enter and the history keys are now tracked as deliberate input, so the folder change they cause is recognized as user navigation instead of a change the gate has to wait out: the folder probe runs at the active cadence for a moment after such a press so the change is seen before the user's next key press, and the post-change suspension lifts on its own once the new view has settled. A folder change that no input precedes (a programmatic renavigation, a network refresh) still waits for the user, and the auto-focused first item still never previews without a key press.
- Wheel-scrolling Explorer no longer leaves the preview stuck on the file that was under the cursor before the scroll, and a wheel scroll now takes over from a keyboard preview instead of leaving it frozen while the list scrolls underneath it. The file the keyboard showed is not latched, so the mouse may preview it again, and the recent-keyboard-input window ends with the handoff, so the keyboard preview only comes back when a navigation key is pressed again.
- After the mouse takes over from a keyboard preview, nothing is previewed until the keyboard is used again or the mouse hovers a different file, so moving the mouse off a keyboard preview no longer re-previews the file that was already shown.
- Preview placement and presentation interact correctly with display changes: the preview window is repositioned before the new frame is installed, because crossing between displays of different scale raises `WM_DPICHANGED`, which resets the preview — installing the frame first discarded it and left the previous display's image stranded on screen — and both the window and the loading spinner are painted before being revealed, so a new hover can no longer flash the previously previewed image for a frame at the new position and size.

## [0.1.14-rc.10] - 2026-09-15

### Added

- Animated PNGs are now recognized inside `.png` files instead of only in `.apng` files. An APNG is an ordinary PNG with an `acTL` chunk ahead of its first image data, so a `.png` file's chunk list is walked — a few dozen header bytes, no decoding — and only a file carrying that chunk takes the animation path; every other `.png` keeps the static path it had before. A `.png` animation with a single frame still falls back to a static preview, a malformed or truncated chunk list is treated as static, and `.apng` files behave exactly as they did.

### Fixed

- Animated WebP, APNG and GIF previews no longer stop part-way and start over before reaching the end. The streaming decoder used to stop for good once the frames it had decoded passed 256 MB, and it also stopped after 300 frames; the player then treated the animation as finished and looped back to the first frame. At 1024x1664 and 20 fps a decoded frame is about 6.5 MB, so that limit landed under two seconds into a seven-second animation — the longer or larger the file, the earlier it restarted. Playback now runs through a sliding window instead of a hard cap: the decoder waits while it is a few frames ahead of what is on screen, the player releases frames it has already shown once 16 MB of them have piled up behind the playhead, and when the file ends the decoder starts it over so the animation loops. Memory is bounded by the window rather than by the length of the animation, the frame cap is gone so the whole file plays however many frames it holds, and an animation small enough to fit is still decoded once and loops from memory exactly as before.

### Changed

- The layered preview window now keeps one memory DC and one DIB section for its lifetime instead of building and tearing both down on every repaint. An animated preview used to allocate, fill and free a full `width * height * 4` buffer plus four GDI objects per frame, which at the animation ceiling is the largest per-frame cost of GIF, APNG and animated WebP playback; the surface is now rebuilt only when the preview size changes, which is exactly the case the old code had no choice but to build a new one for. The window is still painted before it is revealed, so a hover still cannot flash the previous image at the new position.
- Frames are composed straight into that surface. The composed frame no longer exists twice — the per-frame buffer and the copy that moved it into the DIB are both gone — and the composition loop walks the frame row by row, so each pixel's position is a counter instead of a division. Every byte written is the byte the previous path produced, including the zero padding for a short source frame and the transparent, black, white and checkerboard background formulas.
- Video playback no longer re-scans the desktop for the `ffplay` window once it is on screen. The style monitor and the preview loop used to run a full top-level window enumeration every 5-100 ms and every 200 ms respectively for the whole of a playback; they now reuse the window found by the first scan while it still exists and still belongs to the recorded player process, and fall back to a full enumeration exactly when that handle stops being a window (a recreated `ffplay` window) or no longer belongs to the player. The styles applied to the window are unchanged, and the PID-verified stop path was not touched.
- The Explorer hook no longer allocates on every poll tick. The per-tick configuration snapshot no longer clones the off-trigger key string, the off-trigger virtual key is resolved once when the configuration changes instead of being trimmed, lowercased and parsed into a fresh string every tick, and the Explorer window class check compares `CabinetWClass`/`ExplorerWClass` in place instead of lowercasing every top-level window's class name into a new string on every enumeration.
- The Explorer window and folder snapshot is shared rather than copied. Cache hits on the 250 ms snapshot now hand back a reference-counted list instead of cloning the vector and every folder path, which removes the remaining per-call allocations from the hover-resolution and folder-probe paths. When the snapshot is rebuilt did not change: the same 250 ms boundary and the same cache-clearing events.

## [0.1.14-rc.9] - 2026-09-15

### Added

- Image previews now cover the rest of the formats the bundled `image` decoder already reads instead of stopping at the eleven hardcoded extensions: added APNG (`.apng`), Targa (`.tga`), the portable anymap family (`.pbm`, `.pgm`, `.ppm`, `.pam`, `.pnm`), Radiance HDR (`.hdr`), OpenEXR (`.exr`), QOI (`.qoi`) and farbfeld (`.ff`). All of them decode through the same `image` crate that already handles PNG and JPEG, so this adds no dependency, no startup cost and no idle cost — the codecs were already compiled into the binary and only the extension gate kept them out of previews. Measured decode plus resize cost per hover stays in the same band as the existing formats, with EXR (largest files) and farbfeld (16-bit output converted down) at the top of it.
- APNG previews animate like GIFs when the file has more than one frame and fall back to the static image path when it does not. Frames arrive already composited by the decoder — blend and dispose operations are resolved before they reach the preview — so playback reuses the existing animated pipeline unchanged: the first 12 frames or 500 ms are decoded before the window is revealed, the remainder streams in on a background thread during playback, and the 300-frame and 256 MB caps still apply. Frame delays come from the file's exact rational delays and are floored like GIF's so a zero-delay animation cannot spin the render loop.
- Video previews now cover the formats FFmpeg natively demuxes instead of the eight hardcoded extensions: the list grew from `mp4`/`webm`/`mkv`/`avi`/`mov`/`wmv`/`flv`/`m4v` to the whole MPEG-TS family (`.ts`, `.m2ts`, `.mts`, `.m2t`, `.tr`, `.tp`, `.tod`), MPEG-PS and elementary streams (`.mpg`, `.mpeg`, `.mpe`, `.m2p`, `.vob`, `.vro`, `.m1v`, `.m2v`, `.mpv`), Windows and TiVo recordings (`.wtv`, `.dvr-ms`, `.ty`, `.ty+`), raw codec streams (`.h264`, `.h26l`, `.264`, `.avc`, `.h265`, `.hevc`, `.265`, `.h266`, `.vvc`, `.266`, `.vc1`, `.rcv`, `.av1`, `.obu`, `.evc`, `.apv`, `.avs`, `.avs2`, `.avs3`, `.cavs`, `.drc`, `.vc2`), ISO base media variants (`.3gp`, `.3g2`, `.mj2`, `.psp`, `.ismv`, `.f4v`, `.qt`, `.divx`), RealMedia (`.rm`, `.rmvb`), Flash (`.swf`), Ogg (`.ogv`, `.ogm`), MXF/GXF (`.mxf`, `.gxf`), DV (`.dv`, `.dif`), NUT, NSV, IVF, Y4M, MJPEG (`.mjpg`, `.mjpeg`), and the game and camcorder containers FFmpeg reads (`.bik`, `.bk2`, `.smk`, `.roq`, `.mve`, `.cpk`, `.thp`, `.usm`, `.moflex`, `.xmv`, `.mvi`, `.mxg`, `.rsd`, `.str`, `.cin`, `.c93`, `.cdxl`, `.xl`, `.flm`, `.yop`, `.imx`, `.dav`, `.viv`, `.ivr`, `.vw`, `.cdg`, `.pmp`, `.kux`, `.ifv`).
- `.ts` and `.mts` are shared with TypeScript sources, so those two extensions are no longer decided by name alone: the file is probed with a 2 KB read that requires an MPEG-TS sync byte (`0x47`) at a 188-, 192- or 204-byte packet stride and at the alignment the file starts at. Standard transport streams, BDAV/AVCHD `.m2ts` layouts (whose packets begin with a 4-byte timestamp) and DVB captures with FEC all pass, while a TypeScript file no longer starts `ffplay`. The probe only runs for the two ambiguous extensions, so every other hover stays a plain extension lookup.

### Changed

- The video extension list moved into a single `src/video_formats.rs` module exposing `is_video_file`, replacing the identical eight-extension constant and extension check that were duplicated in `explorer_hook.rs` and `preview_window.rs`.

### Fixed

- Video previews no longer risk leaving a stray `ffplay` process, and its frozen window, behind when a stop does not take effect. Stopping playback is now kill-only — the app requests termination and never blocks in `wait`, which could freeze the Explorer hook thread indefinitely on a process stuck in kernel I/O — and the recorded PID is kept until the process is confirmed gone instead of being cleared on the spot. Two checks make that recoverable: the Explorer hook re-checks the recorded PID once a second while nothing is hovered, and the preview thread re-checks it before spawning a new `ffplay`, so a new preview can no longer stack a second player next to one that survived its stop. Both terminate only when the PID still belongs to `ffplay.exe`, read from the same process handle that is terminated so a recycled PID can never hit an unrelated process, and the record is cleared through a compare-exchange only once the process is gone, so a fresh spawn cannot have its PID wiped by a stale check. A player that still refuses to die is retried on the next sweep instead of being forgotten.
- Entering a folder no longer waits for a mouse move before previewing the item under the cursor. Clicks (left/right/middle), Enter and the history keys are now tracked as deliberate input, so the folder change they cause is recognized as user navigation instead of a change the gate has to wait out: the folder probe runs at the active cadence for a moment after such a press so the change is seen before the user's next key press, and the post-change suspension lifts on its own once the new view has settled. A folder change that no input precedes (a programmatic renavigation, a network refresh) still waits for the user, and the auto-focused first item still never previews without a key press.
- A single navigation key press now opens the keyboard preview right after a folder change. The post-folder-change gate compared the focused item against a baseline recorded only once a key was already pressed, so that first press was swallowed as the baseline and the preview only appeared on the second one. The gate now also lifts on a navigation key press seen after the change — the press transition, not the key-down state, so key state left over from the navigation that opened the folder cannot lift it — and the item that press selects previews immediately.

## [0.1.14-rc.8] - 2026-09-15

### Added

- Added `src/wheel_input.rs`: a system-wide low-level mouse hook (`WH_MOUSE_LL`) installed on its own thread with a message pump. It publishes a monotonic wheel-tick counter that the Explorer hook consumes, giving the polling loop a signal for wheel input without touching Explorer or its accessibility providers.

### Fixed

- Wheel-scrolling Explorer no longer leaves the preview stuck on the file that was under the cursor before the scroll. A wheel tick counts as user input, the hover stability window restarts while the wheel is turning, and once the list stops moving the item that landed under the parked cursor is resolved and previewed — or the preview is dropped when that item is not a media file, exactly like a mouse move off a file. A wheel scroll also releases the post-folder-change suspension, like a mouse move.
- A wheel scroll now takes over from a keyboard preview. Scrolling while an arrow-key preview is on screen closes it, releases the pointer freeze, and hands the screen back to the mouse, so the item that lands under the cursor is previewed instead of the keyboard preview staying frozen while the list scrolls underneath it. The file the keyboard showed is not latched, so the mouse may preview it again, and the recent-keyboard-input window ends with the handoff, so the keyboard preview only comes back when a navigation key is pressed again.
- A scroll-driven probe only acts on an item that survives two consecutive probes, so a preview can no longer be shown for a row that is still animating away under Explorer's smooth scrolling, and a scroll that leaves the same file under the cursor keeps the current preview instead of restarting it (no image re-decode, no video or GIF restart). A tick is only counted while the wheel is driving Explorer — the pointer is over it, or over a preview window that covers the pointer while Explorer still receives the wheel — so scrolling a menu, a browser, or the desktop changes nothing.

## [0.1.14-rc.7] - 2026-09-14

### Added

- Added single-instance enforcement: launching the app while it is already running (desktop shortcut, Start menu, startup entry, or the `.exe` directly) now detects the running copy through a session-local named mutex (`Local\rust-hover-preview-single-instance`) and exits immediately instead of adding a second tray icon with its own Explorer hook and preview window. The instance already running is left untouched and keeps serving previews.
- The guard is claimed before DPI setup, COM initialization, config loading, and every background thread, so a duplicate launch does no work at all. The mutex is released when the process ends — including after a crash or a forced kill — so the next launch becomes the primary instance again.

### Fixed

- Keyboard previews now take priority over the pointer. A preview opened with the arrow keys is no longer dismissed when the mouse cursor happens to sit where the preview appears: it stays on top of the parked cursor instead of blinking away and reappearing on every navigation key.
- Every keyboard preview now measures its own on-screen box once it appears, and if that box covers the mouse cursor the pointer-driven triggers are frozen until the cursor is moved on purpose: the hover resolver, the folder probe, the cursor-over-preview check, and mouse hover previews all stay off while the keyboard owns the screen. Cursor jitter below 20 px no longer counts as movement, so a parked or lightly nudged mouse cannot cancel a keyboard preview or its temporary pause.
- The cursor-over-preview check no longer runs on every poll. It is skipped while a keyboard preview is on screen or the pointer is frozen, and outside those states it is a single shared check instead of two per tick that returns immediately without touching the cursor when no preview is on screen.
- After the mouse takes over from a keyboard preview, nothing is previewed until the keyboard is used again or the mouse hovers a different file, so moving the mouse off a keyboard preview no longer re-previews the file that was already shown.
- The first keyboard navigation key now switches straight to the keyboard preview. After a mouse-hover preview, the first key press used to be swallowed as a fresh focus baseline (`last_focused_name` is reset on every mouse move), so the preview only switched to the keyboard one on the second press.

## [0.1.14-rc.6] - 2026-09-14

### Added

- Added a `Preview Scaling` tray submenu with `Fit to Screen`, `400%`, `300%`, `200%`, `150%`, `100% (Default)`, `50%`, and `25%` options, persisted through the new `preview_scale` config key.
- `preview_scale` also accepts hand-edited values in `config.ini`: `fit` (or `fit to screen`) plus any percentage from 1 to 1000 written as `200`, `200%`, or `75`. `0` resets to the 100% default.

### Changed

- Preview sizing is now derived from the requested scale instead of always rendering at the media's native resolution.
- Every requested scale is capped by the space available beside the cursor or focused item, so a preview that would cross the screen edge is scaled down to fit instead of being clipped — in both `Follow Cursor` and `Best Position` modes, and for mouse and keyboard previews.
- `Fit to Screen` scales the preview as large as the display area allows, including enlarging images smaller than the screen.
- Enlarged GIF and animated WebP frames now use a smooth filter instead of nearest-neighbor, avoiding blocky upscaled previews.
- Bumped version to 0.1.14-rc.6 in Cargo.toml and Cargo.lock

## [0.1.14-rc.5] - 2026-09-14

### Changed

- Documentation: documented display-aware preview placement and flicker-free presentation in README.md, and added a preview placement section to ARCHITECTURE.md.
- Corrected the NSIS installer asset name in README.md to `rust-hover-preview_<version>_x64-setup.exe`, matching the files published with each release.

### Fixed

- Preview layout is now bounded to the display nearest the hovered cursor or focused item instead of the whole virtual desktop, so the preview no longer clips onto a neighboring monitor when multiple displays are attached.
- The preview window is now painted before it is revealed, and the loading spinner likewise, so a new hover can no longer flash the previously previewed image for a frame at the new position and size.
- The preview window is now repositioned before the new frame is installed. Crossing between displays of different scale raises `WM_DPICHANGED`, which resets the preview, so installing the frame first discarded it and left the previous display's image stranded on screen.

## [0.1.14-rc.3] - 2026-07-03

### Added

- Added system sleep/resume resilience: the preview window now handles `WM_POWERBROADCAST` events to detect system resume from sleep, resetting the layered window's composition surface (which is destroyed when DWM restarts) and re-asserting topmost/layered styles so previews continue working after wake.
- Added `WM_POWERBROADCAST` handling in the tray window to re-add the tray icon on resume, preventing it from disappearing after DWM/Explorer restart during sleep cycles.
- Cleaned up video playback and background decoding on system suspend/standby (`PBT_APMSUSPEND`/`PBT_APMSTANDBY`) to avoid resource leaks.

### Changed

- Updated installation instructions to present two clear asset options per release (NSIS installer `RustHoverPreview-<version>-setup.exe` and standalone portable binary `rust-hover-preview.exe`) with an upgrade note for existing users.
- Updated config example in README to document the `confirm_file_type` and `webp_playback_fps` settings with explanations.
- Bumped version to 0.1.14-rc.3 in Cargo.toml and Cargo.lock

## [0.1.14-rc.2] - 2026-07-01

### Changed

- Improved Explorer detection by switching from ShellWindows COM enumeration to EnumWindows with class-based matching (CabinetWClass/ExplorerWClass) so idle polling no longer triggers Explorer-side COM provider allocations.
- Simplified the cursor-over-Explorer check to a pure HWND/class walk now that ShellWindows is no longer called during cursor polling.
- Made hover and keyboard focus probing input-aware: introduced input-grace constants and helpers (recent_elapsed_within, should_probe_keyboard_focus, should_probe_hover_resolver, should_probe_stationary_hover) that track the last user/keyboard input times and only run probes when input is recent or the preview is already active.
- Added a stationary_hover_probe_done flag so the stationary-hover probe runs at most once per stable hover and is reset on movement, folder change, or suppression, avoiding repeated accessibility work for the same parked cursor.
- Bumped version to 0.1.14-rc.2 in Cargo.toml and Cargo.lock

## [0.1.14-rc.1] - 2026-06-24

### Changed

- Handled multi-monitor virtual screen and display changes

## [0.1.13] - 2026-06-03

### Changed

- Release 0.1.13: bump version and update changelog

## [0.1.13-rc.2] - 2026-06-03

### Added

- Added features section to TODO.md for document support

### Changed

- Bumped version to 0.1.13-rc.2 in Cargo.toml and Cargo.lock
- Increased cache limits for explorer and video geometry to improve performance
- Removed the tag exclusion pattern from the deploy workflow and simplified tag handling

## [0.1.13-rc.1] - 2026-06-03

### Added

- Added per-monitor DPI awareness configuration (with fallback) to prevent scaling artifacts on high-DPI displays

### Changed

- Removed MSI installer references from deployment workflows and scripts
- Switched MSVC linking to use rust-lld via .cargo/config.toml
- Documentation: renamed TODO section to Known Issues
- Documentation: cleared the resolved high-DPI preview issue from Known Issues
- Documentation: added ARCHITECTURE.md with a full system overview

### Fixed

- Fixed hover preview rendering issues on Windows UI scaling above 100%

## [0.1.12] - 2026-04-30

### Added

- Added shell view index sync item limit and cache clearing for Explorer probe stabilization

### Changed

- Refactored media path resolution and removed unused preview target management
- Enhanced hover preview logic to prevent state hanging with large result sets and improve dismissal behavior

### Fixed

- Fixed hover preview state management for improved reliability when navigating large file sets

## [0.1.11] - 2026-04-30

### Added

- Added WebP playback FPS configuration with sanitization logic to prevent invalid values

### Changed

- Refactored file handling logic in Explorer hook; removed unused cache clearing function and optimized item retrieval
- Improved Windows resource file description and search view metadata handling

## [0.1.10] - 2026-04-29

### Changed

- Restored the default NSIS older-version prompt while keeping the installer-side process shutdown, config default completion, and old-folder migration logic.
- Limited automatic startup registration to first install and first portable run, while leaving later launches and updates to the user's saved setting.
- Removed custom installer `taskkill` actions and returned running-app shutdown to cargo-packager's default Windows installer behavior.
- Updated release metadata for v0.1.10.

## [0.1.9] - 2026-04-29

### Added

- Added transparent-background preview modes for PNG/WebP transparency: transparent, black, white, and checkerboard.
- Added configurable same-file rehover delay with a 750 ms default and tray controls.
- Added installer migration cleanup for the old `%LOCALAPPDATA%\Rust Hover Preview` install folder.

### Changed

- Switched animated WebP playback to Google's libwebp via `webp-animation` for smoother startup and playback.
- Moved per-user installer output to `%LOCALAPPDATA%\rust-hover-preview` while keeping the display name `Rust Hover Preview`.
- Switched configuration path resolution to the `directories` crate while keeping config at `%APPDATA%\rust-hover-preview\config.ini`.
- Updated the tray off-trigger label to show the configured key, such as `Enable Off Trigger Key (Alt)`.
- Updated the tray icon tooltip to show the full `Rust Hover Preview` app name.

### Fixed

- Updated NSIS and WiX installers to stop a running `rust-hover-preview.exe` before replacing installed files during upgrades.
- Added installer-side config default completion so missing `config.ini` parameters are added without changing existing user values.
- Hardened preview hide behavior so stuck video previews are hidden and ffplay is stopped more aggressively.
- Added Explorer COM/UIA slow-probe backoff and cache clearing to reduce runaway polling after Explorer gets sluggish.
- Fixed repeated same-file hover retrigger timing after the preview self-dismisses.
- Fixed transparent PNG/WebP rendering artifacts by compositing alpha explicitly for each transparency mode.
- Removed unused Explorer helper functions that produced dead-code warnings.

## [0.1.8] - 2026-04-29

### Added

- Added optional off-trigger key support so previews can be suppressed while a configured modifier key is held.
- Added active Explorer Shell view detection and Shell view media indexing for more reliable file resolution.
- Added media indexing from configured search roots to broaden Explorer hover matching.

### Changed

- Improved Explorer media path normalization, URL handling, and URL decoding.
- Cached Explorer real folder lookups to reduce repeated Shell resolution work.
- Updated Windows installer build orchestration to use the PowerShell helper in CI.

### Fixed

- Removed branch-triggered release deployment so installer publishing only follows the intended release flow.

## [0.1.7] - 2026-04-28

### Added

- Added Windows installer packaging via `cargo packager`, producing release `.exe` and `.msi` assets in CI.
- Added `build-installers.ps1` helper flow for local installer generation with artifact validation.

### Changed

- Updated deploy workflow release packaging to publish `target/packager/*.exe` and `target/packager/*.msi` on `v*.*.*` tags.
- Refactored preview window, media loading, and video-dimension internals for clearer structure and maintainability.
- Improved installer build orchestration with a WiX fallback path for constrained local sessions.

### Fixed

- Added explicit Windows Installer (`msiserver`) availability checks before MSI generation.
- Improved installer-build error handling and output validation to fail fast when packaging artifacts are missing.

## [0.1.6] - 2026-04-06

### Changed

- Optimized animated GIF/WebP streaming by switching to queue-based frame transfer, reducing duplicate frame copying and memory churn.
- Limited preview animation decoding to the first 300 streamed frames and standardized animation timing to a 30 FPS ceiling for more stable CPU usage.
- Added a new tray option, `Confirm File Type` (off by default), to optionally validate image headers for mislabeled file extensions.

### Fixed

- Reduced severe CPU spikes while hovering animated WebP files by tightening streaming overlay redraw conditions and easing decoder contention.
- Improved rapid-hover behavior across multiple animated files by making decoder workers more cooperative during long streams.
- Fixed hover preview triggering for `.jpeg` files by treating JPEG aliases (`.jpg`, `.jpeg`, `.jpe`, `.jfif`) consistently during media resolution.
- With `Confirm File Type` enabled, mislabeled images (for example PNG content with a JPG/JPEG extension) now load with the correct decoder.
- Prevented the previous hovered image from flashing briefly when switching to a new file.

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

# Changelog

## [0.2.3] - 2026-09-17

### Changed

- Explorer hover previews resolve the file under the pointer by the item under it instead of by its name. The item is read from the view's accessibility tree in one batched call — its box, the name the view shows it under and Explorer's own `ItemIndex` — and that position is turned into a path by the Shell window's active `IFolderView2` through `SIGDN_FILESYSPATH`, which for a search result is the real file wherever it lives.
- The pointer no longer keeps an index warm to answer with: the search-root index, the budgeted descent below the root, the view index and the folder index are the keyboard's and the focused item's alone now, and the view kind no longer selects a resolver. The legacy search-view index, its two caches, its background completion and its own search-root resolver are removed.
- Documentation is up to date with the release: ARCHITECTURE.md's Pointer Resolution section describes the identity route, and the search-results known issue is gone from TODO.md.
- Bumped version to 0.2.3 in Cargo.toml and Cargo.lock

### Fixed

- A hover preview in a search results view no longer waits on background walks: a search of several hundred results stays responsive, keeps the preview that is up, and no longer trips the hook's slow-probe pause.
- Two search results that share a file name are told apart by the item under the pointer, so each of them previews its own file instead of the first one resolving for all of them.
- The whole of an item's row belongs to the pointer, including the padding and empty space a view leaves after its text, so a Details or Content row previews from anywhere on it.

## [0.2.2] - 2026-09-16

### Added

- Added a `Toggle Preview Types` submenu to the tray menu, right under `Enable Preview`, with one gate per kind of preview — `Images`, `Videos`, `Text` and `PDF` — all switched on and written to `config.ini` as `image_preview_enabled`, `video_preview_enabled`, `text_preview_enabled` and `pdf_preview_enabled`. The gates decide whether a kind may be previewed at all, so the lists and settings that choose which files are previewed are untouched, and switching a kind off and back on restores what was configured rather than a default.
- Both halves of the app read the gates — the Explorer hook on the tick that resolved a file, and the renderer when it measures the preview it was handed — so a toggle needs no restart and leaves no cache to clear. A preview of a kind that is switched off goes away on the spot, rebuilt from the hover it came from, while a preview of any other kind is left untouched.
- Added `Select All` to the text preview's own context menu, above `Copy`. It takes the whole frame — the range a `Copy` with nothing selected would take — and the menu's `Copy` is the same function the polled `Ctrl+C` runs, so the two cannot disagree about what is selected.
- A text preview is now recognized by one function, `is_text_preview`, which asks the video gate before the text gate, so a `.ts` or `.mts` file that really is an MPEG transport stream is never handed to the text renderer. The places that used to ask the question separately — the scale a preview is laid out with, the size it is measured at and the box its renderer is handed — ask that one function now, so they cannot disagree about a `.ts` file.

### Changed

- Removed `Enable Text Preview` from the tray menu, which `Toggle Preview Types` → `Text` replaces. The setting is the same one — `text_preview_enabled` is still the key, still checked ahead of the extension and name lists — so a `config.ini` written before the change keeps its answer.
- Documentation is up to date with the release: README.md lists the `Toggle Preview Types` submenu and the four keys, and ARCHITECTURE.md's text-preview and Explorer-hook sections describe the gates, the order they are asked in, where a type change is answered and what ends a keyboard preview's ownership of the screen.
- Bumped version to 0.2.2 in Cargo.toml and Cargo.lock

### Fixed

- Two search results that share a file name are told apart by the view, not by the name: the view's own selection and the item whose cell holds the point are asked before any name lookup, and both answers are still checked against the name before being taken.
- A keyboard preview no longer waits on name lookups it does not need. The walk below the search root and the folder and view indexes are the fallback now, so a key press answers immediately instead of finding nothing while a walk is running.
- A keyboard preview in a search result no longer depends on where the item's text sits: the probe walks the item's whole subtree — bounded, and only when the element really is an item — and asks the item's own box in bands as well, which is where each view puts the text.
- A pointer over a Content view row gets that row's file. The row's second line is metadata and reports to the accessibility tree as a name, so a name is only taken when it could be a file this app previews, and otherwise the walk goes on up to the item that holds the real name.
- A keyboard preview of a wide item is placed beside the display's room rather than beside the item: a Content view row is read as that kind of row, `Best Position` measures the display's room from its middle instead of its edges, and `Follow Cursor` grows its quadrant from that same middle. A mouse preview is placed from the cursor already and was never affected.
- A text preview no longer puts the list beside it out of reach. The region that keeps it alive reaches the preview's nearest point and is as tall as the crossing, instead of being the box around both rectangles, which covered the whole column of the list between the pointer and the preview.
- Text previews no longer answer `Ctrl+A`. The key was polled against the window that has focus, so whether Explorer's highlight ever appeared came down to timing; selecting everything is `Select All` in the preview's own menu now, and `VK_A` is no longer read at all.
- Keyboard previews work in a search whose results come from more than one folder. The item under the focus is asked what file it is — its own value first, then the provider at the item's box and along its name column, then the shell data model and the view's own focused item — instead of resolving a name in a folder that does not hold the file.
- A keyboard preview appears again in a view that reports its list as the focused element. The search results view's list is not a file, so the list's own selection names the item in that case and everything downstream is resolved from it.
- Two files that share a name are two files to the keyboard: the item's own box is part of what tells one focus observation from the next, so the preview follows the item the keyboard is on.
- A name index that is due for a rebuild no longer stops answering while it is rebuilt: the folder index and the search-root index are kept and read while their replacement is walked, and a fresh walk still replaces what it was built from.
- What was cached about a view is dropped when the location changes, so a name looked up after another folder is opened in the window — or the same window is searched again — no longer resolves against the previous location's items.
- A result that lives below the folder a search was started in resolves: the name is walked for below that folder as well, under a 40 ms budget on the hook loop, in the pointer's hover-folder step, in its search resolver and in the keyboard's.
- A search view of several hundred results no longer goes previewless. The hook's slow-probe pause stops the asking and nothing else now: the preview that is up stays, the keyboard's turn is left as it is, and the hover starts again from scratch when the pause ends.
- A Shell view is no longer walked to the end while the hook loop waits for it. The walk on the loop stops when its 150 ms budget is out and the view is walked to the end on a thread of its own, which claims a COM apartment first; a complete index is kept longer than a short one (30 s against 5 s).
- A file a keyboard preview handed over to the mouse is no longer left un-previewable: the hand-over latch is the delay a dismissed hover already carries and expires, rather than latching the file for good.
- Keyboard and mouse previews no longer fight over a list the keyboard is walking. A navigation key press puts the keyboard in charge of the screen until a deliberate pointer move, a wheel tick or a reset ends its turn, and the pointer tolerance that ends that turn is in force for the whole of it.
- Text previews no longer depend on how an extension's case was typed: the extension is lowercased once, where it is read, so a `README.MD` renders as the document it describes, an `.NFO` in capitals keeps its ANSI colors and CP437 decoding, and an `.RTF` is reduced to the text it carries like any other.

## [0.2.1] - 2026-09-16

### Added

- Added text previews: the first screenful of a text file, colored by the syntax definition its extension names. Coloring is `syntect` — the TextMate grammars `bat` and `delta` use — plus `two-face` for the definitions it does not bundle, so a preview is colored the way an editor would color it and a file no grammar claims is still previewed as plain text.
- Text previews are sized to their content instead of to the `Preview Scaling` setting: `measure` resolves the box its first screenful wants, bounded by the display it will be shown on, and the font stays the same size across DPI because the scale comes from the monitor under the hover rather than from a constant.
- Added rendered Markdown: headings, emphasis, inline code, links, lists, block quotes, rules, tables and fenced code blocks colored by the language their info string names. `Markdown Preview` in the tray switches a file between that and `Highlighted Source`, and the choice persists as `markdown_mode`.
- Added `Atom One Light` and `One Dark Pro` as the text preview themes, with `Atom One Light` the default and a `Text Preview Theme` tray submenu to switch. Both are the VS Code themes of the same name, converted once into the TextMate `.tmTheme` format `syntect` reads and embedded in the binary, so no theme file ships beside the executable; switching the theme re-renders the preview that is on screen.
- Added the `[text] extensions` configuration key: the full list of extensions previewed as text, written with the built-in list on first run and read back from the file, so an extension can be added or deleted without a rebuild and without a restart. Entries are normalized on read, `.ts` and `.mts` still resolve to a video first when the file really is an MPEG transport stream, and an empty list turns text previews off.
- Added a `Text Preview Font Size` tray submenu — `100%` through `400%`, with `125%` the default — sizing the text a preview is drawn at, for rendered Markdown as much as for source files. It is written to `config.ini` as `text_font_scale` and honored as a plain percentage, so `150`, `150%` or a step the menu does not offer are all used as written (1 to 1000; `0` resets to 125).
- Added an `Enable Text Preview` tray toggle, above `Text Preview Theme` and on by default. It is persisted as `text_preview_enabled` and consulted ahead of the extension list, so switching text previews off leaves the list intact and switching them back on restores exactly what was configured.
- Added `.rtf` previews by reducing the file to the text it carries — Word's own tables are skipped whole, and paragraphs, tabs, escaped bytes and `\uN` sequences are resolved. Added `.nfo` previews that read the file the way scene art was written: CP437 rather than a Windows code page, ANSI color codes consumed into actual colors, and those colors pulled toward the page they are drawn on.
- Added encoding detection to the text reader: a UTF-8 or UTF-16 byte order mark decides the encoding outright, UTF-16 without a mark is recognized from where its NUL bytes fall, and otherwise the file is read as UTF-8 and falls back to Windows-1252. A NUL byte, or a file that is mostly control characters after decoding, means the file is not text and is skipped rather than shown as a screenful of replacement characters.
- Added the layout rules that make a screenful readable: code, markup and NFO art keep their columns and are clipped at the right edge, while prose wraps between words. A frame is pulled back until it can be filled from where it starts, and a line is kept at the end for a note whenever the frame cannot show everything; one preview reads up to 2 MB and can be scrolled through its first 2000 lines.
- Added `Enable Text Preview Full Mode`, a tray toggle above the text preview theme that decides whether a text preview is something to work with or only something to look at. With it on a preview can be scrolled, selected from and copied, and a pointer can rest on it; with it off the preview closes the moment the pointer touches it.
- Added a second half to the text gate, for the files a repository is recognized by. Dot files are read by the name they are written with, with the dot dropped on both sides, and a new `names` list in `config.ini` carries the extensionless ones — `LICENSE`, `Makefile`, `Dockerfile` — matched whole and without regard to case. A file is previewed when either list claims it.
- Replaced the `Enable Off Trigger Key` tray item with a `Trigger Key (Alt)` submenu holding the two things the key can mean, one of which is always active: **Trigger Key to Disable Preview** (the default) and **Trigger Key to Enable Preview**, where previews wait for the key instead of being suppressed by it. The settings are `trigger_key` and `trigger_key_mode`, and a `config.ini` written before the rename still names its key.
- Added text selection and copying to text previews in full mode: dragging across the text selects it, `Ctrl+C` copies it, and a right-click opens a menu whose single item is `Copy`. A selection is measured in the text rather than on the screen, what is copied is what the frame shows, and the clipboard is the system one, handed over with `SetClipboardData` so a failure at any step leaves the previous contents alone.
- Added text preview scrolling with a scrollbar drawn into the preview itself: the wheel scrolls the text, the thumb can be dragged, and a pointer that is on its way to the preview is no longer read as the user leaving Explorer. The wheel reaches the preview through the low-level mouse hook rather than the window, so Explorer does not scroll its list behind it.
- Added a `text_scroll_far_edge_grace_pixels` setting that gives the pointer between 0 and 1000 logical pixels of grace past the far edge of a text preview in full mode, scaled by the display's DPI, so a hand that overshoots the scrollbar by a few pixels does not take the preview down with it.
- Added a `theme` folder beside `config.ini` whose `.tmTheme` files join the `Text Preview Theme` tray submenu below the two bundled themes, read from the folder every time the menu is built, so a theme dropped in or edited while the app runs needs no restart. A file theme is written as `theme=custom:atom-one-light`, and a file that is missing, unreadable or not a TextMate theme falls back to `Atom One Light` rather than taking the preview with it.
- A text preview no longer styles a whole file to show the first screenful of it. What is parsed once is the text and where its lines start, and styled lines are built per window, on demand, indexed by absolute line number and continued from the nearest checkpoint, so nothing past the window is ever colored and reopening the same file at the same place costs nothing.

### Changed

- `Best Position` now centers a preview on the line it belongs to instead of on the display: the preview's vertical position is the cursor's own row (or the focused item's), while the display edge still wins, and one taller than the display is pinned to the top. `Follow Cursor` is unchanged, and the rule covers every preview kind because it lives in the shared layout both mouse and keyboard previews are computed from.
- The Explorer hook's media gate now accepts text files, which is what makes a text file resolve the way an image does, and what puts text files into the folder and search-root indexes the hook builds. Those indexes keep their existing bounds, so the change is in what can be found, not in how much work a rebuild costs.
- A text preview is rendered into the same BGRA frame an image arrives in, so it rides the existing machinery unchanged: the spinner, the generation check, the background compositing and the reused layered surface. The frame is forced opaque on its alpha byte, so a text preview does not depend on the `transparent_background` mode it is drawn under.
- The parsed document is cached per path, modification time, length, theme and Markdown mode, so a second hover, a repaint and a theme switch cost a layout instead of a read, a decode and a highlight. The parse state that continues a source file from part-way into it is kept per thread, while the styled windows are shared between the thread that measures a preview and the worker that paints it.
- Fonts and GDI objects are created per paint and released with it, with the memory DC handed back its original object before they go, so a hover leaves nothing behind in a process that may run for weeks in the tray.
- Documentation is up to date with the release: README.md covers the text formats, the two bundled themes and the `theme` folder beside them, the Markdown modes, the sizing rules, scrolling a preview and the editable extension list, and ARCHITECTURE.md's Text Previews section describes the measure/render split, the windowed document model, the layout rules and the theme-switch path.
- Bumped version to 0.2.1 in Cargo.toml and Cargo.lock

### Fixed

- Scrolling a text preview in full mode stops at the end of the document, and it reaches it. The pull-back that keeps the last screenful full was counted in lines that fit rather than lines that were drawn, so a wrapped line could leave the last lines of a file unreachable; it is measured in the document's own lines now, with less than a line left below the last one.
- A rendered Markdown document no longer scrolls past the end of what there is to draw. Its scroll position, its scrollbar and the lines it says are left out were counted in the file's own lines while a rendered document is drawn in the lines its renderer makes, so scrolling into the second half of a long `.md` could land on a frame of empty page; the walk that renders the document counts what it drew now, a position past the end resolves to the last screenful, and a thumb drag is mapped in the range the bar was drawn for.

## [0.2.0] - 2026-09-16

### Added

- Added PDF previews: the first page of a `.pdf` file, rendered by the PDF engine that ships with Windows (`Windows.Data.Pdf`) rather than by a bundled renderer. It adds no crate, ships no native library in the installer and asks the user to install nothing, which is what made it the cheaper pick over `pdfium` and over a pure-Rust rasterizer.
- The page is rendered at exactly the pixel size the layout asks for, so it is drawn once at the size it is shown instead of being enlarged afterwards, and the size the engine reports for the page (DIPs, 96 per inch) is what the preview is placed and sized from, so a page that is not A4 or Letter lands correctly instead of at a guessed shape. The engine is asked to encode BMP, so the decode on this side is a header parse instead of a PNG inflate, and the page is painted on an opaque white background.
- A page is measured once: page 1's size is resolved when the hover arrives and cached per path afterwards, failures included, in a cache bounded the way the video geometry cache is. Both the sizing probe and the render initialize a multithreaded apartment on their thread first, which is what WinRT requires.
- A `.pdf` file is confirmed to carry a `%PDF-` header within its first kilobyte before the path reaches the OS renderer, so a mislabeled file never gets parsed by it, and a PDF that cannot be rendered — password-protected, damaged or mislabeled — is simply not previewed instead of showing an error.

### Changed

- PDF previews ignore the `Preview Scaling` setting and always take the largest size the display area allows, since a page is rendered at the preview's size rather than enlarged afterwards and a bigger preview is sharper text rather than a blurrier one. The decision lives in one place (`effective_preview_scale`) so the layout and the render always agree, and the fitted size is still clamped to the free space on the chosen side.
- Documentation is up to date with the release: README.md lists PDF under supported formats and states what is and is not previewed, and ARCHITECTURE.md documents the PDF path, the apartment both rendering threads initialize, the page-size cache and the header check.
- Bumped version to 0.2.0 in Cargo.toml and Cargo.lock

### Fixed

- PDF previews appear when a PDF is hovered in a normal folder view. The document was loaded through `StorageFile`, which rejects the paths the Explorer hook hands over — canonicalizing a shell path returns the verbatim form (`\\?\G:\...`) and `GetFileFromPathAsync` fails on it — so the page size could not be read and the hover showed nothing at all. The file is now read with `std::fs::read` and handed to the engine as an in-memory stream, and the explicit `Storage` feature that route needed is gone from the manifest.

## [0.1.14] - 2026-09-15

### Added

- Added image previews for the rest of the formats the bundled `image` decoder already reads: APNG (`.apng`), Targa (`.tga`), the portable anymap family (`.pbm`, `.pgm`, `.ppm`, `.pam`, `.pnm`), Radiance HDR (`.hdr`), OpenEXR (`.exr`), QOI (`.qoi`) and farbfeld (`.ff`). They add no dependency and no startup or idle cost, since the codecs were already compiled into the binary and only the extension gate kept them out of previews.
- Added APNG playback for files with more than one frame, recognized from the file's content rather than its extension: a `.png` file's chunk list is walked for an `acTL` chunk ahead of its first image data, so a `.png` animation takes the animation path while every other `.png` keeps the static one. A single-frame animation still shows statically, and a malformed or truncated chunk list is treated as static.
- Added video previews for the formats FFmpeg natively demuxes instead of the eight hardcoded extensions: the whole MPEG-TS family, MPEG-PS and elementary streams, Windows and TiVo recordings, raw codec streams, ISO base media variants, RealMedia, Flash, Ogg, MXF/GXF, DV, NUT, NSV, IVF, Y4M, MJPEG, and the game and camcorder containers FFmpeg reads. `.ts` and `.mts` are shared with TypeScript sources, so those two are probed with a 2 KB read that requires an MPEG-TS sync byte at the file's packet stride.
- Added a `Preview Scaling` tray submenu with `Fit to Screen` and percentages from `25%` to `400%`, persisted through the new `preview_scale` config key. The key also accepts hand-edited values (`fit`, `fit to screen`, or any percentage from 1 to 1000 written as `200`, `200%` or `75`), while `0` resets to the `100%` default.
- Added single-instance enforcement: launching the app while it is already running detects the running copy through a session-local named mutex and exits immediately instead of adding a second tray icon with its own Explorer hook and preview window. The guard is claimed before DPI setup, COM initialization, config loading and every background thread, and the mutex is released when the process ends, including after a crash or a forced kill.
- Added `src/wheel_input.rs`: a system-wide low-level mouse hook (`WH_MOUSE_LL`) installed on its own thread with a message pump, publishing a monotonic wheel-tick counter that the Explorer hook consumes, so the polling loop has a signal for wheel input without touching Explorer or its accessibility providers.
- Added system sleep/resume resilience: the preview window handles `WM_POWERBROADCAST` to reset the layered window's composition surface and re-assert its styles, the tray window re-adds its icon after a DWM/Explorer restart, and video playback and background decoding are cleaned up on suspend/standby to avoid resource leaks.

### Changed

- Preview sizing is derived from the requested scale instead of always rendering at the media's native resolution, and every requested scale is capped by the space available beside the cursor or focused item — in both position modes and for mouse and keyboard previews — so a scaled-up preview is reduced to fit rather than clipped by the display edge. `Fit to Screen` scales the preview as large as the display area allows, including enlarging images smaller than the screen, and enlarged GIF and animated WebP frames use a smooth filter instead of nearest-neighbor.
- Preview layout is bounded to the display nearest the hovered cursor or focused item instead of the whole virtual desktop, so the preview no longer clips onto a neighboring monitor when multiple displays are attached or their configuration changes.
- The layered preview window keeps one memory DC and one DIB section for its lifetime instead of building and tearing both down on every repaint, and frames are composed straight into that surface. The per-frame buffer and the copy that moved it into the DIB are both gone, and the surface is rebuilt only when the preview size changes; every byte written is the byte the previous path produced.
- Video playback no longer re-scans the desktop for the `ffplay` window once it is on screen: the style monitor and the preview loop reuse the window the first scan found while it still exists and still belongs to the recorded player process, and fall back to a full enumeration only when that handle stops being a window or no longer belongs to the player.
- Explorer detection switched from ShellWindows COM enumeration to EnumWindows with class-based matching (`CabinetWClass`/`ExplorerWClass`), so idle polling no longer triggers Explorer-side COM provider allocations, and the cursor-over-Explorer check is a pure HWND/class walk now that ShellWindows is never called during cursor polling.
- Hover and keyboard focus probing is input-aware: probes run only when input is recent or a preview is already active, and a stationary-hover latch caps stationary work to one probe per parked cursor.
- Wheel scrolling refreshes the preview without moving the mouse: a wheel tick counts as user input, the item that lands under the parked cursor is resolved once the list stops moving, and a scroll-driven probe only acts on an item that survives two consecutive probes. A tick is only counted while the wheel is driving Explorer.
- Keyboard previews take priority over the pointer: a preview opened with the arrow keys stays topmost over a parked cursor and the pointer-driven triggers stay frozen until the cursor is moved more than 20 px on purpose or the wheel is turned, at which point the keyboard preview is closed, the pointer freeze is released and the screen is handed back to the mouse.
- The Explorer hook no longer allocates on every poll tick: the per-tick configuration snapshot no longer clones the off-trigger key string, the off-trigger virtual key is resolved once when the configuration changes, the Explorer class check compares class names in place, and the 250 ms Explorer window and folder snapshot hands back a reference-counted list instead of cloning the vector and every folder path.
- The video extension list moved into a single `src/video_formats.rs` module exposing `is_video_file`, replacing the identical eight-extension constant and extension check that were duplicated in `explorer_hook.rs` and `preview_window.rs`.
- Documentation is up to date with the release: README.md covers the expanded image and video format support, preview scaling, display-aware placement, flicker-free presentation, wheel-scroll refresh, keyboard preview priority and single-instance behavior, ARCHITECTURE.md documents single instance, preview placement and the video process lifecycle, and the NSIS installer asset name in README.md was corrected to `rust-hover-preview_<version>_x64-setup.exe`.
- Bumped version to 0.1.14 in Cargo.toml and Cargo.lock

### Fixed

- Animated WebP, APNG and GIF previews no longer stop part-way and start over before reaching the end: the streaming decoder's 256 MB and 300-frame caps are gone, and playback runs through a sliding window instead. The decoder waits while it is a few frames ahead, the player releases frames once 16 MB of them have piled up behind the playhead, and the whole file plays however many frames it holds with memory bounded by the window.
- Video previews no longer risk leaving a stray `ffplay` process, and its frozen window, behind when a stop does not take effect. Stopping is kill-only and never blocks in `wait`, the recorded PID is kept until the process is confirmed gone and re-checked by the hook and before a new spawn, and both checks terminate only when the PID still belongs to `ffplay.exe`, read from the same process handle that is terminated.
- A single navigation key press now opens the keyboard preview right after a folder change: the post-folder-change gate also lifts on a navigation key press seen after the change — the press transition, not the key-down state — instead of swallowing that first press as a baseline. After a mouse-hover preview the first key press likewise switches straight to the keyboard preview.
- Entering a folder no longer waits for a mouse move before previewing the item under the cursor: clicks, Enter and the history keys are tracked as deliberate input, so the folder probe runs at the active cadence for a moment after such a press. A folder change that no input precedes still waits for the user, and the auto-focused first item still never previews without a key press.
- Wheel-scrolling Explorer no longer leaves the preview stuck on the file that was under the cursor before the scroll, and a wheel scroll takes over from a keyboard preview instead of leaving it frozen while the list scrolls underneath it. The file the keyboard showed is not latched, so the mouse may preview it again.
- After the mouse takes over from a keyboard preview, nothing is previewed until the keyboard is used again or the mouse hovers a different file, so moving the mouse off a keyboard preview no longer re-previews the file that was already shown.
- Preview placement and presentation interact correctly with display changes: the preview window is repositioned before the new frame is installed, because crossing between displays of different scale raises `WM_DPICHANGED`, which resets the preview, and both the window and the loading spinner are painted before being revealed, so a new hover can no longer flash the previously previewed image at the new position and size.

## [0.1.14-rc.10] - 2026-09-15

### Added

- Animated PNGs are now recognized inside `.png` files instead of only in `.apng` files: the chunk list is walked for an `acTL` chunk ahead of the first image data, so a `.png` animation takes the animation path while every other `.png` keeps the static one. A single-frame animation still falls back to a static preview, and a malformed or truncated chunk list is treated as static.

### Fixed

- Animated WebP, APNG and GIF previews no longer stop part-way and start over before reaching the end. The 256 MB and 300-frame caps are gone and playback runs through a sliding window instead, so memory is bounded by the window rather than by the length of the animation, the whole file plays however many frames it holds, and an animation small enough to fit is still decoded once and loops from memory.

### Changed

- The layered preview window keeps one memory DC and one DIB section for its lifetime instead of building and tearing both down on every repaint, and the surface is rebuilt only when the preview size changes. The window is still painted before it is revealed, so a hover cannot flash the previous image at the new position.
- Frames are composed straight into that surface: the per-frame buffer and the copy that moved it into the DIB are both gone, and the composition loop walks the frame row by row so each pixel's position is a counter instead of a division.
- Video playback no longer re-scans the desktop for the `ffplay` window once it is on screen, reusing the window the first scan found while it still exists and still belongs to the recorded player process. The styles applied to the window and the PID-verified stop path are unchanged.
- The Explorer hook no longer allocates on every poll tick: the off-trigger key string is not cloned per tick and its virtual key is resolved once when the configuration changes, and the Explorer class check compares `CabinetWClass`/`ExplorerWClass` in place instead of lowercasing every top-level window's class name on every enumeration.
- The Explorer window and folder snapshot is shared rather than copied: cache hits on the 250 ms snapshot hand back a reference-counted list instead of cloning the vector and every folder path, and when the snapshot is rebuilt did not change.

## [0.1.14-rc.9] - 2026-09-15

### Added

- Image previews now cover the rest of the formats the bundled `image` decoder already reads: APNG, Targa, the portable anymap family, Radiance HDR, OpenEXR, QOI and farbfeld. All of them decode through the same `image` crate that already handles PNG and JPEG, so this adds no dependency and no startup or idle cost.
- APNG previews animate like GIFs when the file has more than one frame and fall back to the static image path when it does not. Frames arrive already composited by the decoder, so playback reuses the existing animated pipeline unchanged.
- Video previews now cover the formats FFmpeg natively demuxes instead of the eight hardcoded extensions: the MPEG-TS family, MPEG-PS and elementary streams, Windows and TiVo recordings, raw codec streams, ISO base media variants, RealMedia, Flash, Ogg, MXF/GXF, DV, NUT, NSV, IVF, Y4M, MJPEG, and the game and camcorder containers FFmpeg reads.
- `.ts` and `.mts` are shared with TypeScript sources, so those two extensions are no longer decided by name alone: the file is probed with a 2 KB read that requires an MPEG-TS sync byte at a 188-, 192- or 204-byte packet stride and at the alignment the file starts at. The probe only runs for the two ambiguous extensions, so every other hover stays a plain extension lookup.

### Changed

- The video extension list moved into a single `src/video_formats.rs` module exposing `is_video_file`, replacing the identical eight-extension constant and extension check that were duplicated in `explorer_hook.rs` and `preview_window.rs`.

### Fixed

- Video previews no longer risk leaving a stray `ffplay` process, and its frozen window, behind when a stop does not take effect: stopping is kill-only and never blocks in `wait`, the recorded PID is kept until the process is confirmed gone, and both checks terminate only when the PID still belongs to `ffplay.exe`.
- Entering a folder no longer waits for a mouse move before previewing the item under the cursor: clicks, Enter and the history keys are tracked as deliberate input, so the folder change they cause is recognized as user navigation instead of a change the gate has to wait out.
- A single navigation key press now opens the keyboard preview right after a folder change, because the gate lifts on the press transition seen after the change rather than swallowing the first press as a baseline.

## [0.1.14-rc.8] - 2026-09-15

### Added

- Added `src/wheel_input.rs`: a system-wide low-level mouse hook (`WH_MOUSE_LL`) installed on its own thread with a message pump. It publishes a monotonic wheel-tick counter that the Explorer hook consumes, giving the polling loop a signal for wheel input without touching Explorer or its accessibility providers.

### Fixed

- Wheel-scrolling Explorer no longer leaves the preview stuck on the file that was under the cursor before the scroll: a wheel tick counts as user input, the hover stability window restarts while the wheel is turning, and the item that landed under the parked cursor is resolved and previewed once the list stops moving.
- A wheel scroll now takes over from a keyboard preview: scrolling while an arrow-key preview is on screen closes it, releases the pointer freeze and hands the screen back to the mouse. The file the keyboard showed is not latched, so the mouse may preview it again.
- A scroll-driven probe only acts on an item that survives two consecutive probes, and a scroll that leaves the same file under the cursor keeps the current preview instead of restarting it. A tick is only counted while the wheel is driving Explorer, so scrolling a menu, a browser or the desktop changes nothing.

## [0.1.14-rc.7] - 2026-09-14

### Added

- Added single-instance enforcement: launching the app while it is already running — desktop shortcut, Start menu, startup entry, or the `.exe` directly — detects the running copy through a session-local named mutex (`Local\rust-hover-preview-single-instance`) and exits immediately instead of adding a second tray icon with its own Explorer hook and preview window.
- The guard is claimed before DPI setup, COM initialization, config loading and every background thread, so a duplicate launch does no work at all. The mutex is released when the process ends, including after a crash or a forced kill, so the next launch becomes the primary instance again.

### Fixed

- Keyboard previews now take priority over the pointer: a preview opened with the arrow keys is no longer dismissed when the mouse cursor happens to sit where the preview appears, and it stays on top of the parked cursor instead of blinking away on every navigation key.
- Every keyboard preview now measures its own on-screen box once it appears, and if that box covers the mouse cursor the pointer-driven triggers are frozen until the cursor is moved on purpose. Cursor jitter below 20 px no longer counts as movement, so a parked or lightly nudged mouse cannot cancel a keyboard preview.
- The cursor-over-preview check no longer runs on every poll: it is skipped while a keyboard preview is on screen or the pointer is frozen, and outside those states it is a single shared check.
- After the mouse takes over from a keyboard preview, nothing is previewed until the keyboard is used again or the mouse hovers a different file.
- The first keyboard navigation key now switches straight to the keyboard preview instead of being swallowed as a fresh focus baseline.

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

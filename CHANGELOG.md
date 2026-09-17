# Changelog

## [0.2.6]

### Added

- Office previews: hovering a Word, Excel or PowerPoint document shows a page of it, drawn by the installed Office in the background — Word and Excel export page 1 to PDF, PowerPoint exports slide 1 as an image, and a workbook on a machine with no printer is copied out as a picture of its used range — with the page cached under `%LOCALAPPDATA%\rust-hover-preview\cache\office`, so a document previews from its first hover and instantly after. No engine is started until the pointer has rested on the file for two seconds.
- `office_preview_enabled`, `office_render_enabled` and `office_cache_mb` in `config.ini`, an editable `[office] extensions` list, an `Office` entry in the tray's `Preview Types` submenu, and an `Office Preview` submenu carrying `Render With Office`.

### Changed

- Office previews are drawn from the rendered page alone. The picture a document saves inside itself is no longer read at all: a thumbnail a couple of hundred pixels across is either tiny in a preview or an enlargement of something that small, and skipping it keeps the preview thread from parsing a zip or a compound file. The `cfb` reader the legacy formats needed went with it.
- Nothing about an Office preview is held in memory between hovers. A page's own size is read from the cached file each time it is measured, so no cache of the app's can go stale, grow with the documents hovered, or stand between a hover and its preview.
- While a page is being rendered the preview is the spinner's own box — just big enough for it — instead of one shaped like the page that is coming.
- The page that arrives for the hover on screen replaces what is up without the preview being taken down first: the spinner stays while a large picture is decoded and is swapped for the page when it is ready, instead of blinking off and back.
- A picture larger than the box it is shown in is now shrunk with the box filter rather than the smooth one, which is the difference between half a second and a quarter of a second for a worksheet's corner at screen resolution.

### Fixed

- PowerPoint decks preview: a slide is reached through the collection's item — `Slides.Item(1)`, which is what VBA's `Slides(1)` means — because asking `Slides` itself for one is answered with "member not found", which left decks with a spinner and nothing else.
- Excel workbooks preview on a machine with no printer: exporting a page goes through the print pipeline and refuses to run without one, so where `ActivePrinter` reports none — or the export fails — the used range's top-left is copied out as a picture instead. A machine with a printer exports its page as before.
- A render that never returns no longer costs the documents after it: a worker still inside one piece of work after thirty seconds is given up on when the next hover arrives, the Office process this app started is ended with it, and a fresh worker answers that hover. Ending an engine counts as such work, so a quit that hangs is given up on too. A panic inside one render is contained the same way — that document fails, not the tier — and a worker hands its thread id back however its thread ends.
- An Office engine no longer survives its own quit, which is what left an Excel instance running after the first workbook: Excel finishes quitting while its automation client is pumping, so the quit is given a moment of pumping and a process that is still there afterwards — one this app started, verified as still being that application — is ended. Every later render used to attach to the instance that had been left behind, so a workbook preview could stop working for good.
- The picture a workbook is copied out as is taken off the clipboard rather than left on it: leaving it there is what makes Excel ask, on its way out, whether a large amount of information should stay on the clipboard — a dialog no one is present to answer.
- A file that is not a document at all — a text file named `.docx`, say — no longer starts an Office engine: the file's own header has to say it is an OOXML package or an OLE compound file first.
- The box a preview is placed for while its page is being rendered now follows the family: a portrait sheet for Word, a landscape one for Excel, a 16:9 slide for PowerPoint, instead of one portrait page for all three.
- A render that produces nothing now remembers what Office said about it, and an ignored test (`cargo test -- --ignored office_render_smoke_test`) drives a document of each family through the real path and reports every step.

## [0.2.5] - 2026-09-17

### Added

- Archive previews: hovering an archive shows a summary line and a capped tree of what it holds, read from its own table of contents without unpacking anything.
- `archive_preview_enabled` and an `[archive] extensions` list in `config.ini`, plus an `Archives` entry in the tray's `Toggle Preview Types` submenu.
- Archive listings are painted with the text preview's themes, fonts and margins, and follow the `Text Preview Font Size` setting.
- `Avoid Filename` in the tray's `Preview Position` submenu, or `avoid_filename` in `config.ini`: a preview is moved clear of the name of the file it is about, and resized where the display leaves no room for it beside the name. On by default, and applies to both positions.

### Changed

- The tray menu is grouped into categories — `Text Preview`, `Timing`, `Placement` and `Volume` among them — with shorter labels, and `Config.ini` carries the running version.
- The minimum supported Rust version is now 1.88, the `zip` crate's floor.
- The GDI painting a text preview and an archive listing share moved into `text_paint.rs`.
- Archive listings keep room between an entry's icon and its name.
- The tray's `Preview Position` entries carry radio marks, and `Avoid Filename` a checkmark of its own.
- RAR archives are read through RARLAB's UnRAR sources compiled in by the `unrar` crate. See the licence note in README.

## [0.2.4] - 2026-09-17

### Added

- `image_cache_mb` in `config.ini` (default `64`) sets how much memory decoded image previews may be kept in, so hovering back over a folder no longer decodes the same pictures again. `0` turns the cache off, and the value is capped at `2048`.

### Changed

- The preview thread no longer wakes on a timer while nothing is on screen: it waits on the hover channel instead, so an idle app stops waking sixty times a second, and a hover is answered as it arrives rather than on the next tick.
- Animated GIF, APNG and WebP previews open on their first frames instead of waiting for a startup buffer.
- Video previews read a file's dimensions and detect its letterboxing at the same time instead of one after the other.
- A PDF's page size is remembered from the render that read it, so a file is not opened twice for one preview.
- Bumped version to 0.2.4 in Cargo.toml and Cargo.lock

### Fixed

- A PDF is no longer read into memory in full before its first page is rendered, so a large one costs memory in proportion to the page rather than to the file.
- A OneDrive or SharePoint file that has not been downloaded no longer starts downloading when the pointer rests on it.
- A very large image can no longer take the app down with it.
- Hovering no longer stalls indefinitely when Explorer stops answering.

## [0.2.3] - 2026-09-17

### Changed

- Hover previews in a search results view are resolved from the item under the pointer instead of its file name, so a search of any size previews as fast as a folder does.
- Keyboard previews resolve the same way, from the item the focus is on.
- Removed the name-based fallbacks and the search-result indexing behind them: no directory walks, no background scans, no index caches, no MSAA.
- A file nothing can vouch for is left without a preview rather than guessed at by name.
- Documentation updated: ARCHITECTURE.md describes the resolution, and the search-results known issue is gone from TODO.md.
- Text previews wrap long lines instead of cutting them off at the right edge.

### Fixed

- A text preview of a file with long lines is now tall enough to show the rows the line wraps into, instead of cutting them off below the first one.
- Search results that share a file name each preview their own file.
- A search result that lives in another folder previews from the keyboard too, not only under the pointer.
- A Details or Content row previews from anywhere on it, not only from its text.
- A keyboard preview of a Content view row is placed in the empty tail past the row's last column — the edge the row's own text stops at, read from the view — instead of in the middle of the display, and the space to the right of the columns is what sizes it.
- A search result previews in any tab of an Explorer window, and follows the tab when it changes.
- Hovering a folder, an executable or any other item this app does not preview no longer shows another tab's file.
- Bumped version to 0.2.3 in Cargo.toml and Cargo.lock

## [0.2.2] - 2026-09-16

### Added

- New tray menu **Toggle Preview Types** to turn image, video, text, and PDF previews on or off separately.
- **Select All** in the text preview right-click menu.
- Better detection for `.ts` and `.mts` files (video vs. text).

### Changed

- Removed the old **Enable Text Preview** tray item; it's now under **Toggle Preview Types**.
- Documentation updated.
- Version bumped to 0.2.2.

### Fixed

- Search results with the same file name are now told apart correctly.
- Keyboard previews respond faster, especially in search results.
- Text previews no longer accidentally respond to Ctrl+A.
- Previews no longer get stuck or fight between mouse and keyboard.
- Many other small fixes for folders, search, and keyboard navigation.

## [0.2.1] - 2026-09-16

### Added

- Text previews: see the first screen of code, Markdown, logs, `.nfo`, `.rtf`, and more.
- Syntax highlighting and rendered Markdown.
- Light and dark themes, plus custom themes from a `theme` folder.
- Font size options (100%–400%).
- Full mode: scroll, select, and copy text.
- Toggle text previews on/off.
- Better encoding detection.

### Changed

- Text previews size themselves to fit content.
- Previews are faster after first load.
- Documentation updated.
- Version bumped to 0.2.1.

### Fixed

- Scrolling and Markdown preview fixes.

## [0.2.0] - 2026-09-16

### Added

- PDF previews: see the first page of a PDF using Windows’ built-in PDF reader.

### Changed

- PDF previews always use the best available size for sharp text.
- Documentation updated.
- Version bumped to 0.2.0.

### Fixed

- PDF previews now work in normal folder views.

## [0.1.14] - 2026-09-15

### Added

- More image formats: APNG, TGA, HDR, EXR, QOI, and more.
- Animated PNG support.
- Many more video formats.
- **Preview Scaling** menu: fit to screen or 25%–400%.
- Only one copy of the app can run at a time.
- Better sleep/resume behavior.

### Changed

- Previews stay on the correct monitor.
- Smoother, flicker-free previews.
- Keyboard previews take priority over the mouse.
- Wheel scrolling refreshes the preview.
- Documentation updated.
- Version bumped to 0.1.14.

### Fixed

- Many fixes for animated images, video cleanup, keyboard/mouse switching, and folder changes.

## [0.1.14-rc.10] - 2026-09-15

### Added

- Animated PNGs recognized inside `.png` files.

### Fixed

- Animated WebP, APNG, and GIF no longer stop part-way.

### Changed

- Smoother preview window and video playback.
- Explorer hook uses less memory.

## [0.1.14-rc.9] - 2026-09-15

### Added

- More image formats (APNG, TGA, HDR, EXR, QOI, farbfeld).
- Many more video formats.
- Better `.ts`/`.mts` detection.

### Changed

- Video format list moved to one place.

### Fixed

- No stray video processes left behind.
- Folder previews work without moving the mouse.
- First arrow key press shows keyboard preview after folder change.

## [0.1.14-rc.8] - 2026-09-15

### Added

- System-wide mouse wheel detection.

### Fixed

- Wheel scrolling no longer leaves the preview stuck.
- Wheel scroll takes over from keyboard preview.
- Scroll probes only act on stable items.

## [0.1.14-rc.7] - 2026-09-14

### Added

- Single-instance enforcement (only one copy runs).

### Fixed

- Keyboard previews take priority over the mouse.
- Cursor jitter no longer cancels keyboard previews.
- Mouse takes over cleanly from keyboard.

## [0.1.14-rc.6] - 2026-09-14

### Added

- **Preview Scaling** tray menu (Fit to Screen, 25%–400%).
- `preview_scale` config key with hand-editable values.

### Changed

- Previews scale to fit screen edges.
- Smoother upscaling for GIF/WebP.
- Version bumped to 0.1.14-rc.6.

## [0.1.14-rc.5] - 2026-09-14

### Changed

- Documentation for display-aware placement.
- Installer asset name corrected.

### Fixed

- Previews bounded to nearest display.
- No flicker when changing previews.
- Reposition before installing new frame.

## [0.1.14-rc.3] - 2026-07-03

### Added

- Sleep/resume resilience for preview and tray windows.
- Cleanup on suspend/standby.

### Changed

- Installation instructions clarified.
- Config example updated.
- Version bumped to 0.1.14-rc.3.

## [0.1.14-rc.2] - 2026-07-01

### Changed

- Faster Explorer detection (less COM usage).
- Input-aware probing.
- Stationary hover probe runs only once per hover.
- Version bumped to 0.1.14-rc.2.

## [0.1.14-rc.1] - 2026-06-24

### Changed

- Better multi-monitor and display change handling.

## [0.1.13] - 2026-06-03

### Changed

- Release 0.1.13: version bump and changelog update.

## [0.1.13-rc.2] - 2026-06-03

### Added

- Features section in TODO.md for document support.

### Changed

- Increased cache limits for performance.
- Simplified deploy workflow.
- Version bumped to 0.1.13-rc.2.

## [0.1.13-rc.1] - 2026-06-03

### Added

- Per-monitor DPI awareness.

### Changed

- Removed MSI installer references.
- Switched to rust-lld.
- Documentation updates.
- Added ARCHITECTURE.md.

### Fixed

- Fixed high-DPI scaling artifacts.

## [0.1.12] - 2026-04-30

### Added

- Shell view index sync limits and cache clearing.

### Changed

- Refactored media path resolution.
- Enhanced hover preview logic for large result sets.

### Fixed

- Fixed hover preview state management.

## [0.1.11] - 2026-04-30

### Added

- WebP playback FPS configuration.

### Changed

- Refactored file handling in Explorer hook.
- Improved Windows resource description.
- Better search view metadata handling.

## [0.1.10] - 2026-04-29

### Changed

- Restored default NSIS older-version prompt.
- Limited automatic startup to first install/run.
- Removed custom installer taskkill actions.
- Updated release metadata.

## [0.1.9] - 2026-04-29

### Added

- Transparent background modes for PNG/WebP.
- Configurable same-file rehover delay.
- Installer migration cleanup.

### Changed

- Switched to Google's libwebp for animated WebP.
- Moved installer output to `%LOCALAPPDATA%`.
- Config path uses `directories` crate.
- Updated tray labels and tooltip.

### Fixed

- Installers stop running app before upgrade.
- Config default completion.
- Stuck video previews hidden more aggressively.
- Explorer slow-probe backoff.
- Repeated same-file hover timing.
- Transparent PNG/WebP rendering.
- Removed dead code.

## [0.1.8] - 2026-04-29

### Added

- Optional off-trigger key.
- Active Explorer Shell view detection.
- Media indexing from search roots.

### Changed

- Improved path normalization and URL handling.
- Cached Explorer folder lookups.
- Updated installer build orchestration.

### Fixed

- Removed branch-triggered release deployment.

## [0.1.7] - 2026-04-28

### Added

- Windows installer packaging via `cargo packager`.
- `build-installers.ps1` helper.

### Changed

- Updated deploy workflow to publish installers.
- Refactored preview window and media loading.
- Improved installer build orchestration.

### Fixed

- Added Windows Installer availability checks.
- Improved installer error handling.

## [0.1.6] - 2026-04-06

### Changed

- Optimized animated GIF/WebP streaming.
- Limited animation to 300 frames and 30 FPS.
- Added **Confirm File Type** tray option (off by default).

### Fixed

- Reduced CPU spikes for animated WebP.
- Improved rapid-hover behavior.
- Fixed `.jpeg` hover preview.
- Mislabeled images load with correct decoder.
- Prevented previous image flash.

## [0.1.5] - 2026-04-01

### Changed

- Improved Explorer folder caching.
- Optimized media file resolution.

## [0.1.4] - 2026-03-03

### Changed

- Refined folder-navigation input gate.

### Fixed

- Eliminated residual previews after folder open.

## [0.1.3] - 2026-03-03

### Added

- Post-folder-navigation input gate.

### Changed

- Mouse hover target validates against cursor.
- Prioritized media resolution in folder under cursor.

### Fixed

- Prevented auto-preview of first item.
- Keyboard preview starts only after real input.
- Removed dead-code warnings.

## [0.1.2] - 2026-03-03

### Added

- Separate cursor-over detection for image and video previews.
- Startup grace period for video previews.

### Changed

- Refined dismissal logic.
- Improved ffplay window discovery.
- Reasserted topmost styles during video playback.
- Updated ffprobe/ffplay flags.

### Fixed

- Prevented premature video dismissal.
- Reduced false preview closures.

## [0.1.1] - 2026-03-03

### Added

- Keyboard preview support.
- Animated WebP and improved GIF/WebP streaming.
- Loading spinner overlay.
- Tray options: follow cursor and video volume.
- GitHub Actions workflows.

### Changed

- Refactored preview window and media streaming.
- Enhanced animated media playback.
- Optimized CPU polling rates.
- Config uses INI format with auto-save.

### Fixed

- Preview window hides correctly during keyboard hover.
- Corrected preview delay menu wording.
- Improved frame decoding reliability.

### Miscellaneous

- Updated README.
- Initial config generation.

## [0.1.0] - 2026-02-03

### Added

- Image preview on hover: JPG, JPEG, PNG, GIF, BMP, ICO, TIFF, WebP.
- Animated GIF and WebP playback.
- Video preview: MP4, WebM, MKV, AVI, MOV, WMV, FLV, M4V.
- System tray controls.
- INI config file in `%APPDATA%\rust-hover-preview\config.ini`.

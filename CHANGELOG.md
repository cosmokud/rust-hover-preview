# Changelog

## [0.2.14]

### Changed

- New defaults, for new installations: previews are placed at their best position, appear instantly, and wait 200 ms before the same file previews again. Transparency is shown over a checkerboard rather than black — for pictures, drawings and design documents (font specimens stay on a white page) — and `DDS Background` offers only Black and White, starting at white. An installation that already exists keeps the values its `config.ini` holds, so if you see black where you expected a checkerboard, that is your file and not a bug.
- `config.ini` is grouped under the same headings as the tray menu — `General`, `Preview Types`, `Text Preview`, `Timing`, `Placement`, `Scaling`, `Background`, `Volume`, `Performance`, `Advanced` — so a setting is easy to find. Every key the app knows is written there, and the tray marks the item each setting starts at with `(Default)`.
- `config.ini` is checked and tidied every time the app reads it, including right after you edit it: a missing setting comes back, a value the app cannot read is replaced with the value it is actually using, and any line that is not one the app writes is removed. A list you edited yourself is left exactly as it is.
- Settings that older versions stored under a different name are no longer carried over — `svg_scale`, `svg_background`, `svg_preview_enabled`, `off_trigger_key`, `avoid_filename`, `transparent_background`. Those lines are removed and the settings go back to their defaults, so if you are updating from 0.2.12 or older it is worth a look at the **Placement**, **Scaling** and **Background** menus afterwards. Updating from 0.2.13, nothing changes.
- The installer no longer writes into `config.ini` and no longer adds the startup entry: the app does both on its first run, at the path it is running from.
- Bump version to 0.2.14 in `Cargo.toml` and `Cargo.lock`.
- Preview for CorelDRAW `cdr` and Procreate `procreate` documents, under the **Design** kind: both save a picture of the drawing beside their layers, and that picture is what is shown. A CorelDRAW file's picture is a 256-pixel thumbnail, so it is small at Fit to Screen.

### Fixed

- Editing `config.ini` by hand now takes effect as soon as you save it, and a value the app cannot read — a misspelled tone map, a delay past its limit — is written back as the value in use instead of staying in the file.

## [0.2.13]

### Added

- Preview for design documents: Photoshop `psd` and `psb`, Krita `kra`, OpenRaster `ora`, and the project containers other drawing tools save (`sketch`, `fig`, `xd`). A layered document is shown from the finished picture its format keeps inside it, so it arrives as it was saved.
- Preview for Illustrator `ai` files, drawn as a PDF the same way a PDF is.
- `Design` under the tray's **Preview Types**, with `design_preview_enabled` and a `[design]` extension list in `config.ini`.
- `Design Scaling` under the tray's **Scaling** menu, with `design_scale` in `config.ini`: how much of the screen a design preview covers — Fit to Screen (default), or 75%, 50%, 25%, 10%.
- `Design Background` under the tray's **Background** menu, with `design_background` in `config.ini`: what a design preview is drawn over, black by default.
- Vector previews: Windows metafiles (`wmf`, `emf`) and Illustrator files saved as encapsulated PostScript (`eps`, `epsi`), drawn by Windows at the preview's size so they stay sharp. An `.eps` shows the picture its writer saved inside it; one without such a picture shows nothing.
- `Vector` under **Preview Types**, with `Vector Scaling` and `Vector Background` in the tray, and `vector_preview_enabled`, `vector_scale`, `vector_background` and a `[vector]` extension list in `config.ini`. SVG documents are covered by these — the `svg_background`, `svg_scale` and `svg_preview_enabled` keys are still read.
- `ai` files that are not PDFs — Illustrator saved as PostScript — show the preview saved inside them, under the **Design** kind.

### Changed

- `Vector Scaling` now starts at Fit to Screen rather than half the display.
- `svg` and `svgz` moved from the `[image]` list to the `[vector]` list, where the other drawings are; existing `config.ini` files are updated on the next run.
- Bump version to 0.2.13 in `Cargo.toml` and `Cargo.lock`.

### Fixed

- A name in the text list _and_ in a drawing or document list is previewed as that kind again, not as text — the two halves of the app now agree about a file whose name is written down twice.
- SVG previews are back: a document is drawn by the engine again rather than being handed to the readers that replay the other half of the Vector kind, which turned it down.
- A Photoshop document smaller than the screen shows a preview again: it was refused whenever the preview was enlarged past the document's own size.
- EPS files whose preview is a palette picture — the shape Photoshop writes into an EPS — now preview. That shape is read by the app itself, since the picture decoder turns it down.

## [0.2.12]

### Changed

- The scaling options now live in a `Scaling` menu of their own in the tray, right below `Placement`.
- `Font Background` now defaults to white; `font_background` in `config.ini` still sets it.
- Tray labels renamed for consistency: `Image Scaling`, `Video Scaling`, `Avoid Nothing`, and `Videos` in `Codecs`.
- Preview hover checks now put far less load on Windows Explorer, so the file list stays responsive while previews are running.

### Fixed

- Keyboard previews in `Details` and `Content` views now follow `Avoid Filename` and `Avoid Filename Column` instead of behaving like `Avoid Details`: a row of a view is recognized by the columns it draws beside the name rather than by its box being at least half the display across, which read the row of a narrower window as a box item and placed its preview past the whole of its columns at every way of avoiding.
- A failed startup of the Explorer lookup could leave hover previews off for the rest of the run; it is now retried until it works.
- The internal timeout that stops a stalled Explorer from blocking the app is now verified instead of assumed, and refreshes itself if a newer one is ever needed.

## [0.2.11]

### Added

- Video previews without FFmpeg: where `ffplay` is not installed, videos are decoded by the media engine Windows already has — `mp4`, `mov`, `m4v`, `mkv`, `webm`, `avi`, `wmv`, `asf`, `ts`, `m2ts`, `mts`, `3gp`, and more with the codec extensions below — played in the same preview window as everything else, with sound and looping. FFmpeg is now optional rather than required.
- `Codecs` in the tray menu: what this machine has of every engine and codec a preview can lean on, grouped into `Video`, `Images` and `Engines`, with a check or a cross per row and the missing ones greyed. The rows are for reading only — nothing in the menu does anything, and there is no setting behind it.
- `windows-core` as a declared dependency, for the media engine's event callback. Nothing is added to the build graph by it.
- Sample lines for Arabic, Hebrew, Thai and Devanagari, so a font of one of those scripts is drawn as its own script rather than by the first characters its map happens to hold: the Arabic and Hebrew pangrams, and the openings of the Thai pangram written by the Computer Association of Thailand and of the Devanagari one the script is sampled with. A line is still drawn whole or not at all, so a script a font holds every character of is a line and one it does not is the font's own characters, the same rule as before.
- `Font Face` under the tray's **Placement** menu, with `ttc_face` in `config.ini`: which face of a `.ttc` collection a specimen is drawn from — `First Face` (the default) down to `Tenth Face`. A page has no syntax for naming a face inside a collection, so the face is written out as a font of its own, and this is which one; a collection with fewer faces than the setting names is drawn from the last one it has, and the specimen's heading says which face came out — `(2 of 4)` — so what was asked for and what was drawn cannot be taken for each other.
- DDS texture previews now support almost all DDS formats, including compressed, uncompressed, cubemaps, texture arrays, and mip chains; only the first face and first level are shown.
- Signed DDS textures now display correctly: zero is in the middle, and missing color channels show as neutral gray.
- Signed HDR DDS textures now get a preview and use the same HDR tone mapping as other light-based images; negative light is shown as black.
- More DDS formats now preview: packed HDR, depth buffers, old numbered formats, and signed BC4/BC5.
- New `DDS Background` setting lets you choose what DDS textures are previewed over.
- New `hdr_tone_map` and `hdr_exposure` settings control how HDR images are shown.
- New `spinner_delay_ms` setting controls how long to wait before showing the loading spinner.
- New `Videos Scaling` under the tray's **Placement** menu, with `video_scale` in `config.ini`: how large a video is shown, 100% by default.
- New `Animated Scaling` beside it, with `animated_scale` in `config.ini`: how large an animated GIF, WebP or PNG is shown, 100% by default. A check of the file itself tells an animation apart from a still — a GIF or PNG with a single frame follows `Images Scaling` like any other picture.

### Changed

- `README.md` describes what FFmpeg adds rather than requiring it, and lists the free Microsoft Store codec extensions for HEVC, VP9, AV1, MPEG-2 and Ogg.
- Bump version to 0.2.11 in `Cargo.toml` and `Cargo.lock`.
- A specimen's lines carry their own direction, so the Arabic and Hebrew lines are laid out from the right, their full stop ending them where the script ends it rather than where a left-to-right page would; every other line is drawn as it was.
- A font that covers more of the sample lines than the specimen's box was shaped for — a pan-script one, which covers most of them — is drawn at smaller type rather than past the bottom of the box.
- The page a specimen is drawn in is named for the face as well as the backdrop, so switching faces is a page the browser has not seen rather than the one before it answered out of its cache.
- Video files are now checked in the background with a spinner, and the result is remembered so they are not checked again every time.
- Video previews show the spinner until the video player window appears.
- Video players that fail to open or hang are closed automatically.
- The loading spinner is now the same small pointer spinner for every preview type, not a preview-sized box.
- Spinner delay is now unified at 250 ms by default, including Office documents and browser starts.
- EXR and HDR previews are now tone mapped instead of clipped, so bright areas fade naturally instead of burning out.
- DDS files marked as opaque are now shown opaque instead of using an unused alpha channel.
- DDS previews now use the best detail level for the displayed size, reducing work without losing visible detail.
- Truncated or too-short DDS files are handled safely using the real file length.
- DDS is now in the built-in image list, and older config files are updated automatically.
- The tray's `Scaling` is now `Images Scaling`, and videos have a `Videos Scaling` of their own; both stay hand-editable in `config.ini`.
- Animated GIF, WebP and PNG previews follow their own `Animated Scaling`, separate from still images; a `config.ini` written before this gets the animation scale from `preview_scale`.
- `Font Face` is out of the tray menu; `ttc_face` in `config.ini` still picks the face of a `.ttc`.

### Fixed

- Animations no longer stop at the end of their first play. An animation whose frames were partly given back to keep memory in check was left holding its last frame for good, because the file was only decoded again while a decoder was still running and that decoder had already finished; the two are now settled together, so either the animation is still being decoded or the whole of it is in hand and plays from beginning to end, forever. This affected GIF, animated WebP and animated PNG previews.
- Animations with a long frame in them no longer freeze on that frame. Any frame held for a whole second or more — which a GIF is free to ask for, and one of these asks for 1.2 seconds on its first frame — was unreachable: the playhead was treated as having fallen behind once a second had passed, and its clock was reset every tick, so the wait could never end. The frame's own delay is now part of what "behind" means, and a long hold is simply a long hold.
- An animation that fits the memory it is kept in stays whole instead of being taken apart frame by frame: it is decoded once, plays through, and wraps back into the frame it started on rather than being read from disk again on every pass.

## [0.2.10]

### Added

- Font previews for `ttf`, `otf`, `ttc`, `woff` and `woff2`, drawn by the WebView2 engine the SVG previews already use: the name the font calls itself, the pangram _The quick brown fox jumps over the lazy dog._, and a line each for Japanese, Chinese, Korean, Cyrillic and Greek that the font's own character map covers. A line is drawn only where every character of it is in the font — a browser falls back per glyph and says nothing about it — so nothing is ever shown in a system font and passed off as the font; a font of a script there is no line for is shown by the characters its own map holds instead.
- `Font Scaling` under the tray's **Placement** menu, with `font_scale` in `config.ini`: how much of the screen a specimen is drawn over — `Fit to Screen`, or 75%, 50% (the default), 25%, 10%.
- `Fonts` under **Preview Types**, and a `[font]` extension list in `config.ini`; both are written on first run.
- `Font Background` under the tray's **Background** menu, with `font_background` in `config.ini`: a specimen is a page of text, so its ink follows the backdrop — light on black, dark on the light ones, and light with a shadow where there is none to be read against.
- A `.ttc` previews its first face, which is written out as a font of its own beside the browser's profile folder: a page has no syntax for naming a face inside a collection.
- `brotli-decompressor`, for a WOFF2's tables: a specimen is described by two of them — the character map and the name — parsed out of an sfnt, a WOFF, a WOFF2 or a collection's first face on this side, with nothing rasterized.

### Changed

- Build deploy artifacts with a `github` profile (fat LTO, one codegen unit), which with the three changes below takes the released exe from 7,544,320 to 6,676,992 bytes (867 KB, 11.5%); `cargo build --release` keeps thin LTO for the local loop.
- Decode still WebP with the Windows codec at the layout box, removing `image-webp`, and with the bundled libwebp (`webp_image`) where that codec is not installed — which is what a Windows 10 machine usually is; animated WebP is unchanged.
- Remove `rayon` from the build graph, making `.exr` previews decode single-threaded.
- Unify GIF on `gif 0.14` to match `image`.
- Draw SVG previews with WebView2 instead of in-app rasterization, removing `resvg`, `usvg`, `tiny-skia`, and related crates (exe 10.6 MB → 7.2 MB).
- Remove the app’s own animated-document reader, leaving all documents to the engine with no movement check.
- Drop SVG preview on machines without WebView2, since no fallback reader remains.
- Show the spinner immediately on cold-engine document hovers, and not at all when warm.
- Move the engine window with the pointer wait so documents draw where the wait ended.
- Stand down failed engines for one minute instead of five, costing only that hover and clearing its spinner.
- Paint `Checkerboard` for `svg_background` in the engine page so SVG previews keep the standard backdrop.
- Bump version to 0.2.10 in `Cargo.toml` and `Cargo.lock`.

## [0.2.9]

### Added

- HEIC (`.heic`, `.heif`), AVIF (`.avif`) and JPEG XL (`.jxl`) previews, decoded by the codec Windows has rather than by one shipped with the app. The codec extensions they need are listed in the README; a machine without one shows no preview for those files, and a multi-image file (a burst, an animated AVIF) shows its first frame.
- A picture of one of those formats is decoded at the size of the preview rather than the size of the file, so a 48-megapixel photograph costs what its preview costs, and it arrives at the codec as the frame is composed rather than through a resample and two conversions.
- `avif`, `heic`, `heif` and `jxl` in the built-in image list; a `config.ini` written by an earlier version is brought up to it rather than left without them.
- `PDF Scaling` and `Office Scaling` under the tray's **Placement** menu, with `pdf_scale` and `office_scale` in `config.ini`: how much of the screen a page is drawn over — `Fit to Screen` (the default), or 75%, 50%, 25%, 10%.
- `110%` in the tray's **Text Preview → Font Size**.

### Changed

- The `image` crate is compiled with its decoders named instead of its defaults, which takes its AV1 encoder out of the build graph: this path decodes pictures and never writes one.
- PDF and Office pages no longer follow `preview_scale`: each page kind has a share of the display of its own, so a picture's scale leaves a page alone.
- Tray menu regrouped: `Text Preview` sits below `Preview Types`, `Trigger Key` moved into `Timing` above `Delay`, `Confirm File Type` moved into `Performance`, and the two engine idle-time submenus are named for what they set — `Office Engine TTL` and `SVG Engine TTL`.
- Bumped version to 0.2.9 in Cargo.toml and Cargo.lock.

## [0.2.8]

### Added

- New default “Avoid Filename” keeps previews from covering the file name itself.

### Changed

- Tray menu gets an **Avoid** submenu, and `avoid_filename` is replaced by `avoid_mode`.
- Slide previews now use your screen width (640–1920) instead of fixed 1280, while Word/Excel pages stay the same.

### Fixed

- `Avoid Filename` now measures names using the item’s own display zoom, fixing half-covered names on scaled screens.
- Preview spacing and mouse-move distances now scale correctly with screen zoom at 100%, 150%, and 200%.
- Previews now use the screen your active window is on, fixing multi-monitor fullscreen and missing-preview problems.
- If a preview can’t attach to its screen, it now goes to the main screen instead of spanning two monitors.
- Screen/zoom changes or wake-from-sleep now close and reopen browser document previews at the mouse, not on a missing screen.
- The app now notices monitor rescaling or primary-monitor changes, so preview positions update correctly.
- PDF page size is remembered per file version, so a re-exported PDF is measured fresh.
- Screen zoom is now read from the window under the point, with the monitor as backup.
- PDF and Office-exported pages now fit inside the preview box instead of running off-screen.

## [0.2.7]

### Added

- SVG previews (`svg`, `svgz`), drawn at the size the preview is shown at and sized like a picture: `100%` is the size the document asks for, `50%` that at half.
- Animated SVGs are played by the WebView2 runtime Windows 11 ships with, and by this app's own reader where it is not installed. The engine is pointed at a page of this app's own that draws the document as an image, so it fills the preview box at every scale.
- `webview_idle` and a `Performance → Keep Animated SVG Engine` entry, ten minutes by default: the browser is kept warm between hovers rather than started for each one.
- The engine keeps its state in a folder per run, cleared at startup; one that fails is retried, then stood down for five minutes so this app's reader plays the document, and the hover that was up is laid out again. `RHP_WEBVIEW_TRACE` and `RHP_WEBVIEW_PROFILE` are the diagnostics for it.
- `trigger_key_enabled` in `config.ini` and a check in the tray `Trigger Key` submenu, to switch the trigger key off without changing its mode.
- `svg_background`, the backdrop an SVG document is drawn over, in `config.ini` and as `Background → SVG Background` in the tray: Transparent, Black, White or Checkerboard, the same four a picture is offered.
- `svg_scale`, how much of the screen an SVG document is drawn over, in `config.ini` and as `Placement → SVG Scaling` in the tray: `Fit to Screen`, or `75`, `50` (default), `25` and `10` percent of the display. A vector is drawn at whatever size it is asked for, so a document's percentage is of the room the display has rather than of the size the file asks for — half the screen at `50%`, where a picture at `50%` is half of its own size. Both readers follow it: the still frame this app draws, and the engine that plays an animated document.
- `svg_preview_enabled` and an `SVG` entry under `Preview Types`, below `Office`: a document is its own kind even though `svg` and `svgz` are entries of the image list, so the two are switched independently — `Images` off leaves a document previewing, and `SVG` off leaves pictures alone.
- `decode_budget_gb`, what one hover may decode or read for in gigabytes (default `1`), in `config.ini` and as `Performance → Decode Budget` in the tray: a file past it gets no preview instead of memory the app may not get.

### Changed

- Installing over an older version no longer asks: the previous version is uninstalled first and a running copy of the app is closed instead of prompted for.
- The NSIS setup is built from `packaging/nsis/installer.nsi` on a pinned `cargo-packager` version, and the dead `wix` format is gone.
- `cargo build --release` closes a running copy of the app before linking; debug builds are untouched.
- Video previews no longer write `%TEMP%\rust-hover-preview-video.log`, and one left by an earlier version is deleted at startup.
- `cargo fmt --check` and `cargo clippy --all-targets -- -D warnings` pass again.
- The tray's `Background` submenu is above `Volume` and holds an `Image Background` and an `SVG Background` half, each listing Transparent, Black, White and Checkerboard; `transparent_background` is read as both `image_background` and `svg_background`, and written out under the two names.
- An SVG is no longer sized by `preview_scale`: it is drawn at `svg_scale`, half the screen by default, and a picture's scale leaves it alone.
- Disabling Office/SVG now kills its engine immediately (Office at once, browser within a tick), whether toggled in UI or config.ini. Renders refused because the tier was off no longer count against the document, so re-enabling Office previews it without waiting out backoff.

### Fixed

- An animated SVG stopped moving once the engine had been let go for idle — ten minutes after the last one, on the default `webview_idle`: the thread that played it ended with its browser while the app still held its handle, so every document after that was left on its still frame for the rest of the run. The thread now outlives the engine and begins another one for the next document.
- The WebView2 browser no longer outlives the app when the app is killed, and one left behind by an earlier run is ended by the next launch instead of being left holding its profile folder.
- Every process the app starts — an Office engine, the WebView2 browser, an `ffplay` — is put in a Windows job object, so a crash, a kill from Task Manager or a logoff ends them with the app, and is recorded under `%LOCALAPPDATA%\rust-hover-preview\engines` so that the next launch ends whatever the job could not take. Nothing is ever acted on by id alone: a record is used only when the process still carries the image _and_ the start time it was recorded with.
- One engine per Office family is enforced rather than assumed: a family's engine is started only when nothing this app began for it is still running, and a replacement waits for the process it replaces to be gone. Exiting while a document is mid-render no longer leaves the engine behind either.
- Pictures over 40 megapixels preview again: the pixel cap is gone, and every reader — pictures, animated GIF/APNG/WebP, SVG documents, Office exports, themes — now asks the decode budget above before it allocates.

## [0.2.6]

### Added

- Office previews for Word, Excel and PowerPoint, rendered by the installed Office.
- `office_preview_enabled`, `office_cache_mb`, an editable `[office] extensions` list, and an `Office` entry under `Preview Types`.
- Editable `[image]` and `[video]` extension lists in `config.ini`, so which pictures and videos preview is a file edit like the text, archive and office lists.
- A deleted `extensions=` key comes back with its built-in list, in the file as well as in memory.
- `pdf_cache_mb` and `text_cache_mb`, with `PDF` and `Text` entries in the new tray `Cache` submenu; memory-only, capped at `2048`, default `32` for PDF and `0` for text.
- Tray `Cache` submenu for image, text, PDF and Office caches (`0 MB`–`2 GB`), plus `90%`, `80%` and `70%` font sizes.
- A `Performance` submenu holding the caches and a new `Keep Office Engine` setting: `0 seconds`, `1 minute`, `5 minutes`, `10 minutes` (default), `30 minutes`, `1 hour` or `Indefinitely`, with `office_engine_idle` in `config.ini` behind it.

### Changed

- The built-in extension and file-name lists are written in alphabetical order, and a list still holding them is rewritten in that order on upgrade.
- Office previews render pages instead of using saved thumbnails; rendered pages stay in memory (`office_cache_mb`, default `64 MB`).
- A document's page is asked for as soon as it is hovered rather than after a two-second rest, so a preview waits for the render instead of for a timer in front of it.
- Each family keeps its own Office engine, so a folder holding a document, a workbook and a deck starts each application once rather than quitting one and starting another every time the pointer crosses between them; an engine is let go ten minutes after its family was last asked for a page.
- The settings a render needs of an Office instance — dialogs suppressed, macros switched off — are taken for that render and put back when it is over, instead of being held for as long as the engine lives. Nothing of the instance's own is therefore held reconfigured while an engine sits warm between documents, which is what makes `Indefinitely` safe for an instance that is the user's own Word or Excel.
- PDF and Office pages use fit-to-screen sizing, but preview scales below `100%` now reduce it; `100%`+ and image previews are unchanged.
- The waiting spinner is now a transparent haloed arc, appears immediately, follows the pointer, and is placed flush at the pointer's corner.
- A page arriving over an existing preview replaces it without taking the preview down first.
- Worksheet images are copied only as far as needed and fast-shrunk; refused renders are retried, and unreadable pages are re-rendered.
- A document that refuses a page is ignored for two minutes instead of ten.
- `image_cache_mb` defaults to `32` rather than `64`, and `pdf_cache_mb` to `32`, so the decodes and rasters a folder is swept back over are already done; the text cache still starts at `0`.
- Removed the `Office Preview` submenu; `office_render_enabled` is no longer read.
- A workbook's picture — the no-printer answer — is asked for on a machine that has a printer only when an attempt is already a retry: an export that came to nothing there is the instance declining rather than the document failing, and a copy would be refused by the same instance, after the retries the clipboard needs.
- An Office engine let go because it refused a page is ended where it stands rather than asked to quit and pumped for the seconds an unanswered quit takes, so the retry reaches the fresh instance, which is what the retry is for, without the wait in front of it.

### Fixed

- PowerPoint decks and Excel workbooks on printerless machines now preview.
- Office engines no longer survive quit, and failed or unresponsive instances are replaced.
- Workbook clipboard images are released.
- A preview that crossed a display boundary was discarded and stayed gone until the pointer moved; the hover it came from is now replayed, so it is laid out again at the new display's scale and put back.
- An Office page that landed while another hover message was in hand left the preview waiting with nothing to end the wait; the cap now still applies, so the spinner comes down.
- The Office render worker makes itself known before initializing its apartment, so the first request after startup can no longer start a second worker beside it.
- Non-document files no longer start an Office engine.
- The document spinner is correctly sized.
- Replaced videos are re-measured and re-cropped using file version as well as path.
- `config.ini` is written in a fixed order — `[settings]` first, then the sections and their keys alphabetically — instead of being reshuffled on every save.
- A page landing over a waiting spinner no longer flashes the spinner stretched across the page's box: the window is given the page's size and place by the paint itself, in one call, rather than being resized ahead of it.
- A pointer that drifts onto a waiting spinner no longer dismisses the hover it belongs to — the spinner publishes the box it occupies as one that holds the pointer — so the page being waited for is not thrown away with the hover, and the first hover of a document after a folder change is no longer answered with nothing and then with its page on the second.
- Explorer crashing or being ended no longer leaves previews dead until the app is restarted.
- Office's own "Publishing…" progress window no longer flashes over a hover: a render watches for that window in the process it started and hides it, which no setting the engine holds — alerts, screen updating, the automation security mode — was able to do.

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
